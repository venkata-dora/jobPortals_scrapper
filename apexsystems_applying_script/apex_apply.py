#!/usr/bin/env python3
"""Fill Apex Systems applications and leave tabs open for manual submit."""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import random
import re
import sys
import time
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

try:
    from playwright_stealth import stealth_sync as _stealth_sync  # pip install playwright-stealth
    _STEALTH_AVAILABLE = True
except ImportError:
    _STEALTH_AVAILABLE = False

# Set CAPSOLVER_API_KEY env var or pass --capsolver-key to enable automated captcha solving.
# Sign up at capsolver.com — ~$0.80 per 1000 solves.
_CAPSOLVER_API_KEY = os.environ.get("CAPSOLVER_API_KEY", "")
_CAPSOLVER_BASE = "https://api.capsolver.com"


DEFAULT_RESUME = Path(os.environ.get("JOB_PORTAL_RESUME", "resume.docx")).expanduser()
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "output"
DEFAULT_PROFILE_DIR = Path(__file__).resolve().parent / ".browser_profile"

RECAPTCHA_IFRAME_SELECTORS = (
    "iframe[title='reCAPTCHA']",
    "iframe[src*='recaptcha']",
    "iframe[src*='captcha']",
    "iframe[src*='challenge']",
)
RECAPTCHA_CHECKBOX_SELECTORS = (
    ".recaptcha-checkbox-border",
    "#recaptcha-anchor",
    ".recaptcha-checkbox",
    "[role='checkbox']",
)

VERIFICATION_SELECTORS = (
    'iframe[src*="captcha" i]',
    'iframe[src*="challenge" i]',
    '[class*="captcha" i]',
    '[id*="captcha" i]',
    'input[type="checkbox"][name*="captcha" i]',
)


@dataclass(frozen=True)
class Applicant:
    first_name: str = ""
    last_name: str = ""
    email: str = ""


def latest_jobs_file(output_dir: Path) -> Path:
    files = sorted(output_dir.glob("apex_jobs_*.json"), key=lambda path: path.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No apex_jobs_*.json files found in {output_dir}")
    return files[0]


def load_jobs(path: Path) -> list[dict[str, Any]]:
    with path.open() as handle:
        jobs = json.load(handle)
    if not isinstance(jobs, list):
        raise ValueError(f"Expected a list of jobs in {path}")
    return jobs


def apply_stealth(page: Page) -> None:
    if _STEALTH_AVAILABLE:
        _stealth_sync(page)
        print("[stealth] Playwright stealth applied.", flush=True)
    else:
        page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});
            window.chrome = {runtime: {}};
        """)
        print("[stealth] Basic fingerprint masking applied (install playwright-stealth for full coverage).", flush=True)


def _get_checkbox_page_coords(page: Page, iframe_locator: Any) -> tuple[float, float] | None:
    """Compute absolute viewport-level (x, y) center of the reCAPTCHA checkbox."""
    for iframe_sel in RECAPTCHA_IFRAME_SELECTORS:
        try:
            fl = page.frame_locator(iframe_sel)
            for cb_sel in RECAPTCHA_CHECKBOX_SELECTORS:
                try:
                    box = fl.locator(cb_sel).first.bounding_box(timeout=2000)
                    if box:
                        x = box["x"] + box["width"] / 2
                        y = box["y"] + box["height"] / 2
                        print(f"[recaptcha] Checkbox coords: ({x:.0f}, {y:.0f}) via {cb_sel}", flush=True)
                        return x, y
                except Exception:
                    continue
        except Exception:
            continue
    print("[recaptcha] Could not find checkbox coords.", flush=True)
    return None


def _capsolver_request(endpoint: str, payload: dict) -> dict:
    """POST to capsolver API, return parsed JSON response."""
    data = json.dumps(payload).encode()
    req = urllib.request.Request(
        f"{_CAPSOLVER_BASE}{endpoint}",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())


def _capsolver_solve_recaptcha(api_key: str, site_key: str, page_url: str, timeout_s: int = 120) -> str | None:
    """Submit reCAPTCHA v2 task to capsolver, poll until token ready. Returns token or None."""
    create_payload = {
        "clientKey": api_key,
        "task": {
            "type": "ReCaptchaV2TaskProxyless",
            "websiteURL": page_url,
            "websiteKey": site_key,
        },
    }
    try:
        resp = _capsolver_request("/createTask", create_payload)
        if resp.get("errorId", 0) != 0:
            print(f"[capsolver] Create error: {resp.get('errorDescription')}", flush=True)
            return None
        task_id = resp.get("taskId")
        print(f"[capsolver] Task created: {task_id}", flush=True)
    except Exception as exc:
        print(f"[capsolver] Create request failed: {exc}", flush=True)
        return None

    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        time.sleep(5)
        try:
            result = _capsolver_request("/getTaskResult", {"clientKey": api_key, "taskId": task_id})
            if result.get("errorId", 0) != 0:
                print(f"[capsolver] Poll error: {result.get('errorDescription')}", flush=True)
                return None
            status = result.get("status")
            if status == "ready":
                token = result.get("solution", {}).get("gRecaptchaResponse")
                print(f"[capsolver] Token received (len={len(token or '')})", flush=True)
                return token
            print(f"[capsolver] Status: {status} — polling...", flush=True)
        except Exception as exc:
            print(f"[capsolver] Poll error: {exc}", flush=True)
    print("[capsolver] Timed out waiting for solution.", flush=True)
    return None


def _inject_recaptcha_token(page: Page, token: str) -> bool:
    """Inject solved token into hidden textarea and fire the reCAPTCHA callback."""
    try:
        page.evaluate(f"""
            (function() {{
                // Inject token into all g-recaptcha-response textareas
                document.querySelectorAll('[name="g-recaptcha-response"]').forEach(function(el) {{
                    el.value = {json.dumps(token)};
                    el.dispatchEvent(new Event('input', {{bubbles: true}}));
                    el.dispatchEvent(new Event('change', {{bubbles: true}}));
                }});
                // Fire the reCAPTCHA v2 callback if defined
                if (typeof ___grecaptcha_cfg !== 'undefined') {{
                    var clients = ___grecaptcha_cfg.clients;
                    for (var k in clients) {{
                        var client = clients[k];
                        for (var p in client) {{
                            var obj = client[p];
                            if (obj && typeof obj.callback === 'function') {{
                                try {{ obj.callback({json.dumps(token)}); }} catch(e) {{}}
                            }}
                        }}
                    }}
                }}
                // Fallback: call window.onload style callbacks
                if (typeof window.onRecaptchaSuccess === 'function') {{
                    window.onRecaptchaSuccess({json.dumps(token)});
                }}
            }})();
        """)
        print("[capsolver] Token injected into page.", flush=True)
        return True
    except Exception as exc:
        print(f"[capsolver] Token injection failed: {exc}", flush=True)
        return False


def _extract_recaptcha_sitekey(page: Page) -> str | None:
    """Extract data-sitekey from page DOM or iframe src."""
    try:
        sitekey = page.evaluate("""
            (function() {
                // div.g-recaptcha
                var el = document.querySelector('[data-sitekey]');
                if (el) return el.getAttribute('data-sitekey');
                // iframe src param
                var iframe = document.querySelector('iframe[src*="recaptcha"]');
                if (iframe) {
                    var m = iframe.src.match(/[?&]k=([^&]+)/);
                    if (m) return m[1];
                }
                return null;
            })()
        """)
        if sitekey:
            print(f"[capsolver] Found sitekey: {sitekey}", flush=True)
        return sitekey
    except Exception as exc:
        print(f"[capsolver] Sitekey extraction error: {exc}", flush=True)
        return None


def handle_security_checkbox(page: Page, max_attempts: int = 3) -> bool:
    """Solve reCAPTCHA: capsolver API if key set, otherwise real mouse click."""
    print("[recaptcha] Looking for security verification iframe...", flush=True)

    iframe_locator = None
    for selector in RECAPTCHA_IFRAME_SELECTORS:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=2000):
                iframe_locator = loc
                print(f"[recaptcha] Found iframe: {selector}", flush=True)
                break
        except Exception:
            continue

    if iframe_locator is None:
        print("[recaptcha] No reCAPTCHA iframe found.", flush=True)
        return False

    # --- Path A: capsolver API (fully automated) ---
    if _CAPSOLVER_API_KEY:
        print("[recaptcha] capsolver key set — attempting API solve...", flush=True)
        sitekey = _extract_recaptcha_sitekey(page)
        if sitekey:
            page_url = page.url
            token = _capsolver_solve_recaptcha(_CAPSOLVER_API_KEY, sitekey, page_url)
            if token:
                page.wait_for_timeout(random.randint(800, 1800))
                ok = _inject_recaptcha_token(page, token)
                if ok:
                    page.wait_for_timeout(random.randint(1500, 3000))
                    if _recaptcha_is_solved(page):
                        print("[recaptcha] capsolver solve confirmed!", flush=True)
                        return True
                    print("[recaptcha] Token injected but solve not confirmed — may still work.", flush=True)
                    return True
        print("[recaptcha] capsolver failed — falling back to mouse click.", flush=True)

    # --- Path B: real mouse click (works on fresh profiles with good behavioral score) ---
    page.wait_for_timeout(random.randint(800, 1800))

    for attempt in range(1, max_attempts + 1):
        print(f"[recaptcha] Mouse attempt {attempt}/{max_attempts}...", flush=True)
        try:
            coords = _get_checkbox_page_coords(page, iframe_locator)
            if coords is None:
                print("[recaptcha] Could not resolve checkbox coordinates.", flush=True)
                page.wait_for_timeout(1200)
                continue

            x, y = coords

            iframe_box = iframe_locator.bounding_box()
            if iframe_box:
                mid_x = iframe_box["x"] + iframe_box["width"] / 2
                mid_y = iframe_box["y"] + iframe_box["height"] / 2
                _move_mouse_to(page, mid_x, mid_y)
                page.wait_for_timeout(random.randint(300, 700))

            print(f"[recaptcha] Moving to checkbox at ({x:.0f},{y:.0f})...", flush=True)
            _move_mouse_to(page, x, y)
            page.wait_for_timeout(random.randint(200, 500))

            jx = x + random.uniform(-2, 2)
            jy = y + random.uniform(-2, 2)
            page.mouse.click(jx, jy)
            print(f"[recaptcha] Mouse click sent at ({jx:.1f},{jy:.1f}).", flush=True)

            page.wait_for_timeout(random.randint(2000, 3500))

            try:
                challenge = page.locator("iframe[src*='bframe']").first
                if challenge.is_visible(timeout=2000):
                    print("[recaptcha] Image challenge appeared — score too low. Set CAPSOLVER_API_KEY to auto-solve.", flush=True)
                    return False
            except Exception:
                pass

            if _recaptcha_is_solved(page):
                print("[recaptcha] Checkbox confirmed checked — success!", flush=True)
                return True

            print("[recaptcha] Click sent, outcome unclear — continuing.", flush=True)
            return True

        except Exception as exc:
            print(f"[recaptcha] Attempt {attempt} error: {exc}", flush=True)
            page.wait_for_timeout(random.randint(1000, 2500))

    print("[recaptcha] All attempts exhausted.", flush=True)
    return False


# Per-run speed profile: randomized once so each session has different timing fingerprint
_RUN_SPEED = random.uniform(0.7, 1.4)  # <1 = faster session, >1 = slower session
_TYPING_WPM_BASE = random.randint(55, 110)  # chars/min baseline varies per run


def _scaled(ms: int) -> int:
    return max(50, int(ms * _RUN_SPEED))


def human_pause(page: Page, minimum_ms: int = 500, maximum_ms: int = 2500) -> None:
    total = random.randint(_scaled(minimum_ms), _scaled(maximum_ms))
    chunks = random.randint(2, 4)
    for _ in range(chunks):
        page.wait_for_timeout(total // chunks)
    page.wait_for_timeout(total % chunks)


def _idle_mouse_wander(page: Page, viewport_w: int = 1280, viewport_h: int = 768) -> None:
    """Random aimless mouse movements simulating idle reading/thinking."""
    num_moves = random.randint(2, 5)
    for _ in range(num_moves):
        tx = random.randint(int(viewport_w * 0.1), int(viewport_w * 0.9))
        ty = random.randint(int(viewport_h * 0.1), int(viewport_h * 0.8))
        page.mouse.move(tx, ty)
        page.wait_for_timeout(random.randint(80, 300))


def _move_mouse_to(page: Page, x: float, y: float) -> None:
    """Move mouse to target in small steps to mimic natural movement."""
    try:
        cx, cy = random.randint(100, 400), random.randint(100, 400)
    except Exception:
        cx, cy = 200, 300
    steps = random.randint(8, 18)
    # Pick random curve: ease-in-out, overshoot-correct, or linear-noisy
    curve = random.choice(["ease", "overshoot", "noisy"])
    for i in range(1, steps + 1):
        t = i / steps
        if curve == "ease":
            factor = t * t * (3 - 2 * t)
        elif curve == "overshoot":
            factor = t + 0.2 * math.sin(t * math.pi)
        else:
            factor = t + random.uniform(-0.05, 0.05)
        factor = max(0.0, min(1.0, factor))
        jitter_scale = 3 if i < steps - 2 else 1
        mx = cx + (x - cx) * factor + random.uniform(-jitter_scale, jitter_scale)
        my = cy + (y - cy) * factor + random.uniform(-jitter_scale, jitter_scale)
        page.mouse.move(mx, my)
        page.wait_for_timeout(random.randint(6, 22))


def human_mouse_click(page: Page, x: float, y: float) -> None:
    """Move mouse naturally then click with slight coordinate jitter."""
    jx = x + random.uniform(-4, 4)
    jy = y + random.uniform(-4, 4)
    _move_mouse_to(page, jx, jy)
    page.wait_for_timeout(random.randint(_scaled(50), _scaled(200)))
    page.mouse.click(jx, jy)


def human_scroll(page: Page) -> None:
    """Scroll with varied curve and occasional back-scroll to mimic reading."""
    num_scrolls = random.randint(2, 6)
    direction_bias = random.choice([1, 1, 1, -1])  # mostly down
    peak = random.randint(2, max(2, num_scrolls - 1))
    for i in range(num_scrolls):
        # Bell-curve speed: slow start, fast middle, slow end
        t = i / max(num_scrolls - 1, 1)
        bell = math.sin(t * math.pi)
        delta = int((random.randint(60, 180) + bell * 150) * direction_bias)
        page.mouse.wheel(random.uniform(-2, 2), delta)
        page.wait_for_timeout(random.randint(_scaled(80), _scaled(380)))
        if i == peak and random.random() < 0.4:
            # Pause at peak as if reading something interesting
            page.wait_for_timeout(random.randint(500, 1500))
    # 40% chance scroll back up a little
    if random.random() < 0.4:
        page.mouse.wheel(0, -random.randint(30, 150))
        page.wait_for_timeout(random.randint(200, 600))


def human_type(page: Page, locator: Any, text: str) -> None:
    """Type text with realistic human variation: speed bursts, fatigue, typos, thinking pauses."""
    try:
        box = locator.bounding_box()
        if box:
            human_mouse_click(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        else:
            locator.click()
    except Exception:
        locator.click()
    page.wait_for_timeout(random.randint(_scaled(100), _scaled(380)))
    try:
        locator.fill("")
    except Exception:
        pass

    base_ms = int(60000 / max(_TYPING_WPM_BASE * 5, 1))
    # Per-field speed modifier — even within a session, each field feels different
    field_speed = random.uniform(0.75, 1.35)
    # Burst mode: occasionally type a stretch of chars very fast
    in_burst = False
    burst_remaining = 0

    for i, char in enumerate(text):
        # Fatigue: gradually slow down over long text (humans slow down mid-form)
        fatigue = 1.0 + (i / max(len(text), 1)) * random.uniform(0.0, 0.25)

        if in_burst and burst_remaining > 0:
            delay = random.randint(20, 55)
            burst_remaining -= 1
            if burst_remaining == 0:
                in_burst = False
        elif char == " ":
            delay = random.randint(int(base_ms * 0.3 * field_speed), int(base_ms * 0.7 * field_speed))
        elif char in ".,@-_+":
            delay = random.randint(int(base_ms * 1.1 * field_speed * fatigue), int(base_ms * 2.2 * field_speed * fatigue))
        elif char.isupper():
            # Shift key press adds slight overhead
            delay = random.randint(int(base_ms * 0.9 * field_speed), int(base_ms * 1.8 * field_speed * fatigue))
        else:
            delay = random.randint(int(base_ms * 0.5 * field_speed), int(base_ms * 1.5 * field_speed * fatigue))

        # 3% chance: simulate typo — type wrong char then backspace
        if random.random() < 0.03 and char.isalpha():
            wrong = random.choice("qwertyuiopasdfghjklzxcvbnm")
            locator.type(wrong, delay=random.randint(40, 120))
            page.wait_for_timeout(random.randint(80, 250))
            locator.press("Backspace")
            page.wait_for_timeout(random.randint(100, 350))

        locator.type(char, delay=delay)

        # 5% thinking pause (hesitate mid-word)
        if random.random() < 0.05:
            page.wait_for_timeout(random.randint(280, 950))
        # 4% enter burst mode for next 3-7 chars
        elif random.random() < 0.04 and not in_burst:
            in_burst = True
            burst_remaining = random.randint(3, 7)
        # 2% long pause (distracted)
        elif random.random() < 0.02:
            page.wait_for_timeout(random.randint(900, 2200))


def _recaptcha_is_solved(page: Page) -> bool:
    """Return True if reCAPTCHA checkbox is already checked / token filled."""
    try:
        response = page.locator('[name="g-recaptcha-response"]').first.input_value(timeout=300)
        if response.strip():
            return True
    except Exception:
        pass
    for iframe_sel in RECAPTCHA_IFRAME_SELECTORS:
        try:
            fl = page.frame_locator(iframe_sel)
            checked = fl.locator("[aria-checked='true'], .recaptcha-checkbox-checked").first
            if checked.is_visible(timeout=300):
                return True
        except Exception:
            continue
    return False


def verification_is_visible(page: Page) -> bool:
    if _recaptcha_is_solved(page):
        return False
    for selector in VERIFICATION_SELECTORS:
        try:
            if page.locator(selector).first.is_visible(timeout=300):
                return True
        except Exception:
            continue
    try:
        body_text = page.locator("body").inner_text(timeout=1000)
    except Exception:
        return False
    return bool(
        re.search(
            r"(verify (?:that )?you are human|human verification|security verification|complete the challenge)",
            body_text,
            re.I,
        )
    )


def wait_for_submission_confirmation(page: Page, timeout_seconds: int = 30) -> None:
    deadline = time.monotonic() + timeout_seconds
    confirmation_re = re.compile(
        r"(application (?:was |has been )?submitted|thank you for (?:applying|your application)|submission (?:complete|confirmed))",
        re.I,
    )
    while time.monotonic() < deadline:
        try:
            text = page.locator("body").inner_text(timeout=1000)
            if confirmation_re.search(text):
                return
        except Exception:
            pass
        page.wait_for_timeout(1000)
    raise RuntimeError("Apex did not show an application submission confirmation")


def wait_for_manual_verification(page: Page) -> None:
    if not verification_is_visible(page):
        return
    print("[verification] Security check detected. Attempting auto-click...", flush=True)
    page.bring_to_front()
    auto_success = handle_security_checkbox(page)
    if auto_success and not verification_is_visible(page):
        print("[verification] Auto-click resolved verification.", flush=True)
        return
    print("[verification] Auto-click insufficient - please solve manually in the open browser window.", flush=True)
    while verification_is_visible(page):
        page.wait_for_timeout(1000)
    page.wait_for_load_state("domcontentloaded", timeout=30000)
    print("[verification] Manual verification completed; resuming.", flush=True)


def human_click(page: Page, locator: Any, timeout: int = 10000) -> None:
    human_pause(page, 300, 1200)
    try:
        box = locator.bounding_box(timeout=timeout)
        if box:
            cx = box["x"] + box["width"] / 2
            cy = box["y"] + box["height"] / 2
            human_mouse_click(page, cx, cy)
        else:
            locator.click(timeout=timeout)
    except Exception:
        locator.click(timeout=timeout)
    human_pause(page, 400, 1500)
    wait_for_manual_verification(page)


def close_cookie_banner(page: Page) -> None:
    try:
        button = page.locator("#onetrust-accept-btn-handler")
        if button.is_visible(timeout=2500):
            human_click(page, button, timeout=2500)
    except PlaywrightTimeoutError:
        pass


def fill_input(page: Page, name: str, value: str) -> None:
    field = page.locator(f"input[name='{name}']").first
    human_type(page, field, value)
    human_pause(page, 200, 600)


def wait_for_resume_parse(page: Page) -> None:
    continue_button = page.get_by_role("button", name=re.compile(r"continue", re.I)).first
    for _ in range(40):
        try:
            text = continue_button.inner_text(timeout=1000)
            disabled = continue_button.is_disabled(timeout=1000)
            if "continue" in text.lower() and not disabled:
                return
        except Exception:
            pass
        page.wait_for_timeout(1000)
    raise RuntimeError("Resume parse did not finish; Continue button stayed unavailable")


def _build_page_mouse_history(page: Page, viewport_w: int = 1280, viewport_h: int = 768, seconds: int = 6) -> None:
    """Move mouse naturally around the page for several seconds to build reCAPTCHA behavioral score."""
    print("[behavior] Building mouse history for reCAPTCHA scoring...", flush=True)
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        tx = random.randint(int(viewport_w * 0.05), int(viewport_w * 0.95))
        ty = random.randint(int(viewport_h * 0.05), int(viewport_h * 0.90))
        _move_mouse_to(page, tx, ty)
        page.wait_for_timeout(random.randint(200, 700))
        if random.random() < 0.25:
            human_scroll(page)


def fill_application(page: Page, applicant: Applicant, resume_path: Path) -> None:
    close_cookie_banner(page)
    wait_for_manual_verification(page)
    vp = page.viewport_size or {"width": 1280, "height": 768}
    _build_page_mouse_history(page, vp["width"], vp["height"], seconds=random.randint(5, 9))
    human_scroll(page)
    page.locator("input[type='file']").first.set_input_files(str(resume_path))
    wait_for_resume_parse(page)
    human_click(page, page.get_by_role("button", name=re.compile(r"continue", re.I)).first)

    # Resume parsing usually fills these, but set them explicitly so every tab is consistent.
    fill_input(page, "firstName", applicant.first_name)
    fill_input(page, "lastName", applicant.last_name)
    fill_input(page, "email", applicant.email)

    page.get_by_role("button", name=re.compile(r"submit application", re.I)).first.wait_for(timeout=10000)
    wait_for_manual_verification(page)


def write_report(path: Path, rows: Iterable[dict[str, Any]]) -> None:
    rows = list(rows)
    if not rows:
        return
    keys = ["status", "job_id", "title", "posted_date", "apply_url", "message", "screenshot"]
    with path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=keys)
        writer.writeheader()
        writer.writerows(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fill Apex applications, with optional automatic submission.")
    parser.add_argument("--jobs-file", type=Path, help="Filtered apex_jobs_*.json file. Defaults to latest output.")
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--resume", type=Path, default=DEFAULT_RESUME)
    parser.add_argument("--first-name", default=Applicant.first_name)
    parser.add_argument("--last-name", default=Applicant.last_name)
    parser.add_argument("--email", default=Applicant.email)
    parser.add_argument("--limit", type=int, default=4)
    parser.add_argument("--start-at", type=int, default=1, help="1-based job index to start from.")
    parser.add_argument("--keep-open-minutes", type=int, default=120)
    parser.add_argument("--slow-mo", type=int, default=250)
    parser.add_argument("--browser-channel", default="", help="Playwright browser channel (e.g. 'chrome'). Omit to use bundled Chromium (avoids conflicts with open Chrome).")
    parser.add_argument(
        "--profile-dir",
        type=Path,
        default=DEFAULT_PROFILE_DIR,
        help="Persistent browser profile used to retain Apex login sessions and cookies.",
    )
    parser.add_argument(
        "--url",
        action="append",
        default=[],
        help="Application URL to process instead of a scraped jobs file. Repeat for multiple URLs.",
    )
    parser.add_argument(
        "--submit",
        action="store_true",
        help="Click Submit Application after filling. Without this flag, review tabs remain open.",
    )
    parser.add_argument(
        "--capsolver-key",
        default="",
        help="capsolver.com API key for automated reCAPTCHA solving. Also reads CAPSOLVER_API_KEY env var.",
    )
    parser.add_argument(
        "--setup-profile",
        action="store_true",
        help="Open browser for manual Google login / profile setup, then exit. No jobs processed.",
    )
    return parser.parse_args()


def main() -> int:
    global _CAPSOLVER_API_KEY
    args = parse_args()
    if args.capsolver_key:
        _CAPSOLVER_API_KEY = args.capsolver_key
    if _CAPSOLVER_API_KEY:
        print(f"[capsolver] API key set — reCAPTCHA will be solved automatically.", flush=True)
    else:
        print("[capsolver] No API key — will attempt mouse click (may need manual solve).", flush=True)
    if args.setup_profile:
        profile_dir = args.profile_dir.expanduser().resolve()
        profile_dir.mkdir(parents=True, exist_ok=True)
        launch_kwargs: dict = dict(headless=False, slow_mo=250, locale="en-US", timezone_id="America/Chicago")
        if args.browser_channel:
            launch_kwargs["channel"] = args.browser_channel
        print(f"Opening browser with profile: {profile_dir}")
        print("Log into Google (google.com), then close the browser window.")
        with sync_playwright() as pw:
            ctx = pw.chromium.launch_persistent_context(str(profile_dir), **launch_kwargs)
            pg = ctx.new_page()
            pg.goto("https://google.com")
            input("Press Enter here after logging in and closing the browser...")
            ctx.close()
        print("Profile saved. Run script normally now.")
        return 0

    missing = [name for name, value in (("first name", args.first_name), ("last name", args.last_name), ("email", args.email)) if not value]
    if missing:
        print(f"Missing applicant {', '.join(missing)}. Pass the values with flags.", file=sys.stderr)
        return 2
    resume_path = args.resume.expanduser().resolve()
    if not resume_path.exists():
        print(f"Resume not found: {resume_path}", file=sys.stderr)
        return 2

    jobs_file = args.jobs_file
    if args.url:
        jobs = [
            {"title": f"Application {index}", "job_id": f"URL-{index}", "apply_url": url}
            for index, url in enumerate(args.url, start=1)
        ]
    else:
        jobs_file = jobs_file or latest_jobs_file(args.out_dir)
        jobs = load_jobs(jobs_file)
    start = max(args.start_at - 1, 0)
    selected = jobs[start:]
    if args.limit > 0:
        selected = selected[: args.limit]
    if not selected:
        print("No jobs selected.")
        return 0

    applicant = Applicant(args.first_name, args.last_name, args.email)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = args.out_dir / f"apply_{stamp}"
    run_dir.mkdir(parents=True, exist_ok=True)
    report_rows: list[dict[str, Any]] = []

    print(f"Jobs source: {jobs_file or 'command-line URLs'}")
    print(f"Resume: {resume_path}")
    print(f"Selected jobs: {len(selected)}")
    print(f"Mode: {'automatic submit' if args.submit else 'fill tabs and leave open for review'}")

    with sync_playwright() as playwright:
        profile_dir = args.profile_dir.expanduser().resolve()
        profile_dir.mkdir(parents=True, exist_ok=True)
        # Randomize viewport to avoid fingerprinting on fixed dimensions
        viewport_width = random.choice([1280, 1366, 1440, 1536, 1920])
        viewport_height = random.choice([768, 800, 864, 900, 1080])
        launch_kwargs: dict = dict(
            headless=False,
            slow_mo=args.slow_mo,
            viewport={"width": viewport_width, "height": viewport_height},
            locale="en-US",
            timezone_id="America/Chicago",
            accept_downloads=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                f"--window-size={viewport_width},{viewport_height + 74}",
            ],
        )
        if args.browser_channel:
            launch_kwargs["channel"] = args.browser_channel
        context = playwright.chromium.launch_persistent_context(
            str(profile_dir),
            **launch_kwargs,
        )
        first_page = context.new_page()
        apply_stealth(first_page)

        try:
            for offset, job in enumerate(selected, start=start + 1):
                if offset == start + 1:
                    page = first_page
                else:
                    page = context.new_page()
                    apply_stealth(page)
                title = str(job.get("title") or "").strip()
                job_id = str(job.get("job_id") or "").strip()
                apply_url = str(job.get("apply_url") or "").strip()
                screenshot = run_dir / f"{offset:03d}_{job_id.lower()}_filled.png"
                print(f"\n[{offset}/{len(jobs)}] {title} | {job_id}")
                print(apply_url)

                status = "filled"
                message = "Filled and waiting for manual Submit Application."
                try:
                    if not apply_url:
                        raise RuntimeError("Missing apply_url")
                    # Random pre-navigation idle (looks like user pausing before clicking Apply)
                    time.sleep(random.uniform(2.0, 5.0))
                    page.goto(apply_url, wait_until="domcontentloaded", timeout=45000)
                    # Simulate reading/scanning the page after load
                    human_pause(page, 1800, 4000)
                    _idle_mouse_wander(page, viewport_width, viewport_height)
                    human_scroll(page)
                    human_pause(page, 1000, 2500)
                    _idle_mouse_wander(page, viewport_width, viewport_height)
                    wait_for_manual_verification(page)
                    fill_application(page, applicant, resume_path)
                    if args.submit:
                        submit_button = page.get_by_role("button", name=re.compile(r"submit application", re.I)).first
                        human_click(page, submit_button)
                        try:
                            page.wait_for_load_state("domcontentloaded", timeout=30000)
                        except PlaywrightTimeoutError:
                            pass
                        wait_for_manual_verification(page)
                        wait_for_submission_confirmation(page)
                        status = "submitted"
                        message = "Application submitted."
                    page.screenshot(path=str(screenshot), full_page=True)
                    print(f"Status: {status}")
                except Exception as exc:
                    status = "error"
                    message = str(exc)
                    try:
                        page.screenshot(path=str(screenshot), full_page=True)
                    except Exception:
                        pass
                    print(f"Status: error - {message}")

                report_rows.append(
                    {
                        "status": status,
                        "job_id": job_id,
                        "title": title,
                        "posted_date": job.get("posted_date", ""),
                        "apply_url": apply_url,
                        "message": message,
                        "screenshot": str(screenshot),
                    }
                )
                write_report(run_dir / "apply_report.csv", report_rows)
                # Longer inter-job gap so velocity doesn't flag session
                time.sleep(random.uniform(8.0, 20.0))

            print("\nBrowser is staying open for review.")
            print(f"Apply report: {run_dir / 'apply_report.csv'}")
            print(f"It will stay open for {args.keep_open_minutes} minutes unless you close it first.")
            first_page.wait_for_timeout(max(args.keep_open_minutes, 1) * 60 * 1000)
        finally:
            context.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
