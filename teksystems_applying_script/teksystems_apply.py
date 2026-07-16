#!/usr/bin/env python3
"""Fill TEKsystems applications and optionally submit them."""

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
from typing import Any

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

try:
    from playwright_stealth import stealth_sync as _stealth_sync
    _STEALTH_AVAILABLE = True
except ImportError:
    _STEALTH_AVAILABLE = False

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "output"
DEFAULT_PROFILE_DIR = SCRIPT_DIR / ".browser_profile"

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
)

_CAPSOLVER_API_KEY = os.environ.get("CAPSOLVER_API_KEY", "")
_CAPSOLVER_BASE = "https://api.capsolver.com"

# Per-run speed profile — different timing fingerprint each session
_RUN_SPEED = random.uniform(0.7, 1.4)
_TYPING_WPM_BASE = random.randint(55, 110)


@dataclass(frozen=True)
class Applicant:
    first_name: str
    last_name: str
    email: str
    phone: str
    city: str
    state: str
    zip_code: str
    recent_title: str


# ── helpers ───────────────────────────────────────────────────────────────────

def latest_jobs_file(output_dir: Path) -> Path:
    files = sorted(output_dir.glob("teksystems_jobs_*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No teksystems_jobs_*.json files found in {output_dir}")
    return files[0]


def load_jobs(path: Path) -> list[dict[str, Any]]:
    with path.open() as fh:
        jobs = json.load(fh)
    if not isinstance(jobs, list):
        raise ValueError(f"Expected a list of jobs in {path}")
    return jobs


def _scaled(ms: int) -> int:
    return max(50, int(ms * _RUN_SPEED))


# ── stealth ───────────────────────────────────────────────────────────────────

def apply_stealth(page: Page) -> None:
    if _STEALTH_AVAILABLE:
        _stealth_sync(page)
        print("[stealth] playwright-stealth applied.", flush=True)
    else:
        page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});
            window.chrome = {runtime: {}};
        """)
        print("[stealth] Basic fingerprint masking applied.", flush=True)


# ── human mouse ───────────────────────────────────────────────────────────────

def _move_mouse_to(page: Page, x: float, y: float) -> None:
    try:
        cx, cy = random.randint(100, 400), random.randint(100, 400)
    except Exception:
        cx, cy = 200, 300
    steps = random.randint(8, 18)
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
        jitter = 3 if i < steps - 2 else 1
        page.mouse.move(
            cx + (x - cx) * factor + random.uniform(-jitter, jitter),
            cy + (y - cy) * factor + random.uniform(-jitter, jitter),
        )
        page.wait_for_timeout(random.randint(6, 22))


def human_mouse_click(page: Page, x: float, y: float) -> None:
    jx = x + random.uniform(-4, 4)
    jy = y + random.uniform(-4, 4)
    _move_mouse_to(page, jx, jy)
    page.wait_for_timeout(random.randint(_scaled(50), _scaled(200)))
    page.mouse.click(jx, jy)


def _idle_mouse_wander(page: Page, viewport_w: int = 1280, viewport_h: int = 768) -> None:
    for _ in range(random.randint(2, 5)):
        page.mouse.move(
            random.randint(int(viewport_w * 0.1), int(viewport_w * 0.9)),
            random.randint(int(viewport_h * 0.1), int(viewport_h * 0.8)),
        )
        page.wait_for_timeout(random.randint(80, 300))


# ── human actions ─────────────────────────────────────────────────────────────

def human_pause(page: Page, minimum_ms: int = 500, maximum_ms: int = 2500) -> None:
    total = random.randint(_scaled(minimum_ms), _scaled(maximum_ms))
    chunks = random.randint(2, 4)
    for _ in range(chunks):
        page.wait_for_timeout(total // chunks)
    page.wait_for_timeout(total % chunks)


def human_scroll(page: Page) -> None:
    num_scrolls = random.randint(2, 6)
    direction_bias = random.choice([1, 1, 1, -1])
    peak = random.randint(2, max(2, num_scrolls - 1))
    for i in range(num_scrolls):
        bell = math.sin(i / max(num_scrolls - 1, 1) * math.pi)
        delta = int((random.randint(60, 180) + bell * 150) * direction_bias)
        page.mouse.wheel(random.uniform(-2, 2), delta)
        page.wait_for_timeout(random.randint(_scaled(80), _scaled(380)))
        if i == peak and random.random() < 0.4:
            page.wait_for_timeout(random.randint(500, 1500))
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
    field_speed = random.uniform(0.75, 1.35)
    in_burst = False
    burst_remaining = 0

    for i, char in enumerate(text):
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
            delay = random.randint(int(base_ms * 0.9 * field_speed), int(base_ms * 1.8 * field_speed * fatigue))
        else:
            delay = random.randint(int(base_ms * 0.5 * field_speed), int(base_ms * 1.5 * field_speed * fatigue))

        if random.random() < 0.03 and char.isalpha():
            wrong = random.choice("qwertyuiopasdfghjklzxcvbnm")
            locator.type(wrong, delay=random.randint(40, 120))
            page.wait_for_timeout(random.randint(80, 250))
            locator.press("Backspace")
            page.wait_for_timeout(random.randint(100, 350))

        locator.type(char, delay=delay)

        if random.random() < 0.05:
            page.wait_for_timeout(random.randint(280, 950))
        elif random.random() < 0.04 and not in_burst:
            in_burst = True
            burst_remaining = random.randint(3, 7)
        elif random.random() < 0.02:
            page.wait_for_timeout(random.randint(900, 2200))


def human_click(page: Page, locator: Any, timeout: int = 15000) -> None:
    human_pause(page, 300, 1200)
    try:
        box = locator.bounding_box(timeout=timeout)
        if box:
            human_mouse_click(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        else:
            locator.click(timeout=timeout)
    except Exception:
        locator.click(timeout=timeout)
    human_pause(page, 400, 1500)
    wait_for_manual_verification(page)


def _build_page_mouse_history(page: Page, viewport_w: int = 1280, viewport_h: int = 768, seconds: int = 6) -> None:
    print("[behavior] Building mouse history for reCAPTCHA scoring...", flush=True)
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        _move_mouse_to(
            page,
            random.randint(int(viewport_w * 0.05), int(viewport_w * 0.95)),
            random.randint(int(viewport_h * 0.05), int(viewport_h * 0.90)),
        )
        page.wait_for_timeout(random.randint(200, 700))
        if random.random() < 0.25:
            human_scroll(page)


# ── captcha ───────────────────────────────────────────────────────────────────

def _capsolver_request(endpoint: str, payload: dict) -> dict:
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
    try:
        resp = _capsolver_request("/createTask", {
            "clientKey": api_key,
            "task": {"type": "ReCaptchaV2TaskProxyless", "websiteURL": page_url, "websiteKey": site_key},
        })
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
            if result.get("status") == "ready":
                token = result.get("solution", {}).get("gRecaptchaResponse")
                print(f"[capsolver] Token received (len={len(token or '')})", flush=True)
                return token
            print(f"[capsolver] Status: {result.get('status')} — polling...", flush=True)
        except Exception as exc:
            print(f"[capsolver] Poll error: {exc}", flush=True)
    print("[capsolver] Timed out.", flush=True)
    return None


def _inject_recaptcha_token(page: Page, token: str) -> bool:
    try:
        page.evaluate(f"""
            (function() {{
                document.querySelectorAll('[name="g-recaptcha-response"]').forEach(function(el) {{
                    el.value = {json.dumps(token)};
                    el.dispatchEvent(new Event('input', {{bubbles: true}}));
                    el.dispatchEvent(new Event('change', {{bubbles: true}}));
                }});
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
            }})();
        """)
        print("[capsolver] Token injected.", flush=True)
        return True
    except Exception as exc:
        print(f"[capsolver] Injection failed: {exc}", flush=True)
        return False


def _extract_recaptcha_sitekey(page: Page) -> str | None:
    try:
        return page.evaluate("""
            (function() {
                var el = document.querySelector('[data-sitekey]');
                if (el) return el.getAttribute('data-sitekey');
                var iframe = document.querySelector('iframe[src*="recaptcha"]');
                if (iframe) { var m = iframe.src.match(/[?&]k=([^&]+)/); if (m) return m[1]; }
                return null;
            })()
        """)
    except Exception:
        return None


def _recaptcha_is_solved(page: Page) -> bool:
    try:
        if page.locator('[name="g-recaptcha-response"]').first.input_value(timeout=300).strip():
            return True
    except Exception:
        pass
    for sel in RECAPTCHA_IFRAME_SELECTORS:
        try:
            if page.frame_locator(sel).locator("[aria-checked='true'], .recaptcha-checkbox-checked").first.is_visible(timeout=300):
                return True
        except Exception:
            continue
    return False


def _get_checkbox_page_coords(page: Page) -> tuple[float, float] | None:
    for iframe_sel in RECAPTCHA_IFRAME_SELECTORS:
        for cb_sel in RECAPTCHA_CHECKBOX_SELECTORS:
            try:
                box = page.frame_locator(iframe_sel).locator(cb_sel).first.bounding_box(timeout=2000)
                if box:
                    x = box["x"] + box["width"] / 2
                    y = box["y"] + box["height"] / 2
                    print(f"[recaptcha] Checkbox coords: ({x:.0f}, {y:.0f})", flush=True)
                    return x, y
            except Exception:
                continue
    return None


def handle_security_checkbox(page: Page, max_attempts: int = 3) -> bool:
    print("[recaptcha] Looking for security verification...", flush=True)
    iframe_visible = any(
        _safe_visible(page, sel) for sel in RECAPTCHA_IFRAME_SELECTORS
    )
    if not iframe_visible:
        print("[recaptcha] No reCAPTCHA iframe found.", flush=True)
        return False

    if _CAPSOLVER_API_KEY:
        print("[recaptcha] capsolver key set — attempting API solve...", flush=True)
        sitekey = _extract_recaptcha_sitekey(page)
        if sitekey:
            token = _capsolver_solve_recaptcha(_CAPSOLVER_API_KEY, sitekey, page.url)
            if token:
                page.wait_for_timeout(random.randint(800, 1800))
                if _inject_recaptcha_token(page, token):
                    page.wait_for_timeout(random.randint(1500, 3000))
                    print("[recaptcha] capsolver solve done.", flush=True)
                    return True
        print("[recaptcha] capsolver failed — falling back to mouse click.", flush=True)

    page.wait_for_timeout(random.randint(800, 1800))
    for attempt in range(1, max_attempts + 1):
        print(f"[recaptcha] Mouse attempt {attempt}/{max_attempts}...", flush=True)
        try:
            coords = _get_checkbox_page_coords(page)
            if not coords:
                page.wait_for_timeout(1200)
                continue
            x, y = coords
            _move_mouse_to(page, x, y)
            page.wait_for_timeout(random.randint(200, 500))
            page.mouse.click(x + random.uniform(-2, 2), y + random.uniform(-2, 2))
            page.wait_for_timeout(random.randint(2000, 3500))
            try:
                if page.locator("iframe[src*='bframe']").first.is_visible(timeout=2000):
                    print("[recaptcha] Image challenge — set CAPSOLVER_API_KEY to auto-solve.", flush=True)
                    return False
            except Exception:
                pass
            if _recaptcha_is_solved(page):
                print("[recaptcha] Checkbox confirmed!", flush=True)
                return True
            return True
        except Exception as exc:
            print(f"[recaptcha] Attempt {attempt} error: {exc}", flush=True)
            page.wait_for_timeout(random.randint(1000, 2500))

    print("[recaptcha] All attempts exhausted.", flush=True)
    return False


def _safe_visible(page: Page, selector: str) -> bool:
    try:
        return page.locator(selector).first.is_visible(timeout=2000)
    except Exception:
        return False


def verification_is_visible(page: Page) -> bool:
    if _recaptcha_is_solved(page):
        return False
    for sel in VERIFICATION_SELECTORS:
        try:
            if page.locator(sel).first.is_visible(timeout=300):
                return True
        except Exception:
            continue
    try:
        body = page.locator("body").inner_text(timeout=1000)
        return bool(re.search(r"(verify (?:that )?you are human|human verification|security verification|complete the challenge)", body, re.I))
    except Exception:
        return False


def wait_for_manual_verification(page: Page) -> None:
    if not verification_is_visible(page):
        return
    print("[verification] Security check detected. Attempting auto-click...", flush=True)
    page.bring_to_front()
    if handle_security_checkbox(page) and not verification_is_visible(page):
        print("[verification] Auto-click resolved.", flush=True)
        return
    print("[verification] Please solve manually in the browser window.", flush=True)
    while verification_is_visible(page):
        page.wait_for_timeout(1000)
    page.wait_for_load_state("domcontentloaded", timeout=30000)
    print("[verification] Verification completed; resuming.", flush=True)


# ── form helpers ──────────────────────────────────────────────────────────────

def close_cookie_banner(page: Page) -> None:
    candidates = page.locator(
        "#onetrust-reject-all-handler, #onetrust-close-btn-container button, button[aria-label='Close']"
    )
    for i in range(candidates.count()):
        button = candidates.nth(i)
        try:
            if button.is_visible(timeout=500):
                human_pause(page, 300, 800)
                button.click(timeout=3000, force=True)
                human_pause(page, 300, 600)
                return
        except PlaywrightTimeoutError:
            pass


def fill_labeled(page: Page, label: str, value: str) -> None:
    field = page.get_by_label(re.compile(rf"^{re.escape(label)}", re.I)).first
    # LWC/Salesforce inputs reject char-by-char type — use fill() then trigger events
    try:
        box = field.bounding_box(timeout=5000)
        if box:
            human_mouse_click(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        else:
            field.click()
    except Exception:
        field.click()
    human_pause(page, 200, 500)
    field.fill(value, timeout=15000)
    # Fire input/change events so LWC framework picks up the value
    field.evaluate("el => { el.dispatchEvent(new Event('input', {bubbles:true})); el.dispatchEvent(new Event('change', {bubbles:true})); }")
    human_pause(page, 200, 600)


def choose_menu(page: Page, current_or_placeholder: str, choice: str) -> None:
    button = page.get_by_role("button", name=re.compile(current_or_placeholder, re.I)).first
    human_click(page, button)
    choice_pattern = re.compile(
        r"(?:New Jersey|\bNJ\b)" if choice == "New Jersey" else rf"\b{re.escape(choice)}\b",
        re.I,
    )
    option = page.locator('[role="option"]').filter(has_text=choice_pattern).first
    try:
        human_click(page, option)
    except PlaywrightTimeoutError:
        button.press("Home")
        button.press("n" if choice == "New Jersey" else choice[0])
        button.press("Enter")
    human_pause(page, 300, 800)


def _click_radio_if_visible(page: Page, question_pattern: str, answer: str, timeout: int = 3000) -> bool:
    """Find a question by text and click the matching radio. Returns True if handled."""
    try:
        question = page.get_by_text(re.compile(question_pattern, re.I)).first
        if not question.is_visible(timeout=timeout):
            return False
        radio = question.locator(f"xpath=following::*[self::input[@type='radio'] or self::label][normalize-space()='{answer}' or @value='{answer}'][1]")
        if not radio.is_visible(timeout=2000):
            # fallback: search nearby by role
            radio = page.get_by_role("radio", name=re.compile(rf"^{re.escape(answer)}$", re.I)).first
        radio.click(force=True)
        print(f"[form] '{question_pattern}' → '{answer}'", flush=True)
        human_pause(page, 200, 600)
        return True
    except Exception:
        return False


def _handle_optional_questions(page: Page) -> None:
    """Answer optional questions that may or may not appear depending on the job."""
    human_pause(page, 400, 900)

    # Security clearance → No
    for pattern in (r"security clearance", r"active.*clearance", r"clearance.*required"):
        if _click_radio_if_visible(page, pattern, "No"):
            break

    # Sponsorship / visa → No
    for pattern in (r"sponsorship", r"require.*sponsor", r"visa.*sponsor"):
        if _click_radio_if_visible(page, pattern, "No"):
            break

    # Felony conviction → No
    _click_radio_if_visible(page, r"(felony|criminal conviction)", "No")

    # Veteran status → I am not a protected veteran
    for answer in ("I am not a protected veteran", "Not a protected veteran", "No"):
        try:
            q = page.get_by_text(re.compile(r"veteran status", re.I)).first
            if q.is_visible(timeout=2000):
                radio = page.get_by_role("radio", name=re.compile(answer, re.I)).first
                if radio.is_visible(timeout=1500):
                    radio.click(force=True)
                    human_pause(page, 200, 500)
                    print(f"[form] veteran status → '{answer}'", flush=True)
                    break
        except Exception:
            pass

    # Disability → No, I don't have a disability
    try:
        q = page.get_by_text(re.compile(r"disability", re.I)).first
        if q.is_visible(timeout=2000):
            for answer in ("No, I don't have a disability", "I don't have a disability", "No"):
                try:
                    radio = page.get_by_role("radio", name=re.compile(re.escape(answer), re.I)).first
                    if radio.is_visible(timeout=1500):
                        radio.click(force=True)
                        human_pause(page, 200, 500)
                        print(f"[form] disability → '{answer}'", flush=True)
                        break
                except Exception:
                    continue
    except Exception:
        pass

    # Work authorization → Yes
    for pattern in (r"authorized to work", r"legally authorized", r"work authorization"):
        if _click_radio_if_visible(page, pattern, "Yes"):
            break


def fill_application(page: Page, applicant: Applicant, resume_path: Path) -> None:
    close_cookie_banner(page)

    vp = page.viewport_size or {"width": 1280, "height": 768}
    _build_page_mouse_history(page, vp["width"], vp["height"], seconds=random.randint(5, 9))
    human_scroll(page)

    page.locator("input[type='file']").first.set_input_files(str(resume_path))
    human_pause(page, 800, 2000)

    fill_labeled(page, "First Name", applicant.first_name)
    fill_labeled(page, "Last Name", applicant.last_name)
    fill_labeled(page, "Phone Number", applicant.phone)
    fill_labeled(page, "Email", applicant.email)
    fill_labeled(page, "City", applicant.city)
    fill_labeled(page, "Zip Code", applicant.zip_code)

    choose_menu(page, r"(Select a Country|United States).*Show menu", "United States")
    choose_menu(page, r"(Select a State|New Jersey).*Show menu", applicant.state)

    human_pause(page, 500, 1200)
    # Work authorization — Yes
    try:
        page.get_by_role("radio", name="Yes", exact=True).first.click(force=True)
    except Exception:
        pass
    human_pause(page, 300, 800)

    title_question = page.get_by_text(re.compile(r"current or most recent job title", re.I)).first
    title_field = title_question.locator("xpath=following::input[1]")
    try:
        title_field.click(timeout=5000)
        human_pause(page, 200, 500)
        title_field.fill(applicant.recent_title, timeout=10000)
        title_field.evaluate("el => { el.dispatchEvent(new Event('input', {bubbles:true})); el.dispatchEvent(new Event('change', {bubbles:true})); }")
    except Exception:
        pass
    human_pause(page, 300, 800)

    page.get_by_role("radio", name="6-10 years", exact=True).click(force=True)
    human_pause(page, 400, 1000)

    for setting in ("Hybrid", "On-site", "Remote"):
        try:
            checkbox = page.get_by_role("checkbox", name=setting, exact=True)
            if checkbox.get_attribute("aria-checked") != "true":
                human_click(page, checkbox)
            human_pause(page, 200, 600)
        except Exception:
            pass

    # Handle optional questions (security clearance, disability, veteran, etc.)
    _handle_optional_questions(page)

    try:
        page.get_by_role("radio", name="I agree to receive text messages.", exact=True).click(force=True)
        human_pause(page, 500, 1500)
    except Exception:
        pass

    wait_for_manual_verification(page)
    page.get_by_role("button", name=re.compile(r"^submit$", re.I)).wait_for(timeout=15000)


def wait_for_confirmation(page: Page, timeout_seconds: int = 30) -> None:
    confirmation = re.compile(r"(thank you|application (?:has been |was )?submitted|submission complete)", re.I)
    for _ in range(timeout_seconds):
        try:
            if confirmation.search(page.locator("body").inner_text(timeout=1000)):
                return
        except Exception:
            pass
        page.wait_for_timeout(1000)
    raise RuntimeError("TEKsystems did not show an application confirmation")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fill TEKsystems applications, with optional submission.")
    parser.add_argument("--jobs-file", type=Path)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--resume", type=Path, required=True)
    parser.add_argument("--first-name", required=True)
    parser.add_argument("--last-name", required=True)
    parser.add_argument("--email", required=True)
    parser.add_argument("--phone", required=True)
    parser.add_argument("--city", default="Jersey City")
    parser.add_argument("--state", default="New Jersey")
    parser.add_argument("--zip-code", default="08540")
    parser.add_argument("--recent-title", default="Senior Python Developer")
    parser.add_argument("--limit", type=int, default=1)
    parser.add_argument("--start-at", type=int, default=1)
    parser.add_argument("--slow-mo", type=int, default=250)
    parser.add_argument("--browser-channel", default="", help="e.g. 'chrome'. Empty = bundled Chromium.")
    parser.add_argument("--profile-dir", type=Path, default=DEFAULT_PROFILE_DIR)
    parser.add_argument("--keep-open-minutes", type=int, default=30)
    parser.add_argument("--submit", action="store_true")
    parser.add_argument("--capsolver-key", default="", help="capsolver.com API key for automated reCAPTCHA.")
    parser.add_argument("--setup-profile", action="store_true", help="Open browser for Google login, then exit.")
    return parser.parse_args()


def main() -> int:
    global _CAPSOLVER_API_KEY
    args = parse_args()
    if args.capsolver_key:
        _CAPSOLVER_API_KEY = args.capsolver_key
    if _CAPSOLVER_API_KEY:
        print("[capsolver] API key set — reCAPTCHA will be solved automatically.", flush=True)
    else:
        print("[capsolver] No API key — will attempt mouse click.", flush=True)

    if args.setup_profile:
        profile_dir = args.profile_dir.expanduser().resolve()
        profile_dir.mkdir(parents=True, exist_ok=True)
        kwargs: dict = dict(headless=False, slow_mo=250, locale="en-US", timezone_id="America/Chicago")
        if args.browser_channel:
            kwargs["channel"] = args.browser_channel
        print(f"Opening browser: {profile_dir}")
        print("Log into Google (google.com), then press Enter here.")
        with sync_playwright() as pw:
            ctx = pw.chromium.launch_persistent_context(str(profile_dir), **kwargs)
            pg = ctx.new_page()
            pg.goto("https://google.com")
            input("Press Enter after logging in...")
            ctx.close()
        print("Profile saved. Run script normally now.")
        return 0

    resume_path = args.resume.expanduser().resolve()
    if not resume_path.exists():
        print(f"Resume not found: {resume_path}", file=sys.stderr)
        return 2

    jobs_file = args.jobs_file or latest_jobs_file(args.out_dir)
    jobs = load_jobs(jobs_file)
    start = max(args.start_at - 1, 0)
    selected = jobs[start : start + args.limit] if args.limit > 0 else jobs[start:]
    if not selected:
        print("No jobs selected.")
        return 0

    applicant = Applicant(
        args.first_name, args.last_name, args.email, args.phone,
        args.city, args.state, args.zip_code, args.recent_title,
    )
    run_dir = args.out_dir / f"apply_{datetime.now():%Y%m%d_%H%M%S}"
    run_dir.mkdir(parents=True, exist_ok=True)
    report_path = run_dir / "apply_report.csv"
    report_rows: list[dict[str, str]] = []

    print(f"Jobs source: {jobs_file}")
    print(f"Selected jobs: {len(selected)}")
    print(f"Mode: {'submit' if args.submit else 'fill for review'}")

    viewport_width = random.choice([1280, 1366, 1440, 1536, 1920])
    viewport_height = random.choice([768, 800, 864, 900, 1080])

    with sync_playwright() as playwright:
        profile_dir = args.profile_dir.expanduser().resolve()
        profile_dir.mkdir(parents=True, exist_ok=True)
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
        context = playwright.chromium.launch_persistent_context(str(profile_dir), **launch_kwargs)
        first_page = context.pages[0] if context.pages else context.new_page()
        apply_stealth(first_page)

        try:
            for index, job in enumerate(selected, start=start + 1):
                if index != start + 1:
                    gap = random.uniform(8.0, 20.0)
                    print(f"[gap] Waiting {gap:.0f}s before next job...", flush=True)
                    time.sleep(gap)
                    page = context.new_page()
                    apply_stealth(page)
                else:
                    page = first_page

                title = str(job.get("title") or "").strip()
                job_id = str(job.get("job_id") or "").strip()
                job_url = str(job.get("job_url") or "").strip()
                screenshot = run_dir / f"{index:03d}_{job_id.lower()}_filled.png"
                print(f"\n[{index}/{len(jobs)}] {title} | {job_id}")

                status = "filled"
                message = "Application filled and waiting for review."
                try:
                    time.sleep(random.uniform(2.0, 5.0))
                    page.goto(job_url, wait_until="domcontentloaded", timeout=60000)
                    human_pause(page, 1800, 4000)
                    _idle_mouse_wander(page, viewport_width, viewport_height)
                    human_scroll(page)
                    human_pause(page, 1000, 2500)

                    close_cookie_banner(page)

                    original_url = page.url
                    apply_page = None

                    # Try new tab first
                    try:
                        with context.expect_page(timeout=6000) as new_page_info:
                            human_click(page, page.get_by_role("button", name="Apply Now", exact=True))
                        apply_page = new_page_info.value
                        apply_page.wait_for_load_state("domcontentloaded", timeout=30000)
                        apply_stealth(apply_page)
                        print(f"[nav] Apply new tab: {apply_page.url}", flush=True)
                    except Exception:
                        # Same tab — wait for URL to actually change from job listing page
                        print("[nav] No new tab — waiting for URL change...", flush=True)
                        try:
                            page.wait_for_function(
                                f"() => window.location.href !== {json.dumps(original_url)}",
                                timeout=20000,
                            )
                        except Exception:
                            pass
                        page.wait_for_load_state("domcontentloaded", timeout=30000)
                        apply_page = page
                        print(f"[nav] Apply same tab: {apply_page.url}", flush=True)

                    # If URL never changed, Apply Now might need a scroll-into-view click
                    if apply_page.url == original_url:
                        print("[nav] URL unchanged — retrying Apply Now click...", flush=True)
                        btn = apply_page.get_by_role("button", name="Apply Now", exact=True).first
                        btn.scroll_into_view_if_needed(timeout=5000)
                        human_pause(apply_page, 500, 1000)
                        try:
                            with context.expect_page(timeout=8000) as new_page_info:
                                btn.click(force=True)
                            apply_page = new_page_info.value
                            apply_page.wait_for_load_state("domcontentloaded", timeout=30000)
                            apply_stealth(apply_page)
                            print(f"[nav] Apply new tab (retry): {apply_page.url}", flush=True)
                        except Exception:
                            try:
                                page.wait_for_function(
                                    f"() => window.location.href !== {json.dumps(original_url)}",
                                    timeout=15000,
                                )
                                apply_page = page
                            except Exception:
                                apply_page = page
                        print(f"[nav] Final URL: {apply_page.url}", flush=True)

                    fill_application(apply_page, applicant, resume_path)
                    page = apply_page

                    if args.submit:
                        human_click(page, page.get_by_role("button", name=re.compile(r"^submit$", re.I)))
                        try:
                            page.wait_for_load_state("domcontentloaded", timeout=30000)
                        except PlaywrightTimeoutError:
                            pass
                        wait_for_manual_verification(page)
                        wait_for_confirmation(page)
                        status = "submitted"
                        message = "Application submitted and confirmed."

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

                report_rows.append({"status": status, "job_id": job_id, "title": title, "job_url": job_url, "message": message})
                with report_path.open("w", newline="") as fh:
                    writer = csv.DictWriter(fh, fieldnames=report_rows[0].keys())
                    writer.writeheader()
                    writer.writerows(report_rows)

            print(f"\nReport: {report_path}")
            print(f"Browser remains open for {args.keep_open_minutes} minutes.")
            try:
                first_page.wait_for_timeout(max(args.keep_open_minutes, 1) * 60 * 1000)
            except Exception:
                pass
        finally:
            try:
                context.close()
            except Exception:
                pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
