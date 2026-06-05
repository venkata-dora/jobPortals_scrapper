#!/usr/bin/env python3
"""Fill and optionally submit a Beacon Hill application in Chrome."""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


DEFAULT_JOB_URL = "https://bhsg.com/jobs/job/ns-iamd-ns_1778632903-sr-automation-engineer-iam-identity-governance-automation-seattle-washington/"
DEFAULT_RESUME = Path("/Users/saisujan/Desktop/jobs_scraping/venkataD_resume.docx")


@dataclass(frozen=True)
class Applicant:
    first_name: str = "Venkata"
    last_name: str = "Dora"
    email: str = "venkatasworkofficial@gmail.com"
    phone: str = "9735445393"
    gender: str = "Male"
    ethnicity: str = "Not Hispanic or Latino"
    race: str = "Asian"
    protected_veteran: str = "No"
    disability: str = "No, I don't have a disability"


def today_mmddyyyy() -> str:
    return datetime.now().strftime("%m/%d/%Y")


def close_cookie_banner(page: Page) -> None:
    for pattern in (r"accept", r"agree", r"close"):
        try:
            button = page.get_by_role("button", name=re.compile(pattern, re.I)).first
            if button.is_visible(timeout=1200):
                button.click(timeout=2500)
                page.wait_for_timeout(300)
                return
        except PlaywrightTimeoutError:
            pass


def choose_radio_or_checkbox(page: Page, label: str) -> None:
    choice = page.get_by_label(label, exact=True).locator("visible=true").first
    if not choice.is_checked(timeout=2000):
        choice.check(timeout=7000)


def open_application(page: Page, url: str) -> None:
    page.goto(url, wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(2500)
    close_cookie_banner(page)


def show_apply_modal(page: Page) -> None:
    page.evaluate(
        """() => {
            const overlay = document.getElementById('pum-4920');
            const container = document.getElementById('popmake-4920');
            const wrapper = document.getElementById('gform_wrapper_10');
            if (!overlay || !container || !wrapper) {
                throw new Error('Beacon Hill apply modal/form was not found');
            }

            overlay.classList.add('pum-active');
            overlay.setAttribute('aria-modal', 'true');
            Object.assign(overlay.style, {
                display: 'block',
                opacity: '1',
                position: 'fixed',
                inset: '0px',
                overflowY: 'auto',
                zIndex: '1999999999',
                backgroundColor: 'rgba(0, 0, 0, 0.72)'
            });

            Object.assign(container.style, {
                display: 'block',
                opacity: '1',
                visibility: 'visible',
                position: 'relative',
                margin: '24px auto',
                maxWidth: '1100px',
                width: '72vw',
                left: 'auto',
                top: 'auto'
            });

            Object.assign(wrapper.style, {
                display: 'block',
                opacity: '1',
                visibility: 'visible'
            });
            document.body.classList.add('pum-open');
        }"""
    )
    page.wait_for_timeout(500)


def fill_main_application(page: Page, applicant: Applicant, resume_path: Path) -> None:
    for attempt in range(3):
        try:
            page.locator("#gform_10").wait_for(state="attached", timeout=25000)
            break
        except PlaywrightTimeoutError:
            if attempt == 2:
                raise
            page.reload(wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(4000)
    show_apply_modal(page)
    # Target the Gravity Form fields directly by stable IDs. The modal is made
    # visible first so this is watchable in Chrome.
    page.evaluate(
        """applicant => {
            const setValue = (id, value) => {
                const el = document.getElementById(id);
                if (!el) throw new Error(`Missing field: ${id}`);
                el.value = value;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            };
            setValue('input_10_1_3', applicant.first_name);
            setValue('input_10_1_6', applicant.last_name);
            setValue('input_10_2', applicant.email);
            setValue('input_10_3', applicant.phone);
            setValue('input_10_4', 'Yes');
            const selfId = document.getElementById('choice_10_22_0');
            if (!selfId) throw new Error('Missing self-identification choice');
            selfId.checked = true;
            selfId.dispatchEvent(new Event('change', { bubbles: true }));
        }""",
        {
            "first_name": applicant.first_name,
            "last_name": applicant.last_name,
            "email": applicant.email,
            "phone": applicant.phone,
        },
    )
    page.locator("#input_10_5").set_input_files(str(resume_path))
    page.wait_for_timeout(1000)


def fill_self_identification(page: Page, applicant: Applicant, signature_date: str) -> None:
    page.get_by_text(re.compile(r"equal employment opportunity|voluntary self-identification", re.I)).wait_for(
        timeout=20000
    )
    choose_radio_or_checkbox(page, applicant.gender)
    choose_radio_or_checkbox(page, applicant.ethnicity)
    choose_radio_or_checkbox(page, applicant.race)
    choose_radio_or_checkbox(page, applicant.protected_veteran)
    choose_radio_or_checkbox(page, applicant.disability)

    # The self-identification section repeats First/Last after the main applicant name.
    first_fields = page.get_by_label(re.compile(r"^first\s*\*?$", re.I)).locator("visible=true")
    last_fields = page.get_by_label(re.compile(r"^last\s*\*?$", re.I)).locator("visible=true")
    if first_fields.count() > 1:
        first_fields.nth(first_fields.count() - 1).fill(applicant.first_name, timeout=7000)
    if last_fields.count() > 1:
        last_fields.nth(last_fields.count() - 1).fill(applicant.last_name, timeout=7000)

    page.get_by_label(re.compile(r"^date\s*\*?$", re.I)).locator("visible=true").first.fill(signature_date, timeout=7000)
    choose_radio_or_checkbox(page, "Digital Signature Consent")


def submit_main_application(page: Page) -> None:
    page.evaluate(
        """() => {
            const submit = document.getElementById('gform_submit_button_10');
            const form = document.getElementById('gform_10');
            if (!submit || !form) throw new Error('Missing Beacon Hill submit form/button');
            if (form.requestSubmit) form.requestSubmit(submit);
            else submit.click();
        }"""
    )


def submit_visible_application(page: Page) -> None:
    submit = page.get_by_role("button", name=re.compile(r"apply now|submit", re.I)).locator("visible=true").last
    submit.scroll_into_view_if_needed(timeout=5000)
    submit.click(timeout=10000)


def submission_status(page: Page) -> str:
    page.wait_for_timeout(6000)
    body = page.locator("body").inner_text(timeout=10000)
    for pattern in (r"thank you", r"success", r"submitted", r"received"):
        if re.search(pattern, body, re.I):
            return f"possible_success: matched '{pattern}' at {page.url}"
    return f"unknown: no success text found at {page.url}"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fill a Beacon Hill application in Chrome.")
    parser.add_argument("--url", default=DEFAULT_JOB_URL, help="Beacon Hill job detail URL.")
    parser.add_argument("--resume", type=Path, default=DEFAULT_RESUME)
    parser.add_argument("--first-name", default=Applicant.first_name)
    parser.add_argument("--last-name", default=Applicant.last_name)
    parser.add_argument("--email", default=Applicant.email)
    parser.add_argument("--phone", default=Applicant.phone)
    parser.add_argument("--date", default=today_mmddyyyy(), help="Signature date, defaults to today's date.")
    parser.add_argument("--submit", action="store_true", help="Click Apply Now/Submit after filling.")
    parser.add_argument("--demo", action="store_true", help="Show the visible apply modal and leave it open.")
    parser.add_argument("--keep-open-seconds", type=int, default=45)
    parser.add_argument(
        "--screenshot",
        type=Path,
        help="Path for a screenshot after fill/submit. Defaults to beaconhill_applying_script/output/apply_*.png.",
    )
    parser.add_argument("--browser-channel", default="chrome", help="Use installed Google Chrome by default.")
    parser.add_argument("--slow-mo", type=int, default=200)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    resume_path = args.resume.expanduser().resolve()
    if not resume_path.exists():
        print(f"Resume not found: {resume_path}", file=sys.stderr)
        return 2

    screenshot_path = args.screenshot
    if screenshot_path is None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = Path(__file__).resolve().parent / "output" / f"apply_{stamp}.png"
    screenshot_path = screenshot_path.expanduser().resolve()
    screenshot_path.parent.mkdir(parents=True, exist_ok=True)

    applicant = Applicant(
        first_name=args.first_name,
        last_name=args.last_name,
        email=args.email,
        phone=args.phone,
    )

    print(f"Opening Chrome: {args.url}")
    print(f"Resume: {resume_path}")
    print(f"Signature date: {args.date}")
    print("Mode: submit" if args.submit else "Mode: fill only")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(channel=args.browser_channel, headless=False, slow_mo=args.slow_mo)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            open_application(page, args.url)
            fill_main_application(page, applicant, resume_path)
            print("Filled main application fields and resume.")

            if args.demo and not args.submit:
                print("Demo mode: visible modal is open and filled; submit was not clicked.")
            elif args.submit:
                submit_main_application(page)
                print("Clicked main Apply Now.")
                try:
                    fill_self_identification(page, applicant, args.date)
                    print("Filled self-identification and signature date.")
                    submit_visible_application(page)
                    print("Clicked final Apply Now/Submit.")
                except Exception as exc:
                    print(f"Self-identification step was not reached or could not be filled: {exc}")
                print(f"Post-submit status: {submission_status(page)}")
            else:
                print("Submit was not clicked. Re-run with --submit to submit automatically.")

            page.screenshot(path=str(screenshot_path), full_page=True)
            print(f"Screenshot: {screenshot_path}")
            page.wait_for_timeout(max(args.keep_open_seconds, 1) * 1000)
        finally:
            context.close()
            browser.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
