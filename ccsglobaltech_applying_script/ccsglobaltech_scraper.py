#!/usr/bin/env python3
"""Standalone CCS Global Tech scraper."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, parse_posted_date, score_title, write_outputs


API_KEYS = [
    "djZkcG95b2J1aWhSdDVNMFgzRVpoQT09",
    "d0k5T0psZ0FIcS9ZQTBiS2laMytaZz09",
    "aU82TXVod3hKSmQvMCtOUlNlZUtYUT09",
]
CP_ID = "Z3RkUkt2OXZJVld2MjFpOVRSTXoxZz09"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape CCS Global Tech jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def first_value(item: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = clean_text(item.get(key))
        if value:
            return value
    return ""


def scrape_jobs(terms: list[str], timeout: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (compatible; local job search scraper)",
        "Origin": "https://jobsapi.ceipal.com",
        "Referer": "https://jobsapi.ceipal.com/",
    })
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    for api_key in API_KEYS:
        for term in terms:
            url = f"https://careerapi.ceipal.com/{api_key}/CareerPortalJobPostings/?page=1"
            form = {
                "page": "1",
                "api_key": api_key,
                "method": "CareerPortalJobPostings",
                "cp_id": CP_ID,
                "from_career_portal": "1",
                "sortorder": "desc",
                "sortby": "modified",
                "searchkey": term,
            }
            response = session.post(url, data=form, timeout=timeout)
            response.raise_for_status()
            payload = response.json()
            results = payload.get("results") or []
            print(f"CCS Global Tech {term}: found {len(results)} CEIPAL jobs")
            for item in results:
                job_id = str(item.get("job_id") or item.get("id") or item.get("job_code") or "")
                if not job_id or job_id in seen:
                    continue
                seen.add(job_id)
                title = first_value(item, "public_job_title", "position_title")
                raw_text = clean_text(" ".join([
                    first_value(item, "requistion_description", "public_job_desc"),
                    first_value(item, "tax_terms"),
                    first_value(item, "multpile_job_location"),
                ]))
                rank, reasons = score_title(title, raw_text)
                posted = parse_posted_date(item.get("modified") or item.get("created"))
                pay_rates = item.get("pay_rates") or []
                salary = ""
                if pay_rates and isinstance(pay_rates[0], dict):
                    salary = clean_text(" ".join(str(pay_rates[0].get(key) or "") for key in ("min_pay_rate", "max_pay_rate", "pay_rate_currency", "pay_rate_pay_frequency_type")))
                job_url = first_value(item, "campus_portal_job_details_url", "easy_apply_job_login_hc", "apply_job")
                apply_url = first_value(item, "apply_job", "apply_job_without_registration", "apply_job_login") or job_url
                jobs.append(VendorJob(
                    "CCS Global Tech",
                    term,
                    rank,
                    reasons,
                    title,
                    first_value(item, "industry"),
                    first_value(item, "multpile_job_location", "city", "state", "country"),
                    first_value(item, "tax_terms"),
                    salary,
                    posted.date().isoformat() if posted else first_value(item, "modified", "created"),
                    job_id,
                    job_url,
                    apply_url,
                    extract_contact_info(raw_text),
                    raw_text[:900],
                    raw_text,
                ))
    print(f"Extracted {len(jobs)} unique CCS Global Tech jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("ccsglobaltech", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
