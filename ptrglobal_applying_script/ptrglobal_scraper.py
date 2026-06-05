#!/usr/bin/env python3
"""Standalone PTR Global scraper."""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


BASE = "https://pinnaclegroup.wd1.myworkdayjobs.com"
SITE = "PinnacleGroup"
SEARCH_URL = f"{BASE}/wday/cxs/pinnaclegroup/{SITE}/jobs"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape PTR Global jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def workday_external_url(path: str) -> str:
    return f"{BASE}/en-US/{SITE}{path}"


def parse_workday_posted(value: str) -> str:
    text = clean_text(value)
    lowered = text.lower()
    if "30+" in lowered:
        return (datetime.now(timezone.utc) - timedelta(days=31)).date().isoformat()
    match = re.search(r"posted\s+(\d+)\s+days?\s+ago", lowered)
    if match:
        return (datetime.now(timezone.utc) - timedelta(days=int(match.group(1)))).date().isoformat()
    if "today" in lowered:
        return datetime.now(timezone.utc).date().isoformat()
    if "yesterday" in lowered:
        return (datetime.now(timezone.utc) - timedelta(days=1)).date().isoformat()
    return text


def is_usa_location(location: str) -> bool:
    text = clean_text(location).lower()
    if not text:
        return False
    non_usa_markers = ("india", "mexico", "canada", "brazil", "philippines", "remote india", "remote mexico")
    if any(marker in text for marker in non_usa_markers):
        return False
    usa_markers = ("united states", " usa", " us", ", us", "remote us", "remote usa")
    state_pattern = r"\b(?:AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b"
    return any(marker in f" {text}" for marker in usa_markers) or re.search(state_pattern, location) is not None


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)", "Accept": "application/json"})
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    details_fetched = 0
    for term in terms:
        response = session.post(
            SEARCH_URL,
            json={"appliedFacets": {}, "limit": 20, "offset": 0, "searchText": term},
            timeout=timeout,
        )
        response.raise_for_status()
        postings = response.json().get("jobPostings") or []
        print(f"PTR Global {term}: found {len(postings)} Workday jobs")
        for posting in postings:
            path = clean_text(posting.get("externalPath"))
            job_id = clean_text((posting.get("bulletFields") or [""])[0]) or Path(path).name
            if not path or job_id in seen:
                continue
            seen.add(job_id)
            title = clean_text(posting.get("title"))
            location = clean_text(posting.get("locationsText"))
            posted = parse_workday_posted(clean_text(posting.get("postedOn")))
            detail_text = ""
            employment_type = ""
            job_url = workday_external_url(path)
            if details_fetched < max_detail_pages:
                detail_api = f"{BASE}/wday/cxs/pinnaclegroup/{SITE}{path}"
                detail_response = session.get(detail_api, timeout=timeout)
                detail_response.raise_for_status()
                info = detail_response.json().get("jobPostingInfo") or {}
                detail_text = clean_text(info.get("jobDescription"))
                title = clean_text(info.get("title")) or title
                location = clean_text(info.get("location")) or location
                posted = parse_workday_posted(clean_text(info.get("postedOn"))) or posted
                employment_type = clean_text(info.get("timeType"))
                job_url = clean_text(info.get("externalUrl")) or job_url
                details_fetched += 1
            if not is_usa_location(location):
                print(f"  Skipped non-USA location: {title} | {location}")
                continue
            if employment_type.lower() == "full time":
                print(f"  Skipped full-time role: {title}")
                continue
            raw_text = clean_text(" ".join([title, location, employment_type, posted, detail_text]))
            rank, reasons = score_title(title, raw_text)
            jobs.append(VendorJob(
                "PTR Global",
                term,
                rank,
                reasons,
                title,
                "",
                location,
                employment_type,
                "",
                posted,
                job_id,
                job_url,
                job_url,
                extract_contact_info(raw_text),
                raw_text[:900],
                raw_text,
            ))
    print(f"Extracted {len(jobs)} unique PTR Global jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("ptrglobal", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
