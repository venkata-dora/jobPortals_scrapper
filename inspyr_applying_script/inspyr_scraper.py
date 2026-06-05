#!/usr/bin/env python3
"""Standalone INSPYR Solutions scraper."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


API_URL = "https://www.inspyrsolutions.com/wp-json/wp/v2/job-listings"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape INSPYR Solutions jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def rendered(value: Any) -> str:
    if isinstance(value, dict):
        return clean_text(value.get("rendered"))
    return clean_text(value)


def field(text: str, label: str) -> str:
    match = re.search(rf"\b{re.escape(label)}\s*:\s*(.*?)(?=\s+[A-Z][A-Za-z /]+:|$)", text)
    return clean_text(match.group(1)) if match else ""


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)", "Accept": "application/json"})
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    for term in terms:
        response = session.get(API_URL, params={"search": term, "per_page": min(max_detail_pages, 100)}, timeout=timeout)
        response.raise_for_status()
        postings = response.json()
        print(f"INSPYR Solutions {term}: found {len(postings)} REST jobs")
        for item in postings:
            job_id = clean_text(item.get("id"))
            if not job_id or job_id in seen:
                continue
            seen.add(job_id)
            title = rendered(item.get("title"))
            raw_text = clean_text(rendered(item.get("content")))
            rank, reasons = score_title(title, raw_text)
            jobs.append(VendorJob(
                "INSPYR Solutions",
                term,
                rank,
                reasons,
                title,
                "",
                field(raw_text, "Location"),
                field(raw_text, "Position Type") or field(raw_text, "Duration"),
                field(raw_text, "Compensation"),
                clean_text(item.get("date_gmt") or item.get("date")),
                job_id,
                clean_text(item.get("link")),
                clean_text(item.get("link")),
                extract_contact_info(raw_text),
                raw_text[:900],
                raw_text,
            ))
    print(f"Extracted {len(jobs)} unique INSPYR Solutions jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("inspyr", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
