#!/usr/bin/env python3
"""Standalone Optomi scraper."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, parse_posted_date, score_title, write_outputs


JOBS_URL = "https://shazamme.io/Job-Listing/src/php/actions"
DUDA_SITE_ID = "9c0caad4"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Optomi jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def value(data: dict[str, Any], *keys: str) -> str:
    for key in keys:
        text = clean_text(data.get(key))
        if text:
            return text
    return ""


def scrape_jobs(terms: list[str], timeout: int) -> list[VendorJob]:
    response = requests.get(
        JOBS_URL,
        params={"dudaSiteID": DUDA_SITE_ID, "action": "Get Jobs"},
        timeout=timeout,
        headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"},
    )
    response.raise_for_status()
    items = response.json()
    print(f"Optomi: found {len(items)} Shazamme jobs")
    lowered_terms = [term.lower() for term in terms]
    jobs: list[VendorJob] = []
    for item in items:
        data = item.get("data") or {}
        title = value(data, "jobName", "title")
        raw_text = clean_text(" ".join([
            title,
            value(data, "shortDescription"),
            value(data, "fullDescription"),
            value(data, "category"),
            value(data, "subCategory"),
            value(data, "city"),
            value(data, "state"),
            value(data, "country"),
            value(data, "workModel"),
            value(data, "workType"),
        ]))
        if lowered_terms and not any(term in raw_text.lower() for term in lowered_terms):
            continue
        rank, reasons = score_title(title, raw_text)
        posted = parse_posted_date(data.get("changedOnUTC") or data.get("createdOnUTC"))
        slug = clean_text(item.get("page_item_url")) or value(data, "jobID")
        job_url = value(data, "detailsUri", "jobURL") or f"https://www.optomi.com/job-details/{slug}"
        if job_url.startswith("/"):
            job_url = "https://www.optomi.com" + job_url
        location = ", ".join(part for part in [value(data, "city"), value(data, "state"), value(data, "country")] if part)
        salary = value(data, "salaryText", "salary")
        if not salary:
            salary = " ".join(part for part in [value(data, "salaryFrom"), value(data, "salaryTo"), value(data, "currencyCode")] if part)
        jobs.append(VendorJob(
            "Optomi",
            next((term for term in terms if term.lower() in raw_text.lower()), terms[0] if terms else ""),
            rank,
            reasons,
            title,
            value(data, "category", "subCategory"),
            location,
            " ".join(part for part in [value(data, "workType"), value(data, "workModel")] if part),
            salary,
            posted.date().isoformat() if posted else value(data, "changedOnUTC", "createdOnUTC"),
            value(data, "referenceNumber", "jobID") or slug,
            job_url,
            job_url,
            extract_contact_info(raw_text),
            raw_text[:900],
            raw_text,
        ))
    print(f"Extracted {len(jobs)} unique Optomi jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("optomi", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
