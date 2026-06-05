#!/usr/bin/env python3
"""Standalone Venturi Group scraper."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, parse_posted_date, score_title, write_outputs


LIST_URL = "https://api.venturi-group.com/api/v1/jobs"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Venturi Group jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def nested_name(value: Any) -> str:
    if isinstance(value, dict):
        return clean_text(value.get("name"))
    return clean_text(value)


def job_url(job_id: Any) -> str:
    return f"https://api.venturi-group.com/api/v1/jobs/{job_id}"


def scrape_jobs(terms: list[str], timeout: int) -> list[VendorJob]:
    response = requests.get(LIST_URL, timeout=timeout, headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
    response.raise_for_status()
    items = response.json().get("data") or []
    print(f"Venturi Group: found {len(items)} API jobs")
    lowered_terms = [term.lower() for term in terms]
    jobs: list[VendorJob] = []
    for item in items:
        title = clean_text(item.get("title"))
        raw_text = clean_text(" ".join([
            title,
            clean_text(item.get("job_description")),
            clean_text(item.get("job_requirement")),
            nested_name(item.get("location")),
            nested_name(item.get("city")),
            clean_text(item.get("job_type")),
            clean_text(item.get("job_contract")),
        ]))
        if lowered_terms and not any(term in raw_text.lower() for term in lowered_terms):
            continue
        rank, reasons = score_title(title, raw_text)
        posted = parse_posted_date(item.get("created_at"))
        location = ", ".join(part for part in [nested_name(item.get("city")), nested_name(item.get("location"))] if part)
        salary = " ".join(clean_text(item.get(key)) for key in ("salary_min", "salary_max", "salary_type") if clean_text(item.get(key)))
        jobs.append(VendorJob(
            "Venturi Group",
            next((term for term in terms if term.lower() in raw_text.lower()), terms[0] if terms else ""),
            rank,
            reasons,
            title,
            "",
            location,
            clean_text(" ".join([str(item.get("job_type") or ""), str(item.get("job_contract") or "")])),
            salary,
            posted.date().isoformat() if posted else clean_text(item.get("created_at")),
            str(item.get("id") or ""),
            job_url(item.get("id")),
            job_url(item.get("id")),
            extract_contact_info(raw_text),
            raw_text[:900],
            raw_text,
        ))
    print(f"Extracted {len(jobs)} unique Venturi Group jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("venturigroup", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
