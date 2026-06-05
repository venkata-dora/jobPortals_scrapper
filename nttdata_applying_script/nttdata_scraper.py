#!/usr/bin/env python3
"""Standalone NTT DATA scraper."""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


BASE = "https://nttdata.eightfold.ai"
DOMAIN = "nttdata.com"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape NTT DATA jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def start_session(timeout: int) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (compatible; local job search scraper)",
        "Accept": "application/json, text/plain, */*",
        "Referer": f"{BASE}/careers",
    })
    response = session.get(f"{BASE}/careers", timeout=timeout)
    response.raise_for_status()
    match = re.search(r'<meta name="_csrf" content="([^"]+)', response.text)
    if match:
        session.headers.update({"x-csrf-token": match.group(1)})
    return session


def posted_from_epoch(value: Any) -> str:
    if not value:
        return ""
    try:
        return datetime.fromtimestamp(int(value), tz=timezone.utc).date().isoformat()
    except (TypeError, ValueError, OSError):
        return clean_text(value)


def position_url(position: dict[str, Any]) -> str:
    path = clean_text(position.get("positionUrl")) or f"/careers/job/{position.get('id')}"
    return urljoin(BASE, path)


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = start_session(timeout)
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    details_fetched = 0
    for term in terms:
        response = session.get(
            f"{BASE}/api/pcsx/search",
            params={"domain": DOMAIN, "query": term, "start": 0, "num": 20},
            timeout=timeout,
        )
        response.raise_for_status()
        positions = ((response.json().get("data") or {}).get("positions") or [])
        print(f"NTT DATA {term}: found {len(positions)} Eightfold jobs")
        for position in positions:
            job_id = clean_text(position.get("id") or position.get("displayJobId"))
            if not job_id or job_id in seen:
                continue
            seen.add(job_id)
            title = clean_text(position.get("name"))
            location = clean_text(", ".join(position.get("locations") or position.get("standardizedLocations") or []))
            posted = posted_from_epoch(position.get("postedTs") or position.get("creationTs"))
            category = clean_text(position.get("department"))
            employment_type = clean_text(position.get("workLocationOption"))
            raw_text = clean_text(" ".join([title, location, category, employment_type]))
            url = position_url(position)
            if details_fetched < max_detail_pages:
                detail = session.get(
                    f"{BASE}/api/pcsx/position_details",
                    params={"domain": DOMAIN, "position_id": job_id},
                    timeout=timeout,
                )
                detail.raise_for_status()
                info = (detail.json().get("data") or {})
                title = clean_text(info.get("name")) or title
                location = clean_text(", ".join(info.get("locations") or info.get("standardizedLocations") or [])) or location
                category = clean_text(info.get("department")) or category
                employment_type = clean_text(info.get("workLocationOption")) or employment_type
                posted = posted_from_epoch(info.get("postedTs") or info.get("creationTs")) or posted
                url = position_url(info)
                description = clean_text(info.get("jobDescription"))
                raw_text = clean_text(" ".join([title, location, category, employment_type, description]))
                details_fetched += 1
            rank, reasons = score_title(title, raw_text)
            jobs.append(VendorJob(
                "NTT DATA",
                term,
                rank,
                reasons,
                title,
                category,
                location,
                employment_type,
                "",
                posted,
                job_id,
                url,
                url,
                extract_contact_info(raw_text),
                raw_text[:900],
                raw_text,
            ))
    print(f"Extracted {len(jobs)} unique NTT DATA jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("nttdata", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
