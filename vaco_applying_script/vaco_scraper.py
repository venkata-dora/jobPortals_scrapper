#!/usr/bin/env python3
"""Standalone Vaco scraper."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, parse_posted_date, score_title, write_outputs


SEARCH_URL = "https://jobs.vaco.com/api/requisitions/search"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Vaco jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def json_ld_job(html: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            payload = json.loads(script.string or "{}")
        except json.JSONDecodeError:
            continue
        if isinstance(payload, dict) and payload.get("@type") == "JobPosting":
            return payload
    return {}


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)", "X-Requested-With": "XMLHttpRequest"})
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    details_fetched = 0
    for term in terms:
        response = session.get(SEARCH_URL, params={"keywords": term}, timeout=timeout)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        rows = [row for row in soup.select("tr") if row.select_one("td.job-title a[href]")]
        print(f"Vaco {term}: found {len(rows)} search jobs")
        for row in rows:
            link = row.select_one("td.job-title a[href]")
            if not link:
                continue
            job_url = str(link.get("href") or "").strip()
            if not job_url or job_url in seen:
                continue
            seen.add(job_url)
            cells = row.find_all("td")
            title = clean_text(link.get_text(" ", strip=True))
            category = clean_text(cells[1].get_text(" ", strip=True) if len(cells) > 1 else "")
            location = clean_text(cells[2].get_text(" ", strip=True) if len(cells) > 2 else "")
            raw_text = clean_text(row.get_text(" ", strip=True))
            posted = ""
            employment_type = ""
            salary = ""
            if details_fetched < max_detail_pages:
                detail = session.get(job_url, timeout=timeout)
                detail.raise_for_status()
                item = json_ld_job(detail.text)
                if item:
                    title = clean_text(item.get("title")) or title
                    description = clean_text(item.get("description"))
                    raw_text = clean_text(" ".join([raw_text, description]))
                    employment_type = clean_text(item.get("employmentType"))
                    salary = clean_text(item.get("baseSalary"))
                    posted = clean_text(item.get("datePosted"))
                    locations = item.get("jobLocation")
                    if isinstance(locations, list):
                        locations = locations[0] if locations else {}
                    if isinstance(locations, dict):
                        address = locations.get("address")
                        if isinstance(address, dict):
                            location = clean_text(", ".join(str(address.get(key) or "") for key in ("addressLocality", "addressRegion", "addressCountry"))) or location
                details_fetched += 1
            rank, reasons = score_title(title, raw_text)
            jobs.append(VendorJob(
                "Vaco",
                term,
                rank,
                reasons,
                title,
                category,
                location,
                employment_type,
                salary,
                posted,
                Path(job_url.rstrip("/")).parts[-2] if len(Path(job_url.rstrip("/")).parts) > 1 else Path(job_url).name,
                job_url,
                job_url,
                extract_contact_info(raw_text),
                raw_text[:900],
                raw_text,
            ))
    print(f"Extracted {len(jobs)} unique Vaco jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("vaco", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
