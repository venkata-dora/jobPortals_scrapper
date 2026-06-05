#!/usr/bin/env python3
"""Standalone Matlen Silver scraper."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


BASE = "https://matlensilver.com"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Matlen Silver jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def json_ld_jobs(html: str) -> list[dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            payload = json.loads(script.string or "{}")
        except json.JSONDecodeError:
            continue
        items = payload if isinstance(payload, list) else [payload]
        for item in items:
            if isinstance(item, dict) and item.get("@type") == "JobPosting":
                jobs.append(item)
    return jobs


def location_from_ld(item: dict[str, Any]) -> str:
    locations = item.get("jobLocation")
    if isinstance(locations, dict):
        locations = [locations]
    parts = []
    for location in locations or []:
        if not isinstance(location, dict):
            continue
        address = location.get("address")
        if isinstance(address, dict):
            parts.append(clean_text(", ".join(str(address.get(key) or "") for key in ("addressLocality", "addressRegion", "addressCountry"))))
    return clean_text(" / ".join(part for part in parts if part))


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)"})
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    for term in terms:
        response = session.get(f"{BASE}/", params={"s": term}, timeout=timeout)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        links = []
        for link in soup.select('a[href*="/job/"]'):
            href = str(link.get("href") or "").split("#", 1)[0]
            if href and href not in links:
                links.append(href)
        print(f"Matlen Silver {term}: found {len(links)} search job links")
        for job_url in links[:max_detail_pages]:
            slug = Path(job_url.rstrip("/")).name
            if slug in seen:
                continue
            seen.add(slug)
            detail = session.get(job_url, timeout=timeout)
            detail.raise_for_status()
            items = json_ld_jobs(detail.text)
            soup = BeautifulSoup(detail.text, "html.parser")
            page_text = clean_text(soup.get_text(" ", strip=True))
            title = clean_text((soup.find("h1") or soup.find("title") or "").get_text(" ", strip=True) if soup.find(["h1", "title"]) else slug.replace("-", " "))
            raw_text = ""
            location = ""
            employment_type = ""
            salary = ""
            posted = ""
            if items:
                item = items[0]
                title = clean_text(item.get("title")) or title
                description = clean_text(item.get("description"))
                location = location_from_ld(item)
                employment_type = clean_text(item.get("employmentType"))
                posted = clean_text(item.get("datePosted"))
                salary = clean_text(item.get("baseSalary"))
                raw_text = clean_text(" ".join([title, location, employment_type, salary, posted, description]))
            if not raw_text:
                raw_text = page_text
            if not posted:
                posted_match = re.search(r"\bPublished\s+([A-Z][a-z]+ \d{1,2}, \d{4})", raw_text)
                posted = posted_match.group(1) if posted_match else ""
            if not location:
                location_match = re.search(r"\bLocation\s+(.*?)\s+Category\b", raw_text)
                location = clean_text(location_match.group(1)) if location_match else ""
            rank, reasons = score_title(title, raw_text)
            jobs.append(VendorJob(
                "Matlen Silver",
                term,
                rank,
                reasons,
                title,
                "",
                location,
                employment_type,
                salary,
                posted,
                slug,
                job_url,
                job_url,
                extract_contact_info(raw_text),
                raw_text[:900],
                raw_text,
            ))
    print(f"Extracted {len(jobs)} unique Matlen Silver jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("matlensilver", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
