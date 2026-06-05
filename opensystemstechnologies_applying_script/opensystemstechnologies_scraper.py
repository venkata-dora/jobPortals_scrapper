#!/usr/bin/env python3
"""Standalone Open Systems Technologies scraper."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus, urljoin

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


BASE = "https://www.opensystemstech.com"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Open Systems Technologies jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def json_ld_job(html: str) -> dict[str, Any]:
    soup = BeautifulSoup(html, "html.parser")
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            item = json.loads(script.string or "{}")
        except json.JSONDecodeError:
            continue
        if isinstance(item, dict) and item.get("@type") == "JobPosting":
            return item
    return {}


def location_from_ld(item: dict[str, Any]) -> str:
    location = item.get("jobLocation")
    if isinstance(location, list):
        location = location[0] if location else {}
    address = location.get("address") if isinstance(location, dict) else {}
    if isinstance(address, dict):
        return clean_text(", ".join(str(address.get(key) or "") for key in ("addressLocality", "addressRegion", "addressCountry")))
    return ""


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)"})
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    for term in terms:
        response = session.get(f"{BASE}/search", params={"keywords": term}, timeout=timeout)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        links = []
        for link in soup.select('a[href*="/job/"]'):
            url = urljoin(BASE, str(link.get("href") or "")).split("#", 1)[0]
            if url and url not in links:
                links.append(url)
        print(f"Open Systems Technologies {term}: found {len(links)} job links")
        for job_url in links[:max_detail_pages]:
            job_id = Path(job_url.rstrip("/")).parts[-2] if len(Path(job_url.rstrip("/")).parts) > 1 else Path(job_url).name
            if job_id in seen:
                continue
            seen.add(job_id)
            detail = session.get(job_url, timeout=timeout)
            detail.raise_for_status()
            item = json_ld_job(detail.text)
            soup = BeautifulSoup(detail.text, "html.parser")
            title = clean_text(item.get("title")) or clean_text((soup.find("h1") or "").get_text(" ", strip=True) if soup.find("h1") else "")
            raw_text = clean_text(item.get("description")) or clean_text(soup.get_text(" ", strip=True))
            location = location_from_ld(item)
            rank, reasons = score_title(title, raw_text)
            jobs.append(VendorJob(
                "Open Systems Technologies",
                term,
                rank,
                reasons,
                title,
                clean_text(item.get("industry") or item.get("occupationalCategory")),
                location,
                clean_text(item.get("employmentType")),
                clean_text(item.get("baseSalary")),
                clean_text(item.get("datePosted")),
                clean_text((item.get("identifier") or {}).get("value") if isinstance(item.get("identifier"), dict) else item.get("identifier")) or job_id,
                job_url,
                job_url,
                extract_contact_info(raw_text),
                raw_text[:900],
                raw_text,
            ))
    print(f"Extracted {len(jobs)} unique Open Systems Technologies jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("opensystemstechnologies", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
