#!/usr/bin/env python3
"""Standalone Ettain Group scraper.

Ettain is now part of Experis, so this scraper uses the current Experis job feed.
"""

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


ALL_JOBS_URL = "https://www.experis.com/en/all-jobs"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Ettain/Experis jobs into CSV/JSON/Excel.")
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


def location_from_job(item: dict) -> str:
    location = item.get("jobLocation")
    if isinstance(location, list):
        location = location[0] if location else {}
    if isinstance(location, dict):
        address = location.get("address")
        if isinstance(address, dict):
            return clean_text(", ".join(str(address.get(key) or "") for key in ("addressLocality", "addressRegion", "postalCode", "addressCountry")))
    return ""


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)"})
    response = session.get(ALL_JOBS_URL, timeout=timeout)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    links: list[tuple[str, str]] = []
    seen_urls: set[str] = set()
    for anchor in soup.find_all("a", href=True):
        href = str(anchor.get("href") or "")
        if "/en/job/" not in href:
            continue
        url = href if href.startswith("http") else "https://www.experis.com" + href
        if url in seen_urls:
            continue
        seen_urls.add(url)
        links.append((clean_text(anchor.get_text(" ", strip=True)), url))
    print(f"Ettain Group/Experis: found {len(links)} all-jobs links")
    lowered_terms = [term.lower() for term in terms]
    jobs: list[VendorJob] = []
    for title_hint, url in links:
        if len(jobs) >= max_detail_pages:
            break
        hint = " ".join([title_hint, url]).lower()
        if lowered_terms and not any(term in hint for term in lowered_terms):
            continue
        detail = session.get(url, timeout=timeout)
        detail.raise_for_status()
        item = json_ld_job(detail.text)
        if not item:
            continue
        title = clean_text(item.get("title")) or title_hint
        raw_text = clean_text(item.get("description"))
        rank, reasons = score_title(title, raw_text)
        posted = parse_posted_date(item.get("datePosted"))
        jobs.append(VendorJob(
            "Ettain Group",
            next((term for term in terms if term.lower() in " ".join([title, raw_text, url]).lower()), terms[0] if terms else ""),
            rank,
            reasons,
            title,
            clean_text(item.get("industry") or item.get("occupationalCategory")),
            location_from_job(item),
            clean_text(item.get("employmentType")),
            clean_text(item.get("baseSalary")),
            posted.date().isoformat() if posted else clean_text(item.get("datePosted")),
            Path(url.rstrip("/")).parts[-2] if len(Path(url.rstrip("/")).parts) > 1 else Path(url).name,
            url,
            clean_text(item.get("url")) or url,
            extract_contact_info(raw_text),
            raw_text[:900],
            raw_text,
        ))
    print(f"Extracted {len(jobs)} unique Ettain Group/Experis jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("ettaingroup", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
