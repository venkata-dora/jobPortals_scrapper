#!/usr/bin/env python3
"""Standalone Oliver James scraper."""

from __future__ import annotations

import argparse
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


SITEMAP_URL = "https://www.oliverjames.com/job-sitemap.xml"
READER_URL = "https://r.jina.ai/http://r.jina.ai/http://https://www.oliverjames.com/en/job-search/"
JOB_LINK_RE = re.compile(r"## \[([^\]]+)\]\((https://www\.oliverjames\.com/en/jobs/[^)]+)\)")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Oliver James jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def title_from_url(url: str) -> str:
    slug = Path(url.rstrip("/")).name
    return clean_text(re.sub(r"-\d+$", "", slug).replace("-", " ").title())


def sitemap_urls(timeout: int) -> list[str]:
    response = requests.get(SITEMAP_URL, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    root = ET.fromstring(response.text)
    namespace = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}
    return [loc.text or "" for loc in root.findall(".//sm:loc", namespace) if loc.text]


def reader_cards(timeout: int) -> dict[str, str]:
    response = requests.get(READER_URL, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    return {url: clean_text(title) for title, url in JOB_LINK_RE.findall(response.text)}


def scrape_jobs(terms: list[str], timeout: int) -> list[VendorJob]:
    lowered_terms = [term.lower() for term in terms]
    cards = reader_cards(timeout)
    urls = sitemap_urls(timeout)
    print(f"Oliver James: found {len(urls)} sitemap jobs and {len(cards)} reader cards")
    jobs: list[VendorJob] = []
    for url in urls:
        title = cards.get(url) or title_from_url(url)
        raw_text = clean_text(" ".join([title, url]))
        if lowered_terms and not any(term in raw_text.lower() for term in lowered_terms):
            continue
        rank, reasons = score_title(title, raw_text)
        jobs.append(VendorJob(
            "Oliver James",
            next((term for term in terms if term.lower() in raw_text.lower()), terms[0] if terms else ""),
            rank,
            reasons,
            title,
            "",
            "",
            "",
            "",
            "",
            Path(url.rstrip("/")).name,
            url,
            url,
            extract_contact_info(raw_text),
            raw_text[:900],
            raw_text,
        ))
    print(f"Extracted {len(jobs)} unique Oliver James jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("oliverjames", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
