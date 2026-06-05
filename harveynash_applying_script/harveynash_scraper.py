#!/usr/bin/env python3
"""Standalone Harvey Nash scraper."""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


SEARCH_URLS = [
    "https://careers.harveynashusa.com/",
    "https://careers.harveynashusa.com/career-areas/engineering-jobs",
]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Harvey Nash jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int, sleep: float) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)"})
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    detail_count = 0
    lowered_terms = [term.lower() for term in terms]
    for url in SEARCH_URLS:
        response = session.get(url, timeout=timeout)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        cards = soup.select("[data-title][data-location]")
        print(f"Harvey Nash: found {len(cards)} job cards at {url}")
        for card in cards:
            link = card.select_one("a.post-title[href]") or card.select_one("a[href*='job-details']")
            if not link:
                continue
            job_url = str(link.get("href") or "").strip()
            if not job_url or job_url in seen:
                continue
            title = clean_text(card.get("data-title") or link.get_text(" ", strip=True))
            card_text = clean_text(card.get_text(" ", strip=True))
            haystack = " ".join([title, card_text]).lower()
            if lowered_terms and not any(term in haystack for term in lowered_terms):
                continue
            seen.add(job_url)
            detail_text = ""
            if detail_count < max_detail_pages:
                detail_response = session.get(job_url, timeout=timeout)
                detail_response.raise_for_status()
                detail_soup = BeautifulSoup(detail_response.text, "html.parser")
                for tag in detail_soup(["script", "style", "noscript"]):
                    tag.decompose()
                detail_text = clean_text(detail_soup.get_text(" ", strip=True))
                detail_count += 1
                time.sleep(sleep)
            raw_text = clean_text(" ".join([card_text, detail_text]))
            rank, reasons = score_title(title, raw_text)
            salary_el = card.select_one(".post-salary")
            jobs.append(VendorJob(
                "Harvey Nash",
                next((term for term in terms if term.lower() in haystack), terms[0] if terms else ""),
                rank,
                reasons,
                title,
                clean_text(card.get("data-category")),
                clean_text(card.get("data-location")),
                clean_text(" ".join([str(card.get("data-status") or ""), str(card.get("data-type") or "")])),
                clean_text(salary_el.get_text(" ", strip=True) if salary_el else ""),
                "",
                Path(job_url.rstrip("/")).name,
                job_url,
                job_url,
                extract_contact_info(raw_text),
                raw_text[:900],
                raw_text,
            ))
    print(f"Extracted {len(jobs)} unique Harvey Nash jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages, args.sleep)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("harveynash", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
