#!/usr/bin/env python3
"""Standalone Hays scraper."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from urllib.parse import quote_plus, urljoin

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


BASE = "https://www.hays.com"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Hays jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def label_value(text: str, label: str) -> str:
    labels = "Job type|Location|Profession|Industry|Pay|Share Job"
    match = re.search(rf"{re.escape(label)}\s+(.*?)(?=\s+(?:{labels})\b|$)", text, re.I)
    return clean_text(match.group(1)) if match else ""


def compact_card_text(card: BeautifulSoup) -> str:
    return clean_text(card.get_text(" ", strip=True))


def job_cards(soup: BeautifulSoup) -> list[BeautifulSoup]:
    cards = []
    for link in soup.select("a.jobTitle[href], a.job-card[href]"):
        card = link
        for parent in link.parents:
            text = compact_card_text(parent)
            if parent.name in {"article", "li"} or "Job type" in text or "View Details" in text:
                card = parent
                break
        cards.append(card)
    return cards


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)"})
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    for term in terms:
        urls = [
            f"{BASE}/job-search/{quote_plus(term)}-jobs",
            f"{BASE}/jobs?search={quote_plus(term)}",
        ]
        for search_url in urls:
            response = session.get(search_url, timeout=timeout)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            cards = job_cards(soup)
            print(f"Hays {term}: found {len(cards)} cards at {search_url}")
            for card in cards[:max_detail_pages]:
                link = card.select_one("a.jobTitle[href], a.job-card[href]")
                if not link:
                    continue
                job_url = urljoin(BASE, str(link.get("href") or ""))
                job_id_match = re.search(r"_(\d+)", job_url)
                job_id = job_id_match.group(1) if job_id_match else job_url
                if job_id in seen:
                    continue
                seen.add(job_id)
                title = clean_text((link.select_one("h3,h4") or link).get_text(" ", strip=True))
                raw_text = compact_card_text(card)
                location = label_value(raw_text, "Location")
                employment_type = label_value(raw_text, "Job type")
                salary = label_value(raw_text, "Pay")
                posted_match = re.search(r"\b\d{2}-\d{2}-\d{4}\b", raw_text)
                posted = posted_match.group(0) if posted_match else ""
                category = clean_text(" / ".join(part for part in [label_value(raw_text, "Profession"), label_value(raw_text, "Industry")] if part))
                rank, reasons = score_title(title, raw_text)
                jobs.append(VendorJob(
                    "Hays",
                    term,
                    rank,
                    reasons,
                    title,
                    category,
                    location,
                    employment_type,
                    salary,
                    posted,
                    job_id,
                    job_url,
                    job_url,
                    extract_contact_info(raw_text),
                    raw_text[:900],
                    raw_text,
                ))
    print(f"Extracted {len(jobs)} unique Hays jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("hays", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
