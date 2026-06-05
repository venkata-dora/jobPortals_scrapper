#!/usr/bin/env python3
"""Standalone Motion Recruitment scraper."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus, urljoin

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, score_title, write_outputs


BASE = "https://motionrecruitment.com"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Motion Recruitment jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def route_for(term: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", term.lower()).strip("-")
    known = {
        "python": "python",
        "data": "data-engineering",
        "data-engineer": "data-engineering",
        "data-engineering": "data-engineering",
        "machine-learning": "machine-learning-data-science",
        "ai-engineer": "machine-learning-data-science",
        "backend": "software-engineering",
        "full-stack": "software-engineering",
    }
    return known.get(slug, slug)


def fetch(session: requests.Session, url: str, timeout: int) -> str:
    last_error = None
    for candidate in (url, "https://r.jina.ai/http://" + url.removeprefix("https://")):
        try:
            response = session.get(candidate, timeout=timeout)
            response.raise_for_status()
            if response.text.strip():
                return response.text
        except requests.RequestException as exc:
            last_error = exc
    if last_error:
        raise last_error
    return ""


def job_id_from_url(url: str) -> str:
    match = re.search(r"/(\d+)(?:[/?#]|$)", url)
    return match.group(1) if match else Path(url.rstrip("/")).name


def salary_from_ld(item: dict[str, Any]) -> str:
    salary = item.get("baseSalary")
    if not isinstance(salary, dict):
        return clean_text(salary)
    value = salary.get("value")
    currency = clean_text(salary.get("currency"))
    if isinstance(value, dict):
        minimum = value.get("minValue")
        maximum = value.get("maxValue")
        unit = clean_text(value.get("unitText")).lower()
        if minimum and maximum:
            suffix = "/hr" if unit == "hour" else f"/{unit}" if unit else ""
            return f"{currency} {minimum}-{maximum}{suffix}".strip()
    return clean_text(salary)


def location_from_ld(item: dict[str, Any]) -> str:
    location = item.get("jobLocation")
    if isinstance(location, list):
        location = location[0] if location else {}
    address = location.get("address") if isinstance(location, dict) else {}
    if isinstance(address, dict):
        return clean_text(", ".join(str(address.get(key) or "") for key in ("addressLocality", "addressRegion", "addressCountry")))
    return ""


def is_usa_job(item: dict[str, Any]) -> bool:
    location = item.get("jobLocation")
    if isinstance(location, list):
        locations = location
    else:
        locations = [location] if isinstance(location, dict) else []
    for location_item in locations:
        address = location_item.get("address") if isinstance(location_item, dict) else {}
        country = address.get("addressCountry") if isinstance(address, dict) else ""
        normalized = clean_text(country).lower()
        if normalized in {"united states", "us", "usa", "united states of america"}:
            return True
    return False


def json_ld_job(html: str) -> dict[str, Any]:
    soup = BeautifulSoup(html, "html.parser")
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            payload = json.loads(script.string or "{}")
        except json.JSONDecodeError:
            continue
        items = payload if isinstance(payload, list) else [payload]
        for item in items:
            if isinstance(item, dict) and item.get("@type") == "JobPosting":
                return item
    return {}


def job_links_from_listing(html: str, page_url: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    links: list[str] = []
    for item in soup.select('li[class*="JobItem"] a[href*="/tech-jobs/"]'):
        href = str(item.get("href") or "")
        url = urljoin(page_url, href).split("#", 1)[0]
        if re.search(r"/\d+(?:[/?#]|$)", url) and url not in links:
            links.append(url)
    return links


def job_from_detail(html: str, job_url: str, term: str) -> VendorJob | None:
    item = json_ld_job(html)
    if not item:
        return None
    if not is_usa_job(item):
        return None
    title = clean_text(item.get("title"))
    raw_text = clean_text(item.get("description"))
    if not title or not raw_text:
        return None
    rank, reasons = score_title(title, raw_text)
    identifier = item.get("identifier")
    if isinstance(identifier, dict):
        identifier = identifier.get("value")
    return VendorJob(
        "Motion Recruitment",
        term,
        rank,
        reasons,
        title,
        clean_text(item.get("industry") or item.get("occupationalCategory")),
        location_from_ld(item),
        clean_text(item.get("employmentType")),
        salary_from_ld(item),
        clean_text(item.get("datePosted")),
        clean_text(identifier) or job_id_from_url(job_url),
        job_url,
        job_url,
        extract_contact_info(raw_text),
        raw_text[:900],
        raw_text,
    )


def jobs_from_listing_only(html: str, page_url: str, term: str) -> list[VendorJob]:
    jobs: list[VendorJob] = []
    soup = BeautifulSoup(html, "html.parser")
    for item in soup.select('li[class*="JobItem"]'):
        link = item.select_one('a[href*="/tech-jobs/"]')
        title_node = item.select_one('[class*="title"], h2')
        if not link or not title_node:
            continue
        job_url = urljoin(page_url, str(link.get("href") or "")).split("#", 1)[0]
        if not re.search(r"/\d+(?:[/?#]|$)", job_url):
            continue
        title = clean_text(title_node.get_text(" ", strip=True))
        raw_text = clean_text(item.get_text(" ", strip=True))
        rank, reasons = score_title(title, raw_text)
        jobs.append(VendorJob(
            "Motion Recruitment",
            term,
            rank,
            reasons,
            title[:180],
            "",
            "",
            "",
            "",
            "",
            job_id_from_url(job_url),
            job_url,
            job_url,
            extract_contact_info(raw_text),
            raw_text[:900],
            raw_text,
        ))
    return jobs


def scrape_jobs(terms: list[str], timeout: int, max_detail_pages: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (compatible; local job search scraper)",
        "Accept": "text/html,application/xhtml+xml,text/plain",
    })
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    for term in terms:
        page_url = f"{BASE}/tech-jobs/{quote_plus(route_for(term))}"
        try:
            html = fetch(session, page_url, timeout)
        except requests.RequestException as exc:
            print(f"Motion Recruitment {term}: failed to load {page_url}: {exc}", file=sys.stderr)
            continue
        links = job_links_from_listing(html, page_url)
        print(f"Motion Recruitment {term}: found {len(links)} detail links")
        fallback_jobs = jobs_from_listing_only(html, page_url, term)
        fallback_by_url = {job.job_url: job for job in fallback_jobs}
        for job_url in links[:max_detail_pages]:
            if job_url in seen:
                continue
            seen.add(job_url)
            job = None
            try:
                detail_html = fetch(session, job_url, timeout)
                job = job_from_detail(detail_html, job_url, term)
            except requests.RequestException as exc:
                print(f"  Detail fetch failed for {job_url}: {exc}", file=sys.stderr)
            if job:
                jobs.append(job)
            else:
                print(f"  Skipped unverified/non-USA job: {job_url}")
    print(f"Extracted {len(jobs)} unique Motion Recruitment jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout, args.max_detail_pages)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("motionrecruitment", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
