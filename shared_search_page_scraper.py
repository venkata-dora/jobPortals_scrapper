#!/usr/bin/env python3
"""Generic search-page scraper for staffing vendor job boards."""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.parse
from pathlib import Path
from typing import Any, Iterable

import requests
from bs4 import BeautifulSoup

from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, parse_posted_date, score_title, write_outputs


DEFAULT_SEARCH_TERMS = [
    "python",
    "full stack",
    "backend",
    "data engineer",
    "data engineering",
    "etl",
    "ai engineer",
    "machine learning",
    "llm",
    "rag",
]

JOB_LINK_RE = re.compile(r"/(?:job|jobs|career|careers|opening|openings|position|positions)(?:/|\\?|$|-)", re.I)
SKIP_URL_RE = re.compile(r"(?:linkedin|facebook|twitter|instagram|youtube|mailto:|tel:|javascript:)", re.I)
BAD_TITLE_RE = re.compile(r"^(?:apply|view|read more|learn more|search|submit|send resume|internal opportunities|view all jobs?)$", re.I)


def absolute_url(base_url: str, href: str) -> str:
    return urllib.parse.urljoin(base_url, href.strip())


def search_url(template: str, term: str) -> str:
    encoded = urllib.parse.quote_plus(term)
    if "{term}" in template:
        return template.format(term=encoded, raw_term=term)
    return template


def canonical_url(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    kept = [(key, value) for key, value in query if not key.lower().startswith(("utm_", "fbclid", "gclid"))]
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path.rstrip("/") or "/", urllib.parse.urlencode(kept), ""))


def fetch(session: requests.Session, url: str, timeout: int) -> str:
    response = session.get(url, timeout=timeout)
    response.raise_for_status()
    return response.text


def parse_json_ld_jobs(soup: BeautifulSoup, page_url: str, company_name: str, search_term: str) -> list[VendorJob]:
    jobs: list[VendorJob] = []
    for script in soup.find_all("script", type=lambda value: value and "ld+json" in value):
        try:
            payload = json.loads(script.string or "{}")
        except json.JSONDecodeError:
            continue
        items = payload if isinstance(payload, list) else [payload]
        for item in items:
            if isinstance(item, dict) and item.get("@graph"):
                graph = item.get("@graph")
                items.extend(graph if isinstance(graph, list) else [])
                continue
            if not isinstance(item, dict) or item.get("@type") != "JobPosting":
                continue
            title = clean_text(item.get("title"))
            description = clean_text(item.get("description"))
            if not title:
                continue
            rank, reasons = score_title(title, description)
            location = ""
            job_location = item.get("jobLocation")
            if isinstance(job_location, list):
                job_location = job_location[0] if job_location else {}
            if isinstance(job_location, dict):
                address = job_location.get("address")
                if isinstance(address, dict):
                    location = clean_text(", ".join(str(address.get(key) or "") for key in ("addressLocality", "addressRegion", "addressCountry")))
            url = clean_text(item.get("url")) or page_url
            posted = parse_posted_date(item.get("datePosted"))
            jobs.append(VendorJob(
                company_name,
                search_term,
                rank,
                reasons,
                title,
                clean_text(item.get("industry") or item.get("occupationalCategory")),
                location,
                clean_text(item.get("employmentType")),
                clean_text(item.get("baseSalary")),
                posted.date().isoformat() if posted else clean_text(item.get("datePosted")),
                clean_text(item.get("identifier")),
                url,
                url,
                extract_contact_info(description),
                description[:900],
                description,
            ))
    return jobs


def title_from_candidate(anchor: Any) -> str:
    text = clean_text(anchor.get_text(" ", strip=True))
    if text and not BAD_TITLE_RE.search(text):
        return text[:180]
    for parent in anchor.parents:
        if getattr(parent, "name", "") in {"article", "li", "tr", "div"}:
            for selector in ("h1", "h2", "h3", "h4", "[class*=title]", "[class*=job-title]"):
                found = parent.select_one(selector)
                if found:
                    candidate = clean_text(found.get_text(" ", strip=True))
                    if candidate and not BAD_TITLE_RE.search(candidate):
                        return candidate[:180]
    return ""


def extract_candidates(soup: BeautifulSoup, page_url: str, search_term: str) -> list[tuple[str, str, str]]:
    candidates: list[tuple[str, str, str]] = []
    normalized_term = search_term.lower()
    for anchor in soup.find_all("a", href=True):
        href = str(anchor.get("href") or "").strip()
        if not href or SKIP_URL_RE.search(href):
            continue
        url = absolute_url(page_url, href)
        text = title_from_candidate(anchor)
        parent_text = ""
        for parent in anchor.parents:
            if getattr(parent, "name", "") in {"article", "li", "tr"}:
                parent_text = clean_text(parent.get_text(" ", strip=True))
                break
        haystack = " ".join([url, text, parent_text]).lower()
        if JOB_LINK_RE.search(url) or normalized_term in haystack:
            title = text or clean_text(urllib.parse.unquote(Path(urllib.parse.urlsplit(url).path).name).replace("-", " ").replace("_", " "))
            if title:
                candidates.append((canonical_url(url), title, parent_text))
    return candidates


def normalize_candidate(company_name: str, search_term: str, url: str, title: str, summary_text: str, detail_text: str) -> VendorJob:
    raw_text = clean_text(" ".join([summary_text, detail_text]))
    rank, reasons = score_title(title, raw_text)
    posted = ""
    date_match = re.search(r"\b(?:20\d{2}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}/\d{1,2}/20\d{2})\b", raw_text)
    if date_match:
        parsed = parse_posted_date(date_match.group(0))
        posted = parsed.date().isoformat() if parsed else date_match.group(0)
    return VendorJob(
        company_name,
        search_term,
        rank,
        reasons,
        title,
        "",
        "",
        "",
        "",
        posted,
        Path(urllib.parse.urlsplit(url).path).name or url,
        url,
        url,
        extract_contact_info(raw_text),
        raw_text[:900],
        raw_text,
    )


def scrape_search_pages(
    company_name: str,
    search_urls: Iterable[str],
    terms: Iterable[str],
    posted_within_days: int,
    exclude_disallowed_work: bool,
    timeout: int,
    max_detail_pages: int,
    sleep: float,
) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)", "Accept": "text/html,application/xhtml+xml"})
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    details_fetched = 0
    loaded_pages: set[str] = set()
    for term in terms:
        for template in search_urls:
            url = search_url(template, term)
            cache_key = canonical_url(url)
            if "{term}" not in template and cache_key in loaded_pages:
                continue
            loaded_pages.add(cache_key)
            try:
                html = fetch(session, url, timeout)
            except requests.RequestException as exc:
                print(f"{company_name} {term}: failed to load {url}: {exc}", file=sys.stderr)
                continue
            soup = BeautifulSoup(html, "html.parser")
            for json_job in parse_json_ld_jobs(soup, url, company_name, term):
                key = canonical_url(json_job.job_url)
                if key in seen:
                    continue
                seen.add(key)
                jobs.append(json_job)
            candidates = extract_candidates(soup, url, term)
            print(f"{company_name} {term}: found {len(candidates)} candidate links at {url}")
            for candidate_url, title, summary_text in candidates:
                if candidate_url in seen:
                    continue
                seen.add(candidate_url)
                detail_text = ""
                if details_fetched < max_detail_pages:
                    try:
                        detail_html = fetch(session, candidate_url, timeout)
                        detail_soup = BeautifulSoup(detail_html, "html.parser")
                        for tag in detail_soup(["script", "style", "noscript"]):
                            tag.decompose()
                        detail_text = clean_text(detail_soup.get_text(" ", strip=True))
                        details_fetched += 1
                        time.sleep(sleep)
                    except requests.RequestException as exc:
                        print(f"  Detail fetch failed for {candidate_url}: {exc}", file=sys.stderr)
                jobs.append(normalize_candidate(company_name, term, candidate_url, title, summary_text, detail_text))
            time.sleep(sleep)
    print(f"Extracted {len(jobs)} unique {company_name} jobs before filtering")
    return filter_and_sort_jobs(jobs, posted_within_days, exclude_disallowed_work)


def build_parser(company_name: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=f"Scrape {company_name} jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def run_cli(company_name: str, company_slug: str, search_urls: list[str]) -> int:
    args = build_parser(company_name).parse_args()
    out_dir = args.out_dir or Path(__file__).resolve().parent / f"{company_slug}_applying_script" / "output"
    jobs = scrape_search_pages(
        company_name,
        search_urls,
        args.terms or DEFAULT_SEARCH_TERMS,
        args.posted_within_days,
        not args.keep_w2_f2f_onsite_interview,
        args.timeout,
        args.max_detail_pages,
        args.sleep,
    )
    write_outputs(company_slug, jobs, out_dir, args.posted_within_days, args.no_excel)
    return 0
