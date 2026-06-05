#!/usr/bin/env python3
"""Standalone Diverse Lynx scraper."""

from __future__ import annotations

import argparse
import sys
import urllib.parse
from pathlib import Path
from typing import Any

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_search_page_scraper import DEFAULT_SEARCH_TERMS
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, parse_posted_date, score_title, write_outputs


PORTAL_KEY = "9xjdnw687b7a7nvvdyut936kpjlgy0023blrozaecads0pdnwppcswnaaku8ji2g"
PORTAL_URL = f"https://www2.jobdiva.com/portal/?a={PORTAL_KEY}&compid=0"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scrape Diverse Lynx jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--no-excel", action="store_true")
    return parser


def authenticate(session: requests.Session, timeout: int) -> dict[str, Any]:
    response = session.get(
        "https://ws.jobdiva.com/candPortal/rest/auth/a",
        headers={
            "Authorization": "Basic YXhlbG9uOmF4ZWxvbg==",
            "portalID": "1",
            "a": PORTAL_KEY,
            "compid": "0",
        },
        timeout=timeout,
    )
    response.raise_for_status()
    return response.json()


def scrape_jobs(terms: list[str], timeout: int) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; local job search scraper)", "Referer": PORTAL_URL})
    auth = authenticate(session, timeout)
    headers = {
        "portalID": str(auth["portalID"]),
        "token": auth["token"],
        "a": auth["a"],
        "content-type": "application/x-www-form-urlencoded",
    }
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    for term in terms:
        form = urllib.parse.urlencode({
            "portalID": auth["portalID"],
            "from": 1,
            "to": 20,
            "keywords": term,
            "country": "",
            "states": "",
            "city": "",
            "zipcode": "",
            "miles": "",
            "jobCategories": "",
            "jobTypes": "",
            "jobDivisions": "",
            "onsiteFlex": "",
            "qualifications": "",
            "unit": "",
        })
        response = session.post("https://ws.jobdiva.com/candPortal/rest/job/searchjobsportal", headers=headers, data=form, timeout=timeout)
        response.raise_for_status()
        payload = response.json()
        results = payload.get("data") or []
        print(f"Diverse Lynx {term}: found {len(results)} JobDiva jobs")
        for item in results:
            job_id = str(item.get("id") or item.get("refNo") or "")
            if not job_id or job_id in seen:
                continue
            seen.add(job_id)
            title = clean_text(item.get("title"))
            raw_text = clean_text(" ".join([
                clean_text(item.get("jobDescription")),
                clean_text(item.get("positionType")),
                clean_text(item.get("workingRemote")),
            ]))
            rank, reasons = score_title(title, raw_text)
            posted = parse_posted_date(item.get("postDate"))
            salary = clean_text(" ".join(str(item.get(key) or "") for key in ("payRate", "payFrequency")))
            job_url = f"{PORTAL_URL}#/jobs/{job_id}"
            jobs.append(VendorJob(
                "Diverse Lynx",
                term,
                rank,
                reasons,
                title,
                "",
                clean_text(item.get("location")),
                clean_text(item.get("positionType")),
                salary,
                posted.date().isoformat() if posted else clean_text(item.get("postDateStr")),
                clean_text(item.get("refNo")) or job_id,
                job_url,
                job_url,
                extract_contact_info(raw_text),
                raw_text[:900],
                raw_text,
            ))
    print(f"Extracted {len(jobs)} unique Diverse Lynx jobs before filtering")
    return jobs


def main() -> int:
    args = build_parser().parse_args()
    jobs = scrape_jobs(args.terms or DEFAULT_SEARCH_TERMS, args.timeout)
    filtered = filter_and_sort_jobs(jobs, args.posted_within_days, not args.keep_w2_f2f_onsite_interview)
    out_dir = args.out_dir or Path(__file__).resolve().parent / "output"
    write_outputs("diverselynx", filtered, out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
