#!/usr/bin/env python3
"""Standalone Artech scraper using the public JobDiva candidate portal API."""

from __future__ import annotations

import argparse
import sys
import time
import urllib.parse
from pathlib import Path
from typing import Any, Iterable

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_vendor_filters import VendorJob, clean_text, extract_contact_info, filter_and_sort_jobs, parse_posted_date, score_title, write_outputs


PORTAL_A = "kvjdnwtsxgckrpsoozx5qc0oueybw1005779v7x6soig8eyqqmzaubfdl9tcx21s"
PORTAL_URL = f"https://www1.jobdiva.com/portal/?a={PORTAL_A}&compid=0#/"
BASE_URL = "https://ws.jobdiva.com/candPortal/rest/"
DEFAULT_SEARCH_TERMS = ["python", "full stack", "backend", "data engineer", "data engineering", "etl", "ai engineer", "machine learning", "llm", "rag"]


def get_token(session: requests.Session, timeout: int) -> dict[str, Any]:
    headers = {
        "Authorization": "Basic YXhlbG9uOmF4ZWxvbg==",
        "portalID": "1",
        "a": PORTAL_A,
        "compid": "0",
    }
    response = session.get(BASE_URL + "auth/a", headers=headers, timeout=timeout)
    response.raise_for_status()
    return response.json()


def auth_headers(token: dict[str, Any]) -> dict[str, str]:
    return {
        "portalID": str(token.get("portalID") or ""),
        "token": str(token.get("token") or ""),
        "a": str(token.get("a") or PORTAL_A),
        "User-Agent": "Mozilla/5.0",
        "Referer": PORTAL_URL,
    }


def fetch_term(session: requests.Session, headers: dict[str, str], term: str, timeout: int, page_size: int, max_pages: int, sleep: float) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for page in range(max_pages):
        start = page * page_size + 1
        end = start + page_size - 1
        payload = urllib.parse.urlencode({
            "portalID": 1,
            "from": start,
            "to": end,
            "keywords": term,
            "country": "US",
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
        response = session.post(
            BASE_URL + "job/searchjobsportal",
            data=payload,
            headers={**headers, "content-type": "application/x-www-form-urlencoded"},
            timeout=timeout,
        )
        response.raise_for_status()
        data = response.json()
        batch = data.get("data") or []
        rows.extend(batch)
        total = int(data.get("total") or len(rows))
        print(f"Artech {term}: extracted {len(rows)}/{total}")
        if not batch or len(rows) >= total:
            break
        time.sleep(sleep)
    return rows


def fetch_detail(session: requests.Session, headers: dict[str, str], job_id: str, timeout: int) -> dict[str, Any] | None:
    response = session.get(BASE_URL + f"job/getdetailbyjobid/{job_id}?compid=0", headers=headers, timeout=timeout)
    if response.status_code == 404:
        return None
    response.raise_for_status()
    data = response.json()
    return data.get("job") if isinstance(data.get("job"), dict) else None


def location_from_job(job: dict[str, Any]) -> str:
    if job.get("location"):
        return clean_text(job.get("location"))
    main = job.get("mainLocation") if isinstance(job.get("mainLocation"), dict) else {}
    parts = [main.get("city"), main.get("state"), main.get("country")]
    return clean_text(", ".join(str(part).strip() for part in parts if part))


def category_from_job(job: dict[str, Any]) -> str:
    udfs = job.get("jobUDFs") if isinstance(job.get("jobUDFs"), list) else []
    categories = [clean_text(item.get("value")) for item in udfs if isinstance(item, dict) and clean_text(item.get("name")).lower() == "category"]
    return ", ".join(category for category in categories if category)


def posted_date_from_job(job: dict[str, Any]) -> str:
    parsed = parse_posted_date(job.get("postDate"))
    if parsed:
        return parsed.isoformat()
    return clean_text(job.get("postDateStr"))


def normalize(row: dict[str, Any], search_term: str, detail: dict[str, Any] | None) -> VendorJob:
    job = detail or row
    title = clean_text(job.get("title"))
    raw_text = clean_text(job.get("jobDescription"))
    rank, reasons = score_title(title, raw_text)
    job_id = str(job.get("id") or row.get("id") or "")
    ref_no = clean_text(job.get("refNo") or row.get("refNo"))
    salary = clean_text(" / ".join(part for part in [job.get("payRate") or row.get("payRate"), job.get("payFrequency") or row.get("payFrequency")] if part))
    job_url = f"{PORTAL_URL}jobs/{job_id}" if job_id else PORTAL_URL
    contact_info = extract_contact_info(" ".join([
        raw_text,
        clean_text(job.get("primaryRecruiterEmail")),
        clean_text(job.get("primaryRecruiterPhone")),
    ]))
    return VendorJob(
        "Artech",
        search_term,
        rank,
        reasons,
        title,
        category_from_job(job),
        location_from_job(job),
        clean_text(job.get("positionType") or row.get("positionType")),
        salary,
        posted_date_from_job(job),
        ref_no or job_id,
        job_url,
        job_url,
        contact_info,
        raw_text[:900],
        raw_text,
    )


def scrape(terms: Iterable[str], posted_within_days: int, exclude_disallowed_work: bool, timeout: int, page_size: int, max_pages: int, max_detail_pages: int, sleep: float) -> list[VendorJob]:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0", "Accept": "application/json", "Referer": PORTAL_URL})
    headers = auth_headers(get_token(session, timeout))
    seen: set[str] = set()
    jobs: list[VendorJob] = []
    details_fetched = 0
    for term in terms:
        for row in fetch_term(session, headers, term, timeout, page_size, max_pages, sleep):
            job_id = str(row.get("id") or "")
            key = job_id or clean_text(row.get("refNo"))
            if not key or key in seen:
                continue
            seen.add(key)
            detail = None
            if details_fetched < max_detail_pages and job_id:
                detail = fetch_detail(session, headers, job_id, timeout)
                details_fetched += 1
                time.sleep(sleep)
            jobs.append(normalize(row, term, detail))
    print(f"Extracted {len(jobs)} unique Artech jobs before filtering")
    return filter_and_sort_jobs(jobs, posted_within_days, exclude_disallowed_work)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape Artech jobs into CSV/JSON/Excel.")
    parser.add_argument("--term", action="append", dest="terms")
    parser.add_argument("--posted-within-days", type=int, default=4)
    parser.add_argument("--keep-w2-f2f-onsite-interview", action="store_true")
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--page-size", type=int, default=20)
    parser.add_argument("--max-pages", type=int, default=3)
    parser.add_argument("--max-detail-pages", type=int, default=60)
    parser.add_argument("--sleep", type=float, default=0.2)
    parser.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent / "output")
    parser.add_argument("--no-excel", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    jobs = scrape(
        args.terms or DEFAULT_SEARCH_TERMS,
        args.posted_within_days,
        not args.keep_w2_f2f_onsite_interview,
        args.timeout,
        args.page_size,
        args.max_pages,
        args.max_detail_pages,
        args.sleep,
    )
    write_outputs("artech", jobs, args.out_dir, args.posted_within_days, args.no_excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
