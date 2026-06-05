#!/usr/bin/env python3
"""Shared open-tabs helper for standalone vendor folders."""

from __future__ import annotations

import argparse
import json
import sys
import time
import webbrowser
from pathlib import Path
from typing import Any


def latest_jobs_file(output_dir: Path, company_slug: str) -> Path:
    files = sorted(output_dir.glob(f"{company_slug}_jobs_*.json"), key=lambda path: path.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No {company_slug}_jobs_*.json files found in {output_dir}")
    return files[0]


def load_jobs(path: Path) -> list[dict[str, Any]]:
    with path.open(encoding="utf-8") as handle:
        jobs = json.load(handle)
    if not isinstance(jobs, list):
        raise ValueError(f"Expected a list of jobs in {path}")
    return jobs


def run_open_jobs(company_slug: str, company_label: str, default_output_dir: Path) -> int:
    parser = argparse.ArgumentParser(description=f"Open {company_label} job pages.")
    parser.add_argument("--jobs-file", type=Path)
    parser.add_argument("--out-dir", type=Path, default=default_output_dir)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--start-at", type=int, default=1)
    parser.add_argument("--delay", type=float, default=0.5)
    args = parser.parse_args()
    jobs_file = args.jobs_file or latest_jobs_file(args.out_dir, company_slug)
    jobs = load_jobs(jobs_file)
    selected = jobs[max(args.start_at - 1, 0):]
    if args.limit > 0:
        selected = selected[: args.limit]
    print(f"Jobs file: {jobs_file}")
    print(f"Opening {len(selected)} {company_label} tabs...")
    for index, job in enumerate(selected, start=max(args.start_at, 1)):
        url = str(job.get("job_url") or job.get("apply_url") or "").strip()
        title = str(job.get("title") or "").strip()
        if not url:
            print(f"[{index}] skipped, missing url: {title}", file=sys.stderr)
            continue
        print(f"[{index}/{len(jobs)}] {title}")
        webbrowser.open_new_tab(url)
        time.sleep(args.delay)
    return 0
