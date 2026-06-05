#!/usr/bin/env python3
"""Open filtered Artech job pages in browser tabs."""

from __future__ import annotations

import argparse
import json
import sys
import time
import webbrowser
from pathlib import Path
from typing import Any


DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def latest_jobs_file(output_dir: Path) -> Path:
    files = sorted(output_dir.glob("artech_jobs_*.json"), key=lambda path: path.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No artech_jobs_*.json files found in {output_dir}")
    return files[0]


def load_jobs(path: Path) -> list[dict[str, Any]]:
    with path.open(encoding="utf-8") as handle:
        jobs = json.load(handle)
    if not isinstance(jobs, list):
        raise ValueError(f"Expected a list of jobs in {path}")
    return jobs


def main() -> int:
    parser = argparse.ArgumentParser(description="Open Artech job pages.")
    parser.add_argument("--jobs-file", type=Path)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--start-at", type=int, default=1)
    parser.add_argument("--delay", type=float, default=0.5)
    args = parser.parse_args()
    jobs_file = args.jobs_file or latest_jobs_file(args.out_dir)
    jobs = load_jobs(jobs_file)
    selected = jobs[max(args.start_at - 1, 0):]
    if args.limit > 0:
        selected = selected[: args.limit]
    print(f"Jobs file: {jobs_file}")
    print(f"Opening {len(selected)} Artech tabs...")
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


if __name__ == "__main__":
    raise SystemExit(main())
