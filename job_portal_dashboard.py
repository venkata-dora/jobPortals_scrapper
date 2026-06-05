#!/usr/bin/env python3
"""Local control UI for the separate job scraper folders."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import signal
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser
from dataclasses import dataclass, asdict
from datetime import date, datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
CONFIG_PATH = Path(os.environ.get("JOB_DASHBOARD_CONFIG_PATH", ROOT / "job_portal_dashboard_config.json")).expanduser()
APPLIED_PATH = Path(os.environ.get("JOB_DASHBOARD_APPLIED_PATH", ROOT / "job_portal_dashboard_applied.json")).expanduser()
CLICKS_PATH = Path(os.environ.get("JOB_DASHBOARD_CLICKS_PATH", ROOT / "job_portal_dashboard_clicks.json")).expanduser()
SEEN_PATH = Path(os.environ.get("JOB_DASHBOARD_SEEN_PATH", ROOT / "job_portal_dashboard_seen.json")).expanduser()
APPLIED_TTL_SECONDS = 2 * 60 * 60
PYTHON = sys.executable or "python3"


DEFAULT_KEYWORDS = [
    "python developer",
    "python engineer",
    "full stack developer",
    "full stack engineer",
    "backend developer",
    "software engineer python",
    "data engineer",
    "data engineering",
    "etl developer",
    "ai engineer",
    "machine learning engineer",
    "generative ai",
    "llm",
    "rag",
]


@dataclass(frozen=True)
class Vendor:
    slug: str
    label: str
    folder: str
    scraper: str
    opener: str
    prefix: str
    terms_mode: str = "none"
    page_arg: str = ""
    page_default: int = 0
    color: str = ""


VENDORS: list[Vendor] = [
    Vendor("teksystems", "TEKsystems", "teksystems_applying_script", "teksystems_scraper.py", "teksystems_open_jobs.py", "teksystems", "file", color="red"),
    Vendor("apexsystems", "Apex Systems", "apexsystems_applying_script", "apex_scraper.py", "apex_open_jobs.py", "apex", "file", "rows-per-search", 100),
    Vendor("judgegroup", "Judge Group", "judgegroup_applying_script", "judgegroup_scraper.py", "judgegroup_open_jobs.py", "judgegroup", "file", "max-pages", 3),
    Vendor("beaconhill", "Beacon Hill", "beaconhill_applying_script", "beaconhill_scraper.py", "beaconhill_open_jobs.py", "beaconhill", "file", "max-pages", 3),
    Vendor("akkodis", "Akkodis", "akkodis_applying_script", "akkodis_scraper.py", "akkodis_open_jobs.py", "akkodis", "file", "max-detail-pages", 30),
    Vendor("randstad", "Randstad", "randstad_applying_script", "randstad_scraper.py", "randstad_open_jobs.py", "randstad", "file"),
    Vendor("eliassen", "Eliassen", "eliassen_applying_script", "eliassen_scraper.py", "eliassen_open_jobs.py", "eliassen"),
    Vendor("experis", "Experis", "experis_applying_script", "experis_scraper.py", "experis_open_jobs.py", "experis", "append", "max-pages", 3),
    Vendor("brooksource", "Brooksource", "brooksource_applying_script", "brooksource_scraper.py", "brooksource_open_jobs.py", "brooksource"),
    Vendor("kellymitchell", "KellyMitchell", "kellymitchell_applying_script", "kellymitchell_scraper.py", "kellymitchell_open_jobs.py", "kellymitchell", "append", "jobs-per-page", 50),
    Vendor("mitchellmartin", "Mitchell Martin", "mitchellmartin_applying_script", "mitchellmartin_scraper.py", "mitchellmartin_open_jobs.py", "mitchellmartin", "none", "max-jobs", 40, color="red"),
    Vendor("cbts", "CBTS", "cbts_applying_script", "cbts_scraper.py", "cbts_open_jobs.py", "cbts", "file"),
    Vendor("roberthalf", "Robert Half", "roberthalf_applying_script", "roberthalf_scraper.py", "roberthalf_open_jobs.py", "roberthalf", "file", "max-pages", 3),
    Vendor("kforce", "Kforce", "kforce_applying_script", "kforce_scraper.py", "kforce_open_jobs.py", "kforce", "file"),
    Vendor("insightglobal", "Insight Global", "insightglobal_applying_script", "insightglobal_scraper.py", "insightglobal_open_jobs.py", "insightglobal", "append", color="red"),
    Vendor("artech", "Artech", "artech_applying_script", "artech_scraper.py", "artech_open_jobs.py", "artech", "append", "max-pages", 3),
    Vendor("ccsglobaltech", "CCS Global Tech", "ccsglobaltech_applying_script", "ccsglobaltech_scraper.py", "ccsglobaltech_open_jobs.py", "ccsglobaltech", "append"),
    Vendor("diverselynx", "Diverse Lynx", "diverselynx_applying_script", "diverselynx_scraper.py", "diverselynx_open_jobs.py", "diverselynx", "append"),
    Vendor("ettaingroup", "Ettain Group", "ettaingroup_applying_script", "ettaingroup_scraper.py", "ettaingroup_open_jobs.py", "ettaingroup", "append"),
    Vendor("harveynash", "Harvey Nash", "harveynash_applying_script", "harveynash_scraper.py", "harveynash_open_jobs.py", "harveynash", "append"),
    Vendor("hays", "Hays", "hays_applying_script", "hays_scraper.py", "hays_open_jobs.py", "hays", "append"),
    Vendor("inspyr", "INSPYR Solutions", "inspyr_applying_script", "inspyr_scraper.py", "inspyr_open_jobs.py", "inspyr", "append"),
    Vendor("matlensilver", "Matlen Silver", "matlensilver_applying_script", "matlensilver_scraper.py", "matlensilver_open_jobs.py", "matlensilver", "append"),
    Vendor("motionrecruitment", "Motion Recruitment", "motionrecruitment_applying_script", "motionrecruitment_scraper.py", "motionrecruitment_open_jobs.py", "motionrecruitment", "append"),
    Vendor("nttdata", "NTT DATA", "nttdata_applying_script", "nttdata_scraper.py", "nttdata_open_jobs.py", "nttdata", "append"),
    Vendor("oliverjames", "Oliver James", "oliverjames_applying_script", "oliverjames_scraper.py", "oliverjames_open_jobs.py", "oliverjames", "append"),
    Vendor("optomi", "Optomi", "optomi_applying_script", "optomi_scraper.py", "optomi_open_jobs.py", "optomi", "append"),
    Vendor("opensystemstechnologies", "Open Systems Technologies", "opensystemstechnologies_applying_script", "opensystemstechnologies_scraper.py", "opensystemstechnologies_open_jobs.py", "opensystemstechnologies", "append"),
    Vendor("ptrglobal", "PTR Global", "ptrglobal_applying_script", "ptrglobal_scraper.py", "ptrglobal_open_jobs.py", "ptrglobal", "append"),
    Vendor("pyramidconsulting", "Pyramid Consulting", "pyramidconsulting_applying_script", "pyramidconsulting_scraper.py", "pyramidconsulting_open_jobs.py", "pyramidconsulting", "append"),
    Vendor("strategicstaffing", "Strategic Staffing Solutions", "strategicstaffing_applying_script", "strategicstaffing_scraper.py", "strategicstaffing_open_jobs.py", "strategicstaffing", "append"),
    Vendor("vaco", "Vaco", "vaco_applying_script", "vaco_scraper.py", "vaco_open_jobs.py", "vaco", "append"),
    Vendor("venturigroup", "Venturi Group", "venturigroup_applying_script", "venturigroup_scraper.py", "venturigroup_open_jobs.py", "venturigroup", "append"),
]

VENDOR_BY_SLUG = {vendor.slug: vendor for vendor in VENDORS}

ROTATION_PAIRS = [
    ["teksystems", "apexsystems"],
    ["judgegroup", "beaconhill"],
    ["akkodis", "randstad"],
    ["eliassen", "experis"],
    ["brooksource", "kellymitchell"],
    ["mitchellmartin", "cbts"],
    ["roberthalf", "kforce"],
    ["insightglobal", "artech"],
    ["ccsglobaltech", "diverselynx"],
    ["ettaingroup", "harveynash"],
    ["hays", "inspyr"],
    ["matlensilver", "motionrecruitment"],
    ["nttdata", "oliverjames"],
    ["optomi", "opensystemstechnologies"],
    ["ptrglobal", "pyramidconsulting"],
    ["strategicstaffing", "vaco"],
    ["venturigroup"],
]

RUN_LOCK = threading.Lock()
RUNS: dict[str, dict[str, Any]] = {}
PROCESS_LOCK = threading.Lock()
ACTIVE_PROCESSES: dict[int, dict[str, Any]] = {}
SEEN_LOCK = threading.Lock()
STOP_EVENT = threading.Event()


def default_config() -> dict[str, Any]:
    return {
        "posted_within_days": 4,
        "open_limit": 8,
        "start_at": 1,
        "delay": 0.5,
        "keep_open_minutes": 60,
        "keywords": DEFAULT_KEYWORDS,
        "vendor_overrides": {},
        "judge_apply": {
            "resume_path": "resume.docx",
            "first_name": "",
            "last_name": "",
            "email": "",
            "phone": "",
            "country": "United States",
            "street_address": "",
            "city": "",
            "state": "",
            "zip_code": "",
            "keep_open_seconds": 45,
        },
    }


def load_config() -> dict[str, Any]:
    config = default_config()
    if CONFIG_PATH.exists():
        try:
            saved = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(saved, dict):
                config.update(saved)
        except json.JSONDecodeError:
            pass
    config["keywords"] = normalize_keywords(config.get("keywords"))
    return config


def save_config(config: dict[str, Any]) -> None:
    clean = default_config()
    clean.update(config)
    clean["keywords"] = normalize_keywords(clean.get("keywords"))
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(clean, indent=2), encoding="utf-8")


def normalize_keywords(value: Any) -> list[str]:
    if isinstance(value, str):
        parts = value.splitlines()
    elif isinstance(value, list):
        parts = [str(item) for item in value]
    else:
        parts = []
    seen: set[str] = set()
    keywords: list[str] = []
    for part in parts:
        keyword = " ".join(part.strip().split())
        key = keyword.lower()
        if keyword and key not in seen:
            seen.add(key)
            keywords.append(keyword)
    return keywords or DEFAULT_KEYWORDS[:]


def register_process(proc: subprocess.Popen[Any], kind: str, label: str) -> None:
    with PROCESS_LOCK:
        ACTIVE_PROCESSES[proc.pid] = {
            "pid": proc.pid,
            "kind": kind,
            "label": label,
            "started_at": datetime.now().isoformat(timespec="seconds"),
        }


def unregister_process(pid: int) -> None:
    with PROCESS_LOCK:
        ACTIVE_PROCESSES.pop(pid, None)


def launch_process(cmd: list[str], kind: str, label: str, **kwargs: Any) -> subprocess.Popen[Any]:
    proc = subprocess.Popen(cmd, start_new_session=True, **kwargs)
    register_process(proc, kind, label)
    return proc


def active_processes() -> list[dict[str, Any]]:
    with PROCESS_LOCK:
        return list(ACTIVE_PROCESSES.values())


def stop_active_processes() -> dict[str, Any]:
    STOP_EVENT.set()
    processes = active_processes()
    stopped: list[dict[str, Any]] = []
    for info in processes:
        pid = int(info["pid"])
        try:
            os.killpg(pid, signal.SIGTERM)
            stopped.append(info)
        except ProcessLookupError:
            unregister_process(pid)
        except OSError as exc:
            info["error"] = str(exc)
            stopped.append(info)
    return {"stopped": stopped, "count": len(stopped)}


def nth_workday_index(today: date | None = None) -> int:
    today = today or date.today()
    anchor = date(2026, 5, 11)
    step = 1 if today >= anchor else -1
    current = anchor
    count = 0
    while current != today:
        current += timedelta(days=step)
        if current.weekday() < 5:
            count += step
    return count


def active_pair_for(day: date | None = None) -> list[str]:
    index = nth_workday_index(day) % len(ROTATION_PAIRS)
    return ROTATION_PAIRS[index]


def rotation_preview(days: int = 10) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    current = date.today()
    while len(rows) < days:
        if current.weekday() < 5:
            pair = active_pair_for(current)
            rows.append({
                "date": current.isoformat(),
                "weekday": current.strftime("%A"),
                "vendors": [VENDOR_BY_SLUG[slug].label for slug in pair],
                "slugs": pair,
            })
        current += timedelta(days=1)
    return rows


def latest_jobs_file(vendor: Vendor) -> Path | None:
    out_dir = ROOT / vendor.folder / "output"
    files = sorted(out_dir.glob(f"{vendor.prefix}_jobs_*.json"), key=lambda path: path.stat().st_mtime, reverse=True)
    return files[0] if files else None


def load_latest_jobs(vendor: Vendor) -> list[dict[str, Any]]:
    latest = latest_jobs_file(vendor)
    if not latest:
        raise FileNotFoundError(f"No latest jobs file found for {vendor.label}")
    data = load_jobs_file(latest)
    return data


def load_jobs_file(path: Path) -> list[dict[str, Any]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError(f"Expected a list of jobs in {path}")
    return [row for row in data if isinstance(row, dict)]


def job_url(job: dict[str, Any]) -> str:
    return str(job.get("job_url") or job.get("apply_url") or job.get("url") or "").strip()


def job_key(vendor_slug: str, job: dict[str, Any]) -> str:
    identity = job_url(job) or "|".join(
        str(job.get(field) or "").strip().lower()
        for field in ("title", "location", "posted_date", "job_id")
    )
    return hashlib.sha256(f"{vendor_slug}|{identity}".encode("utf-8")).hexdigest()


def job_keys_for_rows(vendor_slug: str, jobs: list[dict[str, Any]]) -> list[str]:
    seen: set[str] = set()
    keys: list[str] = []
    for job in jobs:
        key = job_key(vendor_slug, job)
        if key not in seen:
            seen.add(key)
            keys.append(key)
    return keys


def load_job_keys_from_file(vendor: Vendor, path: Path | None) -> list[str]:
    if not path:
        return []
    try:
        return job_keys_for_rows(vendor.slug, load_jobs_file(path))
    except (OSError, json.JSONDecodeError, ValueError):
        return []


def load_seen_state(path: Path = SEEN_PATH) -> dict[str, Any]:
    if not path.exists():
        return {"vendors": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"vendors": {}}
    if not isinstance(data, dict):
        return {"vendors": {}}
    vendors = data.get("vendors")
    if not isinstance(vendors, dict):
        data["vendors"] = {}
    return data


def save_seen_state(state: dict[str, Any], path: Path = SEEN_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, indent=2, sort_keys=True), encoding="utf-8")


def baseline_keys_for_vendor(vendor: Vendor, fallback_file: Path | None, state: dict[str, Any] | None = None) -> set[str]:
    state = state or load_seen_state()
    vendor_state = state.get("vendors", {}).get(vendor.slug, {})
    saved_keys = vendor_state.get("last_successful_keys") if isinstance(vendor_state, dict) else None
    if isinstance(saved_keys, list):
        return {str(key) for key in saved_keys}
    return set(load_job_keys_from_file(vendor, fallback_file))


def latest_new_keys_for_vendor(slug: str, state: dict[str, Any] | None = None) -> set[str]:
    state = state or load_seen_state()
    vendor_state = state.get("vendors", {}).get(slug, {})
    keys = vendor_state.get("last_new_keys") if isinstance(vendor_state, dict) else None
    if not isinstance(keys, list):
        return set()
    return {str(key) for key in keys}


def update_seen_success(
    vendor: Vendor,
    latest_file: Path | None,
    jobs: list[dict[str, Any]],
    previous_keys: set[str],
    run_id: str,
    ready_at: str,
    path: Path = SEEN_PATH,
) -> dict[str, Any]:
    latest_keys = job_keys_for_rows(vendor.slug, jobs)
    new_keys = [key for key in latest_keys if key not in previous_keys]
    with SEEN_LOCK:
        state = load_seen_state(path)
        vendors = state.setdefault("vendors", {})
        vendors[vendor.slug] = {
            "last_successful_keys": latest_keys,
            "last_new_keys": new_keys,
            "last_new_count": len(new_keys),
            "last_run_id": run_id,
            "ready_at": ready_at,
            "latest_file": str(latest_file.relative_to(ROOT)) if latest_file and latest_file.exists() else "",
            "latest_count": len(jobs),
        }
        save_seen_state(state, path)
    return {
        "new_keys": new_keys,
        "new_count": len(new_keys),
        "ready_run_id": run_id,
        "ready_at": ready_at,
    }


def load_applied_marks() -> dict[str, dict[str, Any]]:
    now = time.time()
    if not APPLIED_PATH.exists():
        return {}
    try:
        data = json.loads(APPLIED_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    marks: dict[str, dict[str, Any]] = {}
    changed = False
    for key, value in data.items():
        if not isinstance(value, dict):
            changed = True
            continue
        marked_at = float(value.get("marked_at") or 0)
        if marked_at and now - marked_at < APPLIED_TTL_SECONDS:
            marks[str(key)] = value
        else:
            changed = True
    if changed:
        save_applied_marks(marks)
    return marks


def save_applied_marks(marks: dict[str, dict[str, Any]]) -> None:
    APPLIED_PATH.parent.mkdir(parents=True, exist_ok=True)
    APPLIED_PATH.write_text(json.dumps(marks, indent=2, sort_keys=True), encoding="utf-8")


def load_click_counts() -> dict[str, dict[str, Any]]:
    today = date.today().isoformat()
    if not CLICKS_PATH.exists():
        return {}
    try:
        data = json.loads(CLICKS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {k: v for k, v in data.items() if isinstance(v, dict) and v.get("date") == today}


def save_click_counts(counts: dict[str, dict[str, Any]]) -> None:
    CLICKS_PATH.parent.mkdir(parents=True, exist_ok=True)
    CLICKS_PATH.write_text(json.dumps(counts, indent=2, sort_keys=True), encoding="utf-8")


def increment_click(key: str) -> int:
    today = date.today().isoformat()
    counts = load_click_counts()
    entry = counts.get(key)
    if entry and entry.get("date") == today:
        entry["count"] = int(entry.get("count") or 0) + 1
    else:
        entry = {"date": today, "count": 1}
    counts[key] = entry
    save_click_counts(counts)
    return int(entry["count"])


def format_job_for_ui(vendor: Vendor, job: dict[str, Any], index: int, marks: dict[str, dict[str, Any]], clicks: dict[str, dict[str, Any]]) -> dict[str, Any]:
    key = job_key(vendor.slug, job)
    mark = marks.get(key) or {}
    marked_at = float(mark.get("marked_at") or 0)
    expires_at = marked_at + APPLIED_TTL_SECONDS if marked_at else 0
    click_entry = clicks.get(key) or {}
    return {
        "key": key,
        "index": index,
        "vendor": vendor.label,
        "slug": vendor.slug,
        "title": str(job.get("title") or "Untitled job"),
        "location": str(job.get("location") or ""),
        "posted_date": str(job.get("posted_date") or job.get("posted_day") or ""),
        "employment_type": str(job.get("employment_type") or ""),
        "url": job_url(job),
        "applied": bool(mark),
        "marked_at": datetime.fromtimestamp(marked_at).isoformat(timespec="seconds") if marked_at else "",
        "expires_at": datetime.fromtimestamp(expires_at).isoformat(timespec="seconds") if expires_at else "",
        "seconds_left": max(int(expires_at - time.time()), 0) if expires_at else 0,
        "open_count": int(click_entry.get("count") or 0),
    }


def latest_jobs_for_ui(slug: str, config: dict[str, Any], payload: dict[str, Any]) -> dict[str, Any]:
    vendor = VENDOR_BY_SLUG[slug]
    jobs = load_latest_jobs(vendor)
    total_jobs = len(jobs)
    state = load_seen_state()
    new_keys = latest_new_keys_for_vendor(slug, state)
    vendor_seen = state.get("vendors", {}).get(slug, {})
    if payload.get("new_only"):
        jobs = [job for job in jobs if job_key(slug, job) in new_keys]
    start_at = max(int(payload.get("start_at") or config.get("start_at") or 1), 1)
    limit = int(payload.get("limit") or config.get("open_limit") or 0)
    selected = jobs[start_at - 1:] if start_at > 1 else jobs
    if limit > 0:
        selected = selected[:limit]
    marks = load_applied_marks()
    clicks = load_click_counts()
    rows = []
    for offset, job in enumerate(selected):
        row = format_job_for_ui(vendor, job, start_at + offset, marks, clicks)
        row["is_new"] = row["key"] in new_keys
        row["ready_run_id"] = str(vendor_seen.get("last_run_id") or "")
        rows.append(row)
    return {
        "vendor": vendor.label,
        "jobs": rows,
        "total": total_jobs,
        "shown_total": len(jobs),
        "new_total": len(new_keys),
        "new_only": bool(payload.get("new_only")),
        "start_at": start_at,
        "limit": limit,
    }


def latest_step_for_vendor(slug: str) -> dict[str, Any] | None:
    for run in reversed(list(RUNS.values())):
        for step in reversed(run.get("steps") or []):
            if step.get("slug") == slug:
                return step
    return None


def vendor_status(vendor: Vendor) -> dict[str, Any]:
    latest = latest_jobs_file(vendor)
    count = 0
    modified = ""
    if latest:
        try:
            data = json.loads(latest.read_text(encoding="utf-8"))
            count = len(data) if isinstance(data, list) else 0
            modified = datetime.fromtimestamp(latest.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        except (OSError, json.JSONDecodeError):
            count = 0
    seen_state = load_seen_state().get("vendors", {}).get(vendor.slug, {})
    step = latest_step_for_vendor(vendor.slug) or {}
    run_status = str(step.get("status") or "")
    new_count = int(step.get("new_count") if step.get("new_count") is not None else seen_state.get("last_new_count") or len(seen_state.get("last_new_keys") or []))
    ready_at = str(step.get("ready_at") or seen_state.get("ready_at") or "")
    return {
        **asdict(vendor),
        "latest_file": str(latest.relative_to(ROOT)) if latest else "",
        "latest_count": count,
        "latest_modified": modified,
        "active_today": vendor.slug in active_pair_for(),
        "run_status": run_status,
        "new_count": new_count,
        "ready_at": ready_at,
        "last_step_status": run_status,
    }


def command_for_scrape(vendor: Vendor, config: dict[str, Any]) -> list[str]:
    cmd = [
        PYTHON,
        str(ROOT / vendor.folder / vendor.scraper),
        "--posted-within-days",
        str(int(config.get("posted_within_days") or 0)),
    ]
    if vendor.terms_mode == "file":
        terms_path = ROOT / vendor.folder / ".dashboard_terms.txt"
        terms_path.write_text("\n".join(normalize_keywords(config.get("keywords"))) + "\n", encoding="utf-8")
        cmd.extend(["--terms-file", str(terms_path)])
    elif vendor.terms_mode == "append":
        for keyword in normalize_keywords(config.get("keywords")):
            cmd.extend(["--term", keyword])
    overrides = config.get("vendor_overrides") if isinstance(config.get("vendor_overrides"), dict) else {}
    vendor_override = overrides.get(vendor.slug) if isinstance(overrides.get(vendor.slug), dict) else {}
    page_value = vendor_override.get(vendor.page_arg) if vendor.page_arg else None
    if vendor.page_arg and page_value not in (None, ""):
        cmd.extend([f"--{vendor.page_arg}", str(int(page_value))])
    return cmd


def command_for_open(vendor: Vendor, config: dict[str, Any], payload: dict[str, Any]) -> list[str]:
    cmd = [
        PYTHON,
        str(ROOT / vendor.folder / vendor.opener),
        "--limit",
        str(int(payload.get("limit") or config.get("open_limit") or 0)),
        "--start-at",
        str(int(payload.get("start_at") or config.get("start_at") or 1)),
    ]
    if "delay" in opener_help(vendor):
        cmd.extend(["--delay", str(float(payload.get("delay") or config.get("delay") or 0.5))])
    if "keep-open-minutes" in opener_help(vendor):
        cmd.extend(["--keep-open-minutes", str(int(payload.get("keep_open_minutes") or config.get("keep_open_minutes") or 60))])
    return cmd


def opener_help(vendor: Vendor) -> str:
    path = ROOT / vendor.folder / vendor.opener
    try:
        return path.read_text(encoding="utf-8")
    except OSError:
        return ""


def start_run(kind: str, vendor_slugs: list[str], config: dict[str, Any]) -> str:
    STOP_EVENT.clear()
    run_id = datetime.now().strftime("%Y%m%d%H%M%S")
    RUNS[run_id] = {
        "id": run_id,
        "kind": kind,
        "status": "running",
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "finished_at": "",
        "vendors": vendor_slugs,
        "steps": [],
        "log": "",
    }
    thread = threading.Thread(target=run_scrapers, args=(run_id, vendor_slugs, config), daemon=True)
    thread.start()
    return run_id


def run_scrapers(run_id: str, vendor_slugs: list[str], config: dict[str, Any]) -> None:
    for slug in vendor_slugs:
        if STOP_EVENT.is_set():
            break
        vendor = VENDOR_BY_SLUG[slug]
        step_started = datetime.now()
        step = {
            "vendor": vendor.label,
            "slug": slug,
            "status": "running",
            "count": 0,
            "output": "",
            "started_at": step_started.isoformat(timespec="seconds"),
            "finished_at": "",
            "duration_seconds": None,
            "summary": "",
        }
        RUNS[run_id]["steps"].append(step)
        before = latest_jobs_file(vendor)
        previous_keys = baseline_keys_for_vendor(vendor, before)
        cmd = command_for_scrape(vendor, config)
        try:
            proc = launch_process(
                cmd,
                "scrape",
                f"Scrape {vendor.label}",
                cwd=ROOT,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            try:
                stdout, stderr = proc.communicate(timeout=900)
            finally:
                unregister_process(proc.pid)
            step_finished = datetime.now()
            after = latest_jobs_file(vendor)
            status = vendor_status(vendor)
            output = ((stdout or "") + "\n" + (stderr or "")).strip()
            stopped = STOP_EVENT.is_set() and proc.returncode and proc.returncode < 0
            ready_at = step_finished.isoformat(timespec="seconds") if proc.returncode == 0 and not stopped else ""
            latest_jobs = load_latest_jobs(vendor) if proc.returncode == 0 and not stopped else []
            new_summary = (
                update_seen_success(vendor, after, latest_jobs, previous_keys, run_id, ready_at)
                if proc.returncode == 0 and not stopped
                else {"new_keys": [], "new_count": 0, "ready_run_id": "", "ready_at": ""}
            )
            step.update({
                "status": "stopped" if stopped else ("done" if proc.returncode == 0 else "failed"),
                "returncode": proc.returncode,
                "count": status["latest_count"] if proc.returncode == 0 else 0,
                "latest_file": status["latest_file"],
                "changed": str(before) != str(after),
                "new_count": new_summary["new_count"],
                "new_keys": new_summary["new_keys"],
                "ready_run_id": new_summary["ready_run_id"],
                "ready_at": new_summary["ready_at"],
                "output": output[-5000:],
                "summary": summarize_scraper_output(output),
                "finished_at": step_finished.isoformat(timespec="seconds"),
                "duration_seconds": round((step_finished - step_started).total_seconds(), 1),
            })
            if STOP_EVENT.is_set():
                break
        except subprocess.TimeoutExpired:
            try:
                os.killpg(proc.pid, signal.SIGTERM)
            except OSError:
                pass
            unregister_process(proc.pid)
            step_finished = datetime.now()
            step.update({
                "status": "failed",
                "output": "Timed out after 900 seconds.",
                "summary": "Timed out after 900 seconds.",
                "returncode": -1,
                "finished_at": step_finished.isoformat(timespec="seconds"),
                "duration_seconds": round((step_finished - step_started).total_seconds(), 1),
            })
        except Exception as exc:  # noqa: BLE001 - surface failures in the UI.
            step_finished = datetime.now()
            step.update({
                "status": "failed",
                "output": str(exc),
                "summary": str(exc),
                "returncode": -1,
                "finished_at": step_finished.isoformat(timespec="seconds"),
                "duration_seconds": round((step_finished - step_started).total_seconds(), 1),
            })
    if STOP_EVENT.is_set():
        RUNS[run_id]["status"] = "stopped"
    else:
        RUNS[run_id]["status"] = "done" if all(step["status"] == "done" for step in RUNS[run_id]["steps"]) else "failed"
    RUNS[run_id]["finished_at"] = datetime.now().isoformat(timespec="seconds")


def summarize_scraper_output(output: str) -> str:
    if not output:
        return ""
    interesting_prefixes = ("Extracted ", "Saved ", "CSV:", "JSON:", "Excel:")
    lines = []
    for line in output.splitlines():
        stripped = line.strip()
        if stripped.startswith(interesting_prefixes):
            lines.append(stripped)
    return " | ".join(lines[-5:])


def open_vendor(slug: str, config: dict[str, Any], payload: dict[str, Any]) -> dict[str, Any]:
    STOP_EVENT.clear()
    vendor = VENDOR_BY_SLUG[slug]
    cmd = command_for_open(vendor, config, payload)
    proc = launch_process(cmd, "open", f"Open {vendor.label}", cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    threading.Thread(target=watch_process, args=(proc,), daemon=True).start()
    return {"status": "started", "pid": proc.pid, "vendor": vendor.label}


def watch_process(proc: subprocess.Popen[Any]) -> None:
    try:
        proc.wait()
    finally:
        unregister_process(proc.pid)


def open_job(slug: str, key: str) -> dict[str, Any]:
    vendor = VENDOR_BY_SLUG[slug]
    jobs = load_latest_jobs(vendor)
    for job in jobs:
        if job_key(slug, job) == key:
            url = job_url(job)
            if not url:
                raise ValueError("Selected job has no URL.")
            webbrowser.open_new_tab(url)
            count = increment_click(key)
            return {"status": "started", "vendor": vendor.label, "title": str(job.get("title") or ""), "job_url": url, "open_count": count}
    raise ValueError("Selected job was not found in the latest output.")


def open_new_jobs(slug: str, config: dict[str, Any]) -> dict[str, Any]:
    vendor = VENDOR_BY_SLUG[slug]
    jobs = load_latest_jobs(vendor)
    new_keys = latest_new_keys_for_vendor(slug)
    opened: list[dict[str, Any]] = []
    delay = max(float(config.get("delay") or 0), 0)
    for job in jobs:
        key = job_key(slug, job)
        if key not in new_keys:
            continue
        url = job_url(job)
        if not url:
            continue
        webbrowser.open_new_tab(url)
        count = increment_click(key)
        opened.append({
            "key": key,
            "title": str(job.get("title") or ""),
            "job_url": url,
            "open_count": count,
        })
        if delay:
            time.sleep(delay)
    return {"status": "started", "vendor": vendor.label, "count": len(opened), "jobs": opened}


def set_applied_mark(slug: str, key: str, applied: bool) -> dict[str, Any]:
    vendor = VENDOR_BY_SLUG[slug]
    jobs = load_latest_jobs(vendor)
    selected = next((job for job in jobs if job_key(slug, job) == key), None)
    if not selected:
        raise ValueError("Selected job was not found in the latest output.")
    marks = load_applied_marks()
    if applied:
        marks[key] = {
            "vendor": vendor.label,
            "slug": slug,
            "title": str(selected.get("title") or ""),
            "url": job_url(selected),
            "marked_at": time.time(),
        }
    else:
        marks.pop(key, None)
    save_applied_marks(marks)
    return {"status": "marked" if applied else "cleared", "key": key, "applied": applied}


def start_judge_apply(config: dict[str, Any], payload: dict[str, Any]) -> dict[str, Any]:
    STOP_EVENT.clear()
    vendor = VENDOR_BY_SLUG["judgegroup"]
    jobs = load_latest_jobs(vendor)
    start_at = max(int(payload.get("start_at") or config.get("start_at") or 1), 1)
    limit = int(payload.get("limit") or config.get("open_limit") or 1)
    selected_jobs = jobs[start_at - 1 :]
    if limit > 0:
        selected_jobs = selected_jobs[:limit]
    selected_jobs = [job for job in selected_jobs if str(job.get("job_url") or job.get("apply_url") or "").strip()]
    if not selected_jobs:
        raise ValueError(f"No Judge Group job found at position {start_at}")

    apply_config = config.get("judge_apply") if isinstance(config.get("judge_apply"), dict) else {}
    script = ROOT / vendor.folder / "judgegroup_apply.py"
    cmd = [
        PYTHON,
        str(script),
        "--start-at",
        str(start_at),
        "--limit",
        str(limit),
        "--resume",
        str(apply_config.get("resume_path") or "resume.docx"),
        "--first-name",
        str(apply_config.get("first_name") or ""),
        "--last-name",
        str(apply_config.get("last_name") or ""),
        "--email",
        str(apply_config.get("email") or ""),
        "--phone",
        str(apply_config.get("phone") or ""),
        "--country",
        str(apply_config.get("country") or "United States"),
        "--street-address",
        str(apply_config.get("street_address") or ""),
        "--city",
        str(apply_config.get("city") or ""),
        "--state",
        str(apply_config.get("state") or ""),
        "--zip-code",
        str(apply_config.get("zip_code") or ""),
        "--keep-open-seconds",
        str(int(payload.get("keep_open_seconds") or apply_config.get("keep_open_seconds") or 45)),
    ]
    if payload.get("submit"):
        cmd.append("--submit")
    proc = launch_process(cmd, "apply", "Judge Group apply", cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    threading.Thread(target=watch_process, args=(proc,), daemon=True).start()
    keys = [job_key(vendor.slug, selected) for selected in selected_jobs]
    if payload.get("submit"):
        for key in keys:
            set_applied_mark(vendor.slug, key, True)
    return {
        "status": "started",
        "pid": proc.pid,
        "vendor": vendor.label,
        "title": str(selected_jobs[0].get("title") or ""),
        "job_url": str(selected_jobs[0].get("job_url") or selected_jobs[0].get("apply_url") or ""),
        "job_key": keys[0],
        "count": len(selected_jobs),
        "job_keys": keys,
        "start_at": start_at,
        "limit": limit,
        "submit": bool(payload.get("submit")),
    }


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/":
            self.send_html(HTML)
        elif parsed.path == "/api/config":
            self.send_json({"config": load_config(), "vendors": [vendor_status(v) for v in VENDORS], "rotation": rotation_preview()})
        elif parsed.path == "/api/status":
            self.send_json({
                "runs": list(RUNS.values())[-10:],
                "vendors": [vendor_status(v) for v in VENDORS],
                "rotation": rotation_preview(),
                "active_processes": active_processes(),
                "stop_requested": STOP_EVENT.is_set(),
            })
        elif parsed.path == "/api/jobs":
            query = urllib.parse.parse_qs(parsed.query)
            slug = str((query.get("vendor") or [""])[0])
            if slug not in VENDOR_BY_SLUG:
                self.send_json({"ok": False, "error": "Unknown vendor."}, status=400)
                return
            try:
                config = load_config()
                result = latest_jobs_for_ui(slug, config, {
                    "start_at": (query.get("start_at") or [""])[0],
                    "limit": (query.get("limit") or [""])[0],
                    "new_only": (query.get("new_only") or [""])[0] in {"1", "true", "yes"},
                })
            except Exception as exc:  # noqa: BLE001
                self.send_json({"ok": False, "error": str(exc)}, status=500)
                return
            self.send_json({"ok": True, **result})
        else:
            self.send_error(404)

    def do_POST(self) -> None:  # noqa: N802
        payload = self.read_json()
        config = load_config()
        if self.path == "/api/config":
            config.update(payload)
            save_config(config)
            self.send_json({"ok": True, "config": load_config()})
            return
        if self.path == "/api/scrape":
            with RUN_LOCK:
                active = any(run.get("status") == "running" for run in RUNS.values())
                if active:
                    self.send_json({"ok": False, "error": "A scrape is already running."}, status=409)
                    return
                mode = payload.get("mode") or "selected"
                if mode == "all":
                    slugs = [vendor.slug for vendor in VENDORS]
                elif mode == "today":
                    slugs = active_pair_for()
                else:
                    slugs = [slug for slug in payload.get("vendors", []) if slug in VENDOR_BY_SLUG]
                if not slugs:
                    self.send_json({"ok": False, "error": "No vendors selected."}, status=400)
                    return
                run_id = start_run(mode, slugs, config)
                self.send_json({"ok": True, "run_id": run_id})
            return
        if self.path == "/api/open":
            slug = str(payload.get("vendor") or "")
            if slug not in VENDOR_BY_SLUG:
                self.send_json({"ok": False, "error": "Unknown vendor."}, status=400)
                return
            try:
                result = open_vendor(slug, config, payload)
            except Exception as exc:  # noqa: BLE001
                self.send_json({"ok": False, "error": str(exc)}, status=500)
                return
            self.send_json({"ok": True, **result})
            return
        if self.path == "/api/stop":
            result = stop_active_processes()
            self.send_json({"ok": True, **result})
            return
        if self.path == "/api/open-job":
            slug = str(payload.get("vendor") or "")
            key = str(payload.get("key") or "")
            if slug not in VENDOR_BY_SLUG or not key:
                self.send_json({"ok": False, "error": "Unknown job."}, status=400)
                return
            try:
                result = open_job(slug, key)
            except Exception as exc:  # noqa: BLE001
                self.send_json({"ok": False, "error": str(exc)}, status=500)
                return
            self.send_json({"ok": True, **result})
            return
        if self.path == "/api/open-new-jobs":
            slug = str(payload.get("vendor") or "")
            if slug not in VENDOR_BY_SLUG:
                self.send_json({"ok": False, "error": "Unknown vendor."}, status=400)
                return
            try:
                result = open_new_jobs(slug, config)
            except Exception as exc:  # noqa: BLE001
                self.send_json({"ok": False, "error": str(exc)}, status=500)
                return
            self.send_json({"ok": True, **result})
            return
        if self.path == "/api/applied":
            slug = str(payload.get("vendor") or "")
            key = str(payload.get("key") or "")
            if slug not in VENDOR_BY_SLUG or not key:
                self.send_json({"ok": False, "error": "Unknown job."}, status=400)
                return
            try:
                result = set_applied_mark(slug, key, bool(payload.get("applied")))
            except Exception as exc:  # noqa: BLE001
                self.send_json({"ok": False, "error": str(exc)}, status=500)
                return
            self.send_json({"ok": True, **result})
            return
        if self.path == "/api/judge-apply":
            try:
                result = start_judge_apply(config, payload)
            except Exception as exc:  # noqa: BLE001
                self.send_json({"ok": False, "error": str(exc)}, status=500)
                return
            self.send_json({"ok": True, **result})
            return
        self.send_error(404)

    def log_message(self, fmt: str, *args: Any) -> None:
        return

    def read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length") or "0")
        if not length:
            return {}
        try:
            data = json.loads(self.rfile.read(length).decode("utf-8"))
            return data if isinstance(data, dict) else {}
        except json.JSONDecodeError:
            return {}

    def send_json(self, payload: dict[str, Any], status: int = 200) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_html(self, html: str) -> None:
        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


HTML = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Job Portal Control</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #1d242b;
      --muted: #65717d;
      --line: #d8dee4;
      --paper: #f7f8fa;
      --panel: #ffffff;
      --brand: #136f63;
      --brand-2: #8a5a12;
      --danger: #a73333;
      --good-bg: #eaf6ef;
      --warn-bg: #fff4df;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--ink);
      background: var(--paper);
    }
    header {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 20px;
      padding: 24px 28px 18px;
      border-bottom: 1px solid var(--line);
      background: var(--panel);
    }
    h1, h2, h3 { margin: 0; letter-spacing: 0; }
    h1 { font-size: 25px; line-height: 1.15; }
    h2 { font-size: 17px; }
    h3 { font-size: 14px; }
    .sub { margin-top: 7px; color: var(--muted); font-size: 14px; }
    main {
      display: grid;
      grid-template-columns: 340px 1fr;
      min-height: calc(100vh - 86px);
    }
    aside {
      padding: 20px;
      border-right: 1px solid var(--line);
      background: #fbfbfc;
    }
    section { padding: 20px 24px 28px; }
    .block {
      border-bottom: 1px solid var(--line);
      padding: 0 0 18px;
      margin-bottom: 18px;
    }
    .block:last-child { border-bottom: 0; }
    label { display: block; font-size: 12px; font-weight: 700; color: #34404b; margin-bottom: 7px; }
    input, textarea, select {
      width: 100%;
      border: 1px solid #c9d1d9;
      background: white;
      border-radius: 6px;
      padding: 9px 10px;
      font: inherit;
      font-size: 14px;
      color: var(--ink);
    }
    textarea { min-height: 190px; resize: vertical; line-height: 1.35; }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .buttons { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }
    .button-row { display: flex; gap: 6px; align-items: center; flex-wrap: wrap; }
    button {
      border: 1px solid #b9c2cb;
      background: white;
      color: var(--ink);
      min-height: 36px;
      border-radius: 6px;
      padding: 0 12px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }
    button.primary { background: var(--brand); color: white; border-color: var(--brand); }
    button.secondary { background: #26323f; color: white; border-color: #26323f; }
    button.warn { background: var(--brand-2); color: white; border-color: var(--brand-2); }
    button.danger { background: var(--danger); color: white; border-color: var(--danger); }
    button.small { min-height: 30px; padding: 0 9px; font-size: 12px; }
    button:disabled { opacity: .55; cursor: not-allowed; }
    table {
      width: 100%;
      border-collapse: collapse;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
    }
    th, td { text-align: left; padding: 9px 12px; border-bottom: 1px solid var(--line); font-size: 13px; vertical-align: middle; }
    th { background: #eef1f4; font-size: 12px; color: #3d4852; }
    td.output-cell { max-width: 310px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .job-count { font-weight: 800; font-size: 14px; }
    .portal-name { font-weight: 800; }
    tr.active td { background: var(--good-bg); }
    tr:last-child td { border-bottom: 0; }
    .toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin-bottom: 14px;
    }
    .pill {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 28px;
      min-height: 24px;
      padding: 2px 8px;
      border-radius: 999px;
      background: #e8edf2;
      font-weight: 800;
      font-size: 12px;
    }
    .active-pill { background: #cdeedb; color: #075f3e; }
    .stop-pill { background: #f3d6d6; color: #7a2020; }
    .ready-pill { background: #dff1ff; color: #17547c; }
    .new-pill { background: #ffe4bc; color: #704300; }
    .failed-pill { background: #f3d6d6; color: #7a2020; }
    .running-pill { background: #e6e0ff; color: #3d2f82; }
    .zero { color: var(--muted); }
    .log {
      margin-top: 18px;
      background: #101820;
      color: #d8e5ee;
      padding: 12px;
      border-radius: 8px;
      min-height: 120px;
      max-height: 260px;
      overflow: auto;
      white-space: pre-wrap;
      font: 12px/1.45 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    }
    .rotation {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 18px;
    }
    .day {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px;
      min-height: 88px;
    }
    .day.today { border-color: #73b894; background: var(--good-bg); }
    .date { font-size: 12px; font-weight: 800; margin-bottom: 6px; }
    .names { color: #34404b; font-size: 13px; line-height: 1.35; }
    .notice {
      background: var(--warn-bg);
      border: 1px solid #e0c891;
      border-radius: 8px;
      padding: 10px 12px;
      font-size: 13px;
      color: #513b10;
      margin-bottom: 14px;
    }
    .judge-panel {
      display: none;
      margin-top: 12px;
      border: 1px solid var(--line);
      background: #f7fbfa;
      border-radius: 8px;
      padding: 10px;
    }
    .judge-panel.visible { display: block; }
    .judge-panel details { margin-top: 8px; }
    .judge-panel summary { cursor: pointer; font-weight: 800; font-size: 13px; color: #34404b; }
    .compact-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 8px; }
    .compact-grid .wide { grid-column: 1 / -1; }
    .jobs-panel {
      margin-top: 18px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
    }
    .jobs-panel-head {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      padding: 12px;
      border-bottom: 1px solid var(--line);
      background: #f3f6f8;
    }
    .jobs-panel-title { font-weight: 850; }
    .jobs-panel-sub { color: var(--muted); font-size: 12px; margin-top: 3px; }
    .job-list { display: grid; }
    .job-row {
      display: grid;
      grid-template-columns: 52px minmax(240px, 1fr) 132px 220px;
      gap: 12px;
      align-items: center;
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
    }
    .job-row.new-job {
      background: #fff9ed;
      box-shadow: inset 3px 0 0 #d78a00;
    }
    .job-row:last-child { border-bottom: 0; }
    .job-title { font-weight: 800; line-height: 1.25; }
    .job-meta { margin-top: 4px; color: var(--muted); font-size: 12px; line-height: 1.35; }
    .applied-pill { background: #d7f0df; color: #0b5b3b; }
    .not-applied-pill { background: #eef1f4; color: #53606b; }
    .empty-state { padding: 14px 12px; color: var(--muted); font-size: 13px; }
    .click-badge {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 18px;
      height: 18px;
      padding: 0 5px;
      border-radius: 999px;
      background: #f0a500;
      color: #fff;
      font-size: 11px;
      font-weight: 800;
      margin-left: 4px;
      vertical-align: middle;
    }
    @media (max-width: 980px) {
      header { align-items: start; flex-direction: column; }
      main { grid-template-columns: 1fr; }
      aside { border-right: 0; border-bottom: 1px solid var(--line); }
      .rotation { grid-template-columns: 1fr 1fr; }
      .job-row { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Job Portal Control</h1>
      <div class="sub">Daily scrape all portals. Actively open two weekday portals from the rotation.</div>
    </div>
    <div class="buttons">
      <button class="primary" id="scrapeAll">Scrape All 15</button>
      <button class="secondary" id="scrapeToday">Scrape Today's 2</button>
      <button class="danger" id="stopAll">Stop</button>
      <button id="refresh">Refresh</button>
    </div>
  </header>
  <main>
    <aside>
      <div class="block">
        <h2>Filters</h2>
        <div class="sub">These values are passed into each scraper that supports them.</div>
        <div class="row" style="margin-top:12px">
          <div>
            <label for="days">Posted within days</label>
            <input id="days" type="number" min="0" step="1">
          </div>
          <div>
            <label for="openLimit">Open limit</label>
            <input id="openLimit" type="number" min="0" step="1">
          </div>
        </div>
        <div class="row" style="margin-top:10px">
          <div>
            <label for="startAt">Start at</label>
            <input id="startAt" type="number" min="1" step="1">
          </div>
          <div>
            <label for="keepOpen">Keep open minutes</label>
            <input id="keepOpen" type="number" min="1" step="1">
          </div>
        </div>
      </div>
      <div class="block">
        <label for="keywords">Keywords</label>
        <textarea id="keywords" spellcheck="false"></textarea>
        <div class="buttons">
          <button class="primary" id="save">Save Controls</button>
        </div>
      </div>
      <div class="block">
        <h2>Open Jobs</h2>
        <div class="sub">Choose a portal, then open latest filtered results.</div>
        <div style="margin-top:12px">
          <label for="openVendor">Portal</label>
          <select id="openVendor"></select>
        </div>
        <div class="buttons">
          <button class="warn" id="openSelected">Open Selected</button>
          <button id="openToday">Open Today's 2</button>
        </div>
        <div class="judge-panel" id="judgePanel">
          <h3>Judge Group Apply</h3>
          <div class="buttons">
            <button class="secondary" id="judgeFill">Fill & Submit Queue</button>
          </div>
          <details>
            <summary>Applicant Settings</summary>
            <div class="compact-grid">
              <div class="wide">
                <label for="judgeResume">Resume path</label>
                <input id="judgeResume" spellcheck="false">
              </div>
              <div>
                <label for="judgeFirst">First name</label>
                <input id="judgeFirst">
              </div>
              <div>
                <label for="judgeLast">Last name</label>
                <input id="judgeLast">
              </div>
              <div class="wide">
                <label for="judgeEmail">Email</label>
                <input id="judgeEmail">
              </div>
              <div>
                <label for="judgePhone">Phone</label>
                <input id="judgePhone">
              </div>
              <div>
                <label for="judgeCountry">Country</label>
                <input id="judgeCountry">
              </div>
              <div class="wide">
                <label for="judgeStreet">Street address</label>
                <input id="judgeStreet">
              </div>
              <div>
                <label for="judgeCity">City</label>
                <input id="judgeCity">
              </div>
              <div>
                <label for="judgeState">State</label>
                <input id="judgeState">
              </div>
              <div>
                <label for="judgeZip">ZIP</label>
                <input id="judgeZip">
              </div>
              <div>
                <label for="judgeKeepOpen">Keep open seconds</label>
                <input id="judgeKeepOpen" type="number" min="1" step="1">
              </div>
            </div>
          </details>
        </div>
      </div>
      <div class="notice">Weekends are skipped for the two-portal rotation. Scrape All 15 is the daily safety net.</div>
    </aside>
    <section>
      <div class="toolbar">
        <h2>Weekday Rotation</h2>
        <span class="pill" id="runState">idle</span>
      </div>
      <div class="rotation" id="rotation"></div>
      <div class="toolbar">
        <h2>Portals</h2>
        <div class="buttons">
          <button id="scrapeSelected">Scrape Checked</button>
        </div>
      </div>
      <table>
        <thead>
          <tr>
            <th><input type="checkbox" id="toggleAll"></th>
            <th>Portal</th>
            <th>Today</th>
            <th>Latest Jobs</th>
            <th>New Jobs</th>
            <th>Latest Output</th>
            <th>Controls</th>
          </tr>
        </thead>
        <tbody id="vendors"></tbody>
      </table>
      <div class="jobs-panel">
        <div class="jobs-panel-head">
          <div>
            <div class="jobs-panel-title" id="jobsTitle">Selected Portal Jobs</div>
            <div class="jobs-panel-sub" id="jobsSub">Applied marks expire after 2 hours.</div>
          </div>
          <div class="button-row">
            <label style="display:flex;align-items:center;gap:6px;margin:0;font-size:12px;font-weight:800;color:#34404b">
              <input id="newOnly" type="checkbox" style="width:auto">
              Show new only
            </label>
            <button id="refreshJobs">Refresh Jobs</button>
          </div>
        </div>
        <div class="job-list" id="jobList">
          <div class="empty-state">Choose a portal to see the latest filtered jobs.</div>
        </div>
      </div>
      <div class="log" id="log">Ready.</div>
    </section>
  </main>
  <script>
    const state = {
      vendors: [],
      rotation: [],
      config: {},
      runs: [],
      configLoaded: false,
      configDirty: false,
      selectedVendor: "",
      checkedVendors: new Set(),
      vendorChecksTouched: false,
      portalJobs: [],
      jobsMeta: {},
      showNewOnly: false,
      stopRequested: false,
      activeProcesses: [],
    };
    const $ = (id) => document.getElementById(id);

    async function api(path, options = {}) {
      const response = await fetch(path, {
        headers: { "Content-Type": "application/json" },
        ...options,
      });
      const data = await response.json();
      if (!response.ok || data.ok === false) throw new Error(data.error || response.statusText);
      return data;
    }

    function readConfig() {
      return {
        posted_within_days: Number($("days").value || 0),
        open_limit: Number($("openLimit").value || 0),
        start_at: Number($("startAt").value || 1),
        keep_open_minutes: Number($("keepOpen").value || 60),
        keywords: $("keywords").value.split(/\n+/).map((x) => x.trim()).filter(Boolean),
        judge_apply: {
          resume_path: $("judgeResume").value.trim(),
          first_name: $("judgeFirst").value.trim(),
          last_name: $("judgeLast").value.trim(),
          email: $("judgeEmail").value.trim(),
          phone: $("judgePhone").value.trim(),
          country: $("judgeCountry").value.trim() || "United States",
          street_address: $("judgeStreet").value.trim(),
          city: $("judgeCity").value.trim(),
          state: $("judgeState").value.trim(),
          zip_code: $("judgeZip").value.trim(),
          keep_open_seconds: Number($("judgeKeepOpen").value || 45),
        },
      };
    }

    async function saveConfig() {
      const data = await api("/api/config", { method: "POST", body: JSON.stringify(readConfig()) });
      state.config = data.config || readConfig();
      state.configDirty = false;
      renderConfig();
      await refreshStatus();
    }

    function renderConfig() {
      if (state.configDirty) return;
      $("days").value = state.config.posted_within_days ?? 4;
      $("openLimit").value = state.config.open_limit ?? 8;
      $("startAt").value = state.config.start_at ?? 1;
      $("keepOpen").value = state.config.keep_open_minutes ?? 60;
      $("keywords").value = (state.config.keywords || []).join("\n");
      const judge = state.config.judge_apply || {};
      $("judgeResume").value = judge.resume_path || "resume.docx";
      $("judgeFirst").value = judge.first_name || "";
      $("judgeLast").value = judge.last_name || "";
      $("judgeEmail").value = judge.email || "";
      $("judgePhone").value = judge.phone || "";
      $("judgeCountry").value = judge.country || "United States";
      $("judgeStreet").value = judge.street_address || "";
      $("judgeCity").value = judge.city || "";
      $("judgeState").value = judge.state || "";
      $("judgeZip").value = judge.zip_code || "";
      $("judgeKeepOpen").value = judge.keep_open_seconds ?? 45;
    }

    function latestName(path) {
      return path ? path.split("/").pop() : "";
    }

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[char]));
    }

    function formatSeconds(seconds) {
      const safe = Math.max(Number(seconds || 0), 0);
      const minutes = Math.floor(safe / 60);
      const hours = Math.floor(minutes / 60);
      const remainder = minutes % 60;
      if (hours > 0) return `${hours}h ${remainder}m left`;
      return `${Math.max(minutes, 1)}m left`;
    }

    function portalRunPill(v) {
      if (v.run_status === "running") return '<span class="pill running-pill">running</span>';
      if (v.run_status === "done" && v.ready_at) return '<span class="pill ready-pill">ready</span>';
      if (v.run_status === "failed") return '<span class="pill failed-pill">failed</span>';
      if (v.run_status === "stopped") return '<span class="pill stop-pill">stopped</span>';
      return "";
    }

    function renderRotation() {
      const today = new Date().toISOString().slice(0, 10);
      $("rotation").innerHTML = state.rotation.slice(0, 5).map((row) => `
        <div class="day ${row.date === today ? "today" : ""}">
          <div class="date">${row.weekday}<br>${row.date}</div>
          <div class="names">${row.vendors.join("<br>")}</div>
        </div>
      `).join("");
    }

    function renderVendors() {
      const previousVendor = state.selectedVendor || $("openVendor").value;
      const previousChecks = new Set([...document.querySelectorAll(".pick:checked")].map((item) => item.value));
      if (state.vendorChecksTouched) state.checkedVendors = previousChecks;

      $("openVendor").innerHTML = state.vendors.map((v) => `<option value="${v.slug}">${v.label}</option>`).join("");
      const validVendor = state.vendors.some((v) => v.slug === previousVendor);
      $("openVendor").value = validVendor ? previousVendor : (state.vendors[0]?.slug || "");
      state.selectedVendor = $("openVendor").value;
      renderJudgePanel();

      $("vendors").innerHTML = state.vendors.map((v) => `
        <tr class="${v.active_today ? "active" : ""}">
          <td><input class="pick" type="checkbox" value="${v.slug}" ${(state.vendorChecksTouched ? state.checkedVendors.has(v.slug) : v.active_today) ? "checked" : ""}></td>
          <td><span class="portal-name" ${v.color === "red" ? 'style="color:#c0392b;font-weight:900"' : ""}>${v.label}</span></td>
          <td>${v.active_today ? '<span class="pill active-pill">active</span>' : ""}</td>
          <td class="${v.latest_count ? "" : "zero"}">
            <span class="job-count">${v.latest_count}</span>
            ${portalRunPill(v)}
          </td>
          <td>
            ${Number(v.new_count || 0) > 0 ? `<span class="pill new-pill">${v.new_count} new</span>` : '<span class="zero">No New</span>'}
          </td>
          <td class="output-cell" title="${v.latest_file || ""}">${v.latest_file ? `${latestName(v.latest_file)} · ${v.latest_modified}` : '<span class="zero">No output yet</span>'}</td>
          <td>
            <div class="button-row">
              <button class="small" data-open="${v.slug}">Open</button>
              <button class="small ${Number(v.new_count || 0) > 0 ? "warn" : ""}" data-open-new="${v.slug}" title="Open only jobs that were not in the previous successful scrape" ${Number(v.new_count || 0) > 0 ? "" : "disabled"}>
                ${Number(v.new_count || 0) > 0 ? `Open New ${v.new_count}` : "No New"}
              </button>
              ${v.slug === "judgegroup" ? '<button class="small secondary" data-judge-fill>Fill & Submit Queue</button>' : ""}
            </div>
          </td>
        </tr>
      `).join("");
      document.querySelectorAll("[data-open]").forEach((button) => {
        button.addEventListener("click", () => openPortal(button.dataset.open));
      });
      document.querySelectorAll("[data-open-latest]").forEach((button) => {
        button.addEventListener("click", () => openPortal(button.dataset.openLatest));
      });
      document.querySelectorAll("[data-open-new]").forEach((button) => {
        button.addEventListener("click", () => openNewJobs(button.dataset.openNew));
      });
      document.querySelectorAll("[data-judge-fill]").forEach((button) => {
        button.addEventListener("click", () => judgeApply());
      });
      document.querySelectorAll(".pick").forEach((box) => {
        box.addEventListener("change", () => {
          state.vendorChecksTouched = true;
          state.checkedVendors = new Set([...document.querySelectorAll(".pick:checked")].map((item) => item.value));
        });
      });
      updateToggleAll();
    }

    function renderJudgePanel() {
      $("judgePanel").classList.toggle("visible", state.selectedVendor === "judgegroup");
    }

    function renderJobsPanel() {
      const vendor = state.vendors.find((item) => item.slug === state.selectedVendor);
      $("jobsTitle").textContent = vendor ? `${vendor.label} Jobs` : "Selected Portal Jobs";
      const total = state.jobsMeta.total ?? 0;
      const newTotal = state.jobsMeta.new_total ?? 0;
      $("jobsSub").textContent = vendor
        ? `Showing ${state.jobsMeta.new_only ? "new jobs only" : "latest jobs"} from #${state.jobsMeta.start_at || $("startAt").value || 1}${state.jobsMeta.limit ? `, limit ${state.jobsMeta.limit}` : ""}. New since previous scrape: ${newTotal}. Total: ${total}.`
        : "Applied marks expire after 2 hours.";
      if (!state.portalJobs.length) {
        $("jobList").innerHTML = '<div class="empty-state">No latest jobs found for this portal yet. Scrape it first, then refresh.</div>';
        return;
      }
      $("jobList").innerHTML = state.portalJobs.map((job) => {
        const meta = [job.employment_type, job.location, job.posted_date].filter(Boolean).join(" | ");
        const applied = job.applied
          ? `<span class="pill applied-pill">Applied · ${escapeHtml(formatSeconds(job.seconds_left))}</span>`
          : '<span class="pill not-applied-pill">Not applied</span>';
        return `
          <div class="job-row ${job.is_new ? "new-job" : ""}">
            <div class="zero">#${job.index}</div>
            <div>
              <div class="job-title">${job.is_new ? '<span class="pill new-pill">New</span> ' : ""}${escapeHtml(job.title)}</div>
              <div class="job-meta">${escapeHtml(meta || job.url || "")}</div>
            </div>
            <div>${applied}</div>
            <div class="button-row">
              <button class="small" data-open-job="${escapeHtml(job.key)}">Open${job.open_count > 0 ? `<span class="click-badge">${job.open_count}</span>` : ""}</button>
              <button class="small secondary" data-mark-applied="${escapeHtml(job.key)}">Mark Applied</button>
              <button class="small" data-clear-applied="${escapeHtml(job.key)}">Clear</button>
            </div>
          </div>
        `;
      }).join("");
      document.querySelectorAll("[data-open-job]").forEach((button) => {
        button.addEventListener("click", () => openSingleJob(button.dataset.openJob));
      });
      document.querySelectorAll("[data-mark-applied]").forEach((button) => {
        button.addEventListener("click", () => markApplied(button.dataset.markApplied, true));
      });
      document.querySelectorAll("[data-clear-applied]").forEach((button) => {
        button.addEventListener("click", () => markApplied(button.dataset.clearApplied, false));
      });
    }

    function renderRuns() {
      const latest = state.runs[state.runs.length - 1];
      const activeCount = state.activeProcesses.length;
      $("runState").textContent = activeCount ? `${activeCount} running` : (latest ? latest.status : "idle");
      $("runState").classList.toggle("stop-pill", state.stopRequested);
      if (!latest) return;
      const freshNote = latest.kind === "today" ? "Fresh scrape from page/API start for today's 2 portals" : "Fresh scrape run";
      const lines = [`Run ${latest.id} - ${latest.kind} - ${latest.status}`, freshNote];
      const completed = (latest.steps || []).filter((step) => ["done", "failed", "stopped"].includes(step.status)).length;
      lines.push(`Progress: ${completed}/${(latest.vendors || []).length} portals completed`);
      for (const step of latest.steps || []) {
        const countText = step.status === "running" ? "scraping from beginning..." : `${step.count ?? 0} jobs`;
        const duration = step.duration_seconds == null ? "" : ` in ${step.duration_seconds}s`;
        const freshness = step.changed ? "new output" : "no new file";
        const newText = step.status === "done" ? `, ${step.new_count || 0} new` : "";
        const readyText = step.ready_at ? ` ready ${step.ready_at}` : "";
        lines.push(`${step.status.padEnd(7)} ${step.vendor}: ${countText}${newText}${duration} (${freshness})${readyText}`);
        if (step.latest_file) lines.push(`        file: ${step.latest_file}`);
        if (step.summary) lines.push(`        ${step.summary}`);
        if (step.status === "failed" && step.output) lines.push(step.output);
      }
      $("log").textContent = lines.join("\n");
    }

    async function refresh() {
      const configData = await api("/api/config");
      state.config = configData.config || {};
      state.configLoaded = true;
      renderConfig();
      await refreshStatus();
    }

    async function refreshStatus() {
      const data = await api("/api/status");
      state.runs = data.runs || [];
      state.vendors = data.vendors || [];
      state.rotation = data.rotation || [];
      state.activeProcesses = data.active_processes || [];
      state.stopRequested = Boolean(data.stop_requested);
      renderRotation();
      renderVendors();
      await loadSelectedJobs();
      renderRuns();
    }

    async function loadSelectedJobs() {
      if (!state.selectedVendor) return;
      const params = new URLSearchParams({
        vendor: state.selectedVendor,
        start_at: String(Number($("startAt").value || 1)),
        limit: String(Number($("openLimit").value || 0)),
      });
      if (state.showNewOnly) params.set("new_only", "1");
      try {
        const data = await api(`/api/jobs?${params.toString()}`);
        state.portalJobs = data.jobs || [];
        state.jobsMeta = { total: data.total, shown_total: data.shown_total, new_total: data.new_total, new_only: data.new_only, start_at: data.start_at, limit: data.limit };
      } catch (error) {
        state.portalJobs = [];
        state.jobsMeta = {};
      }
      renderJobsPanel();
    }

    async function openNewJobs(slug) {
      state.selectedVendor = slug;
      $("openVendor").value = slug;
      state.showNewOnly = true;
      $("newOnly").checked = true;
      renderJudgePanel();
      const data = await api("/api/open-new-jobs", { method: "POST", body: JSON.stringify({ vendor: slug }) });
      await loadSelectedJobs();
      const vendor = state.vendors.find((item) => item.slug === slug);
      const count = data.count || 0;
      $("log").textContent = count
        ? `Opening ${count} new job${count === 1 ? "" : "s"} for ${vendor?.label || slug}.`
        : `No new jobs to open for ${vendor?.label || slug}.`;
    }

    async function scrape(mode, vendors = []) {
      state.stopRequested = false;
      await saveConfig();
      await api("/api/scrape", { method: "POST", body: JSON.stringify({ mode, vendors }) });
      $("log").textContent = mode === "today"
        ? "Fresh scrape started for today's 2 portals from the beginning. Old counts remain visible until new output finishes."
        : "Fresh scrape started. Old counts remain visible until new output finishes.";
      setTimeout(refresh, 1200);
    }

    function selectedVendors() {
      state.vendorChecksTouched = true;
      state.checkedVendors = new Set([...document.querySelectorAll(".pick:checked")].map((item) => item.value));
      return [...state.checkedVendors];
    }

    async function openPortal(slug) {
      state.stopRequested = false;
      await saveConfig();
      const payload = { vendor: slug, limit: Number($("openLimit").value || 0), start_at: Number($("startAt").value || 1), keep_open_minutes: Number($("keepOpen").value || 60) };
      await api("/api/open", { method: "POST", body: JSON.stringify(payload) });
      $("log").textContent = `Opening ${slug} jobs in the browser.`;
    }

    async function openSingleJob(key) {
      const data = await api("/api/open-job", { method: "POST", body: JSON.stringify({ vendor: state.selectedVendor, key }) });
      $("log").textContent = `Opening ${data.title || data.job_url} in the browser. (opened ${data.open_count || 1}x today)`;
      await loadSelectedJobs();
    }

    async function markApplied(key, applied) {
      await api("/api/applied", { method: "POST", body: JSON.stringify({ vendor: state.selectedVendor, key, applied }) });
      await loadSelectedJobs();
      $("log").textContent = applied ? "Marked applied. This mark will clear automatically after 2 hours." : "Applied mark cleared.";
    }

    async function judgeApply() {
      state.stopRequested = false;
      await saveConfig();
      const payload = {
        submit: true,
        start_at: Number($("startAt").value || 1),
        limit: Number($("openLimit").value || 1),
        keep_open_seconds: Number($("judgeKeepOpen").value || 45),
      };
      const data = await api("/api/judge-apply", { method: "POST", body: JSON.stringify(payload) });
      $("log").textContent = `Filling and submitting ${data.count || 1} Judge Group application(s), starting at #${data.start_at || payload.start_at}.`;
      await loadSelectedJobs();
    }

    async function stopAll() {
      state.stopRequested = true;
      const data = await api("/api/stop", { method: "POST", body: JSON.stringify({}) });
      $("log").textContent = data.count
        ? `Stop requested. Terminated ${data.count} running process(es). Already opened browser tabs may remain open.`
        : "Stop requested. No running dashboard processes were found.";
      await refreshStatus();
    }

    $("save").addEventListener("click", saveConfig);
    $("refresh").addEventListener("click", refreshStatus);
    $("stopAll").addEventListener("click", stopAll);
    $("scrapeAll").addEventListener("click", () => scrape("all"));
    $("scrapeToday").addEventListener("click", () => scrape("today"));
    $("scrapeSelected").addEventListener("click", () => scrape("selected", selectedVendors()));
    $("openSelected").addEventListener("click", () => openPortal($("openVendor").value));
    $("openToday").addEventListener("click", async () => {
      state.stopRequested = false;
      for (const vendor of state.vendors.filter((v) => v.active_today)) {
        if (state.stopRequested) break;
        await openPortal(vendor.slug);
      }
    });
    $("toggleAll").addEventListener("change", (event) => {
      state.vendorChecksTouched = true;
      document.querySelectorAll(".pick").forEach((box) => { box.checked = event.target.checked; });
      state.checkedVendors = new Set([...document.querySelectorAll(".pick:checked")].map((item) => item.value));
      updateToggleAll();
    });
    $("openVendor").addEventListener("change", (event) => { state.selectedVendor = event.target.value; });
    $("openVendor").addEventListener("change", renderJudgePanel);
    $("openVendor").addEventListener("change", () => loadSelectedJobs());
    $("refreshJobs").addEventListener("click", () => loadSelectedJobs());
    $("newOnly").addEventListener("change", (event) => {
      state.showNewOnly = event.target.checked;
      loadSelectedJobs();
    });
    ["days", "openLimit", "startAt", "keepOpen", "keywords", "judgeResume", "judgeFirst", "judgeLast", "judgeEmail", "judgePhone", "judgeCountry", "judgeStreet", "judgeCity", "judgeState", "judgeZip", "judgeKeepOpen"].forEach((id) => {
      $(id).addEventListener("input", () => { state.configDirty = true; });
    });
    ["openLimit", "startAt"].forEach((id) => {
      $(id).addEventListener("change", () => loadSelectedJobs());
    });
    $("judgeFill").addEventListener("click", () => judgeApply());
    setInterval(refreshStatus, 5000);
    refresh().catch((error) => { $("log").textContent = error.message; });

    function updateToggleAll() {
      const boxes = [...document.querySelectorAll(".pick")];
      if (!boxes.length) return;
      const checked = boxes.filter((box) => box.checked).length;
      $("toggleAll").checked = checked === boxes.length;
      $("toggleAll").indeterminate = checked > 0 && checked < boxes.length;
    }
  </script>
</body>
</html>
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the local job portal control UI.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    save_config(load_config())
    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"Job portal dashboard: http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping dashboard.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
