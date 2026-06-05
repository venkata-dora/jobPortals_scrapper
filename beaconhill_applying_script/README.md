# Beacon Hill standalone job scraper

Scrapes Beacon Hill jobs from the public WordPress REST API and writes filtered CSV, JSON, Excel, and daily grouped outputs.

The default search is tuned for a non-junior full stack/Python profile, with AI/ML/GenAI/LLM/RAG and data engineering roles included. It defaults to jobs posted in the last 4 days.

## Run

```bash
python3 beaconhill_scraper.py
```

Useful options:

```bash
python3 beaconhill_scraper.py --posted-within-days 4 --max-pages 3
python3 beaconhill_scraper.py --term "python developer" --term "data engineer" --no-excel
python3 beaconhill_open_jobs.py --limit 5
python3 beaconhill_open_jobs.py --limit 5 --apply
```

Fill one application in Chrome and leave it open without submitting:

```bash
python3 beaconhill_apply.py
```

Fill and submit the Seattle Sr Automation Engineer role:

```bash
python3 beaconhill_apply.py --submit --keep-open-seconds 20
```

## Filters

- Keeps contract, temporary/contract, and temp-to-hire/C2H style jobs.
- Allows C2H.
- Excludes W2/W-2, W2-only, "No C2C", no corp-to-corp, F2F, in-person/onsite interview, local-only, permanent, and direct-hire language.
- Excludes junior, entry-level, intern, embedded/hardware-heavy roles.
- Positively ranks Python, backend, full stack, API, AI/ML, LLM, RAG, data engineer, ETL, cloud, AWS, Azure, and related skills.
