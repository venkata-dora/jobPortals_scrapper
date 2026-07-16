# Apex Systems Applying Script

Standalone Apex Systems scraper for consultant job listings.

## Run

From the project root:

```bash
python apexsystems_applying_script/apex_scraper.py
```

Outputs are written to:

```text
apexsystems_applying_script/output/
```

Each run creates:

- one full CSV with recent jobs ranked
- one full JSON with recent jobs ranked
- one Excel workbook with Summary, All Jobs, and day sheets
- a `daily_YYYYMMDD_HHMMSS/` folder with separate CSV/JSON files per posting day

By default, the scraper keeps only jobs posted in the last 4 days, removes
F2F/onsite/full-time/permanent/direct-hire signals, and removes hourly jobs
whose max rate is below `$55/hour`.
W2 jobs are kept.

Run the daily update manually:

```bash
apexsystems_applying_script/run_daily_apex.sh
```

## Fill Applications

Open filled Apex application tabs for manual submit:

```bash
python apexsystems_applying_script/apex_apply.py --limit 4
```

The script uploads your resume, clicks Continue, fills/verifies first name, last
name, and email, then leaves each tab open at `Submit Application`. It uses a
visible Chrome window and a persistent profile at
`apexsystems_applying_script/.browser_profile`, so cookies and login sessions
are retained between runs. If human verification appears, solve it in the open
window and the script resumes after the challenge clears.

Submit automatically after filling:

```bash
python apexsystems_applying_script/apex_apply.py --limit 4 --submit
```

Process one or more application URLs directly:

```bash
python apexsystems_applying_script/apex_apply.py \
  --url "https://example.com/application/1" \
  --url "https://example.com/application/2"
```

Automatic submit is opt-in so the normal command remains safe for review. The
automation reuses normal browser state; it does not spoof browser fingerprints
or bypass verification challenges.

Change posting-date filter:

```bash
python apexsystems_applying_script/apex_scraper.py --posted-within-days 7
```

Disable filters:

```bash
python apexsystems_applying_script/apex_scraper.py --posted-within-days 0 --keep-w2-f2f-onsite-interview
```
