# Job Portals Scraper

Python scrapers and a local dashboard for checking staffing-company job portals for contract-friendly senior software roles.

The filters are tuned for non-junior full stack, Python, backend/API, data engineering, ETL, AI/ML, GenAI, LLM, and RAG roles. They avoid W2, W2-only, no-C2C, F2F, onsite interview, local-only, permanent, direct-hire, junior, entry-level, intern, embedded, and hardware-heavy roles. C2H/contract-to-hire is allowed.

## Supported Portals

The repo includes separate scraper folders for:

- TEKsystems
- Apex Systems
- Judge Group
- Beacon Hill
- Akkodis
- Randstad
- Eliassen
- Experis
- Brooksource
- KellyMitchell
- Mitchell Martin
- CBTS
- Robert Half
- Kforce
- Insight Global
- Artech
- CCS Global Tech
- Diverse Lynx
- Ettain Group
- Harvey Nash
- Hays
- INSPYR Solutions
- Matlen Silver
- Motion Recruitment
- NTT DATA
- Oliver James
- Optomi
- Open Systems Technologies
- PTR Global
- Pyramid Consulting
- Strategic Staffing Solutions
- Vaco
- Venturi Group

## Requirements

- Python 3.10 or newer
- macOS, Linux, or Windows for scraping
- macOS is recommended for the desktop launcher and launchd scheduling scripts
- Chrome is recommended for the browser-opening helpers

Some helper scripts use Playwright to open browser tabs. After installing Python packages, install Playwright browsers if you plan to use those helpers:

```bash
python3 -m playwright install chromium
```

## Fresh Clone Setup

Clone the repo:

```bash
git clone https://github.com/sai-sujan/jobPortals_scrapper.git
cd jobPortals_scrapper
```

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

On Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Or use the setup script:

```bash
./scripts/setup_env.sh
source .venv/bin/activate
```

Optional browser support:

```bash
python -m playwright install chromium
```

Run tests:

```bash
PYTHONPATH=src pytest -q
```

One resume-related test is skipped unless you place a local `resume.docx` in the project root.

## Optional Environment Variables

No secrets are committed. Most scrapers run without any private keys.

Copy the example file if you need optional features:

```bash
cp .env.example .env
```

Then edit `.env` and load it before running commands:

```bash
source .env
```

Optional variables:

- `GROQ_API_KEY`: enables optional AI filtering when you run a scraper with `--ai-filter`
- `KFORCE_AZURE_SEARCH_KEY`: enables the Kforce Azure Search scraper
- `JOB_PORTAL_RESUME`: default resume path for Apex application prefill

If `KFORCE_AZURE_SEARCH_KEY` is not set, the Kforce scraper skips cleanly and writes an empty output file instead of crashing.

## Run The Dashboard

The easiest way to use the project is the local dashboard:

```bash
python3 job_portal_dashboard.py --port 8766
```

Open:

```bash
open http://127.0.0.1:8766
```

On Windows/Linux, open that URL manually in your browser.

The dashboard lets you:

- scrape all 15 portals
- scrape only today's 2-portal weekday rotation
- scrape checked portals
- edit the posted-date window, usually `4` days
- edit the search keywords
- open the latest filtered jobs for one portal
- open the two portals assigned for today's applying rotation
- track temporary applied marks and per-day open counts in the job list
- fill and submit the filtered Judge Group queue when applicant settings are configured

Dashboard settings are saved locally in `job_portal_dashboard_config.json`. That file is ignored by git so each user can keep their own keyword list and controls.

## Docker

Docker gives a clean, repeatable environment for anyone who clones the repo.

Build and run the dashboard:

```bash
docker compose up --build
```

Open:

```bash
open http://127.0.0.1:8766
```

On Windows/Linux, open that URL manually in your browser.

Docker runtime state is stored in `docker_state/`, and scraper outputs are bind-mounted into each portal's `output/` folder. Those paths are ignored by git.

Useful commands:

```bash
make docker-up
make docker-logs
make docker-down
```

Notes:

- Scraping works inside Docker.
- Browser-opening helpers run inside the container, so local desktop browser tab opening is best from the normal local `.venv` workflow.
- Playwright Chromium is installed in the Docker image for scripts that need browser automation.

## macOS Desktop Launcher

The repo includes an AppleScript launcher:

```bash
open_job_portal_dashboard.applescript
```

To create a clickable desktop app on macOS:

```bash
osacompile -o "$HOME/Desktop/Job Portal Dashboard.app" open_job_portal_dashboard.applescript
```

Double-clicking that app starts the dashboard on port `8766` if needed and opens it in the browser.

## Dashboard Controls

`Posted within days`

Filters job postings by age. `4` means keep jobs posted in the last four days. `0` disables the date filter.

`Open limit`

Controls only how many browser tabs open. It does not affect scraping. `0` means open all latest filtered jobs.

`Start at`

Skips earlier jobs when opening tabs. `1` starts from the first result. `6` starts from job 6.

`Keep open minutes`

Used by Playwright-based open scripts to keep the browser alive.

`Keywords`

Search terms passed into portals that support keyword search.

## Recommended Daily Workflow

Run this each weekday:

```bash
python3 job_portal_dashboard.py --port 8766
```

Then in the dashboard:

1. Click `Scrape All 15`
2. Wait for job counts to update
3. Review the two green active portals for the day
4. Click `Open Today's 2`
5. Apply manually from the opened tabs

There are 15 portals, so the two-portal applying rotation takes 8 workdays to cover every portal. Scraping all 15 daily is still the safety net.

## Command-Line Daily Run

Scrape every portal on weekdays:

```bash
python3 run_daily_scrape.py --mode all --weekdays-only
```

Scrape all portals and open today's two portals:

```bash
python3 run_daily_scrape.py --mode all --weekdays-only --open-today
```

Scrape only today's two active portals:

```bash
python3 run_daily_scrape.py --mode today --weekdays-only
```

Logs are written to `logs/`, which is ignored by git.

## Run A Single Portal

Each portal has a separate folder. General pattern:

```bash
python3 <folder>/<scraper>.py --posted-within-days 4
python3 <folder>/<open_script>.py --limit 5
```

Examples:

```bash
python3 teksystems_applying_script/teksystems_scraper.py --posted-within-days 4
python3 teksystems_applying_script/teksystems_open_jobs.py --limit 5

python3 experis_applying_script/experis_scraper.py --posted-within-days 4 --term "python developer" --term "data engineer"
python3 experis_applying_script/experis_open_jobs.py --limit 5

python3 kellymitchell_applying_script/kellymitchell_scraper.py --posted-within-days 4 --term "full stack developer"
python3 kellymitchell_applying_script/kellymitchell_open_jobs.py --limit 5
```

Most scrapers write:

- CSV
- JSON
- Excel workbook
- daily grouped JSON files

Outputs are written under each portal's `output/` folder. Output folders are ignored by git.

## Common Scraper Options

Many scrapers support:

```bash
--posted-within-days 4
--term "python developer"
--terms-file terms.txt
--no-excel
```

Some portal-specific scrapers also support:

```bash
--max-pages 3
--limit 10
--jobs-per-page 50
--ai-filter
```

Use `--help` on any scraper to see exact options:

```bash
python3 judgegroup_applying_script/judgegroup_scraper.py --help
```

## Website Usage Notes

This project reads public job pages, feeds, or public API responses and writes local filtered results. The dashboard does not submit applications unless you explicitly click an apply automation control and provide applicant settings.

Use it respectfully:

- keep request rates moderate
- do not run tight loops
- review each job manually before applying
- respect each website's terms of use
- avoid scraping portals more often than needed

Most open scripts are only helpers that open job pages in a browser. Judge Group also has an optional fill-and-submit queue from the dashboard when you configure applicant settings. Use that only for jobs you intend to apply to.

## CI

GitHub Actions runs on pushes and pull requests:

- install Python dependencies
- compile Python files
- run `PYTHONPATH=src pytest -q`
- build the Docker image

## Legacy Job Checker

The repo also contains an older generalized job checker under `src/job_checker`.

Doctor check:

```bash
PYTHONPATH=src python3 -m job_checker doctor
```

Run:

```bash
PYTHONPATH=src python3 -m job_checker run
```

It writes:

- `data/jobs.sqlite`
- `reports/jobs_latest.xlsx`
- `reports/YYYY-MM-DD_jobs.xlsx`
- `reports/YYYY-MM-DD_jobs.csv`
- `reports/site/index.html`
- `logs/job_checker.log`

Those files are local runtime outputs and are ignored by git.

## Local Files Not Committed

The repo intentionally ignores:

- scraper output folders
- reports
- SQLite databases
- logs
- local dashboard config
- `.env`
- resumes and personal documents
- Excel files
- screenshots and captured HTML pages
- Python caches

This keeps GitHub clean and avoids uploading private/generated data.

## Troubleshooting

`ModuleNotFoundError`

Run:

```bash
pip install -r requirements.txt
```

`playwright` browser errors

Run:

```bash
python -m playwright install chromium
```

Kforce returns zero jobs

Set `KFORCE_AZURE_SEARCH_KEY` if you have access to the key. Without it, Kforce is skipped safely.

Dashboard port is already in use

Use another port:

```bash
python3 job_portal_dashboard.py --port 8767
```

Then open:

```bash
open http://127.0.0.1:8767
```
