# Akkodis Daily Jobs Scraper

Final Akkodis scraper for daily filtered, ranked job lists with Excel output.

## Run

From the project root:

```bash
python akkodis_applying_script/akkodis_scraper.py --term "python developer" --term "ai engineer"
```

Outputs are written to:

```text
akkodis_applying_script/output/
```

Each run creates:

- one full CSV with all recent jobs ranked
- one full JSON with all recent jobs ranked
- one full Excel workbook with a Summary sheet, All Jobs sheet, and separate day sheets
- a `daily_YYYYMMDD_HHMMSS/` folder with separate CSV/JSON files per posting day

Ranking is based on title fit. AI/ML, Python, Django/FastAPI, backend, full-stack, data science, and data engineering titles score higher than generic titles.

By default, the scraper keeps only jobs posted in the last 4 days.
It also removes jobs whose title or description contains no-C2C, F2F, face-to-face interview, onsite interview, in-person interview, local-candidates-only, full-time, permanent, or direct-hire signals.
W2 jobs are kept.
Hourly jobs below `$55/hour` are removed by default.

To change the posting window:

```bash
python akkodis_applying_script/akkodis_scraper.py --posted-within-days 7
```

To disable the posting-date filter:

```bash
python akkodis_applying_script/akkodis_scraper.py --posted-within-days 0
```

To keep F2F/onsite-interview/full-time/permanent jobs for review:

```bash
python akkodis_applying_script/akkodis_scraper.py --keep-w2-f2f-onsite-interview
```

## Daily Excel

Run the daily Akkodis update manually:

```bash
akkodis_applying_script/run_daily_akkodis.sh
```

Install the daily macOS schedule:

```bash
akkodis_applying_script/install_daily_schedule.sh
```

That schedules the scraper for 8:00 AM every day. New Excel files are written to
`akkodis_applying_script/output/`.

Remove the schedule:

```bash
akkodis_applying_script/uninstall_daily_schedule.sh
```

## Default Search Terms

If no `--term` is provided, the scraper searches:

- ai engineer
- ai ml engineer
- aiml engineer
- machine learning engineer
- ml engineer
- generative ai engineer
- gen ai engineer
- llm engineer
- rag engineer
- data scientist
- senior data scientist
- applied data scientist
- machine learning scientist
- data engineer
- python data engineer
- etl developer
- senior python developer
- python developer
- python engineer
- senior python engineer
- backend python engineer
- backend software engineer
- software engineer python
- software developer python
- django developer
- django engineer
- fastapi developer
- api developer
- rest api developer
- full stack software engineer
- full stack developer
- senior software engineer
- software engineer
- cloud engineer
- aws python developer
- azure python developer
