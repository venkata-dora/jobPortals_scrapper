# TEKsystems Standalone Scraper

Scrapes TEKsystems jobs from `careers.teksystems.com` and writes filtered CSV, JSON, Excel, and daily grouped outputs. Defaults to jobs posted in the last 4 days.

The scraper uses the live search page because TEKsystems returns zero or unrelated
results until its distance slider is maximized. It also selects the Contractor
hiring type, expands the results, and reads the job metadata from each live card.

## Run

```bash
python3 teksystems_scraper.py
```

Useful options:

```bash
python3 teksystems_scraper.py --posted-within-days 4
python3 teksystems_scraper.py --term "data engineer" --term "python developer"
python3 teksystems_scraper.py --no-excel
```

Use `--http-only` only as a fallback when a local Chrome, Edge, or Playwright
Chromium browser is unavailable; TEKsystems' static page data is less reliable.

## Open Jobs

```bash
python3 teksystems_open_jobs.py --limit 10
python3 teksystems_open_jobs.py --apply --limit 5
```

## Fill Applications

Fill one application in a visible Chrome window and leave it open for review:

```bash
python3 teksystems_applying_script/teksystems_apply.py \
  --first-name "Venkata" \
  --last-name "Dora" \
  --email "venkatasworkofficial@gmail.com" \
  --phone "9735445393" \
  --resume "/Users/saisujan/Desktop/resumes_contract/venkataD_resume.docx" \
  --limit 1
```

Add `--submit` to submit after the form is filled. Browser sessions are saved
in `teksystems_applying_script/.browser_profile` and reused on later runs.

## Filter Intent

Keeps Python, full stack, backend/API, AI/ML, data engineering, ETL, cloud, and data science roles.

Rejects junior/entry/intern titles, W2-only, no-C2C/no-corp-to-corp, full-time, permanent, direct hire, face-to-face interview, onsite interview, and local-only signals.

Contract-to-hire/C2H is allowed.
