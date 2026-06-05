#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${0}")" && pwd)"
PROJECT_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
cd "${PROJECT_DIR}"

/opt/anaconda3/bin/python apexsystems_applying_script/apex_scraper.py \
  --posted-within-days "${APEX_POSTED_WITHIN_DAYS:-4}" \
  --sleep "${APEX_SLEEP:-0.2}"
