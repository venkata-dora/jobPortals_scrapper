#!/usr/bin/env python3
"""Open filtered INSPYR Solutions job pages in browser tabs."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_vendor_open_jobs import run_open_jobs


if __name__ == "__main__":
    raise SystemExit(run_open_jobs("inspyr", "INSPYR Solutions", Path(__file__).resolve().parent / "output"))
