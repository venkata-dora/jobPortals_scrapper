import tempfile
import unittest
from pathlib import Path

from job_portal_dashboard import (
    Vendor,
    baseline_keys_for_vendor,
    job_keys_for_rows,
    load_seen_state,
    update_seen_success,
)


def job(title: str, url: str) -> dict:
    return {"title": title, "job_url": url}


class DashboardSeenStateTest(unittest.TestCase):
    def setUp(self):
        self.vendor = Vendor("testco", "TestCo", "testco", "test_scraper.py", "test_open.py", "testco")

    def test_no_previous_baseline_marks_all_jobs_new(self):
        with tempfile.TemporaryDirectory() as tmp:
            state_path = Path(tmp) / "seen.json"
            jobs = [job("Python Developer", "https://example.com/1"), job("Data Engineer", "https://example.com/2")]

            result = update_seen_success(self.vendor, None, jobs, set(), "run-1", "2026-05-28T09:00:00", state_path)

            self.assertEqual(result["new_count"], 2)
            self.assertEqual(result["new_keys"], job_keys_for_rows(self.vendor.slug, jobs))

    def test_previous_baseline_excludes_repeated_jobs(self):
        with tempfile.TemporaryDirectory() as tmp:
            state_path = Path(tmp) / "seen.json"
            old_jobs = [job("Python Developer", "https://example.com/1")]
            new_jobs = old_jobs + [job("Data Engineer", "https://example.com/2")]
            previous_keys = set(job_keys_for_rows(self.vendor.slug, old_jobs))

            result = update_seen_success(self.vendor, None, new_jobs, previous_keys, "run-2", "2026-05-28T10:00:00", state_path)

            self.assertEqual(result["new_count"], 1)
            self.assertEqual(result["new_keys"], job_keys_for_rows(self.vendor.slug, [new_jobs[1]]))

    def test_second_scrape_uses_last_successful_output_as_baseline(self):
        with tempfile.TemporaryDirectory() as tmp:
            state_path = Path(tmp) / "seen.json"
            first_jobs = [job("Python Developer", "https://example.com/1")]
            second_jobs = first_jobs + [job("Data Engineer", "https://example.com/2")]

            update_seen_success(self.vendor, None, first_jobs, set(), "run-1", "2026-05-28T09:00:00", state_path)
            state = load_seen_state(state_path)
            previous_keys = baseline_keys_for_vendor(self.vendor, None, state)
            result = update_seen_success(self.vendor, None, second_jobs, previous_keys, "run-2", "2026-05-28T13:00:00", state_path)

            self.assertEqual(result["new_count"], 1)
            self.assertEqual(result["new_keys"], job_keys_for_rows(self.vendor.slug, [second_jobs[1]]))

    def test_failed_or_stopped_scrape_does_not_advance_baseline(self):
        with tempfile.TemporaryDirectory() as tmp:
            state_path = Path(tmp) / "seen.json"
            first_jobs = [job("Python Developer", "https://example.com/1")]
            update_seen_success(self.vendor, None, first_jobs, set(), "run-1", "2026-05-28T09:00:00", state_path)

            # Failed/stopped scrapes should skip update_seen_success entirely.
            state = load_seen_state(state_path)
            previous_keys = baseline_keys_for_vendor(self.vendor, None, state)

            self.assertEqual(previous_keys, set(job_keys_for_rows(self.vendor.slug, first_jobs)))


if __name__ == "__main__":
    unittest.main()
