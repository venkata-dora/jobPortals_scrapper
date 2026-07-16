import unittest
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from shared_vendor_filters import VendorJob, filter_and_sort_jobs, is_within_posted_days, strict_job_filter_reasons


def make_job(title: str, location: str = "New York, NY", raw_text: str = "Contract role.") -> VendorJob:
    return VendorJob(
        source_company="Test",
        search_term="test",
        title_rank=100,
        title_rank_reasons="test",
        title=title,
        category="",
        location=location,
        employment_type="Contract",
        salary="",
        posted_date="",
        job_id=title,
        job_url=f"https://example.com/{title}",
        apply_url="",
        contact_info="",
        description_snippet="",
        raw_text=raw_text,
    )


class SharedVendorFilterTest(unittest.TestCase):
    def test_strict_filter_rejects_foreign_and_irrelevant_jobs(self):
        self.assertIn("No U.S. location", strict_job_filter_reasons("Senior Data Engineer", "Pune, India"))
        self.assertIn("Unrelated specialty", strict_job_filter_reasons("Cyber IAM Platform Engineer", "Irving, TX"))
        self.assertIn("Non-target language", strict_job_filter_reasons(".NET Developer", "Charlotte, NC"))
        self.assertIn(
            "No U.S. location",
            strict_job_filter_reasons(
                "Senior Data Engineer",
                "Pune, India",
                "NTT DATA has offices in the United States.",
            ),
        )

    def test_strict_filter_accepts_us_target_roles_despite_bad_vendor_country_suffix(self):
        reasons = strict_job_filter_reasons(
            "Python Developer",
            ", , Jersey City, NJ, NORWAY",
            "Location: Jersey City, NJ",
        )
        self.assertEqual(reasons, [])

    def test_strict_filter_uses_description_when_location_field_is_empty(self):
        reasons = strict_job_filter_reasons(
            "Senior Software Engineer (AI/ML)",
            "",
            "This position is in Fort Worth, Texas.",
        )
        self.assertEqual(reasons, [])

    def test_today_rejects_yesterday_even_when_less_than_24_hours_old(self):
        eastern = timezone(timedelta(hours=-4))
        now = datetime(2026, 6, 12, 8, 0, tzinfo=eastern)

        self.assertTrue(is_within_posted_days("2026-06-12T01:00:00-04:00", -1, now))
        self.assertFalse(is_within_posted_days("2026-06-11T20:00:00-04:00", -1, now))
        self.assertTrue(is_within_posted_days("2026-06-11T20:00:00-04:00", 1, now))

    def test_java_titles_are_excluded(self):
        jobs = [
            make_job("Java Full Stack Developer"),
            make_job("Senior Java Developer"),
            make_job("Full Stack Java Engineer"),
            make_job("Python Full Stack Developer"),
        ]
        kept = filter_and_sort_jobs(jobs, posted_within_days=4, exclude_disallowed_work=False)
        self.assertEqual([job.title for job in kept], ["Python Full Stack Developer"])

    def test_javascript_title_is_not_treated_as_java(self):
        jobs = [make_job("JavaScript Full Stack Developer")]
        kept = filter_and_sort_jobs(jobs, posted_within_days=4, exclude_disallowed_work=False)
        self.assertEqual([job.title for job in kept], ["JavaScript Full Stack Developer"])

    def test_lead_and_architect_titles_are_excluded(self):
        jobs = [
            make_job("Lead Python Engineer"),
            make_job("Data Architect"),
            make_job("Senior Python Engineer"),
        ]
        kept = filter_and_sort_jobs(jobs, posted_within_days=4, exclude_disallowed_work=False)
        self.assertEqual([job.title for job in kept], ["Senior Python Engineer"])

    def test_security_clearance_jobs_are_excluded(self):
        cases = [
            "Active security clearance required.",
            "Must hold a Secret clearance.",
            "An active TS/SCI is required.",
            "Must be eligible for Public Trust.",
        ]
        for raw_text in cases:
            with self.subTest(raw_text=raw_text):
                reasons = strict_job_filter_reasons("Senior Python Engineer", "Reston, VA", raw_text)
                self.assertIn("Security clearance required", reasons)


if __name__ == "__main__":
    unittest.main()
