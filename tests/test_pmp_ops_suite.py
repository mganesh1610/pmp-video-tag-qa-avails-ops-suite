#!/usr/bin/env python3
"""Unit and boundary tests for PMP Tag QA, Publisher Avails & XLS Operations Suite."""

from __future__ import annotations

import json
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import pmp_tag_qa_avails_engine as eng  # noqa: E402


class TestPMPTagQAAvailsSuite(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.summary = eng.run_pipeline()

    def test_01_combined_public_record_conservation(self) -> None:
        counts = self.summary["record_counts"]
        self.assertEqual(counts["publisher_articles"], 39644)
        self.assertEqual(counts["pmp_auctions"], 200000)
        self.assertEqual(counts["combined_records"], 239644)

    def test_02_preset_avails_24_segments_and_buffer(self) -> None:
        avails = self.summary["avails_metrics"]
        self.assertEqual(avails["preset_segment_count"], 24)
        self.assertGreater(avails["gross_monthly_avails_imps"], avails["bookable_pmp_avails_imps"])
        self.assertGreater(avails["bookable_pmp_avails_imps"], 50_000_000)

    def test_03_pre_launch_tag_qa_defect_interception(self) -> None:
        qa = self.summary["tag_and_creative_qa"]
        self.assertEqual(qa["total_packages_audited"], 120)
        self.assertEqual(qa["seeded_pre_launch_defects_caught"], 24)
        self.assertEqual(qa["post_remediation_launch_ready"], 120)

    def test_04_negative_vast_and_non_ssl_boundary_rejection(self) -> None:
        bad_pkg = {
            "tag_id": "TAG-NEG",
            "am_owner": "AM-Test",
            "deal_id": "PMP-PC-9999",
            "format": "Outstream Video VAST 4.2",
            "dimensions": "640x360",
            "initial_kb": 120,
            "landing_url": "http://insecure.example/page",
            "markup": '<VAST version="4.2"><Ad><InLine></InLine></Ad></VAST>',
        }
        res = eng.validate_ad_tag_package(bad_pkg)
        self.assertFalse(res["passed"])
        self.assertIn("NON_SSL_HTTP_CALL", res["defects"])
        self.assertIn("MISSING_CLICK_MACRO", res["defects"])
        self.assertIn("MISSING_MP4_MEDIAFILE", res["defects"])

    def test_05_zero_dollar_reconciliation_and_openxml_packages(self) -> None:
        self.assertEqual(self.summary["delivery_reconciliation"]["ledger_variance_usd"], 0.0)
        xlsx_p = ROOT / self.summary["deliverables"]["xlsx_workbook"]
        pptx_p = ROOT / self.summary["deliverables"]["pptx_screenshot_deck"]
        self.assertTrue(zipfile.is_zipfile(xlsx_p))
        self.assertTrue(zipfile.is_zipfile(pptx_p))
        with zipfile.ZipFile(xlsx_p, "r") as zf:
            self.assertIn("xl/worksheets/sheet4.xml", zf.namelist())
        with zipfile.ZipFile(pptx_p, "r") as zf:
            self.assertIn("ppt/slides/slide7.xml", zf.namelist())
        saved = json.loads((ROOT / "outputs" / "validation_summary.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["validation_gates_passed"], "12/12")


if __name__ == "__main__":
    unittest.main()
