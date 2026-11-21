#!/usr/bin/env python3
"""Unit tests for agency_billing_io_reconciliation P1 module."""

from __future__ import annotations

import unittest
import zipfile

from agency_billing_io_reconciliation.io_billing_reconciler import OUT_DIR, run_reconciler


class TestAgencyIOBillingReconciler(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.summary = run_reconciler()

    def test_exact_record_count_and_zero_variance(self) -> None:
        self.assertEqual(self.summary["total_verified_records"], 40144)
        self.assertEqual(self.summary["agency_insertion_orders"], 24)
        self.assertEqual(self.summary["on_time_launch_orders"], 24)
        self.assertEqual(self.summary["total_abs_billing_variance_usd"], 0.0)
        self.assertGreaterEqual(self.summary["blended_delivery_attainment_pct"], 96.0)

    def test_excel_formula_workbook_valid(self) -> None:
        xlsx = OUT_DIR / "agency_io_billing_reconciliation.xlsx"
        self.assertTrue(xlsx.exists())
        self.assertTrue(zipfile.is_zipfile(xlsx))
        with zipfile.ZipFile(xlsx) as zf:
            sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("<f>SUM(", sheet)


if __name__ == "__main__":
    unittest.main()
