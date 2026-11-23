#!/usr/bin/env python3
"""Tests for PerfDash (run from the build directory).

Covers: data sanity (realistic ROAS band, positive revenue, Oct festive
peak), workbook formula hygiene (no short range literals, no structured
references), and a LibreOffice headless recalc of the generated workbook
against the independent Python ground truth in data/ground_truth.json.

Run:  python3 -m unittest discover -s tests -v
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "data")
XLSX = os.path.join(BASE, "performance-dashboards.xlsx")


def load_gt():
    with open(os.path.join(DATA, "ground_truth.json")) as f:
        return json.load(f)


def recalc_workbook():
    """LibreOffice headless recalc; returns {sheet: {(row, col): value}}.

    LO will NOT overwrite the source in place, so copy the xlsx into a
    src/ subdir and convert into a separate outdir, then read the
    converted file with data_only=True.
    """
    from openpyxl import load_workbook
    tmp = tempfile.mkdtemp(prefix="perfdash-recalc-")
    srcdir = os.path.join(tmp, "src")
    outdir = os.path.join(tmp, "out")
    os.makedirs(srcdir)
    os.makedirs(outdir)
    shutil.copy(XLSX, os.path.join(srcdir, "performance-dashboards.xlsx"))
    p = subprocess.run(
        ["soffice", "--headless", "--calc", "--convert-to", "xlsx",
         "--outdir", outdir,
         os.path.join(srcdir, "performance-dashboards.xlsx")],
        capture_output=True, text=True, timeout=300)
    if p.returncode != 0:
        raise RuntimeError(f"soffice failed: {p.stderr[-2000:]}")
    converted = os.path.join(outdir, "performance-dashboards.xlsx")
    wb = load_workbook(converted, data_only=True)
    values = {}
    for ws in wb.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    cells[(c.row, c.column)] = c.value
        values[ws.title] = cells
    shutil.rmtree(tmp, ignore_errors=True)
    return values


class TestDataSanity(unittest.TestCase):
    def setUp(self):
        self.gt = load_gt()

    def test_row_count(self):
        self.assertEqual(self.gt["rows"], 495)

    def test_blended_roas_in_realistic_band(self):
        self.assertGreater(self.gt["totals"]["roas"], 2.0)
        self.assertLess(self.gt["totals"]["roas"], 10.0)

    def test_every_channel_roas_in_band(self):
        for c in self.gt["channels"]:
            self.assertGreater(c["roas"], 2.0, msg=c["channel"])
            self.assertLess(c["roas"], 10.0, msg=c["channel"])

    def test_october_is_festive_peak(self):
        revs = {m["month"]: m["revenue"] for m in self.gt["months"]}
        self.assertEqual(max(revs, key=revs.get), "2026-10")

    def test_variance_mix_has_beats_and_misses(self):
        signs = {1 if m["variance"] > 0 else -1 for m in self.gt["months"]}
        self.assertEqual(signs, {1, -1})

    def test_all_revenue_positive(self):
        self.assertGreater(self.gt["totals"]["revenue"], 0)
        for m in self.gt["months"]:
            self.assertGreater(m["revenue"], 0, msg=m["month"])


class TestWorkbookFormulas(unittest.TestCase):
    def test_no_short_range_literals(self):
        # Excel range literals MUST repeat the column letter on both ends:
        # $H$2:$H$5000, never $H$2:$5000 (LibreOffice raises #NAME?).
        from openpyxl import load_workbook
        wb = load_workbook(XLSX)
        bad = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        for m in re.finditer(r"\$[A-Z]+\$\d+:\$\d+", c.value):
                            bad.append(f"{ws.title}!{c.coordinate}: {m.group(0)}")
        self.assertEqual(bad, [])

    def test_no_structured_references(self):
        from openpyxl import load_workbook
        wb = load_workbook(XLSX)
        bad = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and ("[@" in c.value
                                                     or "tbl" in c.value.lower()):
                        bad.append(f"{ws.title}!{c.coordinate}")
        self.assertEqual(bad, [])

    def test_monthly_uses_sumifs_on_raw(self):
        from openpyxl import load_workbook
        wb = load_workbook(XLSX)
        ws = wb["Monthly"]
        self.assertTrue(str(ws["B4"].value).startswith("=SUMIFS(Raw!"))
        self.assertIn("Targets!", str(ws["K4"].value))

    def test_dashboard_kpis_reference_focus_month(self):
        from openpyxl import load_workbook
        wb = load_workbook(XLSX)
        ws = wb["Dashboard"]
        self.assertEqual(ws["B3"].value, "2026-10")
        self.assertIn("$B$3", str(ws["B6"].value))
        self.assertIn("INDEX(Channel!", str(ws["B12"].value))

    def test_focus_month_has_dropdown(self):
        from openpyxl import load_workbook
        wb = load_workbook(XLSX)
        ws = wb["Dashboard"]
        dvs = list(ws.data_validations.dataValidation)
        self.assertEqual(len(dvs), 1)
        self.assertEqual(dvs[0].type, "list")
        self.assertIn("2026-10", dvs[0].formula1)

    def test_three_charts_present(self):
        from openpyxl import load_workbook
        wb = load_workbook(XLSX)
        self.assertEqual(len(wb["Dashboard"]._charts), 3)


class TestLibreOfficeRecalc(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.gt = load_gt()
        cls.values = recalc_workbook()

    def test_no_formula_errors(self):
        errors = []
        for sheet, cells in self.values.items():
            for (r, c), v in cells.items():
                if isinstance(v, str) and v.startswith("#"):
                    errors.append(f"{sheet} R{r}C{c}: {v}")
        self.assertEqual(errors, [])

    def test_raw_total_row(self):
        raw = self.values["Raw"]
        n = self.gt["rows"]
        self.assertAlmostEqual(raw[(n + 2, 8)], self.gt["totals"]["revenue"],
                               delta=1.0)
        self.assertAlmostEqual(raw[(n + 2, 9)], self.gt["totals"]["ad_spend"],
                               delta=1.0)
        self.assertAlmostEqual(raw[(n + 2, 13)], self.gt["totals"]["roas"],
                               delta=0.01)

    def test_monthly_rows(self):
        monthly = self.values["Monthly"]
        for i, m in enumerate(self.gt["months"]):
            r = 4 + i
            self.assertAlmostEqual(monthly[(r, 5)], m["revenue"], delta=1.0,
                                   msg=m["month"] + " revenue")
            self.assertAlmostEqual(monthly[(r, 9)], m["roas"], delta=0.01,
                                   msg=m["month"] + " ROAS")
            self.assertAlmostEqual(monthly[(r, 11)], m["target"], delta=1.0,
                                   msg=m["month"] + " target")
            self.assertAlmostEqual(monthly[(r, 13)], m["variance_pct"],
                                   delta=0.001, msg=m["month"] + " variance %")
        self.assertAlmostEqual(monthly[(11, 5)], self.gt["totals"]["revenue"],
                               delta=1.0)

    def test_channel_rows(self):
        channel = self.values["Channel"]
        for i, c in enumerate(self.gt["channels"]):
            r = 4 + i
            self.assertAlmostEqual(channel[(r, 5)], c["revenue"], delta=1.0,
                                   msg=c["channel"] + " revenue")
            self.assertAlmostEqual(channel[(r, 9)], c["roas"], delta=0.01,
                                   msg=c["channel"] + " ROAS")
        self.assertAlmostEqual(channel[(9, 5)], self.gt["totals"]["revenue"],
                               delta=1.0)
        shares = sum(channel[(4 + i, 11)] for i in range(5))
        self.assertAlmostEqual(shares, 1.0, delta=0.001)

    def test_campaign_rows(self):
        campaign = self.values["Campaign"]
        for i, c in enumerate(self.gt["campaigns"]):
            r = 4 + i
            self.assertAlmostEqual(campaign[(r, 5)], c["revenue"], delta=1.0,
                                   msg=c["campaign"] + " revenue")
            self.assertAlmostEqual(campaign[(r, 9)], c["roas"], delta=0.01,
                                   msg=c["campaign"] + " ROAS")
        self.assertAlmostEqual(campaign[(10, 5)],
                               self.gt["totals"]["revenue"], delta=1.0)

    def test_dashboard_kpis_focus_month(self):
        dash = self.values["Dashboard"]
        fm = self.gt["focus_month"]
        self.assertAlmostEqual(dash[(6, 2)], fm["revenue"], delta=1.0)
        self.assertAlmostEqual(dash[(7, 2)], fm["orders"], delta=0.5)
        self.assertAlmostEqual(dash[(9, 2)], fm["roas"], delta=0.01)
        self.assertAlmostEqual(dash[(10, 2)], fm["ctr"], delta=0.001)
        self.assertEqual(dash[(12, 2)], self.gt["best_channel"])

    def test_all_checks_pass(self):
        checks = self.values["Checks"]
        for r in range(4, 14):
            self.assertEqual(checks[(r, 4)], "PASS", msg=f"Checks row {r}")


if __name__ == "__main__":
    unittest.main()
