#!/usr/bin/env python3
"""Tests for ForecastForge (run from the build directory).

Covers: model math (month-1 figures, loan run-down, opening balance
sheet, annual sums, balance checks, cash tie-out, sensitivity
centre), workbook formula hygiene (no short range literals, no
structured references), and a LibreOffice headless recalc of the
generated workbook against the independent Python ground truth in
data/ground_truth.json.

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
XLSX = os.path.join(BASE, "forecast-model.xlsx")

import model  # noqa: E402


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
    tmp = tempfile.mkdtemp(prefix="forecastforge-recalc-")
    srcdir = os.path.join(tmp, "src")
    outdir = os.path.join(tmp, "out")
    os.makedirs(srcdir)
    os.makedirs(outdir)
    shutil.copy(XLSX, os.path.join(srcdir, "forecast-model.xlsx"))
    p = subprocess.run(
        ["soffice", "--headless", "--calc", "--convert-to", "xlsx",
         "--outdir", outdir, os.path.join(srcdir, "forecast-model.xlsx")],
        capture_output=True, text=True, timeout=300)
    if p.returncode != 0:
        raise RuntimeError(f"soffice failed: {p.stderr[-2000:]}")
    converted = os.path.join(outdir, "forecast-model.xlsx")
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


class TestModelMath(unittest.TestCase):
    def setUp(self):
        self.gt = load_gt()
        self.months = self.gt["months"]

    def test_opening_balance_sheet_balances(self):
        self.assertAlmostEqual(self.gt["opening_check"], 0.0, places=2)

    def test_month1_figures(self):
        m1 = self.months[0]
        self.assertAlmostEqual(m1["revenue"], 950000.0, places=2)
        self.assertAlmostEqual(m1["cogs"], 950000.0 * 0.58, places=2)
        self.assertAlmostEqual(m1["salaries"], 185000.0, places=2)
        self.assertAlmostEqual(m1["marketing"], 28500.0, places=2)
        self.assertAlmostEqual(m1["depreciation"], 35000.0, places=2)
        self.assertAlmostEqual(m1["interest"], 2500000.0 * 0.11 / 12, places=2)
        self.assertGreater(m1["net"], 0)

    def test_depreciation_steps_up_with_capex(self):
        self.assertAlmostEqual(self.months[5]["depreciation"], 35000.0, places=2)
        self.assertAlmostEqual(self.months[6]["depreciation"], 40000.0, places=2)
        self.assertAlmostEqual(self.months[12]["depreciation"],
                               35000.0 + 5000.0 + 10000.0 / 3, places=2)

    def test_loan_runs_down(self):
        self.assertAlmostEqual(self.months[0]["loan_open"], 2500000.0, places=2)
        self.assertAlmostEqual(self.months[35]["loan_open"], 750000.0, places=2)

    def test_annual_sums_match_monthly(self):
        for yi in range(3):
            chunk = self.months[yi * 12:(yi + 1) * 12]
            y = self.gt["years"][yi]
            for key in ("revenue", "cogs", "gross", "opex", "ebitda",
                        "depreciation", "interest", "pbt", "tax", "net"):
                self.assertAlmostEqual(
                    y[key], sum(r[key] for r in chunk), places=2,
                    msg=f"{y['fy']} {key}")

    def test_balance_checks_are_zero(self):
        for b in self.gt["balance_sheet"]:
            self.assertAlmostEqual(b["balance_check"], 0.0, places=2,
                                   msg=b["fy"])

    def test_cash_tie_out(self):
        for c in self.gt["cash_flow"]:
            self.assertAlmostEqual(
                c["closing_cash"], c["opening_cash"] + c["net_cash"],
                places=2, msg=c["fy"])
        for b, c in zip(self.gt["balance_sheet"], self.gt["cash_flow"]):
            self.assertAlmostEqual(b["cash"], c["closing_cash"], places=2,
                                   msg=b["fy"])

    def test_profitable_every_year(self):
        for y in self.gt["years"]:
            self.assertGreater(y["net"], 0, msg=y["fy"])

    def test_sensitivity_centre_equals_y3_net(self):
        self.assertAlmostEqual(
            self.gt["sensitivity_centre"], self.gt["years"][2]["net"],
            places=2)

    def test_sensitivity_cell_matches_closed_form(self):
        # Worst corner (slow growth, high COGS) must still compute and
        # stay below the centre; best corner above it.
        y3 = self.gt["years"][2]
        worst = model.sensitivity_cell(self.months, self.gt["years"],
                                       -0.02, 0.03)
        best = model.sensitivity_cell(self.months, self.gt["years"],
                                      0.02, -0.03)
        centre = model.sensitivity_cell(self.months, self.gt["years"], 0, 0)
        self.assertAlmostEqual(centre, y3["net"], places=2)
        self.assertLess(worst, centre)
        self.assertGreater(best, centre)

    def test_breakeven_below_month1_revenue_eventually(self):
        # Year-1 break-even is tight (~9.0L vs 9.5L month-1 revenue) but
        # revenue outgrows it; by FY2029 monthly revenue is far above.
        r = self.gt["ratios"]
        self.assertLess(r[0]["breakeven_revenue"], self.months[11]["revenue"])
        self.assertLess(r[2]["breakeven_revenue"], self.months[35]["revenue"])


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

    def test_monthly_formulas_reference_assumptions(self):
        from openpyxl import load_workbook
        wb = load_workbook(XLSX)
        ws = wb["Monthly"]
        self.assertIn("Assumptions!", str(ws["C4"].value))
        self.assertIn("Assumptions!", str(ws["F4"].value))
        self.assertIn("Assumptions!", str(ws["M4"].value))
        self.assertIn("Assumptions!", str(ws["N4"].value))

    def test_scenario_selector_is_a_formula(self):
        from openpyxl import load_workbook
        wb = load_workbook(XLSX)
        ws = wb["Assumptions"]
        self.assertEqual(ws["B2"].value, "Base")
        self.assertTrue(str(ws["B8"].value).startswith("=IF("))
        self.assertTrue(str(ws["B12"].value).startswith("=IF("))


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

    def test_annual_revenue_and_net(self):
        annual = self.values["Annual_PL"]
        for i, y in enumerate(self.gt["years"]):
            r = 4 + i
            self.assertAlmostEqual(annual[(r, 2)], y["revenue"], delta=1.0,
                                   msg=y["fy"] + " revenue")
            self.assertAlmostEqual(annual[(r, 11)], y["net"], delta=1.0,
                                   msg=y["fy"] + " net")

    def test_monthly_spot_values(self):
        monthly = self.values["Monthly"]
        for idx in (0, 5, 6, 12, 23, 35):
            m = self.gt["months"][idx]
            r = 4 + idx
            self.assertAlmostEqual(monthly[(r, 3)], m["revenue"], delta=1.0,
                                   msg=f"month {idx + 1} revenue")
            self.assertAlmostEqual(monthly[(r, 17)], m["net"], delta=1.0,
                                   msg=f"month {idx + 1} net")

    def test_balance_checks_zero(self):
        bs = self.values["Balance_Sheet"]
        for col in (2, 3, 4, 5):
            self.assertAlmostEqual(bs[(14, col)], 0.0, delta=1.0,
                                   msg=f"balance check col {col}")

    def test_cash_tie_out(self):
        bs = self.values["Balance_Sheet"]
        cf = self.values["Cash_Flow"]
        for i, b in enumerate(self.gt["balance_sheet"]):
            bs_col = 3 + i
            cf_col = 2 + i
            self.assertAlmostEqual(bs[(4, bs_col)], b["cash"], delta=1.0,
                                   msg=b["fy"] + " BS cash")
            self.assertAlmostEqual(cf[(13, cf_col)], b["cash"], delta=1.0,
                                   msg=b["fy"] + " CF closing cash")

    def test_sensitivity_centre(self):
        sens = self.values["Sensitivity"]
        self.assertAlmostEqual(sens[(7, 5)],
                               self.gt["sensitivity_centre"], delta=1.0,
                               msg="sensitivity centre vs Y3 net")
        self.assertAlmostEqual(sens[(19, 2)], 0.0, delta=1.0,
                               msg="centre minus Annual Y3 net")

    def test_ratios(self):
        ratios = self.values["Ratios"]
        for i, r in enumerate(self.gt["ratios"]):
            col = 2 + i
            self.assertAlmostEqual(ratios[(5, col)], r["net_margin"],
                                   delta=0.001, msg=r["fy"] + " net margin")
            self.assertAlmostEqual(ratios[(9, col)], r["breakeven_revenue"],
                                   delta=1.0, msg=r["fy"] + " break-even")

    def test_dashboard_kpis(self):
        dash = self.values["Dashboard"]
        self.assertAlmostEqual(dash[(5, 2)], self.gt["years"][2]["revenue"],
                               delta=1.0)
        self.assertAlmostEqual(dash[(6, 2)], self.gt["years"][2]["net"],
                               delta=1.0)


if __name__ == "__main__":
    unittest.main()
