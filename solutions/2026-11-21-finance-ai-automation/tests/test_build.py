#!/usr/bin/env python3
"""Tests for FinanceFlow AI (run from the build directory).

Covers: classification rules, counterparty extraction, anomaly
detection, CSV reconciliation, workbook formulas, and a LibreOffice
headless recalc of the generated workbook against the independent
Python ground truth in out/summary.json.

Run:  python3 -m unittest discover -s tests -v
"""
import csv
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "out")
XLSX = os.path.join(BASE, "financeflow-ai-report.xlsx")

import pipeline  # noqa: E402


def read_csv(name):
    with open(os.path.join(OUT, name), newline="") as f:
        return list(csv.DictReader(f))


class TestClassification(unittest.TestCase):
    def test_rule_order_specific_beats_generic(self):
        # "PURCHASE STOCK" via NEFT must not fall through to Sales.
        self.assertEqual(pipeline.classify(
            "NEFT/HINDUSTAN FOODS LTD/PURCHASE STOCK"),
            ("Purchases (COGS)", 1.0))
        self.assertEqual(pipeline.classify("UPI/9812345670/SHARMA KIRANA/SALES"),
                         ("Sales Income", 1.0))

    def test_payment_fallback_is_low_confidence(self):
        self.assertEqual(pipeline.classify("NEFT/ABC/PAYMENT"),
                         ("Sales Income", 0.7))

    def test_unclassified_never_guessed(self):
        for narr in ("UPI/9876543210/UNKNOWN", "MISC CREDIT", "NEFT/XXXX/XXXX"):
            self.assertEqual(pipeline.classify(narr), ("Unclassified", 0.0))

    def test_each_category_classifies(self):
        cases = {
            "NEFT/LANDLORD PATNA/RENT APR 2026": "Rent",
            "NEFT/SALARY - RAMESH KUMAR": "Salaries",
            "ELECTRICITY BILL - NBPDCL APR": "Utilities",
            "UPI/BSNL/BROADBAND APR": "Utilities",
            "PETROL - HP PETROL PUMP PATNA": "Transport & Fuel",
            "GST PMT - CGST/SGST MAY 2026": "GST Paid",
            "BANK CHARGES - SMS ALERT": "Bank Charges",
            "INT CREDIT - SAVINGS A/C": "Interest Income",
            "UPI/SELF/TRANSFER": "Owner Drawings",
        }
        for narr, want in cases.items():
            got, conf = pipeline.classify(narr)
            self.assertEqual(got, want, narr)
            self.assertEqual(conf, 1.0, narr)

    def test_counterparty_extraction(self):
        self.assertEqual(
            pipeline.extract_counterparty("UPI/9812345670/SHARMA KIRANA/SALES"),
            "Sharma Kirana")
        self.assertEqual(
            pipeline.extract_counterparty("NEFT/HINDUSTAN FOODS LTD/PURCHASE STOCK"),
            "Hindustan Foods Ltd")
        self.assertEqual(
            pipeline.extract_counterparty("NEFT/SALARY - RAMESH KUMAR"),
            "Ramesh Kumar")
        self.assertEqual(
            pipeline.extract_counterparty("UPI/9876543210/UNKNOWN"), "Unknown")


class TestPipelineOutputs(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with open(os.path.join(OUT, "summary.json")) as _f:
            cls.summary = json.load(_f)
        cls.txns = read_csv("classified_transactions.csv")
        cls.monthly = read_csv("monthly_summary.csv")
        cls.cats = read_csv("category_summary.csv")
        cls.anoms = read_csv("anomalies.csv")

    def test_counts(self):
        self.assertEqual(self.summary["transactions"], 268)
        self.assertEqual(len(self.txns), 268)
        self.assertEqual(len(self.monthly), 7)   # Apr..Oct 2026
        self.assertEqual(len(self.cats), 11)     # 10 CoA + Unclassified
        self.assertEqual(self.summary["unclassified"], 3)

    def test_anomaly_set(self):
        flags = sorted({a["flag"] for a in self.anoms})
        self.assertEqual(flags, ["BALANCE_MISMATCH", "DUPLICATE",
                                 "LARGE_AMOUNT", "UNCLASSIFIED"])
        self.assertEqual(len(self.anoms), 6)
        by_flag = {}
        for a in self.anoms:
            by_flag[a["flag"]] = by_flag.get(a["flag"], 0) + 1
        self.assertEqual(by_flag, {"BALANCE_MISMATCH": 1, "DUPLICATE": 1,
                                   "LARGE_AMOUNT": 1, "UNCLASSIFIED": 3})

    def test_balance_mismatch_is_exactly_500(self):
        row = [a for a in self.anoms if a["flag"] == "BALANCE_MISMATCH"][0]
        self.assertIn("diff 500.00", row["detail"])

    def test_monthly_reconciles_with_transactions(self):
        # Independent recomputation from the classified rows.
        tot_inc = tot_exp = 0.0
        per_month = {}
        for t in self.txns:
            tot_inc += float(t["credit"])
            tot_exp += float(t["debit"])
            per_month.setdefault(t["month"], [0.0, 0.0, 0])
            per_month[t["month"]][0] += float(t["credit"])
            per_month[t["month"]][1] += float(t["debit"])
            per_month[t["month"]][2] += 1
        self.assertAlmostEqual(tot_inc, self.summary["total_income"], places=2)
        self.assertAlmostEqual(tot_exp, self.summary["total_expenses"], places=2)
        for row in self.monthly:
            inc, exp, cnt = per_month[row["month"]]
            self.assertAlmostEqual(float(row["income"]), inc, places=2)
            self.assertAlmostEqual(float(row["expenses"]), exp, places=2)
            self.assertAlmostEqual(float(row["net"]), inc - exp, places=2)
            self.assertEqual(int(row["txn_count"]), cnt)

    def test_category_totals_sum_to_net(self):
        total = sum(float(c["total"]) for c in self.cats)
        self.assertAlmostEqual(total, self.summary["net"], places=2)

    def test_closing_balance_ties_out(self):
        self.assertAlmostEqual(
            self.summary["opening_balance"] + self.summary["net"],
            self.summary["closing_balance"], places=2)
        self.assertGreater(self.summary["closing_balance"], 0)

    def test_no_formula_errors_in_csvs(self):
        for t in self.txns:
            self.assertIn(t["category"], [c["category"] for c in self.cats])
            self.assertRegex(t["month"], r"^2026-0[4-9]$|^2026-10$")


class TestWorkbook(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from openpyxl import load_workbook
        cls.wb = load_workbook(XLSX)
        with open(os.path.join(OUT, "summary.json")) as _f:
            cls.summary = json.load(_f)

    def test_sheets_present(self):
        self.assertEqual(self.wb.sheetnames,
                         ["Guide", "Raw", "Monthly", "Categories",
                          "Dashboard", "Checks"])

    def test_raw_net_and_month_are_formulas(self):
        raw = self.wb["Raw"]
        self.assertEqual(raw.max_row, 269)  # header + 268 txns
        for r in (2, 100, 269):
            self.assertEqual(raw.cell(row=r, column=9).value,
                             "=H%d-G%d" % (r, r))
            self.assertEqual(raw.cell(row=r, column=10).value,
                             '=TEXT(A%d,"yyyy-mm")' % r)

    def test_monthly_uses_sumifs(self):
        mon = self.wb["Monthly"]
        self.assertIn("SUMIFS", mon.cell(row=2, column=2).value)
        self.assertIn("SUMIFS", mon.cell(row=2, column=3).value)
        self.assertIn("COUNTIFS", mon.cell(row=2, column=5).value)

    def test_dashboard_has_focus_month_and_charts(self):
        dash = self.wb["Dashboard"]
        self.assertEqual(dash["B2"].value, "2026-10")
        self.assertEqual(len(dash._charts), 3)

    def test_no_hardcoded_money_in_monthly(self):
        mon = self.wb["Monthly"]
        for r in range(2, 9):
            for c in (2, 3, 4):
                v = mon.cell(row=r, column=c).value
                self.assertIsInstance(v, str, (r, c, v))
                self.assertTrue(v.startswith("="), (r, c, v))


class TestLibreOfficeRecalc(unittest.TestCase):
    """Recalc the workbook in LibreOffice and compare to ground truth."""

    @classmethod
    def setUpClass(cls):
        from openpyxl import load_workbook
        tmp = tempfile.mkdtemp(prefix="financeflow-recalc-")
        cls.tmp = tmp
        # LibreOffice will not overwrite the source file in place, so the
        # source lives in src/ and the recalculated copy lands in out/.
        srcdir = os.path.join(tmp, "src")
        outdir = os.path.join(tmp, "out")
        os.makedirs(srcdir)
        os.makedirs(outdir)
        src = os.path.join(srcdir, "financeflow-ai-report.xlsx")
        shutil.copy(XLSX, src)
        p = subprocess.run(
            ["soffice", "--headless", "--calc", "--convert-to", "xlsx",
             "--outdir", outdir, src],
            capture_output=True, text=True, timeout=180)
        assert p.returncode == 0, p.stderr[-2000:]
        cls.recalced = load_workbook(os.path.join(outdir,
                                                  "financeflow-ai-report.xlsx"),
                                     data_only=True)
        with open(os.path.join(OUT, "summary.json")) as f:
            cls.summary = json.load(f)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_no_formula_errors(self):
        for ws in self.recalced.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    # openpyxl marks real formula errors with data_type 'e';
                    # literal '#' header text is not an error.
                    self.assertNotEqual(cell.data_type, "e",
                                        "formula error at %s!%s: %r"
                                        % (ws.title, cell.coordinate,
                                           cell.value))

    def test_monthly_matches_ground_truth(self):
        mon = self.recalced["Monthly"]
        monthly = read_csv("monthly_summary.csv")
        for i, row in enumerate(monthly, start=2):
            self.assertAlmostEqual(mon.cell(row=i, column=2).value,
                                   float(row["income"]), places=2)
            self.assertAlmostEqual(mon.cell(row=i, column=3).value,
                                   float(row["expenses"]), places=2)
            self.assertAlmostEqual(mon.cell(row=i, column=4).value,
                                   float(row["net"]), places=2)
            self.assertEqual(mon.cell(row=i, column=5).value,
                             int(row["txn_count"]))
        tot = len(monthly) + 2
        self.assertAlmostEqual(mon.cell(row=tot, column=2).value,
                               self.summary["total_income"], places=2)
        self.assertAlmostEqual(mon.cell(row=tot, column=4).value,
                               self.summary["net"], places=2)

    def test_category_totals_match(self):
        cat = self.recalced["Categories"]
        cats = read_csv("category_summary.csv")
        for i, row in enumerate(cats, start=2):
            self.assertAlmostEqual(cat.cell(row=i, column=3).value,
                                   float(row["total"]), places=2)
            self.assertEqual(cat.cell(row=i, column=4).value,
                             int(row["count"]))
            self.assertAlmostEqual(cat.cell(row=i, column=5).value,
                                   float(row["share_pct"]) / 100, places=4)

    def test_all_checks_pass(self):
        chk = self.recalced["Checks"]
        results = [chk.cell(row=r, column=5).value
                   for r in range(2, 14)]
        self.assertEqual(results, ["PASS"] * 12,
                         "failing checks: %s" % [c for c in results if c != "PASS"])

    def test_dashboard_kpis_for_focus_month(self):
        dash = self.recalced["Dashboard"]
        # Focus month defaults to 2026-10.
        oct_row = [r for r in read_csv("monthly_summary.csv")
                   if r["month"] == "2026-10"][0]
        self.assertAlmostEqual(dash.cell(row=5, column=2).value,
                               float(oct_row["income"]), places=2)
        self.assertAlmostEqual(dash.cell(row=6, column=2).value,
                               float(oct_row["expenses"]), places=2)
        self.assertAlmostEqual(dash.cell(row=7, column=2).value,
                               float(oct_row["net"]), places=2)
        self.assertEqual(dash.cell(row=8, column=2).value,
                         int(oct_row["txn_count"]))
        self.assertAlmostEqual(dash.cell(row=12, column=2).value,
                               self.summary["closing_balance"], places=2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
