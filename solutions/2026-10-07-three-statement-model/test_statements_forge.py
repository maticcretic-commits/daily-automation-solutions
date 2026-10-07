#!/usr/bin/env python3
"""Tests for statements_forge.py — run with: python3 test_statements_forge.py"""

import csv
import os
import re
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from statements_forge import (
    SharedStrings, _col_letter, _cell_ref, _fmt_num, _sumifs,
    generate_sample, load_accounts, sample_accounts, model,
    build_workbook, write_audit_csv, main, SHEET_NAMES,
)


class TestCore(unittest.TestCase):
    def test_col_letter(self):
        self.assertEqual(_col_letter(1), "A")
        self.assertEqual(_col_letter(26), "Z")
        self.assertEqual(_col_letter(27), "AA")

    def test_cell_ref(self):
        self.assertEqual(_cell_ref(5, 2), "B5")

    def test_fmt_num(self):
        self.assertEqual(_fmt_num(1200000.0), "1200000")
        self.assertEqual(_fmt_num(0.5), "0.50")

    def test_shared_strings_dedupe(self):
        s = SharedStrings()
        self.assertEqual(s.add("x"), 0)
        self.assertEqual(s.add("x"), 0)
        self.assertEqual(s.add("y"), 1)
        self.assertIn("<t xml:space=\"preserve\">x</t>", s.xml())

    def test_sumifs_shape(self):
        f = _sumifs("E", cat="Revenue", subcat="Sales", first=3, last=26)
        self.assertTrue(f.startswith("SUMIFS("))
        self.assertIn("'Trial Balance'!$E$3:$E$26", f)
        self.assertIn('"Revenue"', f)
        self.assertIn('"Sales"', f)
        f2 = _sumifs("E", account="Cash & Bank", opening=True)
        self.assertIn("$D$3:$D$26", f2)
        self.assertIn('"Cash & Bank"', f2)


class TestSample(unittest.TestCase):
    def test_generate_deterministic(self):
        tmp = tempfile.mkdtemp()
        a = os.path.join(tmp, "a.csv")
        b = os.path.join(tmp, "b.csv")
        generate_sample(a)
        generate_sample(b)
        self.assertEqual(open(a, "rb").read(), open(b, "rb").read())

    def test_generate_shape(self):
        tmp = tempfile.mkdtemp()
        p = os.path.join(tmp, "s.csv")
        generate_sample(p)
        with open(p, newline="") as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(len(rows), 24)
        self.assertEqual(list(rows[0].keys()),
                         ["Account", "Category", "Subcategory", "Opening", "Closing"])

    def test_opening_identity(self):
        accs = sample_accounts()
        assets = sum(a["opening"] for a in accs if a["category"] == "Asset")
        le = sum(a["opening"] for a in accs if a["category"] in ("Liability", "Equity"))
        self.assertEqual(assets, le)

    def test_closing_identity_and_plugs(self):
        accs = sample_accounts()
        m = model(accs)
        self.assertEqual(m["balance_check"], 0)
        self.assertEqual(m["tie_out"], 0)
        # retained earnings articulates
        self.assertEqual(m["re_close"], m["re_open"] + m["net_income"])
        # cash walks from the cash-flow statement
        self.assertEqual(m["cash_close"], m["cash_open"] + m["net_change"])

    def test_load_rejects_bad_header(self):
        tmp = tempfile.mkdtemp()
        p = os.path.join(tmp, "bad.csv")
        with open(p, "w") as fh:
            fh.write("nope\n1\n")
        with self.assertRaises(ValueError):
            load_accounts(p)


class TestModel(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.m = model(sample_accounts())

    def test_net_income(self):
        self.assertEqual(self.m["net_income"], 1248000)

    def test_revenue_cogs_gross(self):
        self.assertEqual(self.m["revenue"], 18740000)
        self.assertEqual(self.m["cogs"], 11450000)
        self.assertEqual(self.m["gross"], 7290000)

    def test_ebitda_ebit_ebt(self):
        self.assertEqual(self.m["ebitda"], 2834000)
        self.assertEqual(self.m["ebit"], 2209000)
        self.assertEqual(self.m["ebt"], 1697000)

    def test_balance_sheet_totals(self):
        self.assertEqual(self.m["total_assets"], 15898000)
        self.assertEqual(self.m["total_liab"], 7040000)
        self.assertEqual(self.m["total_equity"], 8858000)

    def test_cash_flow_components(self):
        self.assertEqual(self.m["op_cf"], 1383000)
        self.assertEqual(self.m["inv_cf"], -1200000)
        self.assertEqual(self.m["fin_cf"], -1100000)
        self.assertEqual(self.m["net_change"], -917000)
        self.assertEqual(self.m["cash_open"], 3700000)
        self.assertEqual(self.m["cash_close"], 2783000)

    def test_ratios(self):
        self.assertAlmostEqual(self.m["current_ratio"], 9373000 / 3840000, places=6)
        self.assertAlmostEqual(self.m["net_margin"], 1248000 / 18740000, places=6)
        self.assertAlmostEqual(self.m["roe"], 1248000 / 8858000, places=6)
        self.assertAlmostEqual(self.m["debt_equity"], 7040000 / 8858000, places=6)


class TestWorkbook(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp()
        cls.csvp = os.path.join(cls.tmp, "s.csv")
        cls.xlsx = os.path.join(cls.tmp, "m.xlsx")
        cls.audit = os.path.join(cls.tmp, "a.csv")
        generate_sample(cls.csvp)
        cls.accounts = load_accounts(cls.csvp)
        cls.m = model(cls.accounts)
        build_workbook(cls.accounts, "Test Co", cls.xlsx)
        write_audit_csv(cls.m, cls.audit)
        cls.zf = zipfile.ZipFile(cls.xlsx)
        cls.parts = {n: cls.zf.read(n).decode("utf-8")
                     for n in cls.zf.namelist() if n.endswith(".xml")}

    def test_zip_parts_present(self):
        names = self.zf.namelist()
        for i in range(1, 7):
            self.assertIn("xl/worksheets/sheet%d.xml" % i, names)
        for i in range(1, 4):
            self.assertIn("xl/charts/chart%d.xml" % i, names)
        self.assertIn("xl/drawings/drawing1.xml", names)
        self.assertIn("xl/sharedStrings.xml", names)

    def test_sheet_names(self):
        wb = self.parts["xl/workbook.xml"]
        for name in SHEET_NAMES:
            self.assertIn('name="%s"' % name, wb)

    def test_trial_balance_change_formulas(self):
        tb = self.parts["xl/worksheets/sheet1.xml"]
        self.assertIn("<f>E3-D3</f>", tb)
        self.assertIn("<f>E26-D26</f>", tb)
        self.assertIn("<autoFilter", tb)
        self.assertIn('state="frozen"', tb)

    def test_income_statement_formulas(self):
        isx = self.parts["xl/worksheets/sheet2.xml"]
        self.assertIn("<f>SUM(B5:B6)</f>", isx)          # total revenue
        self.assertIn("<f>B7-B12</f>", isx)             # gross profit
        self.assertIn("<f>B13-(B22-B20)</f>", isx)       # EBITDA
        self.assertIn("<f>B28-B29</f>", isx)             # net income
        self.assertIn("SUMIFS(", isx)

    def test_balance_sheet_wiring(self):
        bsx = self.parts["xl/worksheets/sheet3.xml"]
        self.assertIn("<f>B10+B14</f>", bsx)             # total assets
        self.assertIn("<f>B15-B35</f>", bsx)             # balance check
        self.assertIn("'Income Statement'!B30", bsx)    # NI articulation
        self.assertIn('sqref="B36"', bsx)               # check CF

    def test_cashflow_wiring(self):
        cfx = self.parts["xl/worksheets/sheet4.xml"]
        self.assertIn("<f>SUM(B5:B11)</f>", cfx)         # cash from operations
        self.assertIn("<f>B12+B16+B21</f>", cfx)         # net change
        self.assertIn("<f>B23+B24</f>", cfx)             # cash at end
        self.assertIn("'Balance Sheet'!B6", cfx)         # tie-out
        self.assertIn('sqref="B26"', cfx)

    def test_ratios_wiring(self):
        rax = self.parts["xl/worksheets/sheet5.xml"]
        self.assertIn("'Balance Sheet'!B10/'Balance Sheet'!B22", rax)
        self.assertIn("'Income Statement'!B30/'Income Statement'!B7", rax)

    def test_dashboard_kpi_wiring(self):
        dbx = self.parts["xl/worksheets/sheet6.xml"]
        self.assertIn("'Income Statement'!B7", dbx)
        self.assertIn("'Balance Sheet'!B15", dbx)
        self.assertIn("'Ratios'!B13", dbx)

    def test_charts_reference_live_ranges(self):
        for i in range(1, 4):
            ch = self.parts["xl/charts/chart%d.xml" % i]
            self.assertIn("'Dashboard'!", ch)

    def test_no_hardcoded_statement_totals(self):
        # key total cells must contain <f> formulas, not bare values
        isx = self.parts["xl/worksheets/sheet2.xml"]
        m = re.search(r'<c r="B30"[^>]*>(.*?)</c>', isx)
        self.assertIsNotNone(m)
        self.assertIn("<f>", m.group(1))
        bsx = self.parts["xl/worksheets/sheet3.xml"]
        m = re.search(r'<c r="B15"[^>]*>(.*?)</c>', bsx)
        self.assertIsNotNone(m)
        self.assertIn("<f>", m.group(1))

    def test_audit_csv_matches_model(self):
        with open(self.audit, newline="") as fh:
            rows = {(r["Statement"], r["Line"]): float(r["Amount"])
                    for r in csv.DictReader(fh)}
        self.assertEqual(rows[("Income Statement", "Net Income")], 1248000)
        self.assertEqual(rows[("Balance Sheet", "BALANCE CHECK (must be 0)")], 0)
        self.assertEqual(rows[("Cash Flow", "Tie-out vs BS cash (must be 0)")], 0)
        self.assertEqual(rows[("Balance Sheet", "TOTAL ASSETS")], 15898000)

    def test_no_tool_branding_leak(self):
        blob = b"".join(self.zf.read(n) for n in self.zf.namelist())
        self.assertNotIn(b"Muse", blob)


class TestCLI(unittest.TestCase):
    def test_generate_and_build_end_to_end(self):
        tmp = tempfile.mkdtemp()
        csvp = os.path.join(tmp, "s.csv")
        xlsx = os.path.join(tmp, "m.xlsx")
        audit = os.path.join(tmp, "a.csv")
        self.assertEqual(main(["generate", "--out", csvp]), 0)
        self.assertTrue(os.path.exists(csvp))
        self.assertEqual(main(["build", "--csv", csvp, "--name", "CLI Test",
                               "--out", xlsx, "--audit-csv", audit]), 0)
        self.assertTrue(os.path.exists(xlsx))
        self.assertTrue(os.path.exists(audit))


if __name__ == "__main__":
    unittest.main(verbosity=2)
