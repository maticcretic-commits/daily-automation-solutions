#!/usr/bin/env python3
"""Tests for CashFlowForge: writer primitives, sample determinism, workbook
validity, exact formula wiring, and independent Python recomputation of every
metric the workbook computes via Excel formulas."""

import csv
import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cashflow_forge import (  # noqa: E402
    HORIZON, SHEET_NAMES, TOOL_NAME, SharedStrings, _cell_ref, _col_letter,
    _fmt_num, build_workbook, cell_formula, cell_num, cell_str,
    generate_sample, load_project, main, project_metrics, write_monthly_csv,
)


def _sheet_xml(path, idx):
    with zipfile.ZipFile(path) as zf:
        return zf.read("xl/worksheets/sheet%d.xml" % idx).decode("utf-8")


class TestWriterPrimitives(unittest.TestCase):
    def test_col_letter(self):
        self.assertEqual(_col_letter(1), "A")
        self.assertEqual(_col_letter(26), "Z")
        self.assertEqual(_col_letter(27), "AA")
        self.assertEqual(_col_letter(7), "G")

    def test_cell_ref(self):
        self.assertEqual(_cell_ref(2, 4), "D2")
        self.assertEqual(_cell_ref(26, 7), "G26")

    def test_fmt_num(self):
        self.assertEqual(_fmt_num(4800000.0), "4800000")
        self.assertEqual(_fmt_num(0.12), "0.12")
        self.assertEqual(_fmt_num(24), "24")

    def test_shared_strings_dedupe(self):
        sst = SharedStrings()
        a = sst.add("Inflows")
        b = sst.add("Inflows")
        c = sst.add("Outflows")
        self.assertEqual(a, b)
        self.assertNotEqual(a, c)
        self.assertIn("<t>Inflows</t>", sst.xml().replace(
            ' xml:space="preserve"', ""))


class TestSampleAndLoad(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.csv = os.path.join(self.tmp, "sample.csv")

    def test_generate_deterministic(self):
        generate_sample(self.csv)
        with open(self.csv, "rb") as fh:
            first = fh.read()
        generate_sample(self.csv, seed=999)
        with open(self.csv, "rb") as fh:
            second = fh.read()
        self.assertNotEqual(first, second, "different seed must differ")
        generate_sample(self.csv)  # default seed again
        with open(self.csv, "rb") as fh:
            third = fh.read()
        self.assertEqual(first, third, "same seed must be byte-identical")

    def test_load_shape(self):
        generate_sample(self.csv)
        rows = load_project(self.csv)
        self.assertEqual(len(rows), HORIZON)
        self.assertEqual([r[0] for r in rows], list(range(1, HORIZON + 1)))
        self.assertTrue(all(r[1] >= 0 and r[2] >= 0 for r in rows))

    def test_load_rejects_bad_row_count(self):
        with open(self.csv, "w", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["Month", "Inflows", "Outflows"])
            w.writerow([1, 100, 50])
        with self.assertRaises(ValueError):
            load_project(self.csv)


class TestMetricsMath(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.csv = os.path.join(self.tmp, "sample.csv")
        generate_sample(self.csv)
        self.rows = load_project(self.csv)
        self.inv = 2000000.0
        self.rate = 0.12
        self.m = project_metrics(self.rows, self.inv, self.rate)

    def test_totals_match_manual_sums(self):
        self.assertAlmostEqual(self.m["total_inflows"],
                               sum(r[1] for r in self.rows), places=6)
        self.assertAlmostEqual(self.m["total_outflows"],
                               sum(r[2] for r in self.rows), places=6)

    def test_cumulative_final_is_net_position(self):
        self.assertAlmostEqual(self.m["cumulative"][-1],
                               self.m["net_position"], places=6)
        self.assertAlmostEqual(self.m["net_position"],
                               self.m["total_inflows"] - self.m["total_outflows"]
                               - self.inv, places=6)

    def test_npv_is_sum_of_discounted(self):
        self.assertAlmostEqual(self.m["npv"], sum(self.m["discounted"]),
                               places=6)
        # and the M0 term is exactly -investment (factor 1)
        self.assertAlmostEqual(self.m["discounted"][0], -self.inv, places=6)

    def test_payback_is_first_nonnegative_cumulative(self):
        expected = next((t for t, c in enumerate(self.m["cumulative"])
                         if c >= 0), None)
        self.assertEqual(self.m["payback_months"], expected)
        self.assertIsNotNone(expected, "sample project must pay back")

    def test_irr_zeroes_npv(self):
        self.assertIsNotNone(self.m["irr_monthly"])
        r = self.m["irr_monthly"]
        npv = sum(n / ((1.0 + r) ** t)
                  for t, n in enumerate(self.m["net"]))
        self.assertAlmostEqual(npv, 0.0, delta=1.0)  # within ₹1
        self.assertAlmostEqual(self.m["irr_annual"],
                               (1 + r) ** 12 - 1, places=9)

    def test_bcr_positive(self):
        self.assertGreater(self.m["bcr"], 0)


class TestWorkbook(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.csv = os.path.join(self.tmp, "sample.csv")
        generate_sample(self.csv)
        self.rows = load_project(self.csv)
        self.xlsx = os.path.join(self.tmp, "model.xlsx")
        build_workbook(self.rows, "Test Project", 2000000.0, 0.12, self.xlsx)
        self.audit = os.path.join(self.tmp, "monthly.csv")
        write_monthly_csv(project_metrics(self.rows, 2000000.0, 0.12),
                          self.audit)

    def test_zip_parts_present(self):
        with zipfile.ZipFile(self.xlsx) as zf:
            names = set(zf.namelist())
        for part in ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                     "xl/_rels/workbook.xml.rels", "xl/styles.xml",
                     "xl/sharedStrings.xml", "docProps/core.xml",
                     "xl/drawings/drawing1.xml",
                     "xl/drawings/_rels/drawing1.xml.rels",
                     "xl/worksheets/_rels/sheet4.xml.rels"]:
            self.assertIn(part, names)
        for i in range(1, 5):
            self.assertIn("xl/worksheets/sheet%d.xml" % i, names)
        for i in range(1, 4):
            self.assertIn("xl/charts/chart%d.xml" % i, names)

    def test_sheet_names(self):
        with zipfile.ZipFile(self.xlsx) as zf:
            wb = zf.read("xl/workbook.xml").decode("utf-8")
        for name in SHEET_NAMES:
            self.assertIn('name="%s"' % name, wb)

    def test_assumptions_wiring(self):
        xml = _sheet_xml(self.xlsx, 1)
        self.assertIn("<f>B5/12</f>", xml)          # monthly rate formula
        self.assertIn("<v>2000000</v>", xml)       # investment input
        self.assertIn("<v>0.12</v>", xml)          # annual rate input

    def test_cashflow_formulas(self):
        xml = _sheet_xml(self.xlsx, 2)
        # M0 outflow linked to Assumptions
        self.assertIn("<f>'Assumptions'!$B$4</f>", xml)
        # net CF formula on a mid row, cumulative chain, discount factor
        self.assertIn("<f>B10-C10</f>", xml)
        self.assertIn("<f>E9+D10</f>", xml)
        self.assertIn("<f>1/POWER(1+'Assumptions'!$B$6,A10)</f>", xml)
        self.assertIn("<f>D10*F10</f>", xml)
        # red fill on negative net / cumulative cash flow
        self.assertIn('sqref="D2:D26"', xml)
        self.assertIn('sqref="E2:E26"', xml)
        self.assertIn('operator="lessThan"', xml)

    def test_metrics_formulas(self):
        xml = _sheet_xml(self.xlsx, 3)
        self.assertIn("<f>SUM('Cash Flow'!$B$2:$B$26)</f>", xml)
        self.assertIn("<f>SUM('Cash Flow'!$G$2:$G$26)</f>", xml)  # NPV
        self.assertIn("IRR('Cash Flow'!$D$2:$D$26,0.1)", xml)      # IRR
        self.assertIn("MATCH(TRUE,'Cash Flow'!$E$2:$E$26&gt;=0,0)-1", xml)
        self.assertIn("SUMPRODUCT('Cash Flow'!$B$2:$B$26,"
                      "'Cash Flow'!$F$2:$F$26)", xml)
        # every metric value cell must be a formula cell (nothing hard-coded)
        for r in range(3, 11):
            cell = 'r="B%d"' % r
            self.assertIn(cell, xml)
            seg = xml.split(cell)[1].split("</c>")[0]
            self.assertIn("<f>", seg, "B%d is not a formula cell" % r)

    def test_dashboard_kpi_wiring(self):
        xml = _sheet_xml(self.xlsx, 4)
        for ref in ["'Assumptions'!$B$4", "'Metrics'!$B$3", "'Metrics'!$B$6",
                    "'Metrics'!$B$8", "'Metrics'!$B$9", "'Metrics'!$B$5"]:
            self.assertIn("<f>%s</f>" % ref, xml)

    def test_charts_reference_live_ranges(self):
        with zipfile.ZipFile(self.xlsx) as zf:
            c1 = zf.read("xl/charts/chart1.xml").decode("utf-8")
            c2 = zf.read("xl/charts/chart2.xml").decode("utf-8")
            c3 = zf.read("xl/charts/chart3.xml").decode("utf-8")
        self.assertIn("'Cash Flow'!$D$2:$D$26", c1)   # net CF column chart
        self.assertIn("'Cash Flow'!$E$2:$E$26", c2)   # cumulative line
        self.assertIn("'Metrics'!$B$3:$B$4", c3)      # inflows/outflows pie

    def test_audit_csv_matches_metrics(self):
        m = project_metrics(self.rows, 2000000.0, 0.12)
        with open(self.audit, newline="") as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(len(rows), HORIZON + 1)      # M0..M24
        self.assertEqual(rows[0]["Month"], "0")
        for rec, n, c, d in zip(rows, m["net"], m["cumulative"],
                                m["discounted"]):
            self.assertAlmostEqual(float(rec["NetCashFlow"]), n, places=2)
            self.assertAlmostEqual(float(rec["CumulativeCF"]), c, places=2)
            self.assertAlmostEqual(float(rec["DiscountedCF"]), d, places=2)

    def test_no_tool_branding_leak(self):
        # public repo must not mention the AI assistant
        with zipfile.ZipFile(self.xlsx) as zf:
            blob = b"".join(zf.read(n) for n in zf.namelist())
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
                               "--investment", "1000000", "--rate", "0.10",
                               "--out", xlsx, "--audit-csv", audit]), 0)
        self.assertTrue(os.path.exists(xlsx))
        self.assertTrue(os.path.exists(audit))


if __name__ == "__main__":
    unittest.main(verbosity=2)
