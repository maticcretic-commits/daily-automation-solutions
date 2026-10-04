#!/usr/bin/env python3
"""Tests for DashboardForge (dashboard_factory.py)."""

import csv
import hashlib
import os
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dashboard_factory as df

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}

REQUIRED_PARTS = [
    "[Content_Types].xml", "_rels/.rels", "docProps/core.xml",
    "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml",
    "xl/worksheets/sheet3.xml",
    "xl/worksheets/_rels/sheet3.xml.rels",
    "xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels",
    "xl/charts/chart1.xml", "xl/charts/chart2.xml", "xl/charts/chart3.xml",
    "xl/styles.xml", "xl/sharedStrings.xml",
]


def _xml_of(zf, name):
    return ET.fromstring(zf.read(name))


class Fixture(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="dashforge-test-")
        cls.csv_path = os.path.join(cls.tmp, "sample.csv")
        df.generate_sample(cls.csv_path)
        cls.rows = df.load_sales(cls.csv_path)
        cls.xlsx_path = os.path.join(cls.tmp, "dash.xlsx")
        df.build_workbook(cls.rows, cls.xlsx_path)
        with open(cls.xlsx_path, "rb") as fh:
            cls.blob = fh.read()

    def zf(self):
        return zipfile.ZipFile(self.xlsx_path)


# --- input validation -------------------------------------------------------

class TestInputValidation(Fixture):
    def test_sample_is_deterministic(self):
        p2 = os.path.join(self.tmp, "sample2.csv")
        df.generate_sample(p2)
        with open(self.csv_path, "rb") as a, open(p2, "rb") as b:
            self.assertEqual(a.read(), b.read())

    def test_sample_covers_12_months(self):
        months = df.month_starts(self.rows)
        self.assertEqual(len(months), 12)
        self.assertEqual(months[0].month, 1)
        self.assertEqual(months[-1].month, 12)

    def test_revenue_equals_units_times_price(self):
        for r in self.rows:
            self.assertAlmostEqual(r["Revenue"],
                                   round(r["Units"] * r["UnitPrice"], 2))

    def test_missing_column_rejected(self):
        bad = os.path.join(self.tmp, "bad.csv")
        with open(bad, "w", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["OrderID", "Date"])  # missing the rest
            w.writerow(["ORD-1", "2026-01-05"])
        with self.assertRaises(ValueError) as ctx:
            df.load_sales(bad)
        self.assertIn("missing required columns", str(ctx.exception))

    def test_bad_date_rejected(self):
        bad = os.path.join(self.tmp, "baddate.csv")
        with open(self.csv_path) as src, open(bad, "w", newline="") as fh:
            lines = src.readlines()
            fh.write(lines[0])
            fh.write(lines[1].replace(lines[1].split(",")[1], "not-a-date", 1))
        with self.assertRaises(ValueError) as ctx:
            df.load_sales(bad)
        self.assertIn("bad Date", str(ctx.exception))

    def test_bad_units_rejected(self):
        bad = os.path.join(self.tmp, "badunits.csv")
        with open(self.csv_path) as src, open(bad, "w", newline="") as fh:
            lines = src.readlines()
            fh.write(lines[0])
            parts = lines[1].split(",")
            parts[5] = "abc"
            fh.write(",".join(parts))
        with self.assertRaises(ValueError) as ctx:
            df.load_sales(bad)
        self.assertIn("bad Units", str(ctx.exception))

    def test_empty_csv_rejected(self):
        empty = os.path.join(self.tmp, "empty.csv")
        with open(empty, "w") as fh:
            fh.write(",".join(df.REQUIRED_COLUMNS) + "\n")
        with self.assertRaises(ValueError) as ctx:
            df.load_sales(empty)
        self.assertIn("no data rows", str(ctx.exception))


# --- package structure ------------------------------------------------------

class TestPackage(Fixture):
    def test_all_parts_present(self):
        with self.zf() as zf:
            names = set(zf.namelist())
        for part in REQUIRED_PARTS:
            self.assertIn(part, names, "missing part: %s" % part)

    def test_workbook_has_three_sheets(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/workbook.xml")
        sheets = root.findall("main:sheets/main:sheet", NS)
        self.assertEqual([s.get("name") for s in sheets],
                         ["Raw Data", "Monthly Summary", "Dashboard"])

    def test_content_types_covers_parts(self):
        with self.zf() as zf:
            text = zf.read("[Content_Types].xml").decode("utf-8")
        for part in ("sheet1.xml", "sheet2.xml", "sheet3.xml", "chart1.xml",
                     "styles.xml", "sharedStrings.xml", "drawing1.xml"):
            self.assertIn(part, text)

    def test_all_xml_parts_parse(self):
        with self.zf() as zf:
            for name in zf.namelist():
                if name.endswith(".xml"):
                    ET.fromstring(zf.read(name))  # raises if malformed

    def test_styles_has_expected_number_formats(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/styles.xml")
        codes = {n.get("formatCode") for n in
                 root.findall("main:numFmts/main:numFmt", NS)}
        for code in ("$#,##0", "yyyy-mm-dd", "mmm-yy", "0.0%", "#,##0"):
            self.assertIn(code, codes)


# --- sheet 1: raw data ------------------------------------------------------

class TestRawSheet(Fixture):
    def test_row_count(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/worksheets/sheet1.xml")
        rows = root.findall("main:sheetData/main:row", NS)
        self.assertEqual(len(rows), len(self.rows) + 1)

    def test_headers_in_shared_strings(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/sharedStrings.xml")
        texts = {t.text for t in root.findall("main:si/main:t", NS)}
        for h in df.RAW_HEADERS:
            self.assertIn(h, texts)

    def test_autofilter_and_freeze(self):
        with self.zf() as zf:
            text = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        last = "H%d" % (len(self.rows) + 1)
        self.assertIn('<autoFilter ref="A1:%s"/>' % last, text)
        self.assertIn('state="frozen"', text)

    def test_raw_revenue_total_matches_csv(self):
        expected = round(sum(r["Revenue"] for r in self.rows), 2)
        with self.zf() as zf:
            root = _xml_of(zf, "xl/worksheets/sheet1.xml")
        got = 0.0
        for row in root.findall("main:sheetData/main:row", NS)[1:]:
            cells = row.findall("main:c", NS)
            got += float(cells[7].find("main:v", NS).text)  # col H
        self.assertAlmostEqual(got, expected, places=2)


# --- sheet 2: monthly summary -----------------------------------------------

class TestSummarySheet(Fixture):
    def _formulas(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/worksheets/sheet2.xml")
        return [f.text for f in root.findall(
            "main:sheetData/main:row/main:c/main:f", NS)], root

    def test_row_count_header_months_total(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/worksheets/sheet2.xml")
        rows = root.findall("main:sheetData/main:row", NS)
        self.assertEqual(len(rows), 12 + 2)  # header + 12 months + TOTAL

    def test_sumifs_formulas_cover_all_raw_rows(self):
        formulas, _ = self._formulas()
        last = len(self.rows) + 1
        sumifs = [f for f in formulas if f.startswith("SUMIFS(")]
        self.assertEqual(len(sumifs), 24)  # revenue + units x 12 months
        date_range = "'Raw Data'!$B$2:$B$%d" % last
        for f in sumifs:
            self.assertIn(date_range, f)  # criteria range covers all rows
        revenue_f = [f for f in sumifs if "'Raw Data'!$H$" in f]
        units_f = [f for f in sumifs if "'Raw Data'!$F$" in f]
        self.assertEqual(len(revenue_f), 12)
        self.assertEqual(len(units_f), 12)
        countifs = [f for f in formulas if f.startswith("COUNTIFS(")]
        self.assertEqual(len(countifs), 12)
        for f in countifs:
            self.assertIn("$B$2:$B$%d" % last, f)

    def test_revenue_formulas_reference_revenue_column(self):
        formulas, _ = self._formulas()
        last = len(self.rows) + 1
        revenue_f = [f for f in formulas
                     if f.startswith("SUMIFS('Raw Data'!$H$")]
        self.assertEqual(len(revenue_f), 12)
        for f in revenue_f:
            self.assertIn("DATE(2026,", f)
            self.assertIn("EDATE(", f)

    def test_aov_formulas_guard_division_by_zero(self):
        formulas, _ = self._formulas()
        aov = [f for f in formulas if f.startswith("IF(C")]
        self.assertEqual(len(aov), 13)  # 12 months + TOTAL
        for f in aov:
            self.assertTrue(f.startswith("IF(C") and "/C" in f)

    def test_total_row_sums(self):
        formulas, _ = self._formulas()
        totals = [f for f in formulas if f.startswith("SUM(B2:B13)")]
        self.assertEqual(len(totals), 1)
        self.assertIn("SUM(C2:C13)", formulas)
        self.assertIn("SUM(D2:D13)", formulas)

    def test_conditional_formatting_top3(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/worksheets/sheet2.xml")
        rules = root.findall("main:conditionalFormatting/main:cfRule", NS)
        self.assertEqual(len(rules), 1)
        rule = rules[0]
        self.assertEqual(rule.get("type"), "top10")
        self.assertEqual(rule.get("rank"), "3")
        cf = root.find("main:conditionalFormatting", NS)
        self.assertEqual(cf.get("sqref"), "B2:B13")


# --- sheet 3: dashboard -----------------------------------------------------

class TestDashboardSheet(Fixture):
    def test_kpi_formulas_reference_total_row(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/worksheets/sheet3.xml")
        formulas = [f.text for f in root.findall(
            "main:sheetData/main:row/main:c/main:f", NS)]
        for expected in ("'Monthly Summary'!B14", "'Monthly Summary'!C14",
                         "'Monthly Summary'!D14", "'Monthly Summary'!E14"):
            self.assertIn(expected, formulas)

    def test_title_and_kpi_labels_present(self):
        with self.zf() as zf:
            sst = _xml_of(zf, "xl/sharedStrings.xml")
        texts = {t.text for t in sst.findall("main:si/main:t", NS)}
        self.assertIn("Monthly Sales Dashboard", texts)
        for label in ("Total Revenue", "Total Units", "Total Orders",
                      "Avg Order Value"):
            self.assertIn(label, texts)

    def test_title_is_merged(self):
        with self.zf() as zf:
            text = zf.read("xl/worksheets/sheet3.xml").decode("utf-8")
        self.assertIn('<mergeCell ref="A1:H1"/>', text)

    def test_drawing_reference(self):
        with self.zf() as zf:
            text = zf.read("xl/worksheets/sheet3.xml").decode("utf-8")
        self.assertIn('<drawing r:id="rId1"/>', text)


# --- charts -----------------------------------------------------------------

class TestCharts(Fixture):
    def test_three_charts_anchored(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/drawings/drawing1.xml")
        anchors = root.findall("xdr:twoCellAnchor", NS)
        self.assertEqual(len(anchors), 3)

    def test_chart_types_and_titles(self):
        expectations = [
            ("xl/charts/chart1.xml", "c:barChart", "Monthly Revenue"),
            ("xl/charts/chart2.xml", "c:lineChart", "Orders Trend"),
            ("xl/charts/chart3.xml", "c:pieChart", "Revenue by Month"),
        ]
        with self.zf() as zf:
            for part, chart_tag, title in expectations:
                root = _xml_of(zf, part)
                self.assertIsNotNone(root.find("c:chart/c:plotArea/%s" % chart_tag, NS),
                                     "%s missing %s" % (part, chart_tag))
                titles = [t.text for t in
                          root.findall("c:chart/c:title//a:t", NS)]
                self.assertIn(title, titles)

    def test_chart_series_reference_summary_ranges(self):
        with self.zf() as zf:
            root = _xml_of(zf, "xl/charts/chart1.xml")
        refs = [f.text for f in
                root.findall("c:chart/c:plotArea/c:barChart/c:ser//c:f", NS)]
        joined = " ".join(refs)
        self.assertIn("'Monthly Summary'!$A$2:$A$13", joined)  # categories
        self.assertIn("'Monthly Summary'!$B$2:$B$13", joined)  # values


# --- determinism --------------------------------------------------------------

class TestDeterminism(Fixture):
    def test_byte_identical_rebuild(self):
        p2 = os.path.join(self.tmp, "dash2.xlsx")
        df.build_workbook(self.rows, p2)
        with open(p2, "rb") as fh:
            blob2 = fh.read()
        self.assertEqual(hashlib.sha256(self.blob).hexdigest(),
                         hashlib.sha256(blob2).hexdigest())

    def test_cli_end_to_end(self):
        out = os.path.join(self.tmp, "cli.xlsx")
        rc = df.main(["--input", self.csv_path, "--output", out])
        self.assertEqual(rc, 0)
        self.assertTrue(os.path.getsize(out) > 10000)
        with zipfile.ZipFile(out) as zf:
            self.assertIn("xl/charts/chart1.xml", zf.namelist())


if __name__ == "__main__":
    unittest.main(verbosity=2)
