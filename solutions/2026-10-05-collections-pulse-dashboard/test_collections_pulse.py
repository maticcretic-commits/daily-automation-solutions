#!/usr/bin/env python3
"""Tests for collections_pulse.py — stdlib unittest, no dependencies."""
import csv
import datetime
import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import collections_pulse as cp

REF = datetime.date(2026, 10, 5)

MINI_ROWS = [
    # LoanID, Customer, Phone, Branch, Agent, DueDate, EMI, Paid
    ["LN-1", "A One", "9811111111", "Patna", "R. Verma", "2026-10-05", 5000, 5000],  # current, paid
    ["LN-2", "B Two", "9822222222", "Patna", "R. Verma", "2026-09-20", 8000, 0],     # dpd 15, 1-30
    ["LN-3", "C Three", "9833333333", "Gaya", "A. Singh", "2026-08-20", 10000, 4000],# dpd 46, 31-60
    ["LN-4", "D Four", "9844444444", "Gaya", "A. Singh", "2026-06-01", 6000, 0],     # dpd 126, 90+
    ["LN-5", "E Five", "9855555555", "Patna", "S. Gupta", "2026-09-04", 4000, 1000], # dpd 31, 31-60
]


def write_mini_csv(path, rows=MINI_ROWS):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(cp.REQUIRED_COLUMNS)
        w.writerows(rows)


class TestBucketing(unittest.TestCase):
    def test_boundaries(self):
        self.assertEqual(cp.bucket_of(-5), "Current")
        self.assertEqual(cp.bucket_of(0), "Current")
        self.assertEqual(cp.bucket_of(1), "1-30")
        self.assertEqual(cp.bucket_of(30), "1-30")
        self.assertEqual(cp.bucket_of(31), "31-60")
        self.assertEqual(cp.bucket_of(60), "31-60")
        self.assertEqual(cp.bucket_of(61), "61-90")
        self.assertEqual(cp.bucket_of(90), "61-90")
        self.assertEqual(cp.bucket_of(91), "90+")
        self.assertEqual(cp.bucket_of(400), "90+")


class TestPriority(unittest.TestCase):
    def test_longer_overdue_ranks_higher(self):
        self.assertGreater(cp.priority_score(1000, 60), cp.priority_score(1000, 10))

    def test_bigger_due_ranks_higher(self):
        self.assertGreater(cp.priority_score(9000, 10), cp.priority_score(1000, 10))

    def test_formula(self):
        self.assertAlmostEqual(cp.priority_score(6000, 30), 12000.0)


class TestLoadLedger(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.csv = os.path.join(self.tmp, "ledger.csv")
        write_mini_csv(self.csv)
        self.loans = cp.load_ledger(self.csv, REF)

    def test_counts(self):
        self.assertEqual(len(self.loans), 5)

    def test_derived_fields(self):
        by_id = {l["loan_id"]: l for l in self.loans}
        self.assertEqual(by_id["LN-2"]["due"], 8000)
        self.assertEqual(by_id["LN-2"]["dpd"], 15)
        self.assertEqual(by_id["LN-2"]["bucket"], "1-30")
        self.assertEqual(by_id["LN-3"]["dpd"], 46)
        self.assertEqual(by_id["LN-3"]["bucket"], "31-60")
        self.assertEqual(by_id["LN-4"]["bucket"], "90+")
        self.assertEqual(by_id["LN-1"]["due"], 0)
        self.assertEqual(by_id["LN-1"]["dpd"], 0)  # paid in full -> not overdue
        self.assertEqual(by_id["LN-1"]["bucket"], "Current")

    def test_missing_column(self):
        bad = os.path.join(self.tmp, "bad.csv")
        with open(bad, "w", newline="") as fh:
            w = csv.writer(fh); w.writerow(["LoanID", "Customer"])
            w.writerow(["x", "y"])
        with self.assertRaises(ValueError) as cm:
            cp.load_ledger(bad, REF)
        self.assertIn("missing required column", str(cm.exception))

    def test_bad_date_names_row(self):
        bad = os.path.join(self.tmp, "baddate.csv")
        write_mini_csv(bad, rows=[MINI_ROWS[0][:5] + ["not-a-date", 5000, 0]])
        with self.assertRaises(ValueError) as cm:
            cp.load_ledger(bad, REF)
        self.assertIn("row 2", str(cm.exception))
        self.assertIn("DueDate", str(cm.exception))

    def test_bad_number_names_row(self):
        bad = os.path.join(self.tmp, "badnum.csv")
        write_mini_csv(bad, rows=[MINI_ROWS[0][:6] + ["abc", 0]])
        with self.assertRaises(ValueError) as cm:
            cp.load_ledger(bad, REF)
        self.assertIn("row 2", str(cm.exception))

    def test_negative_emi_rejected(self):
        bad = os.path.join(self.tmp, "neg.csv")
        write_mini_csv(bad, rows=[MINI_ROWS[0][:6] + [-5000, 0]])
        with self.assertRaises(ValueError):
            cp.load_ledger(bad, REF)

    def test_paid_over_emi_rejected(self):
        bad = os.path.join(self.tmp, "over.csv")
        write_mini_csv(bad, rows=[MINI_ROWS[0][:6] + [5000, 9000]])
        with self.assertRaises(ValueError):
            cp.load_ledger(bad, REF)

    def test_empty_file_rejected(self):
        bad = os.path.join(self.tmp, "empty.csv")
        open(bad, "w").close()
        with self.assertRaises(ValueError):
            cp.load_ledger(bad, REF)

    def test_branch_order_preserved(self):
        self.assertEqual(cp.branch_names(self.loans), ["Patna", "Gaya"])


class TestReminders(unittest.TestCase):
    def test_overdue_message_contents(self):
        msg = cp.reminder_message({"customer": "A One", "due": 8000,
                                   "dpd": 15, "loan_id": "LN-2",
                                   "branch": "Patna", "agent": "R. Verma"})
        for bit in ("A One", "8,000", "LN-2", "15", "Patna", "R. Verma"):
            self.assertIn(bit, msg)

    def test_current_message_no_pending(self):
        msg = cp.reminder_message({"customer": "A One", "due": 5000,
                                   "dpd": 0, "loan_id": "LN-1",
                                   "branch": "Patna", "agent": "R. Verma"})
        self.assertNotIn("pending", msg)


class TestWorkbook(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.csv = os.path.join(self.tmp, "ledger.csv")
        write_mini_csv(self.csv)
        self.loans = cp.load_ledger(self.csv, REF)
        self.queue = sorted([l for l in self.loans if l["due"] > 0],
                            key=lambda l: l["priority"], reverse=True)
        self.xlsx = os.path.join(self.tmp, "pulse.xlsx")
        cp.build_workbook(self.loans, self.queue, self.xlsx)
        self.zf = zipfile.ZipFile(self.xlsx)
        self.names = self.zf.namelist()

    def test_part_count(self):
        self.assertEqual(len(self.names), 18)

    def test_sheet_names(self):
        wb = self.zf.read("xl/workbook.xml").decode("utf-8")
        for name in cp.SHEET_NAMES:
            self.assertIn('name="%s"' % name, wb)

    def test_five_sheets_present(self):
        for i in range(1, 6):
            self.assertIn("xl/worksheets/sheet%d.xml" % i, self.names)

    def test_three_charts_and_drawing(self):
        for i in range(1, 4):
            self.assertIn("xl/charts/chart%d.xml" % i, self.names)
        self.assertIn("xl/drawings/drawing1.xml", self.names)
        s4 = self.zf.read("xl/worksheets/sheet4.xml").decode("utf-8")
        self.assertIn("<drawing", s4)

    def test_raw_sheet_row_count_and_autofilter(self):
        raw = self.zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn('ref="A1:K6"', raw)  # header + 5 loans
        self.assertIn('state="frozen"', raw)
        for h in ("LoanID", "DPD", "Bucket"):
            self.assertIn(h, self.zf.read("xl/sharedStrings.xml").decode("utf-8"))

    def test_dpd_summary_formulas_cover_all_raw_rows(self):
        s2 = self.zf.read("xl/worksheets/sheet2.xml").decode("utf-8")
        self.assertIn("'Raw Ledger'!$K$2:$K$6", s2)   # bucket range
        self.assertIn("'Raw Ledger'!$I$2:$I$6", s2)   # due range
        for b in ("1-30", "31-60", "90+"):
            self.assertIn('&quot;%s&quot;' % b, s2)
        self.assertIn("AVERAGEIF", s2)
        self.assertIn("TOTAL", self.zf.read("xl/sharedStrings.xml").decode("utf-8"))
        # red conditional formatting targets the 90+ outstanding cell C6
        self.assertIn('sqref="C6"', s2)

    def test_branch_summary_efficiency_guards_division_by_zero(self):
        s3 = self.zf.read("xl/worksheets/sheet3.xml").decode("utf-8")
        self.assertIn("IF(C2=0,0,D2/C2)", s3)
        self.assertIn("Patna", self.zf.read("xl/sharedStrings.xml").decode("utf-8"))

    def test_dashboard_kpis_wired_to_summary_totals(self):
        s4 = self.zf.read("xl/worksheets/sheet4.xml").decode("utf-8")
        # DPD Summary TOTAL row is 7 (5 buckets + header + total)
        self.assertIn("'DPD Summary'!C7", s4)
        self.assertIn("'Branch Summary'!F4", s4)  # 2 branches -> total row 4
        self.assertIn("PAR&gt;30",
                      self.zf.read("xl/sharedStrings.xml").decode("utf-8"))

    def test_queue_ranked_and_priority_formula(self):
        s5 = self.zf.read("xl/worksheets/sheet5.xml").decode("utf-8")
        # ranking logic: highest priority first —
        # LN-4 (6000 due, dpd 126) > LN-3 (6000 due, dpd 46) >
        # LN-2 (8000 due, dpd 15) > LN-5 (3000 due, dpd 31)
        self.assertEqual([l["loan_id"] for l in self.queue],
                         ["LN-4", "LN-3", "LN-2", "LN-5"])
        self.assertIn("G2*(1+H2/30)", s5)
        sst = self.zf.read("xl/sharedStrings.xml").decode("utf-8")
        self.assertIn("Reminder Template", sst)
        # queue has 4 loans with due > 0 (LN-1 fully paid excluded)
        self.assertEqual(len(self.queue), 4)

    def test_charts_reference_summary_ranges(self):
        c1 = self.zf.read("xl/charts/chart1.xml").decode("utf-8")
        self.assertIn("DPD Summary", c1)
        c3 = self.zf.read("xl/charts/chart3.xml").decode("utf-8")
        self.assertIn("Branch Summary", c3)

    def test_deterministic_rebuild_byte_identical(self):
        again = os.path.join(self.tmp, "pulse2.xlsx")
        cp.build_workbook(self.loans, self.queue, again)
        with open(self.xlsx, "rb") as a, open(again, "rb") as b:
            self.assertEqual(a.read(), b.read())

    def test_workbook_opens_without_errors(self):
        # zip integrity of every part
        for name in self.names:
            self.assertTrue(len(self.zf.read(name)) > 0, name)


class TestRemindersCSV(unittest.TestCase):
    def test_csv_columns_and_message(self):
        tmp = tempfile.mkdtemp()
        csvp = os.path.join(tmp, "ledger.csv")
        write_mini_csv(csvp)
        loans = cp.load_ledger(csvp, REF)
        queue = sorted([l for l in loans if l["due"] > 0],
                       key=lambda l: l["priority"], reverse=True)
        out = os.path.join(tmp, "rem.csv")
        cp.write_reminders(queue, out)
        with open(out, encoding="utf-8") as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[0]["LoanID"], "LN-4")
        self.assertIn("126", rows[0]["Message"])


class TestSampleGenerator(unittest.TestCase):
    def test_deterministic(self):
        tmp = tempfile.mkdtemp()
        a, b = os.path.join(tmp, "a.csv"), os.path.join(tmp, "b.csv")
        cp.generate_sample(a, ref_date=REF)
        cp.generate_sample(b, ref_date=REF)
        with open(a, "rb") as fa, open(b, "rb") as fb:
            self.assertEqual(fa.read(), fb.read())

    def test_sample_loads_clean(self):
        tmp = tempfile.mkdtemp()
        p = os.path.join(tmp, "s.csv")
        cp.generate_sample(p, ref_date=REF)
        loans = cp.load_ledger(p, REF)
        self.assertEqual(len(loans), 180)
        buckets = {l["bucket"] for l in loans}
        self.assertTrue({"Current", "90+"} <= buckets)


if __name__ == "__main__":
    unittest.main(verbosity=1)
