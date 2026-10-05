#!/usr/bin/env python3
"""CollectionsPulse — NBFC / lender collections dashboard + follow-up queue factory.

Reads a loan-ledger CSV and generates:
  1. A real .xlsx workbook (genuine Office Open XML, stdlib only — no openpyxl,
     no pandas, no Excel needed to build it) with five sheets:
       - Raw Ledger: every loan + derived Due / DPD / Bucket columns,
         AutoFilter, frozen panes, currency + date formats.
       - DPD Summary: pivot-style bucket table (Current / 1-30 / 31-60 / 61-90 /
         90+) built from live Excel SUMIF/COUNTIF/AVERAGEIF formulas over the
         raw rows, with a TOTAL row and red conditional formatting on the 90+
         bucket's outstanding.
       - Branch Summary: per-branch accounts, EMI due, collected, outstanding
         and collection-efficiency % (SUMIF formulas, division-by-zero guarded),
         green/red conditional formatting on best/worst efficiency.
       - Dashboard: merged title, 4 formula-wired KPIs (Total Outstanding,
         Accounts in Arrears, Collection Efficiency %, PAR>30 %) and 3 native
         Excel charts (column: outstanding by DPD bucket; pie: outstanding
         share; line: branch collection efficiency).
       - Action Queue: today's priority-ranked follow-up list — the N accounts
         with the highest priority score Due*(1+DPD/30), with a formula
         priority column and a ready-to-send bilingual reminder template.
  2. A followup_reminders.csv: LoanID, Customer, Phone, Branch, Agent, Due,
     DPD and a pre-filled Hinglish/English WhatsApp/SMS reminder message —
     the collections team's call list for the day.

Because the summaries use live Excel formulas (not hard-coded values), the
workbook recalculates automatically if the team edits the raw ledger.

Input CSV columns:
  LoanID, Customer, Phone, Branch, Agent, DueDate (YYYY-MM-DD), EMI, Paid

Usage:
  python collections_pulse.py --generate-sample sample_ledger.csv
  python collections_pulse.py --input sample_ledger.csv --output CollectionsPulse.xlsx \\
      --reminders followup_reminders.csv [--ref-date 2026-10-05] [--top 50]
  python test_collections_pulse.py
"""

import csv
import datetime
import os
import random
import sys
import zipfile
from xml.sax.saxutils import escape as _xml_escape

TOOL_NAME = "CollectionsPulse"
SAMPLE_SEED = 20261005

BUCKETS = ["Current", "1-30", "31-60", "61-90", "90+"]

# Raw Ledger column layout (1-based): A LoanID, B Customer, C Phone, D Branch,
# E Agent, F DueDate, G EMI, H Paid, I Due, J DPD, K Bucket
COL_DUE, COL_DPD, COL_BUCKET = 9, 10, 11
COL_BRANCH, COL_EMI, COL_PAID = 4, 7, 8


# ---------------------------------------------------------------------------
# Small XML helpers
# ---------------------------------------------------------------------------

def _esc(text):
    return _xml_escape(str(text), {'"': "&quot;"})


def _col_letter(n):
    """1 -> 'A', 27 -> 'AA'."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _cell_ref(row, col):
    return "%s%d" % (_col_letter(col), row)


def _excel_serial(day):
    """Excel serial date number for a datetime.date (1900 date system)."""
    return (day - datetime.date(1899, 12, 30)).days


class SharedStrings:
    def __init__(self):
        self._index = {}
        self._items = []

    def add(self, text):
        text = str(text)
        if text not in self._index:
            self._index[text] = len(self._items)
            self._items.append(text)
        return self._index[text]

    def xml(self):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 'count="%d" uniqueCount="%d">' % (len(self._items), len(self._items))]
        for item in self._items:
            parts.append('<si><t xml:space="preserve">%s</t></si>' % _esc(item))
        parts.append('</sst>')
        return "".join(parts)


# ---------------------------------------------------------------------------
# Cell XML builders  (style ids are defined in styles_xml() below)
#
#  0 default · 1 header · 2 currency (₹) · 3 date · 4 percent · 5 integer ·
#  6 title · 7 light-blue header · 8 light-blue currency · 9 mmm-yy ·
#  10 bordered text
# ---------------------------------------------------------------------------

def _cell_open(ref, kind, style):
    if kind == "s":
        return '<c r="%s" t="s" s="%d">' % (ref, style)
    if kind == "f":
        return '<c r="%s" s="%d">' % (ref, style)
    return '<c r="%s" s="%d">' % (ref, style)


def cell_str(row, col, sst_idx, style=0):
    return (_cell_open(_cell_ref(row, col), "s", style) +
            '<v>%d</v></c>' % sst_idx)


def cell_num(row, col, value, style=0):
    return (_cell_open(_cell_ref(row, col), "n", style) +
            '<v>%s</v></c>' % _fmt_num(value))


def cell_formula(row, col, formula, style=0):
    # formula: Excel formula WITHOUT the leading '='
    return (_cell_open(_cell_ref(row, col), "f", style) +
            '<f>%s</f><v>0</v></c>' % _esc(formula))


def _fmt_num(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, float):
        return repr(round(value, 10))
    return str(value)

# ---------------------------------------------------------------------------
# styles.xml
# ---------------------------------------------------------------------------

def styles_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="5">
<numFmt numFmtId="164" formatCode="#,##0"/>
<numFmt numFmtId="165" formatCode="\u20b9#,##0"/>
<numFmt numFmtId="166" formatCode="0.0%"/>
<numFmt numFmtId="167" formatCode="yyyy-mm-dd"/>
<numFmt numFmtId="168" formatCode="mmm-yy"/>
</numFmts>
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="16"/><color rgb="FF2E75B6"/><name val="Calibri"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="165" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="168" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>
</cellXfs>
<dxfs count="2">
<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill></dxf>
<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill></dxf>
</dxfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


# ---------------------------------------------------------------------------
# Domain logic: DPD bucketing, priority scoring, reminder templates
# ---------------------------------------------------------------------------

def bucket_of(dpd):
    """Days-past-due -> arrears bucket label."""
    if dpd <= 0:
        return "Current"
    if dpd <= 30:
        return "1-30"
    if dpd <= 60:
        return "31-60"
    if dpd <= 90:
        return "61-90"
    return "90+"


def priority_score(due, dpd):
    """Follow-up priority: outstanding amount scaled by how long it is overdue."""
    return due * (1.0 + max(dpd, 0) / 30.0)


def reminder_message(loan):
    """Bilingual (Hinglish-first) reminder text for one loan dict."""
    due_fmt = "₹{:,.0f}".format(loan["due"])
    if loan["dpd"] <= 0:
        return ("Namaste %s, aapki EMI %s (Loan %s) jald due hai. "
                "Samay par bhugtan ke liye dhanyavaad — %s branch, %s." %
                (loan["customer"], due_fmt, loan["loan_id"],
                 loan["branch"], loan["agent"]))
    return ("Namaste %s, aapki EMI %s (Loan %s) %d din se pending hai. "
            "Kripya jald se jald bhugtan karein — %s branch, %s." %
            (loan["customer"], due_fmt, loan["loan_id"], loan["dpd"],
             loan["branch"], loan["agent"]))


# ---------------------------------------------------------------------------
# Input: load + validate the loan ledger CSV
# ---------------------------------------------------------------------------

REQUIRED_COLUMNS = ["LoanID", "Customer", "Phone", "Branch", "Agent",
                    "DueDate", "EMI", "Paid"]


def _fail(row_no, message):
    raise ValueError("row %d: %s" % (row_no, message))


def load_ledger(csv_path, ref_date):
    """Load and validate the ledger; derive Due, DPD and Bucket per loan."""
    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        if reader.fieldnames is None:
            raise ValueError("empty file: no header row found")
        missing = [c for c in REQUIRED_COLUMNS if c not in reader.fieldnames]
        if missing:
            raise ValueError("missing required column(s): %s" % ", ".join(missing))
        loans = []
        for row_no, rec in enumerate(reader, start=2):
            loan_id = (rec.get("LoanID") or "").strip()
            if not loan_id:
                continue  # tolerate a trailing blank line, skip nameless rows
            try:
                due_date = datetime.date.fromisoformat(
                    (rec.get("DueDate") or "").strip())
            except ValueError:
                _fail(row_no, "bad DueDate %r (want YYYY-MM-DD)"
                      % rec.get("DueDate"))
            for col in ("EMI", "Paid"):
                raw = (rec.get(col) or "").strip().replace(",", "")
                try:
                    val = float(raw)
                except ValueError:
                    _fail(row_no, "bad %s %r (want a number)" % (col, rec.get(col)))
                if val < 0:
                    _fail(row_no, "%s must not be negative" % col)
                rec[col] = val
            if rec["Paid"] > rec["EMI"]:
                _fail(row_no, "Paid (%.2f) exceeds EMI (%.2f)"
                      % (rec["Paid"], rec["EMI"]))
            due = rec["EMI"] - rec["Paid"]
            dpd = (ref_date - due_date).days if due > 0 else 0
            loans.append({
                "loan_id": loan_id,
                "customer": (rec.get("Customer") or "").strip(),
                "phone": (rec.get("Phone") or "").strip(),
                "branch": (rec.get("Branch") or "").strip(),
                "agent": (rec.get("Agent") or "").strip(),
                "due_date": due_date,
                "emi": rec["EMI"],
                "paid": rec["Paid"],
                "due": due,
                "dpd": dpd,
                "bucket": bucket_of(dpd),
                "priority": priority_score(due, dpd),
            })
    if not loans:
        raise ValueError("no loan rows found")
    return loans


def branch_names(loans):
    seen, order = set(), []
    for loan in loans:
        if loan["branch"] not in seen:
            seen.add(loan["branch"])
            order.append(loan["branch"])
    return order


# ---------------------------------------------------------------------------
# Deterministic sample ledger generator
# ---------------------------------------------------------------------------

def generate_sample(csv_path, seed=SAMPLE_SEED, ref_date=None):
    ref_date = ref_date or datetime.date(2026, 10, 5)
    rng = random.Random(seed)
    branches = ["Patna", "Gaya", "Muzaffarpur", "Bhagalpur"]
    agents = {"Patna": ["R. Verma", "S. Gupta"], "Gaya": ["A. Singh"],
              "Muzaffarpur": ["P. Yadav", "K. Mishra"], "Bhagalpur": ["D. Kumar"]}
    first = ["Amit", "Priya", "Rahul", "Sunita", "Vikas", "Neha", "Suresh",
             "Kavita", "Manoj", "Pooja", "Ramesh", "Anita", "Sanjay", "Meera",
             "Deepak", "Shalini", "Arun", "Geeta", "Naresh", "Ritu"]
    last = ["Kumar", "Singh", "Yadav", "Sharma", "Verma", "Gupta", "Mishra",
            "Pandit", "Rai", "Chaudhary"]
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(REQUIRED_COLUMNS)
        for i in range(1, 181):
            branch = branches[(i - 1) % len(branches)]
            emi = rng.choice([2500, 3000, 4000, 5000, 6000, 8000, 10000, 12000, 15000])
            roll = rng.random()
            if roll < 0.45:        # fully paid, on time
                paid, back = emi, rng.randint(0, 10)
            elif roll < 0.65:      # recently due, unpaid
                paid, back = 0.0, rng.randint(1, 25)
            elif roll < 0.85:      # partial payment
                paid, back = round(emi * rng.choice([0.25, 0.5, 0.75]), 2), rng.randint(5, 70)
            else:                  # long overdue, nothing paid
                paid, back = 0.0, rng.randint(61, 150)
            w.writerow([
                "LN-%05d" % i,
                "%s %s" % (rng.choice(first), rng.choice(last)),
                "98%08d" % rng.randint(10000000, 99999999),
                branch,
                rng.choice(agents[branch]),
                (ref_date - datetime.timedelta(days=back)).isoformat(),
                emi, paid,
            ])
    return csv_path

# ---------------------------------------------------------------------------
# Package-level parts (5 sheets: Raw Ledger, DPD Summary, Branch Summary,
# Dashboard, Action Queue; 3 charts live on the Dashboard)
# ---------------------------------------------------------------------------

SHEET_NAMES = ["Raw Ledger", "DPD Summary", "Branch Summary", "Dashboard",
               "Action Queue"]


def content_types_xml():
    sheets = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.worksheet+xml"/>' % i for i in range(1, 6))
    charts = "".join(
        '<Override PartName="/xl/charts/chart%d.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'drawingml.chart+xml"/>' % i for i in range(1, 4))
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
%s
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
%s
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>""" % (sheets, charts)


def root_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def workbook_xml():
    sheets = "".join(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (name, i, i)
        for i, name in enumerate(SHEET_NAMES, start=1))
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
%s
</sheets>
<calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>""" % sheets


def workbook_rels_xml():
    rels = "".join(
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet%d.xml"/>' % (i, i) for i in range(1, 6))
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
%s
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>""" % rels


def sheet4_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"""


def drawing_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
</Relationships>"""


def core_props_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>%s</dc:creator>
<dc:title>NBFC Collections Dashboard</dc:title>
<cp:keywords>collections,nbfc,dpd,excel dashboard,automation</cp:keywords>
</cp:coreProperties>""" % TOOL_NAME


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def _cols_xml(widths):
    cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>'
                   % (i + 1, i + 1, w) for i, w in enumerate(widths))
    return "<cols>%s</cols>" % cols


def row(cells):
    return "<row>%s</row>" % "".join(cells)


def header_row(sst, texts, style=1):
    return row([cell_str(1, i + 1, sst.add(t), style)
                for i, t in enumerate(texts)])


# -- Sheet 1: Raw Ledger ------------------------------------------------------

RAW_HEADERS = ["LoanID", "Customer", "Phone", "Branch", "Agent", "DueDate",
               "EMI", "Paid", "Due", "DPD", "Bucket"]
RAW_WIDTHS = [12, 20, 14, 14, 14, 12, 12, 12, 12, 8, 10]


def build_raw_rows(sst, loans):
    rows = [header_row(sst, RAW_HEADERS)]
    for r, loan in enumerate(loans, start=2):
        rows.append(row([
            cell_str(r, 1, sst.add(loan["loan_id"]), 10),
            cell_str(r, 2, sst.add(loan["customer"]), 10),
            cell_str(r, 3, sst.add(loan["phone"]), 10),
            cell_str(r, 4, sst.add(loan["branch"]), 10),
            cell_str(r, 5, sst.add(loan["agent"]), 10),
            cell_num(r, 6, _excel_serial(loan["due_date"]), 3),
            cell_num(r, 7, loan["emi"], 2),
            cell_num(r, 8, loan["paid"], 2),
            cell_num(r, 9, loan["due"], 2),
            cell_num(r, 10, loan["dpd"], 5),
            cell_str(r, 11, sst.add(loan["bucket"]), 10),
        ]))
    return rows


def raw_sheet_xml(sst, loans):
    n = len(loans)
    last = "K%d" % (n + 1)
    body = "".join(build_raw_rows(sst, loans))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '%s<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews>'
            '<autoFilter ref="A1:%s"/>'
            '<sheetData>%s</sheetData>'
            '</worksheet>' % (_cols_xml(RAW_WIDTHS), last, body))


# -- Sheet 2: DPD Summary -----------------------------------------------------

# Columns: A Bucket, B Accounts, C Outstanding, D Avg DPD
def build_dpd_summary_sheet(sst, n_raw_rows):
    raw_k = "'Raw Ledger'!$K$2:$K$%d" % (n_raw_rows + 1)
    raw_i = "'Raw Ledger'!$I$2:$I$%d" % (n_raw_rows + 1)
    raw_j = "'Raw Ledger'!$J$2:$J$%d" % (n_raw_rows + 1)
    rows = [header_row(sst, ["DPD Bucket", "Accounts", "Outstanding (₹)", "Avg DPD"])]
    for i, bucket in enumerate(BUCKETS, start=2):
        crit = '"%s"' % bucket
        rows.append(row([
            cell_str(i, 1, sst.add(bucket), 7),
            cell_formula(i, 2, 'COUNTIF(%s,%s)' % (raw_k, crit), 5),
            cell_formula(i, 3, 'SUMIF(%s,%s,%s)' % (raw_k, crit, raw_i), 2),
            cell_formula(i, 4, 'IF(COUNTIF(%s,%s)=0,0,AVERAGEIF(%s,%s,%s))'
                         % (raw_k, crit, raw_k, crit, raw_j), 5),
        ]))
    total = len(BUCKETS) + 2
    rows.append(row([
        cell_str(total, 1, sst.add("TOTAL"), 8),
        cell_formula(total, 2, "SUM(B2:B%d)" % (total - 1), 5),
        cell_formula(total, 3, "SUM(C2:C%d)" % (total - 1), 8),
        cell_str(total, 4, sst.add("—"), 10),
    ]))
    cf = ('<conditionalFormatting sqref="C6"><cfRule type="cellIs" dxfId="1" '
          'priority="1" operator="greaterThan"><formula><v>0</v></formula>'
          '</cfRule></conditionalFormatting>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '%s<sheetData>%s</sheetData>%s</worksheet>'
            % (_cols_xml([14, 12, 18, 12]), "".join(rows), cf))

# -- Sheet 3: Branch Summary ---------------------------------------------------

# Columns: A Branch, B Accounts, C Total EMI, D Collected, E Outstanding,
# F Collection Efficiency %
def build_branch_summary_sheet(sst, loans, n_raw_rows):
    branches = branch_names(loans)
    raw_d = "'Raw Ledger'!$D$2:$D$%d" % (n_raw_rows + 1)
    raw_g = "'Raw Ledger'!$G$2:$G$%d" % (n_raw_rows + 1)
    raw_h = "'Raw Ledger'!$H$2:$H$%d" % (n_raw_rows + 1)
    raw_i = "'Raw Ledger'!$I$2:$I$%d" % (n_raw_rows + 1)
    rows = [header_row(sst, ["Branch", "Accounts", "Total EMI (₹)",
                             "Collected (₹)", "Outstanding (₹)",
                             "Collection Efficiency"])]
    for i, branch in enumerate(branches, start=2):
        crit = '"%s"' % branch
        rows.append(row([
            cell_str(i, 1, sst.add(branch), 7),
            cell_formula(i, 2, 'COUNTIF(%s,%s)' % (raw_d, crit), 5),
            cell_formula(i, 3, 'SUMIF(%s,%s,%s)' % (raw_d, crit, raw_g), 2),
            cell_formula(i, 4, 'SUMIF(%s,%s,%s)' % (raw_d, crit, raw_h), 2),
            cell_formula(i, 5, 'SUMIF(%s,%s,%s)' % (raw_d, crit, raw_i), 2),
            cell_formula(i, 6, 'IF(C%d=0,0,D%d/C%d)' % (i, i, i), 4),
        ]))
    total = len(branches) + 2
    rows.append(row([
        cell_str(total, 1, sst.add("TOTAL"), 8),
        cell_formula(total, 2, "SUM(B2:B%d)" % (total - 1), 5),
        cell_formula(total, 3, "SUM(C2:C%d)" % (total - 1), 2),
        cell_formula(total, 4, "SUM(D2:D%d)" % (total - 1), 2),
        cell_formula(total, 5, "SUM(E2:E%d)" % (total - 1), 8),
        cell_formula(total, 6, 'IF(C%d=0,0,D%d/C%d)' % (total, total, total), 4),
    ]))
    data_last = total - 1
    cf = ('<conditionalFormatting sqref="F2:F%d">'
          '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">'
          '<formula><v>0.85</v></formula></cfRule>'
          '<cfRule type="cellIs" dxfId="1" priority="2" operator="lessThan">'
          '<formula><v>0.6</v></formula></cfRule>'
          '</conditionalFormatting>' % data_last)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '%s<sheetData>%s</sheetData>%s</worksheet>'
            % (_cols_xml([14, 10, 16, 16, 16, 22]), "".join(rows), cf))


# -- Charts (live on the Dashboard sheet) -------------------------------------

def _chart_title_xml(title):
    return ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:r><a:t>%s</a:t></a:r></a:p></c:rich></c:tx>'
            '<c:layout/></c:title>' % _esc(title))


def _chart_series_xml(idx, name, cat_range, val_range):
    return ('<c:ser><c:idx val="%d"/><c:order val="%d"/>'
            '<c:tx><c:v>%s</c:v></c:tx>'
            '<c:cat><c:strRef><c:f>%s</c:f></c:strRef></c:cat>'
            '<c:val><c:numRef><c:f>%s</c:f></c:numRef></c:val>'
            '</c:ser>' % (idx, idx, _esc(name), _esc(cat_range), _esc(val_range)))


_CHART_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
               'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
               '<c:chart><c:autoTitleDeleted val="1"/>')


def chart_column_xml(title, cat_range, val_range, name):
    return (_CHART_HEAD + _chart_title_xml(title) +
            '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>'
            '<c:grouping val="clustered"/>'
            + _chart_series_xml(0, name, cat_range, val_range) +
            '<c:gapWidth val="80"/></c:barChart></c:plotArea>'
            '<c:legend><c:legendPos val="b"/></c:legend>'
            '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>')


def chart_pie_xml(title, cat_range, val_range, name):
    return (_CHART_HEAD + _chart_title_xml(title) +
            '<c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/>'
            + _chart_series_xml(0, name, cat_range, val_range) +
            '</c:pieChart></c:plotArea>'
            '<c:legend><c:legendPos val="b"/></c:legend>'
            '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>')


def chart_line_xml(title, cat_range, val_range, name):
    return (_CHART_HEAD + _chart_title_xml(title) +
            '<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/>'
            + _chart_series_xml(0, name, cat_range, val_range) +
            '</c:lineChart></c:plotArea>'
            '<c:legend><c:legendPos val="b"/></c:legend>'
            '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>')


def drawing_xml():
    anchors = []
    # (col,row) top-left and bottom-right in 0-based EMU-free anchor form
    boxes = [(0, 6, 8, 21), (8, 6, 16, 21), (0, 21, 16, 36)]
    for i, (c1, r1, c2, r2) in enumerate(boxes, start=1):
        anchors.append(
            '<xdr:twoCellAnchor>'
            '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:graphicFrame macro="">'
            '<xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name="Chart %d"/>'
            '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            'r:id="rId%d"/></a:graphicData></a:graphic>'
            '</xdr:graphicFrame><xdr:clientData/>'
            '</xdr:twoCellAnchor>' % (c1, r1, c2, r2, i + 1, i, i))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            + "".join(anchors) + '</xdr:wsDr>')


# -- Sheet 4: Dashboard ---------------------------------------------------------

def dashboard_sheet_xml(sst, n_branches):
    b_total = len(BUCKETS) + 2          # DPD Summary TOTAL row
    br_total = n_branches + 2           # Branch Summary TOTAL row
    rows = []
    rows.append('<row><c r="A1" t="s" s="6"><v>%d</v></c></row>'
                % sst.add("Collections Dashboard — as of today"))
    rows.append(row([cell_str(2, 1, sst.add("KPI"), 1),
                     cell_str(2, 2, sst.add("Value"), 1)]))
    # KPIs (rows 3..6), formulas wired to the summary TOTAL rows
    kpis = [
        ("Total Outstanding (₹)",
         "'DPD Summary'!C%d" % b_total, 2),
        ("Accounts in Arrears",
         "'DPD Summary'!B%d-'DPD Summary'!B2" % b_total, 5),
        ("Collection Efficiency",
         "'Branch Summary'!F%d" % br_total, 4),
        ("PAR>30 (share of outstanding 31+ DPD)",
         "('DPD Summary'!C4+'DPD Summary'!C5+'DPD Summary'!C6)/'DPD Summary'!C%d" % b_total, 4),
    ]
    for i, (label, formula, style) in enumerate(kpis, start=3):
        rows.append(row([
            cell_str(i, 1, sst.add(label), 7),
            cell_formula(i, 2, formula, style),
        ]))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '%s<sheetData>%s</sheetData>'
            '<drawing r:id="rId1"/>'
            '</worksheet>' % (_cols_xml([42, 20]), "".join(rows)))


# -- Sheet 5: Action Queue ------------------------------------------------------

QUEUE_HEADERS = ["Rank", "LoanID", "Customer", "Phone", "Branch", "Agent",
                 "Due (₹)", "DPD", "Bucket", "Priority", "Reminder Template"]
QUEUE_WIDTHS = [7, 12, 20, 14, 14, 14, 12, 8, 10, 14, 60]


def build_queue_sheet(sst, queue):
    rows = [header_row(sst, QUEUE_HEADERS)]
    for rank, loan in enumerate(queue, start=2):
        rows.append(row([
            cell_num(rank, 1, rank - 1, 5),
            cell_str(rank, 2, sst.add(loan["loan_id"]), 10),
            cell_str(rank, 3, sst.add(loan["customer"]), 10),
            cell_str(rank, 4, sst.add(loan["phone"]), 10),
            cell_str(rank, 5, sst.add(loan["branch"]), 10),
            cell_str(rank, 6, sst.add(loan["agent"]), 10),
            cell_num(rank, 7, loan["due"], 2),
            cell_num(rank, 8, loan["dpd"], 5),
            cell_str(rank, 9, sst.add(loan["bucket"]), 10),
            cell_formula(rank, 10, "G%d*(1+H%d/30)" % (rank, rank), 5),
            cell_str(rank, 11, sst.add(reminder_message(loan)), 10),
        ]))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '%s<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews>'
            '<sheetData>%s</sheetData></worksheet>'
            % (_cols_xml(QUEUE_WIDTHS), "".join(rows)))

# ---------------------------------------------------------------------------
# Workbook assembly
# ---------------------------------------------------------------------------

def _add_part(zf, name, text):
    zf.writestr(name, text.encode("utf-8"))


def build_workbook(loans, queue, output_path, title="Collections Dashboard"):
    sst = SharedStrings()
    n = len(loans)
    n_branches = len(branch_names(loans))
    b_total = len(BUCKETS) + 2

    chart1 = chart_column_xml(
        "Outstanding by DPD Bucket",
        "'DPD Summary'!$A$2:$A$6", "'DPD Summary'!$C$2:$C$6", "Outstanding")
    chart2 = chart_pie_xml(
        "Outstanding Share by Bucket",
        "'DPD Summary'!$A$2:$A$6", "'DPD Summary'!$C$2:$C$6", "Share")
    chart3 = chart_line_xml(
        "Collection Efficiency by Branch",
        "'Branch Summary'!$A$2:$A$%d" % (n_branches + 1),
        "'Branch Summary'!$F$2:$F$%d" % (n_branches + 1), "Efficiency")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        _add_part(zf, "[Content_Types].xml", content_types_xml())
        _add_part(zf, "_rels/.rels", root_rels_xml())
        _add_part(zf, "xl/workbook.xml", workbook_xml())
        _add_part(zf, "xl/_rels/workbook.xml.rels", workbook_rels_xml())
        _add_part(zf, "xl/worksheets/sheet1.xml", raw_sheet_xml(sst, loans))
        _add_part(zf, "xl/worksheets/sheet2.xml",
                   build_dpd_summary_sheet(sst, n))
        _add_part(zf, "xl/worksheets/sheet3.xml",
                   build_branch_summary_sheet(sst, loans, n))
        _add_part(zf, "xl/worksheets/sheet4.xml",
                   dashboard_sheet_xml(sst, n_branches))
        _add_part(zf, "xl/worksheets/_rels/sheet4.xml.rels", sheet4_rels_xml())
        _add_part(zf, "xl/worksheets/sheet5.xml", build_queue_sheet(sst, queue))
        _add_part(zf, "xl/drawings/drawing1.xml", drawing_xml())
        _add_part(zf, "xl/drawings/_rels/drawing1.xml.rels", drawing_rels_xml())
        _add_part(zf, "xl/charts/chart1.xml", chart1)
        _add_part(zf, "xl/charts/chart2.xml", chart2)
        _add_part(zf, "xl/charts/chart3.xml", chart3)
        _add_part(zf, "xl/styles.xml", styles_xml())
        _add_part(zf, "xl/sharedStrings.xml", sst.xml())
        _add_part(zf, "docProps/core.xml", core_props_xml())
    return output_path


# ---------------------------------------------------------------------------
# Follow-up reminders CSV
# ---------------------------------------------------------------------------

def write_reminders(queue, csv_path):
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["LoanID", "Customer", "Phone", "Branch", "Agent",
                    "Due", "DPD", "Message"])
        for loan in queue:
            w.writerow([loan["loan_id"], loan["customer"], loan["phone"],
                        loan["branch"], loan["agent"],
                        "%.2f" % loan["due"], loan["dpd"],
                        reminder_message(loan)])
    return csv_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _parse_args(argv):
    args = {"ref_date": None, "top": 50}
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--generate-sample" and i + 1 < len(argv):
            args["generate_sample"] = argv[i + 1]; i += 2
        elif a == "--input" and i + 1 < len(argv):
            args["input"] = argv[i + 1]; i += 2
        elif a == "--output" and i + 1 < len(argv):
            args["output"] = argv[i + 1]; i += 2
        elif a == "--reminders" and i + 1 < len(argv):
            args["reminders"] = argv[i + 1]; i += 2
        elif a == "--ref-date" and i + 1 < len(argv):
            try:
                args["ref_date"] = datetime.date.fromisoformat(argv[i + 1])
            except ValueError:
                sys.exit("error: --ref-date must be YYYY-MM-DD")
            i += 2
        elif a == "--top" and i + 1 < len(argv):
            args["top"] = int(argv[i + 1]); i += 2
        else:
            sys.exit("error: unknown argument %r" % a)
    return args


def main(argv=None):
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    if "generate_sample" in args:
        generate_sample(args["generate_sample"],
                        ref_date=args["ref_date"])
        print("sample ledger written: %s" % args["generate_sample"])
        return 0
    if "input" not in args or "output" not in args:
        sys.exit("error: need --input and --output (or --generate-sample)")
    ref_date = args["ref_date"] or datetime.date.today()
    loans = load_ledger(args["input"], ref_date)
    due_loans = [l for l in loans if l["due"] > 0]
    queue = sorted(due_loans, key=lambda l: l["priority"],
                   reverse=True)[:max(args["top"], 1)]
    build_workbook(loans, queue, args["output"])
    total_due = sum(l["due"] for l in loans)
    print("loans: %d | in arrears: %d | total outstanding: ₹%s"
          % (len(loans), len(due_loans), _fmt_num(round(total_due, 2))))
    print("dashboard written: %s" % args["output"])
    if "reminders" in args:
        write_reminders(queue, args["reminders"])
        print("reminders written: %s (%d)" % (args["reminders"], len(queue)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
