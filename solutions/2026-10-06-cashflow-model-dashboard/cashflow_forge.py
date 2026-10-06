#!/usr/bin/env python3
"""
CashFlowForge - dynamic project cash-flow model + investor dashboard factory.

Reads a monthly project cash-flow CSV (Month, Inflows, Outflows) plus a few
assumptions (project name, initial investment, discount rate, horizon) and
emits a genuine multi-sheet .xlsx workbook (real OOXML, stdlib only - no
openpyxl/pandas):

  1. Assumptions  - editable inputs; every downstream cell is formula-linked
  2. Cash Flow    - M0..M24 schedule: Net CF, Cumulative CF, discount factor
                    and Discounted CF are ALL live Excel formulas
  3. Metrics      - Total Inflows/Outflows, Net Position, NPV, monthly +
                    annualized IRR (Excel IRR()), payback month (MATCH),
                    benefit-cost ratio (SUMPRODUCT)
  4. Dashboard    - 6 formula-wired investor KPIs + 3 native Excel charts:
                    monthly net cash flow (column), cumulative J-curve (line),
                    inflows vs outflows (pie); red fill on negative cash flow

Plus a `monthly_cashflow.csv` audit export computed independently in Python.

Built for the DashboardForge wedge: the kind of deliverable a finance /
NBFC client expects from a "dynamic Excel cash-flow model + summary
dashboard" engagement (monthly inflows/outflows, IRR, payback, investor
summary) - every number traces to a formula, nothing is hard-coded.

Usage:
    python3 cashflow_forge.py generate --out sample_project_2026.csv
    python3 cashflow_forge.py build --csv sample_project_2026.csv \
        --name "Coastal Aquaculture - Phase 1" --investment 4800000 \
        --rate 0.12 --out cashflow_model.xlsx
"""

import argparse
import csv
import datetime
import random
import sys
import zipfile
from xml.sax.saxutils import escape as _xml_escape

TOOL_NAME = "CashFlowForge"
HORIZON = 24                       # months M1..M24 (plus M0 for the investment)
SAMPLE_SEED = 20261006
SHEET_NAMES = ["Assumptions", "Cash Flow", "Metrics", "Dashboard"]

# ---------------------------------------------------------------------------
# OOXML writer core (stdlib only)
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


# Style ids (defined in styles_xml()):
#  0 default · 1 header · 2 currency (₹) · 3 date · 4 percent · 5 integer ·
#  6 title · 7 light-blue header · 8 light-blue currency · 9 mmm-yy ·
#  10 bordered text

def _cell_open(ref, kind, style):
    if kind == "s":
        return '<c r="%s" t="s" s="%d">' % (ref, style)
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


def styles_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="5">
<numFmt numFmtId="164" formatCode="#,##0"/>
<numFmt numFmtId="165" formatCode="\\u20b9#,##0"/>
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
# Domain logic: sample project + CSV loading
# ---------------------------------------------------------------------------

def generate_sample(csv_path, seed=SAMPLE_SEED):
    """Deterministic aquaculture-project cash-flow sample.

    M1-M6 : pond construction / equipment capex, token stocking revenue.
    M7-M24: harvest cycles - inflows ramp with monsoon-season dips,
    outflows = feed + labour + maintenance.
    """
    rng = random.Random(seed)
    rows = []
    for m in range(1, HORIZON + 1):
        if m <= 6:
            capex = 420000 - m * 45000
            opex = 95000 + rng.randint(-8000, 8000)
            inflow = 40000 + rng.randint(-6000, 6000) if m >= 4 else 0
            rows.append((m, inflow, capex + opex))
        else:
            season = 1.0 - 0.22 * (1 if m in (13, 14, 15, 16) else 0)  # monsoon dip
            ramp = min(1.0, 0.55 + (m - 6) * 0.055)                      # ramp-up
            inflow = int(980000 * season * ramp + rng.randint(-30000, 30000))
            outflow = int(560000 * season * (0.75 + 0.25 * ramp)
                          + rng.randint(-20000, 20000))
            rows.append((m, max(inflow, 0), max(outflow, 0)))
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["Month", "Inflows", "Outflows"])
        for m, inflow, outflow in rows:
            w.writerow([m, inflow, outflow])
    return csv_path


def load_project(csv_path):
    """-> list of (month:int, inflows:float, outflows:float) for M1..MHORIZON."""
    rows = []
    with open(csv_path, newline="", encoding="utf-8") as fh:
        for rec in csv.DictReader(fh):
            rows.append((int(rec["Month"]), float(rec["Inflows"]),
                         float(rec["Outflows"])))
    if len(rows) != HORIZON:
        raise ValueError("expected %d monthly rows, got %d"
                         % (HORIZON, len(rows)))
    if [r[0] for r in rows] != list(range(1, HORIZON + 1)):
        raise ValueError("Month column must run 1..%d" % HORIZON)
    return rows


def project_metrics(rows, investment, annual_rate):
    """Independent Python recomputation of every workbook metric."""
    mr = annual_rate / 12.0
    net, cum, disc = [], [], []
    running = 0.0
    total_in = total_out = 0.0
    # M0 first: the investment outflow
    for t, inflow, outflow in [(0, 0.0, investment)] + \
            [(m, i, o) for m, i, o in rows]:
        n = inflow - outflow
        running += n
        factor = 1.0 / ((1.0 + mr) ** t)
        net.append(n)
        cum.append(running)
        disc.append(n * factor)
        if t > 0:
            total_in += inflow
            total_out += outflow
    npv = sum(disc)
    # payback: first month index t where cumulative >= 0
    payback = next((t for t, c in enumerate(cum) if c >= 0), None)
    # IRR via bisection on monthly rate
    def npv_at(r):
        return sum(n / ((1.0 + r) ** t) for t, n in enumerate(net))
    lo, hi = -0.9999, 10.0
    irr_m = None
    if npv_at(lo) * npv_at(hi) < 0:
        for _ in range(200):
            mid = (lo + hi) / 2.0
            if npv_at(lo) * npv_at(mid) <= 0:
                hi = mid
            else:
                lo = mid
        irr_m = (lo + hi) / 2.0
    irr_a = (1.0 + irr_m) ** 12 - 1.0 if irr_m is not None else None
    pv_in = sum(i / ((1.0 + mr) ** t)
                for t, (m, i, o) in
                enumerate([(0, 0.0, 0.0)] + [(mm, ii, oo) for mm, ii, oo in rows]))
    pv_out = sum(o / ((1.0 + mr) ** t)
                 for t, (m, i, o) in
                 enumerate([(0, 0.0, investment)] + [(mm, ii, oo) for mm, ii, oo in rows]))
    return {
        "total_inflows": total_in,
        "total_outflows": total_out,
        "net_position": running,
        "npv": npv,
        "irr_monthly": irr_m,
        "irr_annual": irr_a,
        "payback_months": payback,
        "bcr": (pv_in / pv_out) if pv_out else None,
        "net": net, "cumulative": cum, "discounted": disc,
    }


# ---------------------------------------------------------------------------
# Package parts
# ---------------------------------------------------------------------------

def content_types_xml():
    sheets = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.worksheet+xml"/>' % i for i in range(1, 5))
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
        'Target="worksheets/sheet%d.xml"/>' % (i, i) for i in range(1, 5))
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
%s
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
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
<dc:title>Project Cash-Flow Model + Investor Dashboard</dc:title>
<cp:keywords>cash flow,npv,irr,payback,excel dashboard,finance,automation</cp:keywords>
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


WS_OPEN = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

# -- Sheet 1: Assumptions --------------------------------------------------------

def build_assumptions_sheet(sst, name, investment, annual_rate):
    rows = []
    rows.append(row([cell_str(1, 1, sst.add("Project Assumptions"), 6)]))
    rows.append(row([cell_str(2, 1, sst.add("Parameter"), 1),
                     cell_str(2, 2, sst.add("Value"), 1)]))
    rows.append(row([cell_str(3, 1, sst.add("Project name"), 10),
                     cell_str(3, 2, sst.add(name), 10)]))
    rows.append(row([cell_str(4, 1, sst.add("Initial investment (₹)"), 10),
                     cell_num(4, 2, investment, 2)]))
    rows.append(row([cell_str(5, 1, sst.add("Annual discount rate"), 10),
                     cell_num(5, 2, annual_rate, 4)]))
    rows.append(row([cell_str(6, 1, sst.add("Monthly discount rate"), 10),
                     cell_formula(6, 2, "B5/12", 4)]))
    rows.append(row([cell_str(7, 1, sst.add("Horizon (months)"), 10),
                     cell_num(7, 2, HORIZON, 5)]))
    rows.append(row([cell_str(8, 1, sst.add("Currency"), 10),
                     cell_str(8, 2, sst.add("INR"), 10)]))
    rows.append(row([cell_str(10, 1, sst.add(
        "Edit the blue cells above - every sheet recalculates automatically. "
        "Nothing on this workbook is hard-coded."), 0)]))
    return (WS_OPEN + _cols_xml([26, 34]) +
            '<sheetData>%s</sheetData></worksheet>' % "".join(rows))


# -- Sheet 2: Cash Flow ----------------------------------------------------------
# Row 2 = M0 (investment outflow), rows 3..26 = M1..M24.
# Cols: A Month# | B Inflows | C Outflows | D Net CF | E Cumulative |
#       F Discount Factor | G Discounted CF
CF_HEADERS = ["Month", "Inflows (₹)", "Outflows (₹)", "Net Cash Flow (₹)",
              "Cumulative CF (₹)", "Discount Factor", "Discounted CF (₹)"]
CF_WIDTHS = [8, 16, 16, 18, 18, 16, 18]
N_CF_ROWS = HORIZON + 1          # M0..M24
CF_LAST = N_CF_ROWS + 1         # last data row number (row 26)


def build_cashflow_sheet(sst, rows):
    out = [header_row(sst, CF_HEADERS)]
    monthly = "'Assumptions'!$B$6"      # monthly discount rate (absolute)
    for r, (m, inflow, outflow) in enumerate([(0, 0.0, None)] +
                                             [(m, i, o) for m, i, o in rows],
                                             start=2):
        cells = [cell_num(r, 1, m, 5)]
        cells.append(cell_num(r, 2, inflow, 2))
        if outflow is None:  # M0: investment outflow, linked to Assumptions
            cells.append(cell_formula(r, 3, "'Assumptions'!$B$4", 2))
        else:
            cells.append(cell_num(r, 3, outflow, 2))
        cells.append(cell_formula(r, 4, "B%d-C%d" % (r, r), 2))
        if r == 2:
            cells.append(cell_formula(r, 5, "D2", 2))
        else:
            cells.append(cell_formula(r, 5, "E%d+D%d" % (r - 1, r), 2))
        cells.append(cell_formula(r, 6,
                                  "1/POWER(1+%s,A%d)" % (monthly, r), 4))
        cells.append(cell_formula(r, 7, "D%d*F%d" % (r, r), 2))
        out.append(row(cells))
    cf = ('<conditionalFormatting sqref="D2:D%d">'
          '<cfRule type="cellIs" dxfId="1" priority="1" operator="lessThan">'
          '<formula><v>0</v></formula></cfRule>'
          '</conditionalFormatting>'
          '<conditionalFormatting sqref="E2:E%d">'
          '<cfRule type="cellIs" dxfId="1" priority="2" operator="lessThan">'
          '<formula><v>0</v></formula></cfRule>'
          '</conditionalFormatting>' % (CF_LAST, CF_LAST))
    return (WS_OPEN + _cols_xml(CF_WIDTHS) +
            '<sheetData>%s</sheetData>%s</worksheet>' % ("".join(out), cf))


# -- Sheet 3: Metrics ------------------------------------------------------------

METRIC_DEFS = [
    # (label, formula, style)
    ("Total Inflows (₹)", "SUM('Cash Flow'!$B$2:$B$%d)", 2),
    ("Total Outflows (₹)", "SUM('Cash Flow'!$C$2:$C$%d)", 2),
    ("Net Cash Position (₹)", "SUM('Cash Flow'!$D$2:$D$%d)", 2),
    ("NPV @ discount rate (₹)", "SUM('Cash Flow'!$G$2:$G$%d)", 2),
    ("IRR - monthly", 'IFERROR(IRR(\'Cash Flow\'!$D$2:$D$%d,0.1),"n/a")', 4),
    ("IRR - annualized", 'IFERROR(POWER(1+IRR(\'Cash Flow\'!$D$2:$D$%d,0.1),12)-1,"n/a")', 4),
    ("Payback (months)",
     'IFERROR(MATCH(TRUE,\'Cash Flow\'!$E$2:$E$%d>=0,0)-1,"Not within horizon")', 5),
    ("Benefit-Cost Ratio",
     "IF(SUMPRODUCT('Cash Flow'!$C$2:$C$%d,'Cash Flow'!$F$2:$F$%d)=0,\"n/a\","
     "SUMPRODUCT('Cash Flow'!$B$2:$B$%d,'Cash Flow'!$F$2:$F$%d)/"
     "SUMPRODUCT('Cash Flow'!$C$2:$C$%d,'Cash Flow'!$F$2:$F$%d))", 0),
]


def build_metrics_sheet(sst):
    n = CF_LAST
    rows = [row([cell_str(1, 1, sst.add("Investment Metrics"), 6)]),
            row([cell_str(2, 1, sst.add("Metric"), 1),
                 cell_str(2, 2, sst.add("Value"), 1),
                 cell_str(2, 3, sst.add("How it is computed"), 1)])]
    how = [
        "SUM of monthly inflows (M1..M24)",
        "SUM of monthly outflows incl. M0 investment",
        "SUM of monthly net cash flows",
        "SUM of discounted cash flows (M0..M24)",
        "Excel IRR() over the M0..M24 net-CF series",
        "Annualized: (1 + monthly IRR)^12 - 1",
        "First month with cumulative CF >= 0 (MATCH)",
        "PV(inflows) / PV(outflows) via SUMPRODUCT",
    ]
    for i, ((label, formula, style), note) in enumerate(zip(METRIC_DEFS, how),
                                                       start=3):
        rows.append(row([
            cell_str(i, 1, sst.add(label), 10),
            cell_formula(i, 2, formula % ((n,) * formula.count("%d")), style),
            cell_str(i, 3, sst.add(note), 0),
        ]))
    cf = ('<conditionalFormatting sqref="B6">'
          '<cfRule type="cellIs" dxfId="1" priority="1" operator="lessThan">'
          '<formula><v>0</v></formula></cfRule>'
          '</conditionalFormatting>')
    return (WS_OPEN + _cols_xml([28, 22, 52]) +
            '<sheetData>%s</sheetData>%s</worksheet>' % ("".join(rows), cf))

# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------

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
    # (col,row) top-left and bottom-right in 0-based anchor form
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


# -- Sheet 4: Dashboard ----------------------------------------------------------

def dashboard_sheet_xml(sst, project_name):
    rows = []
    rows.append(row([cell_str(1, 1, sst.add("Investor Dashboard"), 6)]))
    rows.append(row([cell_str(2, 1, sst.add(project_name), 0)]))
    labels = ["Initial Investment (₹)", "Total Inflows (₹)", "NPV (₹)",
              "IRR (Annual)", "Payback (Months)", "Net Position (₹)"]
    rows.append(row([cell_str(3, i + 1, sst.add(t), 7)
                     for i, t in enumerate(labels)]))
    # KPI wiring: B4 of Metrics is Total Outflows; investment lives on
    # Assumptions!B4 - so the first KPI points there directly.
    refs = ["'Assumptions'!$B$4", "'Metrics'!$B$3", "'Metrics'!$B$6",
            "'Metrics'!$B$8", "'Metrics'!$B$9", "'Metrics'!$B$5"]
    styles = [8, 8, 8, 4, 5, 8]
    rows.append(row([cell_formula(4, i + 1, ref, st)
                     for i, (ref, st) in enumerate(zip(refs, styles))]))
    rows.append(row([cell_str(5, 1, sst.add(
        "Every KPI above is a live formula linked to the Assumptions, "
        "Cash Flow and Metrics sheets - change an input and the whole "
        "dashboard recalculates. Nothing is hard-coded."), 0)]))
    cf = ('<conditionalFormatting sqref="C4">'
          '<cfRule type="cellIs" dxfId="1" priority="1" operator="lessThan">'
          '<formula><v>0</v></formula></cfRule>'
          '</conditionalFormatting>')
    return (WS_OPEN + _cols_xml([20, 20, 20, 16, 18, 20]) +
            '<sheetData>%s</sheetData>%s</worksheet>' % ("".join(rows), cf))


# ---------------------------------------------------------------------------
# Workbook assembly + CSV audit export
# ---------------------------------------------------------------------------

def _add_part(zf, name, text):
    zf.writestr(name, text.encode("utf-8"))


def build_workbook(rows, name, investment, annual_rate, output_path):
    sst = SharedStrings()
    n = CF_LAST  # 26
    chart1 = chart_column_xml(
        "Monthly Net Cash Flow",
        "'Cash Flow'!$A$2:$A$%d" % n, "'Cash Flow'!$D$2:$D$%d" % n,
        "Net Cash Flow")
    chart2 = chart_line_xml(
        "Cumulative Cash Flow (J-Curve)",
        "'Cash Flow'!$A$2:$A$%d" % n, "'Cash Flow'!$E$2:$E$%d" % n,
        "Cumulative")
    chart3 = chart_pie_xml(
        "Inflows vs Outflows",
        "'Metrics'!$A$3:$A$4", "'Metrics'!$B$3:$B$4", "Share")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        _add_part(zf, "[Content_Types].xml", content_types_xml())
        _add_part(zf, "_rels/.rels", root_rels_xml())
        _add_part(zf, "xl/workbook.xml", workbook_xml())
        _add_part(zf, "xl/_rels/workbook.xml.rels", workbook_rels_xml())
        _add_part(zf, "xl/worksheets/sheet1.xml",
                   build_assumptions_sheet(sst, name, investment, annual_rate))
        _add_part(zf, "xl/worksheets/sheet2.xml",
                   build_cashflow_sheet(sst, rows))
        _add_part(zf, "xl/worksheets/sheet3.xml", build_metrics_sheet(sst))
        _add_part(zf, "xl/worksheets/sheet4.xml",
                   dashboard_sheet_xml(sst, name))
        _add_part(zf, "xl/worksheets/_rels/sheet4.xml.rels", sheet4_rels_xml())
        _add_part(zf, "xl/drawings/drawing1.xml", drawing_xml())
        _add_part(zf, "xl/drawings/_rels/drawing1.xml.rels", drawing_rels_xml())
        _add_part(zf, "xl/charts/chart1.xml", chart1)
        _add_part(zf, "xl/charts/chart2.xml", chart2)
        _add_part(zf, "xl/charts/chart3.xml", chart3)
        _add_part(zf, "xl/styles.xml", styles_xml())
        _add_part(zf, "xl/sharedStrings.xml", sst.xml())
        _add_part(zf, "docProps/core.xml", core_props_xml())
    return output_path


def write_monthly_csv(metrics, csv_path):
    """Independent Python-side audit export of the M0..M24 schedule."""
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["Month", "NetCashFlow", "CumulativeCF", "DiscountedCF"])
        for t, (n, c, d) in enumerate(zip(metrics["net"],
                                          metrics["cumulative"],
                                          metrics["discounted"])):
            w.writerow([t, "%.2f" % n, "%.2f" % c, "%.2f" % d])
    return csv_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _parse_args(argv):
    p = argparse.ArgumentParser(prog="cashflow_forge",
                                description="Dynamic cash-flow model + "
                                            "investor dashboard factory")
    sub = p.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("generate", help="write a deterministic sample CSV")
    g.add_argument("--out", default="sample_project_2026.csv")
    g.add_argument("--seed", type=int, default=SAMPLE_SEED)

    b = sub.add_parser("build", help="build the .xlsx workbook + audit CSV")
    b.add_argument("--csv", required=True)
    b.add_argument("--name", default="Coastal Aquaculture - Phase 1")
    b.add_argument("--investment", type=float, default=2000000.0)
    b.add_argument("--rate", type=float, default=0.12)
    b.add_argument("--out", default="cashflow_model.xlsx")
    b.add_argument("--audit-csv", default="monthly_cashflow.csv")
    return p.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv or sys.argv[1:])
    if args.cmd == "generate":
        path = generate_sample(args.out, seed=args.seed)
        print("sample CSV -> %s" % path)
        return 0
    rows = load_project(args.csv)
    metrics = project_metrics(rows, args.investment, args.rate)
    build_workbook(rows, args.name, args.investment, args.rate, args.out)
    write_monthly_csv(metrics, args.audit_csv)
    print("workbook -> %s" % args.out)
    print("audit CSV -> %s" % args.audit_csv)
    print("NPV ₹%s | IRR(ann) %.1f%% | payback %s months" % (
        _fmt_num(round(metrics["npv"])),
        (metrics["irr_annual"] or 0) * 100,
        metrics["payback_months"]
        if metrics["payback_months"] is not None else "n/a"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
