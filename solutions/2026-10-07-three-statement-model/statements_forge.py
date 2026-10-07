#!/usr/bin/env python3
"""
StatementsForge - three-statement financial model factory.

Reads a trial-balance CSV (Account, Category, Subcategory, Opening, Closing)
and emits a genuine multi-sheet .xlsx workbook (real OOXML, stdlib only - no
openpyxl/pandas):

  1. Trial Balance  - raw input with a formula Change column, AutoFilter,
                      frozen panes
  2. Income Statement - every line a live SUMIFS on the Trial Balance;
                      Gross Profit / EBITDA / EBIT / Net Income all formulas
  3. Balance Sheet  - Assets = Liabilities + Equity with a live balance
                      check; Retained Earnings articulated as
                      Opening + Net Income (linked to the Income Statement)
  4. Cash Flow      - indirect method, all movements formula-linked to the
                      Trial Balance Change column; ending cash tied back to
                      the Balance Sheet with a live tie-out check
  5. Ratios         - liquidity, leverage, profitability + common-size IS,
                      all formula-linked
  6. Dashboard      - 5 formula-wired KPIs + 3 native Excel charts
                      (revenue vs cost structure, asset mix, cash walk)

Plus an `audit.csv` computed independently in Python (same numbers the
formulas must produce).

Built for the DashboardForge wedge: the deliverable a finance client expects
from a "three-statement Excel model + presentation-ready outputs" engagement -
every number traces to a formula, nothing is hard-coded.

Usage:
    python3 statements_forge.py generate --out sample_company_2026.csv
    python3 statements_forge.py build --csv sample_company_2026.csv \
        --name "Sharma Textiles Pvt Ltd" --out three_statement_model.xlsx
"""

import argparse
import csv
import sys
import zipfile
from xml.sax.saxutils import escape as _xml_escape

TOOL_NAME = "StatementsForge"
COMPANY = "Sharma Textiles Pvt Ltd"
FY_LABEL = "FY 2025-26"
SAMPLE_SEED = 20261007
SHEET_NAMES = ["Trial Balance", "Income Statement", "Balance Sheet",
               "Cash Flow", "Ratios", "Dashboard"]

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
        return "%.2f" % value
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
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
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
# Domain logic: sample company trial balance + independent model
# ---------------------------------------------------------------------------
# (account, category, subcategory, opening, closing); None closing means the
# generator computes it: Cash opening is the balancing plug, Retained
# Earnings closing = opening + Net Income, Cash closing = opening + net
# change in cash from the indirect cash-flow build. Everything else is fixed.

_BASE_ACCOUNTS = [
    ("Cash & Bank", "Asset", "Current", None, None),
    ("Trade Receivables", "Asset", "Current", 2400000, 2950000),
    ("Inventory", "Asset", "Current", 3100000, 3420000),
    ("Prepaid Expenses", "Asset", "Current", 180000, 220000),
    ("Plant & Machinery - Gross", "Asset", "Non-current", 8500000, 9700000),
    ("Accumulated Depreciation", "Asset", "Non-current", -2550000, -3175000),
    ("Trade Payables", "Liability", "Current", 1900000, 2230000),
    ("Short-term Borrowings", "Liability", "Current", 1500000, 1200000),
    ("Statutory Dues Payable", "Liability", "Current", 320000, 410000),
    ("Term Loan", "Liability", "Non-current", 4000000, 3200000),
    ("Share Capital", "Equity", "Share Capital", 5000000, 5000000),
    ("Retained Earnings", "Equity", "Retained", 2610000, None),
    ("Sales Revenue", "Revenue", "Sales", 0, 18500000),
    ("Other Income", "Revenue", "Other", 0, 240000),
    ("Raw Material Consumed", "Expense", "COGS", 0, 9800000),
    ("Direct Labour", "Expense", "COGS", 0, 1650000),
    ("Employee Costs", "Expense", "Operating", 0, 2240000),
    ("Rent", "Expense", "Operating", 0, 720000),
    ("Utilities & Maintenance", "Expense", "Operating", 0, 486000),
    ("Selling & Distribution", "Expense", "Operating", 0, 612000),
    ("Depreciation", "Expense", "Operating", 0, 625000),
    ("Other Admin Expenses", "Expense", "Operating", 0, 398000),
    ("Interest Expense", "Expense", "Finance", 0, 512000),
    ("Tax Expense", "Expense", "Tax", 0, 449000),
]

_CSV_HEADER = ["Account", "Category", "Subcategory", "Opening", "Closing"]


def _by_account(accounts, name):
    for a in accounts:
        if a["account"] == name:
            return a
    raise KeyError(name)


def sample_accounts():
    """Deterministic sample trial balance; plugs computed, identity asserted."""
    accounts = [{"account": a, "category": c, "subcategory": s,
                 "opening": o, "closing": cl}
                for a, c, s, o, cl in _BASE_ACCOUNTS]

    def subtotal(cat, sub, which):
        return sum(a[which] for a in accounts
                   if a["category"] == cat and a["subcategory"] == sub
                   and a[which] is not None)

    # Cash opening plugs the opening balance-sheet identity.
    assets_open = sum(a["opening"] for a in accounts
                      if a["category"] == "Asset" and a["account"] != "Cash & Bank")
    liab_open = sum(a["opening"] for a in accounts if a["category"] == "Liability")
    equity_open = sum(a["opening"] for a in accounts if a["category"] == "Equity")
    cash_open = liab_open + equity_open - assets_open
    assert cash_open > 0, "plug produced non-positive cash opening"
    _by_account(accounts, "Cash & Bank")["opening"] = cash_open
    assets_open += cash_open
    assert assets_open == liab_open + equity_open, "opening identity broken"

    # P&L first: Net Income drives the retained-earnings plug.
    revenue = subtotal("Revenue", "Sales", "closing") + subtotal("Revenue", "Other", "closing")
    cogs = subtotal("Expense", "COGS", "closing")
    opex = subtotal("Expense", "Operating", "closing")
    interest = subtotal("Expense", "Finance", "closing")
    tax = subtotal("Expense", "Tax", "closing")
    net_income = revenue - cogs - opex - interest - tax
    assert net_income > 0, "sample must be profitable"
    re = _by_account(accounts, "Retained Earnings")
    re["closing"] = re["opening"] + net_income

    # Indirect cash-flow movements -> cash closing plug.
    def chg(name):
        a = _by_account(accounts, name)
        return a["closing"] - a["opening"]

    # Depreciation add-back is the Depreciation expense line itself.
    operating = (net_income
                 + _by_account(accounts, "Depreciation")["closing"]
                 - chg("Trade Receivables")
                 - chg("Inventory")
                 - chg("Prepaid Expenses")
                 + chg("Trade Payables")
                 + chg("Statutory Dues Payable"))
    investing = -chg("Plant & Machinery - Gross")
    financing = chg("Short-term Borrowings") + chg("Term Loan")
    net_change = operating + investing + financing
    _by_account(accounts, "Cash & Bank")["closing"] = cash_open + net_change

    # Closing identity must hold by construction (double-entry).
    assets_close = sum(a["closing"] for a in accounts if a["category"] == "Asset")
    liab_close = sum(a["closing"] for a in accounts if a["category"] == "Liability")
    equity_close = sum(a["closing"] for a in accounts if a["category"] == "Equity")
    assert assets_close == liab_close + equity_close, "closing identity broken"
    return accounts


def generate_sample(csv_path):
    accounts = sample_accounts()
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(_CSV_HEADER)
        for a in accounts:
            w.writerow([a["account"], a["category"], a["subcategory"],
                        a["opening"], a["closing"]])
    return csv_path


def load_accounts(csv_path):
    with open(csv_path, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    if not rows or list(rows[0].keys()) != _CSV_HEADER:
        raise ValueError("expected header: %s" % ",".join(_CSV_HEADER))
    accounts = []
    for r in rows:
        accounts.append({"account": r["Account"].strip(),
                         "category": r["Category"].strip(),
                         "subcategory": r["Subcategory"].strip(),
                         "opening": float(r["Opening"]),
                         "closing": float(r["Closing"])})
    if len(accounts) < 10:
        raise ValueError("need at least 10 accounts, got %d" % len(accounts))
    return accounts


def model(accounts):
    """Independent Python-side computation of the full three statements."""
    def cat_sum(cat, sub=None, which="closing"):
        return sum(a[which] for a in accounts
                   if a["category"] == cat and (sub is None or a["subcategory"] == sub))

    def acct(name, which="closing"):
        return _by_account(accounts, name)[which]

    def chg(name):
        return acct(name, "closing") - acct(name, "opening")

    revenue = cat_sum("Revenue")
    cogs = cat_sum("Expense", "COGS")
    opex = cat_sum("Expense", "Operating")
    ebitda = revenue - cogs - (opex - acct("Depreciation"))
    ebit = revenue - cogs - opex
    ebt = ebit - cat_sum("Expense", "Finance")
    net_income = ebt - cat_sum("Expense", "Tax")

    cur_assets = cat_sum("Asset", "Current")
    ncur_assets = cat_sum("Asset", "Non-current")
    total_assets = cur_assets + ncur_assets
    cur_liab = cat_sum("Liability", "Current")
    ncur_liab = cat_sum("Liability", "Non-current")
    total_liab = cur_liab + ncur_liab
    share_cap = cat_sum("Equity", "Share Capital")
    re_open = acct("Retained Earnings", "opening")
    re_close = re_open + net_income
    total_equity = share_cap + re_close
    balance_check = total_assets - (total_liab + total_equity)

    op_cf = (net_income + acct("Depreciation")
             - chg("Trade Receivables") - chg("Inventory") - chg("Prepaid Expenses")
             + chg("Trade Payables") + chg("Statutory Dues Payable"))
    inv_cf = -chg("Plant & Machinery - Gross")
    fin_cf = chg("Short-term Borrowings") + chg("Term Loan")
    net_change = op_cf + inv_cf + fin_cf
    cash_open = acct("Cash & Bank", "opening")
    cash_close = cash_open + net_change
    tie_out = cash_close - acct("Cash & Bank")

    return {
        "revenue": revenue, "cogs": cogs, "gross": revenue - cogs,
        "opex": opex, "ebitda": ebitda, "ebit": ebit, "ebt": ebt,
        "net_income": net_income,
        "cur_assets": cur_assets, "ncur_assets": ncur_assets,
        "total_assets": total_assets,
        "cur_liab": cur_liab, "ncur_liab": ncur_liab, "total_liab": total_liab,
        "share_cap": share_cap, "re_open": re_open, "re_close": re_close,
        "total_equity": total_equity, "balance_check": balance_check,
        "op_cf": op_cf, "inv_cf": inv_cf, "fin_cf": fin_cf,
        "net_change": net_change, "cash_open": cash_open,
        "cash_close": cash_close, "tie_out": tie_out,
        "current_ratio": cur_assets / cur_liab,
        "quick_ratio": (cur_assets - acct("Inventory")) / cur_liab,
        "debt_equity": total_liab / total_equity,
        "gross_margin": (revenue - cogs) / revenue,
        "ebitda_margin": ebitda / revenue,
        "net_margin": net_income / revenue,
        "roe": net_income / total_equity,
        "roa": net_income / total_assets,
    }


# ---------------------------------------------------------------------------
# Workbook packaging
# ---------------------------------------------------------------------------

def content_types_xml():
    sheets = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.worksheet+xml"/>' % i for i in range(1, 7))
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
        'Target="worksheets/sheet%d.xml"/>' % (i, i) for i in range(1, 7))
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
%s
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>""" % rels


def sheet6_rels_xml():
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
<dc:title>Three-Statement Financial Model + Executive Dashboard</dc:title>
<cp:keywords>three statement model,income statement,balance sheet,cash flow,excel dashboard,finance,automation</cp:keywords>
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


WS_OPEN = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

TB = "'Trial Balance'"
IS = "'Income Statement'"
BS = "'Balance Sheet'"
CF = "'Cash Flow'"
RA = "'Ratios'"


def _sumifs(amount_col, cat=None, subcat=None, account=None,
            first=3, last=26, opening=False):
    """Build a SUMIFS on the Trial Balance sheet (closing col E, opening D)."""
    col = "D" if opening else amount_col
    crit = ["%s!$%s$%d:$%s$%d" % (TB, col, first, col, last)]
    if cat:
        crit += ["%s!$B$%d:$B$%d" % (TB, first, last), '"%s"' % cat]
    if subcat:
        crit += ["%s!$C$%d:$C$%d" % (TB, first, last), '"%s"' % subcat]
    if account:
        crit += ["%s!$A$%d:$A$%d" % (TB, first, last), '"%s"' % account]
    return "SUMIFS(" + ",".join(crit) + ")"


def build_trial_balance_sheet(sst, accounts, company):
    first, last = 3, 2 + len(accounts)
    rowmap = {}
    rows = [row([cell_str(1, 1, sst.add("Trial Balance — " + FY_LABEL), 6)]),
            row([cell_str(1, 2, sst.add(company), 0)])]
    rows.append(row([cell_str(2, i + 1, sst.add(h), 1)
                     for i, h in enumerate(["Account", "Category", "Subcategory",
                                            "Opening (₹)", "Closing (₹)",
                                            "Change (₹)"])]))
    for j, a in enumerate(accounts):
        r = first + j
        rowmap[a["account"]] = r
        rows.append(row([
            cell_str(r, 1, sst.add(a["account"]), 10),
            cell_str(r, 2, sst.add(a["category"]), 10),
            cell_str(r, 3, sst.add(a["subcategory"]), 10),
            cell_num(r, 4, a["opening"], 2),
            cell_num(r, 5, a["closing"], 2),
            cell_formula(r, 6, "E%d-D%d" % (r, r), 2),
        ]))
    rows.append(row([cell_str(last + 2, 1, sst.add(
        "Edit any Opening/Closing value above - the Change column and all "
        "three statements recalculate automatically. Nothing downstream is "
        "hard-coded."), 0)]))
    views = ('<sheetViews><sheetView workbookViewId="0">'
             '<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" '
             'state="frozen"/>'
             '<selection pane="bottomLeft" activeCell="A3" sqref="A3"/>'
             '</sheetView></sheetViews>')
    af = '<autoFilter ref="A2:F%d"/>' % last
    return ((WS_OPEN + _cols_xml([30, 14, 16, 18, 18, 18]) + views + af +
             '<sheetData>%s</sheetData></worksheet>' % "".join(rows)),
            rowmap, first, last)


def build_income_statement_sheet(sst, company, first, last):
    S = lambda cat, sub: _sumifs("E", cat=cat, subcat=sub, first=first, last=last)
    rows = [row([cell_str(1, 1, sst.add("Income Statement — " + FY_LABEL), 6)]),
            row([cell_str(2, 1, sst.add(company), 0)])]
    R = {}
    def line(r, label, formula, style=2, note=""):
        cells = [cell_str(r, 1, sst.add(label), 10),
                 cell_formula(r, 2, formula, style)]
        if note:
            cells.append(cell_str(r, 3, sst.add(note), 0))
        rows.append(row(cells))
    def section(r, label):
        rows.append(row([cell_str(r, 1, sst.add(label), 7)]))
    def total(r, label, formula, note=""):
        line(r, label, formula, style=8, note=note)

    section(4, "REVENUE")
    line(5, "Sales Revenue", S("Revenue", "Sales"), note="from Trial Balance")
    line(6, "Other Income", S("Revenue", "Other"), note="from Trial Balance")
    total(7, "Total Revenue", "SUM(B5:B6)")
    section(9, "COST OF GOODS SOLD")
    line(10, "Raw Material Consumed", _sumifs("E", account="Raw Material Consumed", first=first, last=last))
    line(11, "Direct Labour", _sumifs("E", account="Direct Labour", first=first, last=last))
    total(12, "Total COGS", "SUM(B10:B11)")
    total(13, "GROSS PROFIT", "B7-B12")
    section(15, "OPERATING EXPENSES")
    opex_rows = []
    for r, acct in [(16, "Employee Costs"), (17, "Rent"),
                    (18, "Utilities & Maintenance"),
                    (19, "Selling & Distribution"), (20, "Depreciation"),
                    (21, "Other Admin Expenses")]:
        line(r, acct, _sumifs("E", account=acct, first=first, last=last))
        opex_rows.append(r)
    total(22, "Total Operating Expenses", "SUM(B16:B21)")
    total(23, "EBITDA", "B13-(B22-B20)", note="excludes depreciation")
    total(24, "EBIT", "B13-B22")
    section(26, "FINANCE & TAX")
    line(27, "Interest Expense", _sumifs("E", account="Interest Expense", first=first, last=last))
    total(28, "Profit Before Tax", "B24-B27")
    line(29, "Tax Expense", _sumifs("E", account="Tax Expense", first=first, last=last))
    total(30, "NET INCOME", "B28-B29")
    rows.append(row([cell_str(32, 1, sst.add(
        "Every figure above is a live SUMIFS formula on the Trial Balance - "
        "change one account balance and the whole statement recalculates."), 0)]))
    R.update({"rev": "B7", "cogs": "B12", "gross": "B13", "opex": "B22",
              "ebitda": "B23", "ebit": "B24", "ebt": "B28", "ni": "B30"})
    return (WS_OPEN + _cols_xml([30, 20, 44]) +
            '<sheetData>%s</sheetData></worksheet>' % "".join(rows), R)


def build_balance_sheet_sheet(sst, company, first, last, is_refs):
    A = lambda acct: _sumifs("E", account=acct, first=first, last=last)
    AO = lambda acct: _sumifs("E", account=acct, first=first, last=last, opening=True)
    rows = [row([cell_str(1, 1, sst.add("Balance Sheet — as at 31 Mar 2026"), 6)]),
            row([cell_str(2, 1, sst.add(company), 0)])]
    def section(r, label):
        rows.append(row([cell_str(r, 1, sst.add(label), 7)]))
    def sub(r, label):
        rows.append(row([cell_str(r, 1, sst.add(label), 1)]))
    def line(r, label, formula, style=2):
        rows.append(row([cell_str(r, 1, sst.add(label), 10),
                         cell_formula(r, 2, formula, style)]))
    def total(r, label, formula):
        line(r, label, formula, style=8)

    section(4, "ASSETS")
    sub(5, "Current Assets")
    line(6, "Cash & Bank", A("Cash & Bank"))
    line(7, "Trade Receivables", A("Trade Receivables"))
    line(8, "Inventory", A("Inventory"))
    line(9, "Prepaid Expenses", A("Prepaid Expenses"))
    total(10, "Total Current Assets", "SUM(B6:B9)")
    sub(11, "Non-current Assets")
    line(12, "Plant & Machinery (Gross)", A("Plant & Machinery - Gross"))
    line(13, "Less: Accumulated Depreciation", A("Accumulated Depreciation"))
    total(14, "Net Non-current Assets", "SUM(B12:B13)")
    total(15, "TOTAL ASSETS", "B10+B14")
    section(17, "LIABILITIES")
    sub(18, "Current Liabilities")
    line(19, "Trade Payables", A("Trade Payables"))
    line(20, "Short-term Borrowings", A("Short-term Borrowings"))
    line(21, "Statutory Dues Payable", A("Statutory Dues Payable"))
    total(22, "Total Current Liabilities", "SUM(B19:B21)")
    sub(23, "Non-current Liabilities")
    line(24, "Term Loan", A("Term Loan"))
    total(25, "Total Non-current Liabilities", "B24")
    total(26, "TOTAL LIABILITIES", "B22+B25")
    section(28, "EQUITY")
    line(29, "Share Capital", A("Share Capital"))
    line(30, "Retained Earnings — Opening", AO("Retained Earnings"))
    line(31, "Add: Net Income for the year", "%s!%s" % (IS, is_refs["ni"]))
    total(32, "Retained Earnings — Closing", "B30+B31")
    total(33, "TOTAL EQUITY", "B29+B32")
    total(35, "TOTAL LIABILITIES & EQUITY", "B26+B33")
    rows.append(row([cell_str(36, 1, sst.add("BALANCE CHECK (must be 0)"), 10),
                     cell_formula(36, 2, "B15-B35", 5)]))
    rows.append(row([cell_str(37, 1, sst.add(
        "Green = balanced. Retained Earnings articulates: opening balance "
        "plus this year's Net Income from the Income Statement."), 0)]))
    cf = ('<conditionalFormatting sqref="B36">'
          '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal">'
          '<formula>0</formula></cfRule>'
          '<cfRule type="cellIs" dxfId="1" priority="2" operator="notEqual">'
          '<formula>0</formula></cfRule>'
          '</conditionalFormatting>')
    R = {"ca": "B10", "nca": "B14", "assets": "B15", "cl": "B22",
         "ncl": "B25", "liab": "B26", "equity": "B33", "tot_le": "B35",
         "check": "B36", "cash": "B6", "inv": "B8"}
    return (WS_OPEN + _cols_xml([32, 20]) +
            '<sheetData>%s</sheetData>%s</worksheet>' % ("".join(rows), cf), R)


def build_cashflow_sheet(sst, company, rowmap, is_refs, bs_refs):
    F = lambda acct: "%s!$F$%d" % (TB, rowmap[acct])
    rows = [row([cell_str(1, 1, sst.add("Cash Flow Statement (Indirect Method) — " + FY_LABEL), 6)]),
            row([cell_str(2, 1, sst.add(company), 0)])]
    def section(r, label):
        rows.append(row([cell_str(r, 1, sst.add(label), 7)]))
    def line(r, label, formula, style=2, note=""):
        cells = [cell_str(r, 1, sst.add(label), 10),
                 cell_formula(r, 2, formula, style)]
        if note:
            cells.append(cell_str(r, 3, sst.add(note), 0))
        rows.append(row(cells))
    def total(r, label, formula):
        line(r, label, formula, style=8)

    section(4, "OPERATING ACTIVITIES")
    line(5, "Net Income", "%s!%s" % (IS, is_refs["ni"]), note="from Income Statement")
    line(6, "Add: Depreciation & Amortisation",
         _sumifs("E", account="Depreciation"))
    line(7, "Less: Increase in Trade Receivables", "-(%s)" % F("Trade Receivables"))
    line(8, "Less: Increase in Inventory", "-(%s)" % F("Inventory"))
    line(9, "Less: Increase in Prepaid Expenses", "-(%s)" % F("Prepaid Expenses"))
    line(10, "Add: Increase in Trade Payables", F("Trade Payables"))
    line(11, "Add: Increase in Statutory Dues", F("Statutory Dues Payable"))
    total(12, "Cash from Operations", "SUM(B5:B11)")
    section(14, "INVESTING ACTIVITIES")
    line(15, "Purchase of Plant & Machinery", "-(%s)" % F("Plant & Machinery - Gross"))
    total(16, "Cash used in Investing", "B15")
    section(18, "FINANCING ACTIVITIES")
    line(19, "Repayment of Short-term Borrowings", F("Short-term Borrowings"))
    line(20, "Repayment of Term Loan", F("Term Loan"))
    total(21, "Cash used in Financing", "SUM(B19:B20)")
    total(23, "Net Change in Cash", "B12+B16+B21")
    line(24, "Cash at Beginning of Year",
         _sumifs("E", account="Cash & Bank", opening=True))
    total(25, "CASH AT END OF YEAR", "B23+B24")
    rows.append(row([cell_str(26, 1, sst.add("Tie-out vs Balance Sheet cash (must be 0)"), 10),
                     cell_formula(26, 2, "B25-%s!%s" % (BS, bs_refs["cash"]), 5)]))
    rows.append(row([cell_str(27, 1, sst.add(
        "Every movement is a live formula on the Trial Balance Change "
        "column - no hard-coded deltas."), 0)]))
    cf = ('<conditionalFormatting sqref="B26">'
          '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal">'
          '<formula>0</formula></cfRule>'
          '<cfRule type="cellIs" dxfId="1" priority="2" operator="notEqual">'
          '<formula>0</formula></cfRule>'
          '</conditionalFormatting>')
    R = {"op": "B12", "inv": "B16", "fin": "B21", "net_change": "B23",
         "cash_beg": "B24", "cash_end": "B25", "tie": "B26"}
    return (WS_OPEN + _cols_xml([36, 20, 40]) +
            '<sheetData>%s</sheetData>%s</worksheet>' % ("".join(rows), cf), R)


def build_ratios_sheet(sst, company, is_refs, bs_refs):
    rows = [row([cell_str(1, 1, sst.add("Key Ratios & Common-Size Analysis — " + FY_LABEL), 6)]),
            row([cell_str(2, 1, sst.add(company), 0)])]
    def section(r, label):
        rows.append(row([cell_str(r, 1, sst.add(label), 7)]))
    def ratio(r, label, formula, note):
        rows.append(row([cell_str(r, 1, sst.add(label), 10),
                         cell_formula(r, 2, formula, 4),
                         cell_str(r, 3, sst.add(note), 0)]))
    section(4, "LIQUIDITY & LEVERAGE")
    ratio(5, "Current Ratio",
          "%s!%s/%s!%s" % (BS, bs_refs["ca"], BS, bs_refs["cl"]),
          "current assets / current liabilities; > 1.5 is comfortable")
    ratio(6, "Quick Ratio",
          "(%s!%s-%s!%s)/%s!%s" % (BS, bs_refs["ca"], BS, bs_refs["inv"], BS, bs_refs["cl"]),
          "excludes inventory; > 1.0 is healthy")
    ratio(7, "Debt to Equity",
          "%s!%s/%s!%s" % (BS, bs_refs["liab"], BS, bs_refs["equity"]),
          "total liabilities / equity; lower = less leveraged")
    ratio(8, "Equity Ratio",
          "%s!%s/%s!%s" % (BS, bs_refs["equity"], BS, bs_refs["assets"]),
          "share of assets funded by owners")
    section(10, "PROFITABILITY")
    ratio(11, "Gross Margin",
          "%s!%s/%s!%s" % (IS, is_refs["gross"], IS, is_refs["rev"]),
          "gross profit / revenue")
    ratio(12, "EBITDA Margin",
          "%s!%s/%s!%s" % (IS, is_refs["ebitda"], IS, is_refs["rev"]),
          "operating cash earnings / revenue")
    ratio(13, "Net Margin",
          "%s!%s/%s!%s" % (IS, is_refs["ni"], IS, is_refs["rev"]),
          "net income / revenue")
    ratio(14, "Return on Equity",
          "%s!%s/%s!%s" % (IS, is_refs["ni"], BS, bs_refs["equity"]),
          "net income / equity")
    ratio(15, "Return on Assets",
          "%s!%s/%s!%s" % (IS, is_refs["ni"], BS, bs_refs["assets"]),
          "net income / total assets")
    section(17, "COMMON-SIZE INCOME STATEMENT (% of Revenue)")
    for r, label, ref in [(18, "Revenue", is_refs["rev"]),
                          (19, "Cost of Goods Sold", is_refs["cogs"]),
                          (20, "Gross Profit", is_refs["gross"]),
                          (21, "Operating Expenses", is_refs["opex"]),
                          (22, "Net Income", is_refs["ni"])]:
        ratio(r, label, "%s!%s/%s!%s" % (IS, ref, IS, is_refs["rev"]),
              "each line as % of revenue")
    rows.append(row([cell_str(24, 1, sst.add(
        "All ratios are live formulas - revalue one account and every ratio "
        "moves with it."), 0)]))
    return (WS_OPEN + _cols_xml([24, 14, 52]) +
            '<sheetData>%s</sheetData></worksheet>' % "".join(rows),
            {"net_margin": "B13"})


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


def drawing_xml():
    anchors = []
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


def build_dashboard_sheet(sst, company, is_refs, bs_refs, cf_refs, ra_refs):
    rows = [row([cell_str(1, 1, sst.add("Executive Dashboard — " + FY_LABEL), 6)]),
            row([cell_str(2, 1, sst.add(company), 0)])]
    labels = ["Revenue (₹)", "Net Income (₹)", "Total Assets (₹)",
              "Cash & Bank (₹)", "Net Margin"]
    rows.append(row([cell_str(4, i + 1, sst.add(t), 7) for i, t in enumerate(labels)]))
    refs = ["%s!%s" % (IS, is_refs["rev"]), "%s!%s" % (IS, is_refs["ni"]),
            "%s!%s" % (BS, bs_refs["assets"]), "%s!%s" % (BS, bs_refs["cash"]),
            "%s!%s" % (RA, ra_refs["net_margin"])]
    styles = [8, 8, 8, 8, 4]
    rows.append(row([cell_formula(5, i + 1, ref, st)
                     for i, (ref, st) in enumerate(zip(refs, styles))]))
    rows.append(row([cell_str(6, 1, sst.add(
        "Every KPI above is a live formula linked to the three statements - "
        "change one account balance and the whole dashboard recalculates. "
        "Nothing is hard-coded."), 0)]))
    rows.append(row([cell_str(8, 1, sst.add("Chart data (live formulas)"), 7)]))
    mix = [("Total Revenue", is_refs["rev"]), ("Cost of Goods Sold", is_refs["cogs"]),
           ("Operating Expenses", is_refs["opex"]), ("Net Income", is_refs["ni"])]
    for i, (label, ref) in enumerate(mix):
        r = 9 + i
        rows.append(row([cell_str(r, 1, sst.add(label), 10),
                         cell_formula(r, 2, "%s!%s" % (IS, ref), 2)]))
    for i, (label, ref) in enumerate([("Current Assets", bs_refs["ca"]),
                                      ("Non-current Assets (net)", bs_refs["nca"])]):
        r = 14 + i
        rows.append(row([cell_str(r, 1, sst.add(label), 10),
                         cell_formula(r, 2, "%s!%s" % (BS, ref), 2)]))
    walk = [("Opening Cash", cf_refs["cash_beg"]), ("Operating Activities", cf_refs["op"]),
            ("Investing Activities", cf_refs["inv"]),
            ("Financing Activities", cf_refs["fin"]),
            ("Closing Cash", cf_refs["cash_end"])]
    for i, (label, ref) in enumerate(walk):
        r = 17 + i
        rows.append(row([cell_str(r, 1, sst.add(label), 10),
                         cell_formula(r, 2, "%s!%s" % (CF, ref), 2)]))
    return (WS_OPEN + _cols_xml([28, 20]) +
            '<sheetData>%s</sheetData></worksheet>' % "".join(rows))


# ---------------------------------------------------------------------------
# Workbook assembly + audit export
# ---------------------------------------------------------------------------

def _add_part(zf, name, text):
    zf.writestr(name, text.encode("utf-8"))


def build_workbook(accounts, company, output_path):
    sst = SharedStrings()
    tb_xml, rowmap, first, last = build_trial_balance_sheet(sst, accounts, company)
    is_xml, is_refs = build_income_statement_sheet(sst, company, first, last)
    bs_xml, bs_refs = build_balance_sheet_sheet(sst, company, first, last, is_refs)
    cf_xml, cf_refs = build_cashflow_sheet(sst, company, rowmap, is_refs, bs_refs)
    ra_xml, ra_refs = build_ratios_sheet(sst, company, is_refs, bs_refs)
    db_xml = build_dashboard_sheet(sst, company, is_refs, bs_refs, cf_refs, ra_refs)

    chart1 = chart_column_xml("Revenue vs Cost Structure — " + FY_LABEL,
                              "'Dashboard'!$A$9:$A$12", "'Dashboard'!$B$9:$B$12",
                              "Amount (₹)")
    chart2 = chart_pie_xml("Asset Mix", "'Dashboard'!$A$14:$A$15",
                           "'Dashboard'!$B$14:$B$15", "Assets (₹)")
    chart3 = chart_column_xml("Cash Walk — " + FY_LABEL,
                              "'Dashboard'!$A$17:$A$21", "'Dashboard'!$B$17:$B$21",
                              "Cash (₹)")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        _add_part(zf, "[Content_Types].xml", content_types_xml())
        _add_part(zf, "_rels/.rels", root_rels_xml())
        _add_part(zf, "xl/workbook.xml", workbook_xml())
        _add_part(zf, "xl/_rels/workbook.xml.rels", workbook_rels_xml())
        _add_part(zf, "xl/worksheets/sheet1.xml", tb_xml)
        _add_part(zf, "xl/worksheets/sheet2.xml", is_xml)
        _add_part(zf, "xl/worksheets/sheet3.xml", bs_xml)
        _add_part(zf, "xl/worksheets/sheet4.xml", cf_xml)
        _add_part(zf, "xl/worksheets/sheet5.xml", ra_xml)
        _add_part(zf, "xl/worksheets/sheet6.xml", db_xml)
        _add_part(zf, "xl/worksheets/_rels/sheet6.xml.rels", sheet6_rels_xml())
        _add_part(zf, "xl/drawings/drawing1.xml", drawing_xml())
        _add_part(zf, "xl/drawings/_rels/drawing1.xml.rels", drawing_rels_xml())
        _add_part(zf, "xl/charts/chart1.xml", chart1)
        _add_part(zf, "xl/charts/chart2.xml", chart2)
        _add_part(zf, "xl/charts/chart3.xml", chart3)
        _add_part(zf, "xl/styles.xml", styles_xml())
        _add_part(zf, "xl/sharedStrings.xml", sst.xml())
        _add_part(zf, "docProps/core.xml", core_props_xml())
    return output_path


def write_audit_csv(m, csv_path):
    """Independent Python-side audit export of the full model."""
    lines = [
        ("Income Statement", "Total Revenue", m["revenue"]),
        ("Income Statement", "Total COGS", m["cogs"]),
        ("Income Statement", "Gross Profit", m["gross"]),
        ("Income Statement", "Total Operating Expenses", m["opex"]),
        ("Income Statement", "EBITDA", m["ebitda"]),
        ("Income Statement", "EBIT", m["ebit"]),
        ("Income Statement", "Profit Before Tax", m["ebt"]),
        ("Income Statement", "Net Income", m["net_income"]),
        ("Balance Sheet", "Total Current Assets", m["cur_assets"]),
        ("Balance Sheet", "Net Non-current Assets", m["ncur_assets"]),
        ("Balance Sheet", "TOTAL ASSETS", m["total_assets"]),
        ("Balance Sheet", "Total Current Liabilities", m["cur_liab"]),
        ("Balance Sheet", "Total Non-current Liabilities", m["ncur_liab"]),
        ("Balance Sheet", "TOTAL LIABILITIES", m["total_liab"]),
        ("Balance Sheet", "Share Capital", m["share_cap"]),
        ("Balance Sheet", "Retained Earnings - Closing", m["re_close"]),
        ("Balance Sheet", "TOTAL EQUITY", m["total_equity"]),
        ("Balance Sheet", "BALANCE CHECK (must be 0)", m["balance_check"]),
        ("Cash Flow", "Cash from Operations", m["op_cf"]),
        ("Cash Flow", "Cash used in Investing", m["inv_cf"]),
        ("Cash Flow", "Cash used in Financing", m["fin_cf"]),
        ("Cash Flow", "Net Change in Cash", m["net_change"]),
        ("Cash Flow", "Cash at Beginning", m["cash_open"]),
        ("Cash Flow", "CASH AT END", m["cash_close"]),
        ("Cash Flow", "Tie-out vs BS cash (must be 0)", m["tie_out"]),
        ("Ratios", "Current Ratio", m["current_ratio"]),
        ("Ratios", "Quick Ratio", m["quick_ratio"]),
        ("Ratios", "Debt to Equity", m["debt_equity"]),
        ("Ratios", "Gross Margin", m["gross_margin"]),
        ("Ratios", "EBITDA Margin", m["ebitda_margin"]),
        ("Ratios", "Net Margin", m["net_margin"]),
        ("Ratios", "Return on Equity", m["roe"]),
        ("Ratios", "Return on Assets", m["roa"]),
    ]
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["Statement", "Line", "Amount"])
        for stmt, line_name, amount in lines:
            w.writerow([stmt, line_name, "%.2f" % amount])
    return csv_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _parse_args(argv):
    p = argparse.ArgumentParser(prog="statements_forge",
                                description="Three-statement financial model "
                                            "+ executive dashboard factory")
    sub = p.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("generate", help="write a deterministic sample CSV")
    g.add_argument("--out", default="sample_company_2026.csv")

    b = sub.add_parser("build", help="build the .xlsx workbook + audit CSV")
    b.add_argument("--csv", required=True)
    b.add_argument("--name", default=COMPANY)
    b.add_argument("--out", default="three_statement_model.xlsx")
    b.add_argument("--audit-csv", default="audit.csv")
    return p.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv or sys.argv[1:])
    if args.cmd == "generate":
        path = generate_sample(args.out)
        print("sample CSV -> %s" % path)
        return 0
    accounts = load_accounts(args.csv)
    m = model(accounts)
    build_workbook(accounts, args.name, args.out)
    write_audit_csv(m, args.audit_csv)
    print("workbook -> %s" % args.out)
    print("audit CSV -> %s" % args.audit_csv)
    print("Net Income ₹%s | Total Assets ₹%s | Cash ₹%s | balance check %s" % (
        _fmt_num(round(m["net_income"])), _fmt_num(round(m["total_assets"])),
        _fmt_num(round(m["cash_close"])), _fmt_num(round(m["balance_check"]))))
    return 0


if __name__ == "__main__":
    sys.exit(main())
