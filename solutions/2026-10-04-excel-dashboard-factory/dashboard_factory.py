#!/usr/bin/env python3
"""DashboardForge — stdlib-only Excel dashboard factory.

Reads a raw monthly-sales CSV and emits a fully formatted, multi-sheet
.xlsx workbook (real OOXML, no third-party libraries):

  Sheet 1 "Raw Data"        — the sales table: header styling, AutoFilter,
                              frozen panes, currency/date/number formats.
  Sheet 2 "Monthly Summary" — pivot-style summary with real Excel SUMIFS /
                              COUNTIFS formulas, a TOTAL row and
                              conditional formatting (top-3 months).
  Sheet 3 "Dashboard"       — title, KPI cells wired to the summary with
                              formulas, and three native Excel charts
                              (clustered column, line, pie) via DrawingML.

Usage:
    python dashboard_factory.py --input sales.csv --output SalesDashboard.xlsx
    python dashboard_factory.py --generate-sample sample_sales_2026.csv
    python dashboard_factory.py --input sample_sales_2026.csv --output out.xlsx

The formulas are real Excel formulas: they calculate on open in Excel /
LibreOffice. No macros or VBA are required.
"""

import argparse
import csv
import datetime
import os
import random
import sys
import zipfile
from xml.sax.saxutils import escape as _xml_escape

TOOL_NAME = "DashboardForge"
SAMPLE_SEED = 20261004

REQUIRED_COLUMNS = ["OrderID", "Date", "Region", "Category", "Product",
                    "Units", "UnitPrice"]

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


# ---------------------------------------------------------------------------
# Shared strings table
# ---------------------------------------------------------------------------

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
<numFmt numFmtId="165" formatCode="$#,##0"/>
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
<dxfs count="1">
<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill></dxf>
</dxfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


# ---------------------------------------------------------------------------
# Package-level parts
# ---------------------------------------------------------------------------

def content_types_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""


def root_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def workbook_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Raw Data" sheetId="1" r:id="rId1"/>
<sheet name="Monthly Summary" sheetId="2" r:id="rId2"/>
<sheet name="Dashboard" sheetId="3" r:id="rId3"/>
</sheets>
<calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>"""


def workbook_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""


def sheet3_rels_xml():
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
<dc:title>Monthly Sales Dashboard</dc:title>
<cp:keywords>sales dashboard,excel,power bi,automation</cp:keywords>
</cp:coreProperties>""" % TOOL_NAME

# ---------------------------------------------------------------------------
# Input: load + validate the raw sales CSV
# ---------------------------------------------------------------------------

def load_sales(csv_path):
    """Read and validate the raw sales CSV. Returns a list of dict rows."""
    if not os.path.exists(csv_path):
        raise ValueError("input file not found: %s" % csv_path)
    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        if reader.fieldnames is None:
            raise ValueError("input CSV has no header row: %s" % csv_path)
        missing = [c for c in REQUIRED_COLUMNS if c not in reader.fieldnames]
        if missing:
            raise ValueError("input CSV is missing required columns: %s "
                             "(found: %s)" % (", ".join(missing),
                                              ", ".join(reader.fieldnames)))
        rows = []
        for lineno, rec in enumerate(reader, start=2):
            try:
                day = datetime.date.fromisoformat(rec["Date"].strip())
            except (ValueError, AttributeError):
                raise ValueError("row %d: bad Date %r (expected YYYY-MM-DD)"
                                 % (lineno, rec.get("Date")))
            try:
                units = int(str(rec["Units"]).strip())
            except ValueError:
                raise ValueError("row %d: bad Units %r (expected integer)"
                                 % (lineno, rec.get("Units")))
            try:
                unit_price = float(str(rec["UnitPrice"]).strip())
            except ValueError:
                raise ValueError("row %d: bad UnitPrice %r (expected number)"
                                 % (lineno, rec.get("UnitPrice")))
            if units <= 0:
                raise ValueError("row %d: Units must be > 0" % lineno)
            if unit_price <= 0:
                raise ValueError("row %d: UnitPrice must be > 0" % lineno)
            rows.append({
                "OrderID": str(rec["OrderID"]).strip(),
                "Date": day,
                "Region": str(rec["Region"]).strip(),
                "Category": str(rec["Category"]).strip(),
                "Product": str(rec["Product"]).strip(),
                "Units": units,
                "UnitPrice": unit_price,
                "Revenue": round(units * unit_price, 2),
            })
    if not rows:
        raise ValueError("input CSV has no data rows: %s" % csv_path)
    return rows


def month_starts(rows):
    """Sorted list of first-of-month dates covering the data."""
    months = sorted({datetime.date(r["Date"].year, r["Date"].month, 1)
                     for r in rows})
    return months


# ---------------------------------------------------------------------------
# Deterministic sample data
# ---------------------------------------------------------------------------

def generate_sample(csv_path, seed=SAMPLE_SEED, year=2026):
    """Write a deterministic 12-month sample sales CSV."""
    rng = random.Random(seed)
    regions = ["North", "South", "East", "West"]
    catalog = {
        "Electronics": [("Nova X1 Phone", 299.0), ("Volt Laptop 14", 749.0),
                        ("Echo Buds Pro", 79.0)],
        "Apparel": [("Urban Tee", 24.0), ("Denim Jacket", 89.0),
                    ("Runner Shoes", 119.0)],
        "Home": [("Aroma Diffuser", 39.0), ("Steel Cookware Set", 149.0),
                 ("LED Desk Lamp", 45.0)],
        "Grocery": [("Organic Honey 500g", 12.0), ("Almond Pack 1kg", 18.0),
                    ("Cold-Pressed Oil 1L", 15.0)],
    }
    reps = ["A. Sharma", "P. Verma", "R. Iyer", "S. Khan", "M. D'Souza"]
    categories = list(catalog)
    order_no = 1000
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["OrderID", "Date", "Region", "Category", "Product",
                         "Units", "UnitPrice", "SalesRep"])
        for month in range(1, 13):
            # seasonal shape: festive spike Oct-Dec, dip in Feb
            seasonal = {1: 0.85, 2: 0.7, 3: 0.9, 4: 0.95, 5: 1.0, 6: 0.95,
                        7: 1.0, 8: 1.05, 9: 1.1, 10: 1.35, 11: 1.5,
                        12: 1.45}[month]
            n_orders = int(rng.uniform(38, 52) * seasonal)
            for _ in range(n_orders):
                order_no += 1
                cat = rng.choice(categories)
                product, base_price = rng.choice(catalog[cat])
                units = rng.randint(1, 12)
                price = round(base_price * rng.uniform(0.9, 1.15), 2)
                day = rng.randint(1, 28)
                writer.writerow([
                    "ORD-%d" % order_no,
                    "%d-%02d-%02d" % (year, month, day),
                    rng.choice(regions), cat, product, units, price,
                    rng.choice(reps),
                ])
    return csv_path


# ---------------------------------------------------------------------------
# Worksheet builders
# ---------------------------------------------------------------------------

RAW_HEADERS = ["OrderID", "Date", "Region", "Category", "Product",
               "Units", "UnitPrice", "Revenue"]
RAW_WIDTHS = [12, 12, 10, 13, 22, 8, 12, 13]

SUMMARY_HEADERS = ["Month", "Revenue", "Units", "Orders", "Avg Order Value"]


def _cols_xml(widths):
    parts = ["<cols>"]
    for i, w in enumerate(widths, start=1):
        parts.append('<col min="%d" max="%d" width="%s" customWidth="1"/>'
                     % (i, i, w))
    parts.append("</cols>")
    return "".join(parts)


def row(cells):
    return "<row>" + "".join(cells) + "</row>"


def build_raw_sheet_rows(sst, rows):
    header = [cell_str(1, c, sst.add(h), style=1)
              for c, h in enumerate(RAW_HEADERS, start=1)]
    out = [row(header)]
    for r, rec in enumerate(rows, start=2):
        out.append(row([
            cell_str(r, 1, sst.add(rec["OrderID"]), style=10),
            cell_num(r, 2, _excel_serial(rec["Date"]), style=3),
            cell_str(r, 3, sst.add(rec["Region"]), style=10),
            cell_str(r, 4, sst.add(rec["Category"]), style=10),
            cell_str(r, 5, sst.add(rec["Product"]), style=10),
            cell_num(r, 6, rec["Units"], style=5),
            cell_num(r, 7, rec["UnitPrice"], style=2),
            cell_num(r, 8, rec["Revenue"], style=2),
        ]))
    return "".join(out)


def raw_sheet_xml(sst, rows):
    body = build_raw_sheet_rows(sst, rows)
    last = _cell_ref(len(rows) + 1, len(RAW_HEADERS))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            '</sheetView></sheetViews>'
            + _cols_xml(RAW_WIDTHS) +
            '<sheetData>' + body + '</sheetData>'
            '<autoFilter ref="A1:%s"/>' % last +
            '<pageSetup fitToWidth="1" fitToHeight="0" orientation="landscape"/>'
            '</worksheet>')


def build_summary_sheet(sst, months, n_raw_rows):
    """Sheet 2: monthly summary with real SUMIFS/COUNTIFS formulas.

    n_raw_rows: number of data rows on the Raw Data sheet (excl. header).
    """
    raw_last = n_raw_rows + 1
    date_col = "'Raw Data'!$B$2:$B$%d" % raw_last
    units_col = "'Raw Data'!$F$2:$F$%d" % raw_last
    revenue_col = "'Raw Data'!$H$2:$H$%d" % raw_last

    header = [cell_str(1, c, sst.add(h), style=1)
              for c, h in enumerate(SUMMARY_HEADERS, start=1)]
    out = [row(header)]
    for i, month in enumerate(months):
        r = i + 2
        month_ge = '">="&DATE(%d,%d,1)' % (month.year, month.month)
        month_lt = '"<"&EDATE(DATE(%d,%d,1),1)' % (month.year, month.month)
        crit = '%s,%s,%s,%s' % (date_col, month_ge, date_col, month_lt)
        out.append(row([
            cell_num(r, 1, _excel_serial(month), style=9),
            cell_formula(r, 2, "SUMIFS(%s,%s)" % (revenue_col, crit), style=2),
            cell_formula(r, 3, "SUMIFS(%s,%s)" % (units_col, crit), style=5),
            cell_formula(r, 4, "COUNTIFS(%s)" % crit, style=5),
            cell_formula(r, 5, "IF(C%d=0,0,B%d/C%d)" % (r, r, r), style=2),
        ]))
    total_row = len(months) + 2
    out.append(row([
        cell_str(total_row, 1, sst.add("TOTAL"), style=1),
        cell_formula(total_row, 2, "SUM(B2:B%d)" % (total_row - 1), style=2),
        cell_formula(total_row, 3, "SUM(C2:C%d)" % (total_row - 1), style=5),
        cell_formula(total_row, 4, "SUM(D2:D%d)" % (total_row - 1), style=5),
        cell_formula(total_row, 5, "IF(C%d=0,0,B%d/C%d)" % (total_row,
                                                            total_row,
                                                            total_row),
                      style=2),
    ]))
    body = "".join(out)
    cf_range = "B2:B%d" % (total_row - 1)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetViews><sheetView workbookViewId="0" showGridLines="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews>'
            + _cols_xml([14, 16, 12, 12, 16]) +
            '<sheetData>' + body + '</sheetData>'
            '<conditionalFormatting sqref="%s">'
            '<cfRule type="top10" dxfId="0" priority="1" rank="3" percent="0"/>'
            '</conditionalFormatting>' % cf_range +
            '<pageSetup fitToWidth="1" fitToHeight="1" orientation="landscape"/>'
            '</worksheet>')

# ---------------------------------------------------------------------------
# Charts (DrawingML)
# ---------------------------------------------------------------------------

_CHART_NS = ('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
             'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def _chart_title_xml(title):
    return ('<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>%s</a:t></a:r>'
            '</a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>'
            % _esc(title))


def _chart_series_xml(idx, name_cell, cat_range, val_range):
    return ('<c:ser><c:idx val="%d"/><c:order val="%d"/>'
            '<c:tx><c:strRef><c:f>%s</c:f></c:strRef></c:tx>'
            '<c:cat><c:strRef><c:f>%s</c:f></c:strRef></c:cat>'
            '<c:val><c:numRef><c:f>%s</c:f></c:numRef></c:val>'
            '</c:ser>' % (idx, idx, _esc(name_cell), _esc(cat_range),
                          _esc(val_range)))


_AXES_XML = ('<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/>'
             '</c:scaling><c:delete val="0"/><c:axPos val="b"/>'
             '<c:tickLblPos val="nextTo"/><c:crossAx val="222"/></c:catAx>'
             '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/>'
             '</c:scaling><c:delete val="0"/><c:axPos val="l"/>'
             '<c:numFmt formatCode="$#,##0" sourceLinked="0"/>'
             '<c:tickLblPos val="nextTo"/><c:crossAx val="111"/></c:valAx>')


def chart_column_xml(title, name_cell, cat_range, val_range):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<c:chartSpace %s><c:chart>%s<c:plotArea><c:layout/>'
            '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>'
            '%s<c:gapWidth val="150"/><c:axId val="111"/><c:axId val="222"/>'
            '</c:barChart>%s</c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
            '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
            % (_CHART_NS, _chart_title_xml(title),
               _chart_series_xml(0, name_cell, cat_range, val_range),
               _AXES_XML))


def chart_line_xml(title, name_cell, cat_range, val_range):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<c:chartSpace %s><c:chart>%s<c:plotArea><c:layout/>'
            '<c:lineChart><c:grouping val="standard"/>'
            '%s<c:smooth val="1"/><c:axId val="111"/><c:axId val="222"/>'
            '</c:lineChart>%s</c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
            '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
            % (_CHART_NS, _chart_title_xml(title),
               _chart_series_xml(0, name_cell, cat_range, val_range),
               _AXES_XML))


def chart_pie_xml(title, name_cell, cat_range, val_range):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<c:chartSpace %s><c:chart>%s<c:plotArea><c:layout/>'
            '<c:pieChart><c:varyColors val="1"/>'
            '%s<c:dLbls><c:showPercent val="1"/><c:showLeaderLines val="1"/>'
            '</c:dLbls></c:pieChart></c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
            '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
            % (_CHART_NS, _chart_title_xml(title),
               _chart_series_xml(0, name_cell, cat_range, val_range)))


def drawing_xml():
    anchors = []
    # (chart rel id, from col/row, to col/row) — 0-based
    placements = [("rId1", 0, 5, 8, 21),
                  ("rId2", 0, 23, 8, 39),
                  ("rId3", 0, 41, 8, 57)]
    for i, (rid, c1, r1, c2, r2) in enumerate(placements):
        anchors.append(
            '<xdr:twoCellAnchor>'
            '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
            '<xdr:cNvPr id="%d" name="Chart %d"/></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'r:id="%s"/></a:graphicData></a:graphic></xdr:graphicFrame>'
            '<xdr:clientData/></xdr:twoCellAnchor>'
            % (c1, r1, c2, r2, i + 2, i + 1, rid))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            + "".join(anchors) + '</xdr:wsDr>')


# ---------------------------------------------------------------------------
# Sheet 3: Dashboard
# ---------------------------------------------------------------------------

def dashboard_sheet_xml(sst, total_row):
    """KPI cells wired to the Monthly Summary TOTAL row + chart drawing."""
    title = sst.add("Monthly Sales Dashboard")
    subtitle = sst.add("Auto-generated by %s — source: Raw Data sheet; "
                       "all figures are live Excel formulas." % TOOL_NAME)
    kpi_labels = ["Total Revenue", "Total Units", "Total Orders",
                  "Avg Order Value"]
    kpi_formulas = ["'Monthly Summary'!B%d" % total_row,
                    "'Monthly Summary'!C%d" % total_row,
                    "'Monthly Summary'!D%d" % total_row,
                    "'Monthly Summary'!E%d" % total_row]
    kpi_styles = [8, 5, 5, 2]
    cols = [s for pair in zip([1, 3, 5, 7], [2, 4, 6, 8]) for s in pair]
    out = [row([cell_str(1, 1, title, style=6)]),
           row([cell_str(2, 1, subtitle, style=0)])]
    label_cells, value_cells = [], []
    for i, (label, formula, vstyle) in enumerate(
            zip(kpi_labels, kpi_formulas, kpi_styles)):
        c = cols[2 * i]
        label_cells.append(cell_str(3, c, sst.add(label), style=7))
        value_cells.append(cell_formula(4, c, formula, style=vstyle))
    out.append(row(label_cells))
    out.append(row(value_cells))
    body = "".join(out)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetViews><sheetView workbookViewId="0" showGridLines="0"/>'
            '</sheetViews>'
            + _cols_xml([18, 18, 18, 18, 18, 18, 18, 18]) +
            '<sheetData>' + body + '</sheetData>'
            '<mergeCells count="1"><mergeCell ref="A1:H1"/></mergeCells>'
            '<drawing r:id="rId1"/>'
            '<pageSetup fitToWidth="1" fitToHeight="1" orientation="landscape"/>'
            '</worksheet>')


# ---------------------------------------------------------------------------
# Package assembly
# ---------------------------------------------------------------------------

def _add_part(zf, name, text):
    info = zipfile.ZipInfo(name, date_time=(2026, 10, 4, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    zf.writestr(info, text.encode("utf-8"))


def build_workbook(rows, output_path, title="Monthly Sales Dashboard"):
    months = month_starts(rows)
    n_months = len(months)
    total_row = n_months + 2
    last_summary = n_months + 1

    sst = SharedStrings()
    raw_xml = raw_sheet_xml(sst, rows)
    summary_xml = build_summary_sheet(sst, months, len(rows))
    dash_xml = dashboard_sheet_xml(sst, total_row)

    cat_range = "'Monthly Summary'!$A$2:$A$%d" % last_summary
    chart1 = chart_column_xml("Monthly Revenue", "'Monthly Summary'!$B$1",
                              cat_range,
                              "'Monthly Summary'!$B$2:$B$%d" % last_summary)
    chart2 = chart_line_xml("Orders Trend", "'Monthly Summary'!$D$1",
                            cat_range,
                            "'Monthly Summary'!$D$2:$D$%d" % last_summary)
    chart3 = chart_pie_xml("Revenue by Month", "'Monthly Summary'!$B$1",
                           cat_range,
                           "'Monthly Summary'!$B$2:$B$%d" % last_summary)

    with zipfile.ZipFile(output_path, "w") as zf:
        _add_part(zf, "[Content_Types].xml", content_types_xml())
        _add_part(zf, "_rels/.rels", root_rels_xml())
        _add_part(zf, "docProps/core.xml", core_props_xml())
        _add_part(zf, "xl/workbook.xml", workbook_xml())
        _add_part(zf, "xl/_rels/workbook.xml.rels", workbook_rels_xml())
        _add_part(zf, "xl/worksheets/sheet1.xml", raw_xml)
        _add_part(zf, "xl/worksheets/sheet2.xml", summary_xml)
        _add_part(zf, "xl/worksheets/sheet3.xml", dash_xml)
        _add_part(zf, "xl/worksheets/_rels/sheet3.xml.rels", sheet3_rels_xml())
        _add_part(zf, "xl/drawings/drawing1.xml", drawing_xml())
        _add_part(zf, "xl/drawings/_rels/drawing1.xml.rels",
                   drawing_rels_xml())
        _add_part(zf, "xl/charts/chart1.xml", chart1)
        _add_part(zf, "xl/charts/chart2.xml", chart2)
        _add_part(zf, "xl/charts/chart3.xml", chart3)
        _add_part(zf, "xl/styles.xml", styles_xml())
        _add_part(zf, "xl/sharedStrings.xml", sst.xml())
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="%s: build a formatted multi-sheet Excel sales "
                    "dashboard from a raw sales CSV (stdlib only)." % TOOL_NAME)
    parser.add_argument("--input", help="raw sales CSV path")
    parser.add_argument("--output", default="SalesDashboard.xlsx",
                        help="output .xlsx path")
    parser.add_argument("--generate-sample", metavar="PATH",
                        help="write deterministic sample sales CSV and exit")
    args = parser.parse_args(argv)

    if args.generate_sample:
        path = generate_sample(args.generate_sample)
        print("sample CSV written: %s" % path)
        return 0
    if not args.input:
        parser.error("--input is required (or use --generate-sample)")
    rows = load_sales(args.input)
    out = build_workbook(rows, args.output)
    months = month_starts(rows)
    total_revenue = sum(r["Revenue"] for r in rows)
    print("rows: %d | months: %d | total revenue: $%s" %
          (len(rows), len(months), f"{total_revenue:,.2f}"))
    print("dashboard written: %s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
