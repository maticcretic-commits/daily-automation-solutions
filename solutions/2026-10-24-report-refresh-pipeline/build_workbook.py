#!/usr/bin/env python3
"""RefreshForge — rebuild the formula-driven refresh-report workbook.

Reads history/*.csv + refresh_log.json (see refresh.py) and writes
refresh-report.xlsx:

  Guide         — what this is, the pipeline, Power Automate / Power BI map
  Refresh_Log   — every pipeline run (audit values written by refresh.py)
  Raw_Sales     — cumulative sales history + formula revenue column
  Raw_Production— cumulative batches + formula yield column
  Raw_Inventory — cumulative movements + formula signed-qty column
  Raw_Issues    — cumulative issue log (values)
  Daily         — per-day SUMIFS/COUNTIFS summaries (Oct 2026)
  Dashboard     — KPI cards (all formula-linked) + 3 native charts
  Checks        — integrity checks, all must read TRUE

Every computed cell is a live Excel formula over headroom ranges
($A$2:$A$2000). A future refresh appends rows to history/, the workbook is
rebuilt, and every KPI/chart/check recalculates — the same shape as the
client's Power Automate scheduled flow. Plain A1 ranges, no Table objects,
so it recalculates in Excel, LibreOffice and Google Sheets.
"""
import csv
import json
import os
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
HIST = os.path.join(BASE, "history")
OUT = os.path.join(BASE, "refresh-report.xlsx")
LOG = os.path.join(BASE, "refresh_log.json")

H = 2000          # headroom: data rows live in 2..H on every Raw sheet
NDAYS = 31        # Daily covers all of Oct 2026

HDR_FILL = PatternFill("solid", fgColor="1F4E5F")
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color="1F4E5F")
SUB_FONT = Font(bold=True, size=11, color="1F4E5F")
MONEY_FMT = '₹#,##0'
QTY_FMT = '#,##0'
PCT_FMT = '0.0%'
DATE_FMT = 'dd-mmm-yy'
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006")


def d(s):
    return date.fromisoformat(s)


def read_hist(name):
    path = os.path.join(HIST, name)
    if not os.path.exists(path):
        return []
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def read_log():
    if not os.path.exists(LOG):
        return []
    with open(LOG) as f:
        return json.load(f)["runs"]


def read_skus():
    path = os.path.join(BASE, "data", "skus.csv")
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def write_values(ws, rows, cols, date_cols=()):
    """Write raw history rows as values starting at row 2. Returns row count."""
    for r, row in enumerate(rows, start=2):
        for c, col in enumerate(cols, start=1):
            v = row[col]
            if col in date_cols:
                v = d(v)
            elif v.lstrip("-").replace(".", "", 1).isdigit():
                v = float(v) if "." in v else int(v)
            cell = ws.cell(row=r, column=c, value=v)
            if col in date_cols:
                cell.number_format = DATE_FMT
    return len(rows)


def build(base_dir=None):
    base = os.path.abspath(base_dir or BASE)
    global HIST, LOG, OUT
    HIST = os.path.join(base, "history")
    LOG = os.path.join(base, "refresh_log.json")
    OUT = os.path.join(base, "refresh-report.xlsx")

    sales = read_hist("sales.csv")
    prod = read_hist("production.csv")
    inv = read_hist("inventory.csv")
    issues = read_hist("issues.csv")
    runs = read_log()
    skus = read_skus()
    opening_total = sum(int(s["opening_qty"]) for s in skus)

    wb = Workbook()

    # ================= Guide =================
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "RefreshForge — automated report-refresh pipeline (demo)"
    ws["A1"].font = TITLE_FONT
    lines = [
        "Fictional demo data: 'Sharma Foods Processing, Patna' (food plant).",
        "Four departmental exports land in drop/; refresh.py validates, dedupes",
        "and appends only new rows to history/, then this workbook is rebuilt.",
        "",
        "RUN THE PIPELINE",
        "1. python3 generate.py   — (re)create the synthetic drop/ files",
        "2. python3 refresh.py    — validate + dedupe + append + audit log",
        "3. python3 build_workbook.py — rebuild this workbook",
        "Re-running refresh.py on unchanged files appends ZERO rows (idempotent);",
        "dropping new daily exports in and re-running appends exactly the delta.",
        "",
        "POWER AUTOMATE MAPPING (what the client asked for)",
        "Recurrence trigger (every 30 min) → List files in SharePoint folder →",
        "For each: Get file content → Parse CSV → Apply to each row:",
        "  validate (this pipeline's validate_row) → Condition: key seen?",
        "  yes → skip · no → Add row to history table → Compose refresh log →",
        "  Refresh Power BI dataset. The Refresh_Log sheet is the audit the flow",
        "  would write; powerbi/MEASURES.md holds the DAX for the .pbix.",
        "",
        "FORMULAS, NOT VALUES",
        "Every computed cell is a live formula over $2:$2000 headroom ranges, so",
        "future refreshes flow straight through to Daily, Dashboard and Checks.",
        "Checks!B must all read TRUE after any recalculation.",
    ]
    for i, line in enumerate(lines, start=3):
        ws[f"A{i}"] = line
    ws.column_dimensions["A"].width = 95

    # ================= Refresh_Log =================
    rl = wb.create_sheet("Refresh_Log")
    headers = ["run_id", "timestamp", "source", "rows_seen", "rows_new",
               "duplicates", "quarantined", "status"]
    for c, h in enumerate(headers, start=1):
        rl.cell(row=1, column=c, value=h)
    style_header(rl, 1, len(headers))
    r = 2
    for run in runs:
        for name in ("sales", "production", "inventory", "issues"):
            s = run["sources"].get(name)
            if not s:
                continue
            rl.cell(row=r, column=1, value=run["run_id"])
            rl.cell(row=r, column=2, value=run["timestamp"])
            rl.cell(row=r, column=3, value=name)
            rl.cell(row=r, column=4, value=s["seen"]).number_format = QTY_FMT
            rl.cell(row=r, column=5, value=s["new"]).number_format = QTY_FMT
            rl.cell(row=r, column=6,
                    value=s["duplicates"]).number_format = QTY_FMT
            rl.cell(row=r, column=7,
                    value=s["quarantined"]).number_format = QTY_FMT
            rl.cell(row=r, column=8, value=run["status"])
            r += 1
    rl.freeze_panes = "A2"
    if r > 2:
        rl.auto_filter.ref = f"A1:H{r - 1}"
    for c, w in zip("ABCDEFGH", [9, 20, 11, 11, 10, 11, 12, 20]):
        rl.column_dimensions[c].width = w
    lastlog = max(r - 1, 2)

    # ================= Raw sheets =================
    def raw_sheet(title, value_cols, all_cols, date_cols, formulas, widths,
                  rows):
        w = wb.create_sheet(title)
        for c, h in enumerate(all_cols, start=1):
            w.cell(row=1, column=c, value=h)
        style_header(w, 1, len(all_cols))
        n = write_values(w, rows, value_cols, date_cols)
        for rr in range(2, H + 1):
            for c, f in formulas.items():
                w.cell(row=rr, column=c, value=f(rr))
        w.freeze_panes = "A2"
        w.auto_filter.ref = f"A1:{get_column_letter(len(all_cols))}{H}"
        for c, wd in zip([get_column_letter(i + 1)
                          for i in range(len(all_cols))], widths):
            w.column_dimensions[c].width = wd
        return w, n

    rs, n_sales = raw_sheet(
        "Raw_Sales",
        ["date", "order_id", "customer", "sku", "qty", "unit_price"],
        ["date", "order_id", "customer", "sku", "qty", "unit_price",
         "revenue"], {"date"},
        {7: lambda r: f"=E{r}*F{r}"},
        [12, 10, 20, 11, 8, 11, 13], sales)
    for rr in range(2, H + 1):
        rs.cell(row=rr, column=5).number_format = QTY_FMT
        rs.cell(row=rr, column=6).number_format = MONEY_FMT
        rs.cell(row=rr, column=7).number_format = MONEY_FMT

    rp, n_prod = raw_sheet(
        "Raw_Production",
        ["date", "batch_id", "line", "sku", "qty_produced", "qty_rejected"],
        ["date", "batch_id", "line", "sku", "qty_produced", "qty_rejected",
         "yield"], {"date"},
        {7: lambda r: f"=IF((E{r}+F{r})=0,1,E{r}/(E{r}+F{r}))"},
        [12, 10, 9, 11, 13, 12, 9], prod)
    for rr in range(2, H + 1):
        rp.cell(row=rr, column=5).number_format = QTY_FMT
        rp.cell(row=rr, column=6).number_format = QTY_FMT
        rp.cell(row=rr, column=7).number_format = PCT_FMT

    ri, n_inv = raw_sheet(
        "Raw_Inventory",
        ["date", "txn_id", "sku", "movement", "qty"],
        ["date", "txn_id", "sku", "movement", "qty", "signed_qty"], {"date"},
        {6: lambda r: f'=IF(D{r}="IN",E{r},IF(D{r}="OUT",-E{r},0))'},
        [12, 10, 11, 10, 8, 11], inv)
    for rr in range(2, H + 1):
        ri.cell(row=rr, column=5).number_format = QTY_FMT
        ri.cell(row=rr, column=6).number_format = QTY_FMT

    rq, n_issues = raw_sheet(
        "Raw_Issues",
        ["date", "issue_id", "category", "severity", "status",
         "description"],
        ["date", "issue_id", "category", "severity", "status",
         "description"], {"date"},
        {},
        [12, 10, 11, 10, 9, 34], issues)

    # ================= Daily =================
    dy = wb.create_sheet("Daily")
    dheaders = ["date", "sales_revenue", "production_qty", "yield",
                "inventory_net", "issues", "open_issues"]
    for c, h in enumerate(dheaders, start=1):
        dy.cell(row=1, column=c, value=h)
    style_header(dy, 1, len(dheaders))
    D1, D2 = 2, 1 + NDAYS
    for i in range(NDAYS):
        r = D1 + i
        dy.cell(row=r, column=1,
                value=date(2026, 10, 1 + i)).number_format = DATE_FMT
        dy.cell(row=r, column=2, value=(
            f"=SUMIFS(Raw_Sales!$G$2:$G${H},Raw_Sales!$A$2:$A${H},$A{r})")
            ).number_format = MONEY_FMT
        dy.cell(row=r, column=3, value=(
            f"=SUMIFS(Raw_Production!$E$2:$E${H},"
            f"Raw_Production!$A$2:$A${H},$A{r})")).number_format = QTY_FMT
        dy.cell(row=r, column=4, value=(
            f"=IF((SUMIFS(Raw_Production!$E$2:$E${H},Raw_Production!$A$2:$A${H},$A{r})"
            f"+SUMIFS(Raw_Production!$F$2:$F${H},Raw_Production!$A$2:$A${H},$A{r}))=0,1,"
            f"SUMIFS(Raw_Production!$E$2:$E${H},Raw_Production!$A$2:$A${H},$A{r})"
            f"/(SUMIFS(Raw_Production!$E$2:$E${H},Raw_Production!$A$2:$A${H},$A{r})"
            f"+SUMIFS(Raw_Production!$F$2:$F${H},Raw_Production!$A$2:$A${H},$A{r})))")
            ).number_format = PCT_FMT
        dy.cell(row=r, column=5, value=(
            f"=SUMIFS(Raw_Inventory!$F$2:$F${H},"
            f"Raw_Inventory!$A$2:$A${H},$A{r})")).number_format = QTY_FMT
        dy.cell(row=r, column=6, value=(
            f"=COUNTIFS(Raw_Issues!$A$2:$A${H},$A{r})")
            ).number_format = QTY_FMT
        dy.cell(row=r, column=7, value=(
            f'=COUNTIFS(Raw_Issues!$A$2:$A${H},$A{r},'
            f'Raw_Issues!$E$2:$E${H},"Open")')).number_format = QTY_FMT
    dy.conditional_formatting.add(f"F{D1}:F{D2}",
                                  CellIsRule(operator="greaterThan",
                                             formula=["2"], fill=RED_FILL,
                                             font=RED_FONT))
    dy.conditional_formatting.add(f"D{D1}:D{D2}",
                                  CellIsRule(operator="lessThan",
                                             formula=["0.95"], fill=RED_FILL,
                                             font=RED_FONT))
    dy.freeze_panes = "A2"
    for c, w in zip("ABCDEFG", [12, 14, 14, 9, 13, 8, 11]):
        dy.column_dimensions[c].width = w

    # ================= Dashboard =================
    dash = wb.create_sheet("Dashboard")
    dash["A1"] = "Dashboard — KPIs (all formula-linked)"
    dash["A1"].font = TITLE_FONT
    kpis = [
        ("MTD sales revenue", f"=SUM(Raw_Sales!$G$2:$G${H})", MONEY_FMT),
        ("MTD production qty", f"=SUM(Raw_Production!$E$2:$E${H})", QTY_FMT),
        ("Average yield",
         f"=IF((SUM(Raw_Production!$E$2:$E${H})+SUM(Raw_Production!$F$2:$F${H}))=0,1,"
         f"SUM(Raw_Production!$E$2:$E${H})/(SUM(Raw_Production!$E$2:$E${H})"
         f"+SUM(Raw_Production!$F$2:$F${H})))", PCT_FMT),
        ("Open issues",
         f'=COUNTIF(Raw_Issues!$E$2:$E${H},"Open")', QTY_FMT),
        ("Inventory on hand (units)",
         f"=B20+SUM(Raw_Inventory!$F$2:$F${H})", QTY_FMT),
        ("Rows in history",
         f"=COUNTA(Raw_Sales!$B$2:$B${H})+COUNTA(Raw_Production!$B$2:$B${H})"
         f"+COUNTA(Raw_Inventory!$B$2:$B${H})+COUNTA(Raw_Issues!$B$2:$B${H})",
         QTY_FMT),
        ("Last refresh run",
         f"=INDEX(Refresh_Log!$B$2:$B${lastlog},COUNTA(Refresh_Log!$A$2:$A${lastlog}))",
         "General"),
        ("Last refresh status",
         f"=INDEX(Refresh_Log!$H$2:$H${lastlog},COUNTA(Refresh_Log!$A$2:$A${lastlog}))",
         "General"),
        ("Rows added in last refresh",
         f"=SUMIFS(Refresh_Log!$E$2:$E${lastlog},Refresh_Log!$A$2:$A${lastlog},"
         f"INDEX(Refresh_Log!$A$2:$A${lastlog},COUNTA(Refresh_Log!$A$2:$A${lastlog})))",
         QTY_FMT),
    ]
    for i, (label, formula, fmt) in enumerate(kpis, start=3):
        dash[f"A{i}"] = label
        dash[f"A{i}"].font = Font(bold=True)
        dash[f"B{i}"] = formula
        dash[f"B{i}"].number_format = fmt
        dash[f"B{i}"].font = Font(size=13, color="1F4E5F")
    dash.column_dimensions["A"].width = 28
    dash.column_dimensions["B"].width = 20
    dash.column_dimensions["C"].width = 16

    # Issues by category (pie source)
    dash["A13"] = "Issues by category"
    dash["A13"].font = SUB_FONT
    for c, h in enumerate(["Category", "Count", "Open"], start=1):
        dash.cell(row=14, column=c, value=h)
        dash.cell(row=14, column=c).fill = HDR_FILL
        dash.cell(row=14, column=c).font = HDR_FONT
    cats = ["Quality", "Machine", "Material", "Safety"]
    for i, cat in enumerate(cats, start=15):
        dash[f"A{i}"] = cat
        dash[f"B{i}"] = f"=COUNTIFS(Raw_Issues!$C$2:$C${H},$A{i})"
        dash[f"B{i}"].number_format = QTY_FMT
        dash[f"C{i}"] = (f"=COUNTIFS(Raw_Issues!$C$2:$C${H},$A{i},"
                         f"Raw_Issues!$E$2:$E${H},\"Open\")")
        dash[f"C{i}"].number_format = QTY_FMT

    # Opening stock (static input, written by the pipeline)
    dash["A20"] = "Opening stock (units, static input)"
    dash["A20"].font = SUB_FONT
    dash["B20"] = opening_total
    dash["B20"].number_format = QTY_FMT

    # Production by line (bar source)
    dash["A22"] = "Production by line"
    dash["A22"].font = SUB_FONT
    for c, h in enumerate(["Line", "Qty produced", "Batches"], start=1):
        dash.cell(row=23, column=c, value=h)
        dash.cell(row=23, column=c).fill = HDR_FILL
        dash.cell(row=23, column=c).font = HDR_FONT
    for i, line in enumerate(["Line-A", "Line-B"], start=24):
        dash[f"A{i}"] = line
        dash[f"B{i}"] = (
            f"=SUMIFS(Raw_Production!$E$2:$E${H},"
            f"Raw_Production!$C$2:$C${H},$A{i})")
        dash[f"B{i}"].number_format = QTY_FMT
        dash[f"C{i}"] = (
            f"=COUNTIFS(Raw_Production!$C$2:$C${H},$A{i})")
        dash[f"C{i}"].number_format = QTY_FMT

    # --- charts ---
    line = LineChart()
    line.title = "Daily sales revenue vs production qty"
    line.style = 2
    line.height, line.width = 7.5, 15
    line.add_data(Reference(dy, min_col=2, max_col=3, min_row=D1 - 1,
                            max_row=D2), from_rows=False,
                  titles_from_data=True)
    line.set_categories(Reference(dy, min_col=1, max_col=1, min_row=D1,
                                  max_row=D2))
    dash.add_chart(line, "E3")

    pie = PieChart()
    pie.title = "Issues by category"
    pie.style = 2
    pie.height, pie.width = 7.5, 12
    pie.add_data(Reference(dash, min_col=2, max_col=2, min_row=15,
                           max_row=18), titles_from_data=False)
    pie.set_categories(Reference(dash, min_col=1, max_col=1, min_row=15,
                                 max_row=18))
    dash.add_chart(pie, "E19")

    bar = BarChart()
    bar.title = "Production qty by line"
    bar.style = 2
    bar.height, bar.width = 7.5, 12
    bar.add_data(Reference(dash, min_col=2, max_col=2, min_row=24,
                           max_row=25), titles_from_data=False)
    bar.set_categories(Reference(dash, min_col=1, max_col=1, min_row=24,
                                 max_row=25))
    bar.series[0].title = SeriesLabel(v="Qty produced")
    dash.add_chart(bar, "E35")

    # ================= Checks =================
    ch = wb.create_sheet("Checks")
    ch["A1"] = "Checks — all must be TRUE after recalculation"
    ch["A1"].font = TITLE_FONT
    ch["A2"], ch["B2"] = "Check", "Result"
    style_header(ch, 2, 2)
    n1 = 1 + n_sales
    n2 = 1 + n_prod
    n3 = 1 + n_inv
    n4 = 1 + n_issues
    checks = [
        ("No duplicate sales keys in history",
         f"=SUMPRODUCT((Raw_Sales!$B$2:$B${n1}<>\"\")*"
         f"(COUNTIFS(Raw_Sales!$B$2:$B${n1},Raw_Sales!$B$2:$B${n1})>1))=0"),
        ("No duplicate production keys in history",
         f"=SUMPRODUCT((Raw_Production!$B$2:$B${n2}<>\"\")*"
         f"(COUNTIFS(Raw_Production!$B$2:$B${n2},Raw_Production!$B$2:$B${n2})>1))=0"),
        ("No duplicate inventory keys in history",
         f"=SUMPRODUCT((Raw_Inventory!$B$2:$B${n3}<>\"\")*"
         f"(COUNTIFS(Raw_Inventory!$B$2:$B${n3},Raw_Inventory!$B$2:$B${n3})>1))=0"),
        ("No duplicate issue keys in history",
         f"=SUMPRODUCT((Raw_Issues!$B$2:$B${n4}<>\"\")*"
         f"(COUNTIFS(Raw_Issues!$B$2:$B${n4},Raw_Issues!$B$2:$B${n4})>1))=0"),
        ("Revenue column = qty x unit_price (all rows)",
         f"=SUMPRODUCT(ABS(Raw_Sales!$G$2:$G${n1}-"
         f"Raw_Sales!$E$2:$E${n1}*Raw_Sales!$F$2:$F${n1}))<0.01"),
        ("Yield column consistent (all rows)",
         f"=SUMPRODUCT(ABS(Raw_Production!$G$2:$G${n2}-IF(("
         f"Raw_Production!$E$2:$E${n2}+Raw_Production!$F$2:$F${n2})=0,1,"
         f"Raw_Production!$E$2:$E${n2}/(Raw_Production!$E$2:$E${n2}"
         f"+Raw_Production!$F$2:$F${n2}))))<0.001"),
        ("Signed-qty column consistent (all rows)",
         f"=SUMPRODUCT(ABS(Raw_Inventory!$F$2:$F${n3}-IF("
         f"Raw_Inventory!$D$2:$D${n3}=\"IN\",Raw_Inventory!$E$2:$E${n3},IF("
         f"Raw_Inventory!$D$2:$D${n3}=\"OUT\",-Raw_Inventory!$E$2:$E${n3},0))))<0.01"),
        ("Daily sales ties to raw history",
         f"=ABS(SUM(Daily!$B$2:$B${D2})-SUM(Raw_Sales!$G$2:$G${n1}))<0.01"),
        ("Daily production + issues tie to raw history",
         f"=AND(ABS(SUM(Daily!$C$2:$C${D2})-SUM(Raw_Production!$E$2:$E${n2}))<0.01,"
         f"SUM(Daily!$F$2:$F${D2})=COUNTA(Raw_Issues!$B$2:$B${n4}))"),
        ("Inventory KPI = opening + signed movements",
         f"=ABS(Dashboard!$B$7-(Dashboard!$B$20+SUM(Raw_Inventory!$F$2:$F${n3})))<0.01"),
        ("Every refresh run ended OK",
         f"=COUNTIF(Refresh_Log!$H$2:$H${lastlog},\"OK\")"
         f"+COUNTIF(Refresh_Log!$H$2:$H${lastlog},\"OK_WITH_QUARANTINE\")"
         f"=COUNTA(Refresh_Log!$A$2:$A${lastlog})"),
    ]
    for i, (label, f) in enumerate(checks, start=3):
        ch[f"A{i}"] = label
        ch[f"B{i}"] = f
        ch[f"B{i}"].font = Font(bold=True)
    ch.column_dimensions["A"].width = 52
    ch.column_dimensions["B"].width = 12

    wb.save(OUT)
    print(f"wrote {OUT}: {len(wb.sheetnames)} sheets, sales={n_sales} "
          f"prod={n_prod} inv={n_inv} issues={n_issues}")


if __name__ == "__main__":
    import sys
    build(sys.argv[1] if len(sys.argv) > 1 else None)
