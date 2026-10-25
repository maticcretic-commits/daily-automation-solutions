#!/usr/bin/env python3
"""Build the Owner Dashboard workbook (Maple Ridge Builders Ltd., CAD).

Reads data/*.csv (QBO-style exports, SYNTHETIC demo data) and writes a genuine
formula-driven .xlsx: every statement figure is a live A1-style formula over
the raw sheets (no hard-coded totals, no Table objects, no structured refs —
LibreOffice-safe). Refresh = paste new QBO exports into Raw_QBO / AR / AP
and the dashboard follows.

Sheets: Guide, Dashboard (6 KPI cards + 3 native charts), Monthly_PL,
Raw_QBO, AR_Aging, AP_Aging, Cash, Checks (8 live tie-out checks).
"""
import csv
import datetime
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

MONTHS = ["2026-05", "2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
CATEGORIES = [
    "Revenue - Residential", "Revenue - Commercial",
    "Cost - Labour", "Cost - Materials", "Cost - Subcontractors",
    "Overhead - Office", "Overhead - Insurance", "Overhead - Vehicle",
]
AS_OF = datetime.date(2026, 10, 25)
BUCKETS = ["Current", "1-30", "31-60", "61+"]

TITLE_FILL = PatternFill("solid", fgColor="1F4E79")
TITLE_FONT = Font(bold=True, color="FFFFFF", size=12)
HDR_FONT = Font(bold=True)
MONEY = '#,##0'
PCT = '0.0%'
DATEFMT = 'yyyy-mm-dd'


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def style_title(ws, row, ncols, text):
    ws.cell(row=row, column=1, value=text).font = TITLE_FONT
    ws.cell(row=row, column=1).fill = TITLE_FILL
    for c in range(2, ncols + 1):
        ws.cell(row=row, column=c).fill = TITLE_FILL


def build():
    qbo = read_csv("qbo_pl.csv")
    ar = read_csv("ar.csv")
    ap = read_csv("ap.csv")
    opening = int(read_csv("cash.csv")[0]["Amount"])

    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    style_title(ws, 1, 2, "Owner Dashboard — Maple Ridge Builders Ltd. (CAD)")
    guide = [
        "WHAT THIS IS: the monthly owner dashboard a construction CFO would hand",
        "an owner: P&L, cash, receivables and payables in one workbook, built from",
        "QuickBooks Online exports plus spreadsheet inputs. Built as a demo around",
        "a live freelance brief (construction financial dashboard, QuickBooks +",
        "spreadsheets, simple for owners).",
        "",
        "DATA: all figures are SYNTHETIC demo data (seed 20261025), invented for",
        "this demo — not a real client's books. Amounts in Canadian dollars.",
        "",
        "HOW TO REFRESH: paste the new month's QBO P&L export into Raw_QBO",
        "(columns Month / Category / Amount), new invoices into AR_Aging and",
        "new bills into AP_Aging, then update the As-of date cell (AR J2 / AP I2).",
        "Monthly_PL, Dashboard, Cash and Checks all recalculate from formulas.",
        "",
        "RULES THE WORKBOOK FOLLOWS: no hard-coded totals (every total is a",
        "SUM/SUMIFS over raw rows); every dashboard KPI traces to a source cell;",
        "AR/AP aging buckets are formulas off the As-of date; the Checks sheet",
        "must read TRUE throughout after any refresh.",
    ]
    for i, line in enumerate(guide, start=2):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 110

    # ---------------- Raw_QBO ----------------
    ws = wb.create_sheet("Raw_QBO")
    style_title(ws, 1, 3, "Raw QBO P&L export (paste new exports below row 49)")
    for c, h in enumerate(["Month", "Category", "Amount"], start=1):
        ws.cell(row=2, column=c, value=h).font = HDR_FONT
    for r, row in enumerate(qbo, start=3):
        ws.cell(row=r, column=1, value=row["Month"])
        ws.cell(row=r, column=2, value=row["Category"])
        cell = ws.cell(row=r, column=3, value=int(row["Amount"]))
        cell.number_format = MONEY
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 14

    # ---------------- Monthly_PL ----------------
    ws = wb.create_sheet("Monthly_PL")
    style_title(ws, 1, 8, "Monthly P&L — every figure a SUMIFS over Raw_QBO")
    ws.cell(row=2, column=1, value="Category").font = HDR_FONT
    for c, m in enumerate(MONTHS, start=2):
        ws.cell(row=2, column=c, value=m).font = HDR_FONT
    ws.cell(row=2, column=8, value="YTD").font = HDR_FONT
    for r, cat in enumerate(CATEGORIES, start=3):
        ws.cell(row=r, column=1, value=cat)
        for c in range(2, 8):
            col = get_column_letter(c)
            cell = ws.cell(
                row=r, column=c,
                value=(f"=SUMIFS(Raw_QBO!$C$2:$C$500,Raw_QBO!$A$2:$A$500,"
                       f"{col}$2,Raw_QBO!$B$2:$B$500,$A{r})"))
            cell.number_format = MONEY
        cell = ws.cell(row=r, column=8, value=f"=SUM(B{r}:G{r})")
        cell.number_format = MONEY
    # Totals: revenue rows 3-4, costs rows 5-10.
    for c in range(2, 9):
        col = get_column_letter(c)
        for r, formula in [
                (11, f"=SUM({col}3:{col}4)"),
                (12, f"=SUM({col}5:{col}10)"),
                (13, f"={col}11+{col}12")]:
            cell = ws.cell(row=r, column=c, value=formula)
            cell.number_format = MONEY
        cell = ws.cell(row=14, column=c,
                       value=f"=IF({col}11=0,0,{col}13/{col}11)")
        cell.number_format = PCT
    ws.cell(row=11, column=1, value="Total Revenue").font = HDR_FONT
    ws.cell(row=12, column=1, value="Total Costs").font = HDR_FONT
    ws.cell(row=13, column=1, value="Gross Profit").font = HDR_FONT
    ws.cell(row=14, column=1, value="Gross Margin %").font = HDR_FONT
    ws.column_dimensions["A"].width = 26
    for c in range(2, 9):
        ws.column_dimensions[get_column_letter(c)].width = 14

    # ---------------- AR_Aging ----------------
    ws = wb.create_sheet("AR_Aging")
    style_title(ws, 1, 8, "Receivables — aging buckets are formulas off the As-of date")
    ws.cell(row=2, column=10, value="As-of date").font = HDR_FONT
    asof = ws.cell(row=2, column=11, value=AS_OF)
    asof.number_format = DATEFMT
    headers = ["InvoiceID", "Customer", "IssueDate", "DueDate", "Amount",
               "PaidAmount", "Outstanding", "Bucket"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=2, column=c, value=h).font = HDR_FONT
    for r, row in enumerate(ar, start=3):
        ws.cell(row=r, column=1, value=row["InvoiceID"])
        ws.cell(row=r, column=2, value=row["Customer"])
        for c, key in [(3, "IssueDate"), (4, "DueDate")]:
            y, m, d = (int(x) for x in row[key].split("-"))
            cell = ws.cell(row=r, column=c, value=datetime.date(y, m, d))
            cell.number_format = DATEFMT
        for c, key in [(5, "Amount"), (6, "PaidAmount")]:
            cell = ws.cell(row=r, column=c, value=int(row[key]))
            cell.number_format = MONEY
        cell = ws.cell(row=r, column=7, value=f"=E{r}-F{r}")
        cell.number_format = MONEY
        ws.cell(
            row=r, column=8,
            value=(f'=IF(G{r}=0,"Paid",IF(D{r}>=$K$2,"Current",'
                   f'IF($K$2-D{r}<=30,"1-30",IF($K$2-D{r}<=60,"31-60","61+"))))'))
    ws.cell(row=34, column=1, value="Bucket").font = HDR_FONT
    ws.cell(row=34, column=2, value="Outstanding").font = HDR_FONT
    for i, b in enumerate(BUCKETS):
        r = 35 + i
        ws.cell(row=r, column=1, value=b)
        cell = ws.cell(row=r, column=2,
                       value=f"=SUMIFS($G$3:$G$32,$H$3:$H$32,$A{r})")
        cell.number_format = MONEY
    ws.cell(row=39, column=1, value="Total").font = HDR_FONT
    cell = ws.cell(row=39, column=2, value="=SUM(B35:B38)")
    cell.number_format = MONEY
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 26
    for c in ["C", "D"]:
        ws.column_dimensions[c].width = 12
    for c in ["E", "F", "G"]:
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["H"].width = 10

    # ---------------- AP_Aging ----------------
    ws = wb.create_sheet("AP_Aging")
    style_title(ws, 1, 7, "Payables — aging buckets are formulas off the As-of date")
    ws.cell(row=2, column=9, value="As-of date").font = HDR_FONT
    asof = ws.cell(row=2, column=10, value=AS_OF)
    asof.number_format = DATEFMT
    headers = ["BillID", "Vendor", "DueDate", "Amount", "PaidAmount",
               "Outstanding", "Bucket"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=2, column=c, value=h).font = HDR_FONT
    for r, row in enumerate(ap, start=3):
        ws.cell(row=r, column=1, value=row["BillID"])
        ws.cell(row=r, column=2, value=row["Vendor"])
        y, m, d = (int(x) for x in row["DueDate"].split("-"))
        cell = ws.cell(row=r, column=3, value=datetime.date(y, m, d))
        cell.number_format = DATEFMT
        for c, key in [(4, "Amount"), (5, "PaidAmount")]:
            cell = ws.cell(row=r, column=c, value=int(row[key]))
            cell.number_format = MONEY
        cell = ws.cell(row=r, column=6, value=f"=D{r}-E{r}")
        cell.number_format = MONEY
        ws.cell(
            row=r, column=7,
            value=(f'=IF(F{r}=0,"Paid",IF(C{r}>=$J$2,"Current",'
                   f'IF($J$2-C{r}<=30,"1-30",IF($J$2-C{r}<=60,"31-60","61+"))))'))
    ws.cell(row=28, column=1, value="Bucket").font = HDR_FONT
    ws.cell(row=28, column=2, value="Outstanding").font = HDR_FONT
    for i, b in enumerate(BUCKETS):
        r = 29 + i
        ws.cell(row=r, column=1, value=b)
        cell = ws.cell(row=r, column=2,
                       value=f"=SUMIFS($F$3:$F$26,$G$3:$G$26,$A{r})")
        cell.number_format = MONEY
    ws.cell(row=33, column=1, value="Total").font = HDR_FONT
    cell = ws.cell(row=33, column=2, value="=SUM(B29:B32)")
    cell.number_format = MONEY
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 12
    for c in ["D", "E", "F"]:
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["G"].width = 10

    # ---------------- Cash ----------------
    ws = wb.create_sheet("Cash")
    style_title(ws, 1, 2, "Cash position — opening + collected - paid")
    rows = [
        ("Opening bank balance (2026-05-01)", opening, None),
        ("AR collected (paid amounts)", None, "=SUM(AR_Aging!$F$3:$F$32)"),
        ("AP paid", None, "=SUM(AP_Aging!$E$3:$E$26)"),
        ("Cash on hand", None, "=B3+B4-B5"),
    ]
    for i, (label, val, formula) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=label)
        cell = ws.cell(row=i, column=2, value=formula if formula else val)
        cell.number_format = MONEY
        if i == 6:
            ws.cell(row=i, column=1).font = HDR_FONT
            cell.font = HDR_FONT
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 16

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard", 1)  # second tab, after Guide
    style_title(ws, 1, 2, "Owner Dashboard — Maple Ridge Builders Ltd. (CAD)")
    kpis = [
        ("YTD Revenue", "=Monthly_PL!H11", MONEY),
        ("Gross Margin %", "=Monthly_PL!H14", PCT),
        ("Cash on Hand", "=Cash!B6", MONEY),
        ("AR Outstanding", "=SUM(AR_Aging!G3:G32)", MONEY),
        ("AP Outstanding", "=SUM(AP_Aging!F3:F26)", MONEY),
        ("AR Overdue 61+ days", "=AR_Aging!B38", MONEY),
    ]
    for i, (label, formula, fmt) in enumerate(kpis, start=3):
        ws.cell(row=i, column=1, value=label).font = HDR_FONT
        cell = ws.cell(row=i, column=2, value=formula)
        cell.number_format = fmt
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 18

    col_chart = BarChart()
    col_chart.type = "col"
    col_chart.title = "Monthly revenue vs costs"
    data = Reference(wb["Monthly_PL"], min_col=2, min_row=11, max_col=7,
                     max_row=12)
    cats = Reference(wb["Monthly_PL"], min_col=2, min_row=2, max_col=7)
    col_chart.add_data(data, from_rows=True, titles_from_data=False)
    col_chart.set_categories(cats)
    col_chart.series[0].title = SeriesLabel(v="Revenue")
    col_chart.series[1].title = SeriesLabel(v="Costs")
    ws.add_chart(col_chart, "A10")

    line_chart = LineChart()
    line_chart.title = "Gross margin % by month"
    data = Reference(wb["Monthly_PL"], min_col=2, min_row=14, max_col=7)
    line_chart.add_data(data, from_rows=True, titles_from_data=False)
    line_chart.set_categories(cats)
    line_chart.series[0].title = SeriesLabel(v="Gross margin %")
    ws.add_chart(line_chart, "A26")

    pie = PieChart()
    pie.title = "AR aging mix (outstanding)"
    data = Reference(wb["AR_Aging"], min_col=2, min_row=35, max_row=38)
    cats2 = Reference(wb["AR_Aging"], min_col=1, min_row=35, max_row=38)
    pie.add_data(data, titles_from_data=False)
    pie.set_categories(cats2)
    ws.add_chart(pie, "J10")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    style_title(ws, 1, 2, "Tie-out checks — every row must read TRUE")
    ws.cell(row=2, column=1, value="Check").font = HDR_FONT
    ws.cell(row=2, column=2, value="Result").font = HDR_FONT
    rev = ('SUMIFS(Raw_QBO!$C$2:$C$500,Raw_QBO!$B$2:$B$500,'
           '"Revenue - Residential")+SUMIFS(Raw_QBO!$C$2:$C$500,'
           'Raw_QBO!$B$2:$B$500,"Revenue - Commercial")')
    checks = [
        ("PL YTD revenue = raw revenue rows",
         f"=ROUND(Monthly_PL!H11-({rev}),2)=0"),
        ("PL YTD costs = raw total - raw revenue",
         f"=ROUND(Monthly_PL!H12-(SUM(Raw_QBO!$C$2:$C$500)-({rev})),2)=0"),
        ("AR bucket total = AR outstanding",
         "=ROUND(AR_Aging!B39-SUM(AR_Aging!G3:G32),2)=0"),
        ("AP bucket total = AP outstanding",
         "=ROUND(AP_Aging!B33-SUM(AP_Aging!F3:F26),2)=0"),
        ("Dashboard revenue = PL YTD revenue",
         "=Dashboard!B3=Monthly_PL!H11"),
        ("Dashboard cash = Cash sheet",
         "=Dashboard!B5=Cash!B6"),
        ("Gross margin within 0-100%",
         "=AND(Monthly_PL!H14>=0,Monthly_PL!H14<=1)"),
        ("No negative revenue month",
         "=MIN(Monthly_PL!B11:G11)>=0"),
    ]
    for i, (label, formula) in enumerate(checks, start=3):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=formula)
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 10

    out = os.path.join(HERE, "qbo-owner-dashboard.xlsx")
    wb.save(out)
    print("wrote", out)


if __name__ == "__main__":
    build()
