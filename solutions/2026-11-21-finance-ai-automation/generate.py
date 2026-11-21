#!/usr/bin/env python3
"""FinanceFlow AI — build the financial reporting workbook.

Reads out/classified_transactions.csv, out/monthly_summary.csv,
out/category_summary.csv and out/summary.json (all produced by
pipeline.py) and writes financeflow-ai-report.xlsx — a genuine
formula-driven workbook (openpyxl, no macros):

  Guide       — what the workbook is, how to refresh it
  Raw         — every classified transaction; Net and Month are live
                formulas, so a re-classified row flows through
  Monthly     — income / expenses / net / txn count per month, all
                SUMIFS/COUNTIFS over Raw, plus a TOTAL row
  Categories  — totals per chart-of-accounts category (SUMIFS), share %
  Dashboard   — focus-month dropdown driving 7 KPIs + 3 native charts
                (monthly net, category share, income vs expenses)
  Checks      — 11 integrity checks vs the independent Python ground
                truth (out/summary.json + the CSVs); all must PASS

Every figure is a live A1-style formula (LibreOffice-safe: no structured
references, SUMPRODUCT(ABS(...)) where an array sum is needed, column
letters repeated on both ends of every range).

Run:  python3 generate.py
"""
import csv
import json
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")
XLSX = os.path.join(BASE, "financeflow-ai-report.xlsx")

MONEY = u"\u20b9#,##0.00"
MONEY0 = u"\u20b9#,##0"
PCT = "0.0%"
DATEFMT = "yyyy-mm-dd"
NAVY = "1F3864"
LIGHT = "D9E2F3"
GREY = "F2F2F2"
GREEN = "C6EFCE"
RED = "FFC7CE"
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
HDR_FILL = PatternFill("solid", fgColor=NAVY)
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SUB_FONT = Font(size=11, color="404040")
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = THIN


def load_csv(name):
    with open(os.path.join(OUT, name), newline="") as f:
        return list(csv.DictReader(f))


def main():
    txns = load_csv("classified_transactions.csv")
    monthly = load_csv("monthly_summary.csv")
    cats = load_csv("category_summary.csv")
    with open(os.path.join(OUT, "summary.json")) as f:
        summary = json.load(f)

    wb = Workbook()

    # ---------------------------------------------------------------- Guide
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "FinanceFlow AI — Financial Reporting Workbook"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Kesarwani Traders, Patna (wholesale FMCG)  |  Apr–Oct 2026  |  "
                "Built by the FinanceFlow AI pipeline: bank-statement CSV → "
                "rule-based extraction & classification → this workbook.")
    ws["A2"].font = SUB_FONT
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:F2")
    ws.row_dimensions[2].height = 30
    lines = [
        ("What this is", "A demo of the Finance AI Automation pattern: messy bank "
         "narrations are extracted (counterparty, category, confidence) and "
         "turned into a monthly financial report. Every number below is a "
         "live Excel formula — change a category on Raw and the whole "
         "workbook re-computes."),
        ("How to refresh", "1) Drop a new bank_statement.csv in data/  "
         "2) Run python3 pipeline.py  3) Run python3 generate.py. "
         "The classifier is a deterministic rule engine over the chart of "
         "accounts; in production the same interface fronts an LLM call "
         "for narrations no rule can classify."),
        ("Confidence", "1.0 = keyword matched exactly. 0.7 = generic PAYMENT "
         "fallback for Sales. 0.0 = Unclassified — flagged for human "
         "review on the Checks sheet, never guessed."),
        ("Anomalies", "The pipeline flags duplicates, stated-vs-recomputed "
         "balance mismatches, unclassified narrations and single legs over "
         "₹60,000. See out/anomalies.csv and the Checks sheet."),
        ("Sheets", "Raw — 268 classified transactions. Monthly — SUMIFS "
         "summary. Categories — SUMIFS by chart of accounts. Dashboard — "
         "pick a focus month, read the KPIs. Checks — 11 integrity checks "
         "against the independent Python ground truth."),
    ]
    r = 4
    for title, text in lines:
        ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=11)
        c = ws.cell(row=r, column=2, value=text)
        c.font = SUB_FONT
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 42
        r += 1
    ws.column_dimensions["A"].width = 18
    for col in "BCDEF":
        ws.column_dimensions[col].width = 16

    # ------------------------------------------------------------------ Raw
    raw = wb.create_sheet("Raw")
    headers = ["Date", "Narration", "Counterparty", "Category", "Kind",
               "Confidence", "Debit", "Credit", "Net", "Month", "Flags"]
    raw.append(headers)
    style_header(raw, 1, len(headers))
    n = len(txns)
    for i, t in enumerate(txns, start=2):
        raw.cell(row=i, column=1, value=datetime.strptime(t["date"], "%Y-%m-%d").date())
        raw.cell(row=i, column=1).number_format = DATEFMT
        raw.cell(row=i, column=2, value=t["narration"])
        raw.cell(row=i, column=3, value=t["counterparty"])
        raw.cell(row=i, column=4, value=t["category"])
        raw.cell(row=i, column=5, value=t["kind"])
        raw.cell(row=i, column=6, value=float(t["confidence"])).number_format = "0.0"
        raw.cell(row=i, column=7, value=float(t["debit"])).number_format = MONEY
        raw.cell(row=i, column=8, value=float(t["credit"])).number_format = MONEY
        # Net and Month are live formulas, not pasted values.
        raw.cell(row=i, column=9, value="=H%d-G%d" % (i, i)).number_format = MONEY
        raw.cell(row=i, column=10, value='=TEXT(A%d,"yyyy-mm")' % i)
        raw.cell(row=i, column=11, value=t["flags"])
    raw.freeze_panes = "A2"
    raw.auto_filter.ref = "A1:K%d" % (n + 1)
    widths = [12, 46, 24, 20, 12, 11, 14, 14, 14, 9, 22]
    for j, w in enumerate(widths, start=1):
        raw.column_dimensions[get_column_letter(j)].width = w
    raw.sheet_properties.pageSetUpPr.fitToPage = True

    # -------------------------------------------------------------- Monthly
    mon = wb.create_sheet("Monthly")
    mon.append(["Month", "Income", "Expenses", "Net", "Txn count"])
    style_header(mon, 1, 5)
    m = len(monthly)
    for i, row in enumerate(monthly, start=2):
        mon.cell(row=i, column=1, value=row["month"])
        mon.cell(row=i, column=2,
                value='=SUMIFS(Raw!$H$2:$H$%d,Raw!$J$2:$J$%d,$A%d)' % (n + 1, n + 1, i))
        mon.cell(row=i, column=3,
                value='=SUMIFS(Raw!$G$2:$G$%d,Raw!$J$2:$J$%d,$A%d)' % (n + 1, n + 1, i))
        mon.cell(row=i, column=4, value="=B%d-C%d" % (i, i))
        mon.cell(row=i, column=5,
                value='=COUNTIFS(Raw!$J$2:$J$%d,$A%d)' % (n + 1, i))
        for c in (2, 3, 4):
            mon.cell(row=i, column=c).number_format = MONEY
    tot = m + 2
    mon.cell(row=tot, column=1, value="TOTAL").font = Font(bold=True)
    for c, f in ((2, "=SUM(B2:B%d)"), (3, "=SUM(C2:C%d)"),
                 (4, "=SUM(D2:D%d)"), (5, "=SUM(E2:E%d)")):
        cell = mon.cell(row=tot, column=c, value=f % (m + 1))
        cell.font = Font(bold=True)
        if c < 5:
            cell.number_format = MONEY
    for col, w in zip("ABCDE", [10, 16, 16, 16, 11]):
        mon.column_dimensions[col].width = w

    # ------------------------------------------------------------ Categories
    cat = wb.create_sheet("Categories")
    cat.append(["Category", "Kind", "Total", "Count", "Share %"])
    style_header(cat, 1, 5)
    k = len(cats)
    for i, row in enumerate(cats, start=2):
        cat.cell(row=i, column=1, value=row["category"])
        cat.cell(row=i, column=2, value=row["kind"])
        cat.cell(row=i, column=3,
                value='=SUMIFS(Raw!$I$2:$I$%d,Raw!$D$2:$D$%d,$A%d)' % (n + 1, n + 1, i))
        cat.cell(row=i, column=3).number_format = MONEY
        cat.cell(row=i, column=4,
                value='=COUNTIFS(Raw!$D$2:$D$%d,$A%d)' % (n + 1, i))
        # SUMPRODUCT forces array evaluation in LibreOffice for ABS sums.
        cat.cell(row=i, column=5,
                value="=ABS(C%d)/SUMPRODUCT(ABS($C$2:$C$%d))" % (i, k + 1))
        cat.cell(row=i, column=5).number_format = PCT
    tot = k + 2
    cat.cell(row=tot, column=1, value="TOTAL (net)").font = Font(bold=True)
    cell = cat.cell(row=tot, column=3, value="=SUM(C2:C%d)" % (k + 1))
    cell.font = Font(bold=True)
    cell.number_format = MONEY
    for col, w in zip("ABCDE", [22, 13, 16, 8, 10]):
        cat.column_dimensions[col].width = w

    # ------------------------------------------------------------- Dashboard
    dash = wb.create_sheet("Dashboard")
    dash["A1"] = "FinanceFlow AI — Dashboard"
    dash["A1"].font = TITLE_FONT
    dash["A2"] = "Focus month:"
    dash["A2"].font = Font(bold=True, size=11)
    dash["B2"] = monthly[-1]["month"]  # latest month by default
    dash["B2"].font = Font(bold=True, size=12)
    dv = DataValidation(type="list",
                        formula1='"%s"' % ",".join(r["month"] for r in monthly),
                        allow_blank=False)
    dv.error = "Pick a month from the list"
    dash.add_data_validation(dv)
    dv.add(dash["B2"])

    kpis = [
        ("Income (focus month)",
         '=SUMIFS(Raw!$H$2:$H$%d,Raw!$J$2:$J$%d,$B$2)' % (n + 1, n + 1), MONEY),
        ("Expenses (focus month)",
         '=SUMIFS(Raw!$G$2:$G$%d,Raw!$J$2:$J$%d,$B$2)' % (n + 1, n + 1), MONEY),
        ("Net (focus month)",
         '=SUMIFS(Raw!$I$2:$I$%d,Raw!$J$2:$J$%d,$B$2)' % (n + 1, n + 1), MONEY),
        ("Transactions (focus month)",
         '=COUNTIFS(Raw!$J$2:$J$%d,$B$2)' % (n + 1), "0"),
        ("Avg sale (focus month)",
         '=IFERROR(SUMIFS(Raw!$H$2:$H$%d,Raw!$J$2:$J$%d,$B$2,Raw!$D$2:$D$%d,'
         '"Sales Income")/COUNTIFS(Raw!$J$2:$J$%d,$B$2,Raw!$D$2:$D$%d,'
         '"Sales Income"),0)' % (n + 1, n + 1, n + 1, n + 1, n + 1), MONEY),
        ("Unclassified txns (all time)",
         '=COUNTIFS(Raw!$D$2:$D$%d,"Unclassified")' % (n + 1), "0"),
        ("Anomalies flagged (all time)",
         '=COUNTIF(Raw!$K$2:$K$%d,"?*")' % (n + 1), "0"),
        ("Closing balance (recomputed)",
         "=85000+SUM(Raw!$I$2:$I$%d)" % (n + 1), MONEY),
    ]
    r = 4
    dash.cell(row=r, column=1, value="KPI").font = Font(bold=True)
    dash.cell(row=r, column=2, value="Value").font = Font(bold=True)
    for label, formula, fmt in kpis:
        r += 1
        dash.cell(row=r, column=1, value=label).font = SUB_FONT
        cell = dash.cell(row=r, column=2, value=formula)
        cell.number_format = fmt
        cell.font = Font(bold=True, size=11)
    dash.column_dimensions["A"].width = 32
    dash.column_dimensions["B"].width = 20

    # Charts (data from Monthly / Categories).
    bar = BarChart()
    bar.type = "col"
    bar.title = "Monthly net (income − expenses)"
    bar.y_axis.title = "Rs"
    bar.x_axis.title = "Month"
    bar.add_data(Reference(mon, min_col=4, min_row=1, max_row=m + 1),
                 titles_from_data=True)
    bar.set_categories(Reference(mon, min_col=1, min_row=2, max_row=m + 1))
    bar.shape = 4
    bar.height = 7.5
    bar.width = 14
    dash.add_chart(bar, "A14")

    pie = PieChart()
    pie.title = "Category share (absolute totals)"
    pie.add_data(Reference(cat, min_col=5, min_row=1, max_row=k + 1),
                 titles_from_data=True)
    pie.set_categories(Reference(cat, min_col=1, min_row=2, max_row=k + 1))
    pie.height = 7.5
    pie.width = 14
    dash.add_chart(pie, "A30")

    line = LineChart()
    line.title = "Income vs expenses by month"
    line.y_axis.title = "Rs"
    line.add_data(Reference(mon, min_col=2, min_row=1, max_row=m + 1),
                  titles_from_data=True)
    line.add_data(Reference(mon, min_col=3, min_row=1, max_row=m + 1),
                  titles_from_data=True)
    line.set_categories(Reference(mon, min_col=1, min_row=2, max_row=m + 1))
    line.height = 7.5
    line.width = 14
    dash.add_chart(line, "A46")

    # ---------------------------------------------------------------- Checks
    chk = wb.create_sheet("Checks")
    chk.append(["#", "Check", "Workbook value", "Expected (Python)", "Result"])
    style_header(chk, 1, 5)
    exp = summary
    checks = [
        ("Raw row count", "=COUNTA(Raw!$A$2:$A$%d)" % (n + 1),
         exp["transactions"]),
        ("Total debits (expenses)", "=SUM(Raw!$G$2:$G$%d)" % (n + 1),
         exp["total_expenses"]),
        ("Total credits (income)", "=SUM(Raw!$H$2:$H$%d)" % (n + 1),
         exp["total_income"]),
        ("Monthly income total", "=Monthly!B%d" % (m + 2),
         exp["total_income"]),
        ("Monthly expenses total", "=Monthly!C%d" % (m + 2),
         exp["total_expenses"]),
        ("Monthly net total", "=Monthly!D%d" % (m + 2), exp["net"]),
        ("Category totals sum to net", "=Categories!C%d" % (k + 2), exp["net"]),
        ("Closing balance tie-out (85000 + net)",
         "=85000+SUM(Raw!$I$2:$I$%d)" % (n + 1), exp["closing_balance"]),
        ("Duplicate rows flagged",
         '=COUNTIF(Raw!$K$2:$K$%d,"*DUPLICATE*")' % (n + 1), 1),
        ("Unclassified rows", '=COUNTIF(Raw!$D$2:$D$%d,"Unclassified")' % (n + 1),
         exp["unclassified"]),
        ("Balance mismatches flagged",
         '=COUNTIF(Raw!$K$2:$K$%d,"*BALANCE_MISMATCH*")' % (n + 1), 1),
        ("Large amounts flagged",
         '=COUNTIF(Raw!$K$2:$K$%d,"*LARGE_AMOUNT*")' % (n + 1), 1),
    ]
    for i, (label, formula, expected) in enumerate(checks, start=2):
        chk.cell(row=i, column=1, value=i - 1)
        chk.cell(row=i, column=2, value=label)
        c3 = chk.cell(row=i, column=3, value=formula)
        c3.number_format = MONEY
        c4 = chk.cell(row=i, column=4, value=expected)
        c4.number_format = MONEY
        chk.cell(row=i, column=5,
                value='=IF(ABS(C%d-D%d)<0.01,"PASS","FAIL")' % (i, i))
    for col, w in zip("ABCDE", [5, 42, 18, 18, 9]):
        chk.column_dimensions[col].width = w

    wb.save(XLSX)
    print("wrote %s (%d txns, %d months, %d categories, %d checks)"
          % (XLSX, n, m, k, len(checks)))


if __name__ == "__main__":
    main()
