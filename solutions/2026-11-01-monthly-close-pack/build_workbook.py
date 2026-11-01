#!/usr/bin/env python3
"""Build monthly-close-pack.xlsx from data/*.csv — every figure a live formula.

Sheets: Guide, Dashboard, PL_Monthly, Balance_Sheet, CashFlow,
Budget_vs_Actual, Raw_Transactions, Raw_Budget, Assumptions, Checks.
Rules honoured (learned the hard way): plain A1 ranges only, range
literals repeat the column letter on both ends ($E$2:$E$200), no Table
objects, chart series titles come from header cells (titles_from_data).
Dates are real date values. The balance sheet articulates: receivables
close at the uncollected share of October revenue, payables at the
unpaid share of October expenses (Assumptions), cash follows the
indirect cash-flow walk — so the balance check is 0 by construction,
and Checks verifies it after a headless recalc.
"""
import csv
import datetime
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter  # noqa: F401

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
OUT = os.path.join(HERE, "monthly-close-pack.xlsx")

NAVY = "1F3864"
LIGHT = "D9E2F3"
MONEY = "#,##0"
PCT = "0.0%"

CATEGORIES = [
    ("Consulting fees", "Revenue"), ("Monthly retainers", "Revenue"),
    ("Training workshops", "Revenue"), ("Contractor costs", "Direct"),
    ("Salaries", "Operating"), ("Rent", "Operating"),
    ("Software & tools", "Operating"), ("Travel", "Operating"),
    ("Marketing", "Operating"), ("Professional fees", "Operating"),
    ("Utilities", "Operating"), ("Other expenses", "Operating"),
]


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def dval(text):
    return datetime.date.fromisoformat(text)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def main():
    txns = read_csv("transactions.csv")
    budgets = read_csv("budget.csv")
    assert len(budgets) == 120, len(budgets)
    n_txn = len(txns)
    txn_last = n_txn + 1
    bud_last = len(budgets) + 1
    # SUMIFS ranges over the raw sheets (exact data extents, A1 style).
    TX_AMT = f"Raw_Transactions!$E$2:$E${txn_last}"
    TX_TYP = f"Raw_Transactions!$D$2:$D${txn_last}"
    TX_CAT = f"Raw_Transactions!$C$2:$C${txn_last}"
    TX_MON = f"Raw_Transactions!$F$2:$F${txn_last}"
    BU_AMT = f"Raw_Budget!$E$2:$E${bud_last}"
    BU_TYP = f"Raw_Budget!$D$2:$D${bud_last}"
    BU_CAT = f"Raw_Budget!$C$2:$C${bud_last}"
    BU_MON = f"Raw_Budget!$B$2:$B${bud_last}"

    wb = Workbook()

    # ---- Guide ----
    ws = wb.active
    ws.title = "Guide"
    guide_lines = [
        ("Monthly Close Pack — Verma Consulting Services Pvt. Ltd., Vadodara", True),
        ("Demo monthly reporting pack built around a monthly financial", False),
        ("reporting brief: P&L, balance sheet with integrated cash-flow view,", False),
        ("and budget-versus-actual, delivered by the 5th business day.", False),
        ("Every figure is a live formula over the raw sheets — nothing is", False),
        ("hard-coded. Drop a new month's transactions into Raw_Transactions", False),
        ("and budgets into Raw_Budget, extend the SUMIFS ranges, and the", False),
        ("whole pack recomputes. Standards live on Assumptions.", False),
        ("", False),
        ("Standards: depreciation Rs 5,000/month; receivables close at 20%", False),
        ("of October revenue (uncollected); payables close at 10% of", False),
        ("October expenses (unpaid); equipment bought Mar 2026 Rs 75,000;", False),
        ("owner drawings Rs 200,000. Cash follows the indirect cash-flow", False),
        ("walk, so the balance sheet balances by construction (Checks).", False),
        ("", False),
        ("All data are fictional and invented for this demo. No real client", False),
        ("data. Money is INR (Rs); period Jan–Oct 2026, as of 31 Oct 2026.", False),
    ]
    for i, (text, bold) in enumerate(guide_lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        c.font = Font(bold=bold, size=13 if i == 1 else 11,
                      color=NAVY if i == 1 else "000000")
    widths(ws, {"A": 78})

    # ---- Dashboard ----
    ws = wb.create_sheet("Dashboard")
    ws.cell(row=1, column=1,
            value="Monthly Close Pack — Dashboard (as of 2026-10-31)").font = \
        Font(bold=True, size=14, color=NAVY)
    ws.cell(row=2, column=1,
            value="All values are live formulas over PL_Monthly / Balance_Sheet / "
                  "Budget_vs_Actual.").font = Font(italic=True, color="595959")
    ws.cell(row=3, column=1, value="KPI").font = Font(bold=True)
    ws.cell(row=3, column=2, value="Value").font = Font(bold=True)
    ws.cell(row=3, column=3, value="Note").font = Font(bold=True)
    style_header(ws, 3, 3)
    kpis = [
        ("YTD revenue (Jan–Oct 2026)", "=PL_Monthly!B12", MONEY,
         "Sum of monthly revenue"),
        ("YTD net income", "=PL_Monthly!G12", MONEY, "After all expenses + depreciation"),
        ("Net margin", "=B5/B4", PCT, "Net income / revenue"),
        ("Cash at 31 Oct 2026", "=Balance_Sheet!D2", MONEY, "Indirect cash-flow walk"),
        ("Accounts receivable", "=Balance_Sheet!D3", MONEY, "20% of October revenue"),
        ("Total expense variance vs budget", "=Budget_vs_Actual!F15", PCT,
         "Positive = over budget"),
        ("Revenue vs budget", "=SUM(Budget_vs_Actual!C2:C4)/SUM(Budget_vs_Actual!D2:D4)",
         PCT, "Actual / budget, revenue categories"),
        ("Largest category variance %", "=MAX(Budget_vs_Actual!F2:F13)", PCT,
         "Most over-budget category (signed)"),
    ]
    for i, (label, formula, fmt, note) in enumerate(kpis, start=4):
        ws.cell(row=i, column=1, value=label)
        c = ws.cell(row=i, column=2, value=formula)
        c.number_format = fmt
        ws.cell(row=i, column=3, value=note)
    widths(ws, {"A": 36, "B": 20, "C": 46})

    # ---- PL_Monthly ----
    ws = wb.create_sheet("PL_Monthly")
    headers = ["Month", "Revenue", "Direct costs", "Gross profit",
               "Operating expenses", "Depreciation", "Net income",
               "Net margin %", "Budget revenue", "Cumulative net income"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, len(headers))
    for m in range(1, 11):
        r = m + 1
        ws.cell(row=r, column=1, value=datetime.date(2026, m, 1)).number_format = "mmm yyyy"
        ws.cell(row=r, column=2,
                value=f'=SUMIFS({TX_AMT},{TX_TYP},"Revenue",{TX_MON},MONTH($A{r}))')
        ws.cell(row=r, column=3,
                value=f'=SUMIFS({TX_AMT},{TX_TYP},"Direct",{TX_MON},MONTH($A{r}))')
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
        ws.cell(row=r, column=5,
                value=f'=SUMIFS({TX_AMT},{TX_TYP},"Operating",{TX_MON},MONTH($A{r}))')
        ws.cell(row=r, column=6, value="=Assumptions!$B$7")
        ws.cell(row=r, column=7, value=f"=D{r}-E{r}-F{r}")
        ws.cell(row=r, column=8, value=f"=IF(B{r}=0,0,G{r}/B{r})")
        ws.cell(row=r, column=9,
                value=f'=SUMIFS({BU_AMT},{BU_TYP},"Revenue",{BU_MON},MONTH($A{r}))')
        ws.cell(row=r, column=10, value=f"=SUM($G$2:G{r})")
    ws.cell(row=12, column=1, value="TOTAL (Jan–Oct 2026)").font = Font(bold=True)
    for col in (2, 3, 4, 5, 6, 7, 9, 10):
        letter = get_column_letter(col)
        c = ws.cell(row=12, column=col, value=f"=SUM({letter}2:{letter}11)")
        c.font = Font(bold=True)
    ws.cell(row=12, column=8, value="=IF(B12=0,0,G12/B12)").font = Font(bold=True)
    for r in list(range(2, 13)):
        for col in (2, 3, 4, 5, 6, 7, 9, 10):
            ws.cell(row=r, column=col).number_format = MONEY
        ws.cell(row=r, column=8).number_format = PCT
    ws.freeze_panes = "A2"
    widths(ws, {"A": 14, "B": 16, "C": 15, "D": 15, "E": 19, "F": 14,
                "G": 15, "H": 13, "I": 16, "J": 20})

    # ---- Balance_Sheet ----
    ws = wb.create_sheet("Balance_Sheet")
    for j, h in enumerate(["Item", "Opening (1 Jan 2026)", "Movement",
                           "Closing (31 Oct 2026)", "Notes"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 5)
    bs_rows = [
        ("Cash and bank", "=Assumptions!$B$2", "=CashFlow!$B$8", None,
         "Indirect cash-flow walk"),
        ("Accounts receivable", "=Assumptions!$B$3", None,
         "=Assumptions!$B$9*PL_Monthly!$B$11", "20% of October revenue"),
        ("Equipment (net)", "=Assumptions!$B$4", None,
         "=B4+Assumptions!$B$8-SUM(PL_Monthly!$F$2:$F$11)",
         "Opening + purchases − depreciation"),
        ("TOTAL ASSETS", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)", ""),
        ("Accounts payable", "=Assumptions!$B$5", None,
         "=Assumptions!$B$10*(PL_Monthly!$C$11+PL_Monthly!$E$11)",
         "10% of October expenses"),
        ("TOTAL LIABILITIES", "=B6", "=C6", "=D6", ""),
        ("Capital (opening equity)", "=Assumptions!$B$6", 0, None, ""),
        ("Current-year profit", 0, "=PL_Monthly!$G$12", None, "YTD net income"),
        ("Owner drawings", 0, "=-Assumptions!$B$11", None, ""),
        ("TOTAL EQUITY", "=SUM(B8:B10)", "=SUM(C8:C10)", "=SUM(D8:D10)", ""),
        ("TOTAL LIABILITIES & EQUITY", "=B7+B11", "=C7+C11", "=D7+D11", ""),
    ]
    for i, (label, opening, movement, closing, note) in enumerate(bs_rows, start=2):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=opening)
        if movement is not None:
            ws.cell(row=i, column=3, value=movement)
        else:
            ws.cell(row=i, column=3, value=f"=D{i}-B{i}")
        if closing is not None:
            ws.cell(row=i, column=4, value=closing)
        else:
            ws.cell(row=i, column=4, value=f"=B{i}+C{i}")
        ws.cell(row=i, column=5, value=note)
        if label.startswith("TOTAL"):
            for col in range(1, 5):
                ws.cell(row=i, column=col).font = Font(bold=True)
    ws.cell(row=13, column=1,
            value="Balance check (Assets − Liabilities & Equity — must be 0)")
    ws.cell(row=13, column=4, value="=D5-D12").font = Font(bold=True)
    for r in range(2, 14):
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).number_format = MONEY
    widths(ws, {"A": 44, "B": 20, "C": 16, "D": 21, "E": 38})

    # ---- CashFlow ----
    ws = wb.create_sheet("CashFlow")
    ws.cell(row=1, column=1, value="Indirect cash-flow walk (Jan–Oct 2026)")
    ws.cell(row=1, column=2, value="Amount")
    style_header(ws, 1, 2)
    cf_rows = [
        ("Net income", "=PL_Monthly!$G$12"),
        ("Add: depreciation", "=SUM(PL_Monthly!$F$2:$F$11)"),
        ("(Increase) in accounts receivable", "=-Balance_Sheet!C3"),
        ("Increase in accounts payable", "=Balance_Sheet!C6"),
        ("Less: equipment purchased", "=-Assumptions!$B$8"),
        ("Less: owner drawings", "=-Assumptions!$B$11"),
    ]
    for i, (label, formula) in enumerate(cf_rows, start=2):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=formula)
    ws.cell(row=8, column=1, value="Net change in cash").font = Font(bold=True)
    ws.cell(row=8, column=2, value="=SUM(B2:B7)").font = Font(bold=True)
    ws.cell(row=9, column=1, value="Opening cash (1 Jan 2026)")
    ws.cell(row=9, column=2, value="=Assumptions!$B$2")
    ws.cell(row=10, column=1, value="Closing cash (31 Oct 2026)").font = Font(bold=True)
    ws.cell(row=10, column=2, value="=B8+B9").font = Font(bold=True)
    ws.cell(row=11, column=1, value="Tie-out (closing cash − balance-sheet cash — must be 0)")
    ws.cell(row=11, column=2, value="=B10-Balance_Sheet!D2").font = Font(bold=True)
    for r in range(2, 12):
        ws.cell(row=r, column=2).number_format = MONEY
    widths(ws, {"A": 52, "B": 18})

    # ---- Budget_vs_Actual ----
    ws = wb.create_sheet("Budget_vs_Actual")
    for j, h in enumerate(["Category", "Type", "Actual", "Budget", "Variance",
                           "Variance %", "Status"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 7)
    for i, (cat, typ) in enumerate(CATEGORIES, start=2):
        ws.cell(row=i, column=1, value=cat)
        ws.cell(row=i, column=2, value=typ)
        ws.cell(row=i, column=3,
                value=f'=SUMIFS({TX_AMT},{TX_CAT},$A{i})')
        ws.cell(row=i, column=4,
                value=f'=SUMIFS({BU_AMT},{BU_CAT},$A{i})')
        ws.cell(row=i, column=5, value=f"=C{i}-D{i}")
        ws.cell(row=i, column=6, value=f"=IF(D{i}=0,0,E{i}/D{i})")
        ws.cell(row=i, column=7, value=(
            f'=IF($B{i}="Revenue",IF($F{i}<0,"UNDER BUDGET","ON TRACK"),'
            f'IF($F{i}>0.05,"OVER BUDGET","ON TRACK"))'))
    ws.cell(row=14, column=1, value="TOTAL (all categories)").font = Font(bold=True)
    ws.cell(row=14, column=3, value="=SUM(C2:C13)").font = Font(bold=True)
    ws.cell(row=14, column=4, value="=SUM(D2:D13)").font = Font(bold=True)
    ws.cell(row=14, column=5, value="=C14-D14").font = Font(bold=True)
    ws.cell(row=14, column=6, value="=IF(D14=0,0,E14/D14)").font = Font(bold=True)
    ws.cell(row=15, column=1, value="Total expenses (Direct + Operating)").font = Font(bold=True)
    ws.cell(row=15, column=3, value="=SUM(C5:C13)").font = Font(bold=True)
    ws.cell(row=15, column=4, value="=SUM(D5:D13)").font = Font(bold=True)
    ws.cell(row=15, column=5, value="=C15-D15").font = Font(bold=True)
    ws.cell(row=15, column=6, value="=IF(D15=0,0,E15/D15)").font = Font(bold=True)
    for r in range(2, 16):
        for col in (3, 4, 5):
            ws.cell(row=r, column=col).number_format = MONEY
        ws.cell(row=r, column=6).number_format = PCT
    widths(ws, {"A": 22, "B": 11, "C": 15, "D": 15, "E": 15, "F": 12, "G": 14})

    # ---- Raw_Transactions ----
    ws = wb.create_sheet("Raw_Transactions")
    for j, h in enumerate(["Date", "Description", "Category", "Type",
                           "Amount", "MonthNum"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 6)
    for i, t in enumerate(txns, start=2):
        ws.cell(row=i, column=1, value=dval(t["date"])).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2, value=t["description"])
        ws.cell(row=i, column=3, value=t["category"])
        ws.cell(row=i, column=4, value=t["type"])
        ws.cell(row=i, column=5, value=int(t["amount"])).number_format = MONEY
        ws.cell(row=i, column=6, value=f"=MONTH(A{i})")
    ws.auto_filter.ref = f"A1:F{txn_last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 12, "B": 42, "C": 22, "D": 11, "E": 14, "F": 10})

    # ---- Raw_Budget ----
    ws = wb.create_sheet("Raw_Budget")
    for j, h in enumerate(["Month", "MonthNum", "Category", "Type", "Budget"],
                          start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 5)
    for i, b in enumerate(budgets, start=2):
        ws.cell(row=i, column=1, value=dval(b["month"])).number_format = "mmm yyyy"
        ws.cell(row=i, column=2, value=f"=MONTH(A{i})")
        ws.cell(row=i, column=3, value=b["category"])
        ws.cell(row=i, column=4, value=b["type"])
        ws.cell(row=i, column=5, value=int(b["budget"])).number_format = MONEY
    ws.auto_filter.ref = f"A1:E{bud_last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 12, "B": 10, "C": 22, "D": 11, "E": 14})

    # ---- Assumptions ----
    ws = wb.create_sheet("Assumptions")
    ws.cell(row=1, column=1, value="Assumption").font = Font(bold=True)
    ws.cell(row=1, column=2, value="Value").font = Font(bold=True)
    style_header(ws, 1, 2)
    assumptions = [
        ("Opening cash (1 Jan 2026)", 850000, MONEY),
        ("Opening accounts receivable", 120000, MONEY),
        ("Opening equipment (net)", 400000, MONEY),
        ("Opening accounts payable", 60000, MONEY),
        ("Opening equity", 1310000, MONEY),
        ("Depreciation per month", 5000, MONEY),
        ("Equipment purchased (Mar 2026)", 75000, MONEY),
        ("Receivables: share of October revenue uncollected", 0.20, PCT),
        ("Payables: share of October expenses unpaid", 0.10, PCT),
        ("Owner drawings (Jan–Oct 2026)", 200000, MONEY),
    ]
    for i, (label, value, fmt) in enumerate(assumptions, start=2):
        ws.cell(row=i, column=1, value=label)
        c = ws.cell(row=i, column=2, value=value)
        c.number_format = fmt
    ws.cell(row=12, column=1, value="As of date")
    ws.cell(row=12, column=2, value=datetime.date(2026, 10, 31)).number_format = "yyyy-mm-dd"
    ws.cell(row=13, column=1, value="Reporting period")
    ws.cell(row=13, column=2, value="Jan–Oct 2026")
    widths(ws, {"A": 48, "B": 18})

    # ---- Checks ----
    ws = wb.create_sheet("Checks")
    for j, h in enumerate(["Check", "Result", "Expected", "Status"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 4)
    checks = [
        ("Balance sheet balances (gap = 0)", "=Balance_Sheet!D13", 0),
        ("Cash-flow tie-out (gap = 0)", "=CashFlow!B11", 0),
        ("P&L revenue = SUMIFS revenue total",
         f"=PL_Monthly!B12-SUMIFS({TX_AMT},{TX_TYP},\"Revenue\")", 0),
        ("BvA actuals = P&L totals",
         "=SUM(Budget_vs_Actual!C2:C13)-(PL_Monthly!B12+PL_Monthly!C12+PL_Monthly!E12)", 0),
        ("BvA budgets = SUMIFS budget total",
         f"=SUM(Budget_vs_Actual!D2:D13)-SUM({BU_AMT})", 0),
        ("Receivables = 20% of October revenue",
         "=Balance_Sheet!D3-Assumptions!B9*PL_Monthly!B11", 0),
        ("Payables = 10% of October expenses",
         "=Balance_Sheet!D6-Assumptions!B10*(PL_Monthly!C11+PL_Monthly!E11)", 0),
        ("Equipment = opening + purchases − depreciation",
         "=Balance_Sheet!D4-(Assumptions!B4+Assumptions!B8-SUM(PL_Monthly!F2:F11))", 0),
        ("Dashboard revenue = P&L total", "=Dashboard!B4-PL_Monthly!B12", 0),
        ("All 10 months present", "=COUNTA(PL_Monthly!A2:A11)", 10),
    ]
    for i, (label, formula, expected) in enumerate(checks, start=2):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=formula).number_format = MONEY
        ws.cell(row=i, column=3, value=expected)
        if i < 11:
            ws.cell(row=i, column=4, value=f'=IF(ABS(B{i})<=0.01,"PASS","FAIL")')
        else:
            ws.cell(row=i, column=4, value=f'=IF(B{i}=C{i},"PASS","FAIL")')
    widths(ws, {"A": 52, "B": 16, "C": 12, "D": 10})

    # ---- Dashboard charts (all source sheets exist now) ----
    dash = wb["Dashboard"]
    line = LineChart()
    line.title = "Revenue vs budget by month"
    line.y_axis.title = "Rs"
    line.add_data(Reference(wb["PL_Monthly"], min_col=2, min_row=1, max_row=11),
                  titles_from_data=True)
    line.add_data(Reference(wb["PL_Monthly"], min_col=9, min_row=1, max_row=11),
                  titles_from_data=True)
    line.set_categories(Reference(wb["PL_Monthly"], min_col=1, min_row=2, max_row=11))
    dash.add_chart(line, "A13")
    bar = BarChart()
    bar.title = "Cumulative net income"
    bar.y_axis.title = "Rs"
    bar.add_data(Reference(wb["PL_Monthly"], min_col=10, min_row=1, max_row=11),
                 titles_from_data=True)
    bar.set_categories(Reference(wb["PL_Monthly"], min_col=1, min_row=2, max_row=11))
    dash.add_chart(bar, "A28")
    pie = PieChart()
    pie.title = "Operating + direct expense mix (actual)"
    pie.add_data(Reference(wb["Budget_vs_Actual"], min_col=3, min_row=5,
                           max_row=13), titles_from_data=False)
    pie.set_categories(Reference(wb["Budget_vs_Actual"], min_col=1,
                                 min_row=5, max_row=13))
    dash.add_chart(pie, "A43")

    wb.save(OUT)
    print(f"wrote {OUT} ({n_txn} transactions)")


if __name__ == "__main__":
    main()
