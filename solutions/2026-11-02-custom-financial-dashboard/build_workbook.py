#!/usr/bin/env python3
"""Build finance-dashboard.xlsx from data/*.csv — every figure a live formula.

Sheets: Guide, Dashboard, Monthly, Income, Expenses, Sales,
Investments, Assumptions, Checks.
Rules honoured (learned the hard way): plain A1 ranges only, range
literals repeat the column letter on both ends ($D$2:$D$201), no Table
objects, chart series titles come from header cells (titles_from_data),
dates are real date values, charts are added only after every source
sheet exists. No sheet protection anywhere — every formula is unlocked
and inspectable, as the brief asks. Raw entry sheets carry 200
pre-filled formula rows of headroom, so new manual entries pasted
inside the headroom are picked up without editing any formula; a
blank date yields month key 189912, which never matches a 2026 month.
"""
import csv
import datetime
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
OUT = os.path.join(HERE, "finance-dashboard.xlsx")
HEADROOM = 200

NAVY = "1F3864"
LIGHT = "D9E2F3"
MONEY = "#,##0"
MONEY2 = "#,##0.00"
PCT = "0.0%"

MONTHS = [(202601 + i, label) for i, label in enumerate(
    ["Jan 2026", "Feb 2026", "Mar 2026", "Apr 2026", "May 2026",
     "Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026", "Oct 2026"])]

EXP_CATEGORIES = ["Housing", "Food", "Transport", "Utilities",
                  "Healthcare", "Education", "Entertainment", "Shopping",
                  "Subscriptions", "Other"]


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
    income = read_csv("income.csv")
    expenses = read_csv("expenses.csv")
    sales = read_csv("sales.csv")
    investments = read_csv("investments.csv")

    inc_last = 1 + len(income) + HEADROOM
    exp_last = 1 + len(expenses) + HEADROOM
    sal_last = 1 + len(sales) + HEADROOM
    inv_last = 1 + len(investments)  # holdings end; totals row follows
    inv_total_row = inv_last + 1

    INC_AMT = f"Income!$D$2:$D${inc_last}"
    INC_KEY = f"Income!$E$2:$E${inc_last}"
    EXP_CAT = f"Expenses!$B$2:$B${exp_last}"
    EXP_AMT = f"Expenses!$D$2:$D${exp_last}"
    EXP_KEY = f"Expenses!$E$2:$E${exp_last}"
    SAL_REV = f"Sales!$F$2:$F${sal_last}"
    SAL_COST = f"Sales!$G$2:$G${sal_last}"
    SAL_KEY = f"Sales!$I$2:$I${sal_last}"

    wb = Workbook()

    # ---- Guide ----
    ws = wb.active
    ws.title = "Guide"
    guide = [
        ("FinancePulse — personal financial dashboard (demo)", True),
        ("Built around a real Freelancer brief: a manual-entry interactive",
         False),
        ("dashboard showing income / expenses / cash balance, sales / costs /",
         False),
        ("profit, and investments / returns / portfolio value — with unlocked",
         False),
        ("formulas and a hand-off guide. Fictional data: A. Sharma, Patna,",
         False),
        ("Jan–Oct 2026. No real client data; nothing here is client work.", False),
        ("", False),
        ("HOW TO USE IT (manual entry)", True),
        ("1. Type new rows into Income, Expenses or Sales — inside the shaded",
         False),
        ("   headroom rows. Amount, revenue, cost, profit and month-key",
         False),
        ("   columns are formulas: they fill themselves in. Do not type over",
         False),
        ("   a formula column; type only Date / Source / Category / Amount",
         False),
        ("   (or Product / Qty / prices on Sales).", False),
        ("2. On Investments, edit Units / Buy price / Current price in place;",
         False),
        ("   add a holding by inserting a row above the TOTAL row so the",
         False),
        ("   totals pick it up.", False),
        ("3. Monthly, Dashboard and Checks update themselves. If a Checks row",
         False),
        ("   ever reads FAIL, stop and find the broken entry — do not paste",
         False),
        ("   values over a formula to force it green.", False),
        ("4. No sheet is protected and no macro is used. Every formula is",
         False),
        ("   visible in the formula bar — that is deliberate, per the brief.", False),
        ("", False),
        ("WHAT EACH SHEET IS", True),
        ("Dashboard — 12 KPIs + 3 charts. Monthly — month-by-month income,",
         False),
        ("expenses, net, sales, gross profit, cumulative cash, savings rate.",
         False),
        ("Income / Expenses / Sales — the only sheets you type into. Sales",
         False),
        ("also carries the side business. Investments — holdings, invested,",
         False),
        ("value, gain. Assumptions — opening cash. Checks — 10 live checks.",
         False),
    ]
    for i, (text, bold) in enumerate(guide, start=1):
        c = ws.cell(row=i, column=1, value=text)
        c.font = Font(bold=bold, size=12 if i == 1 else 11)
    widths(ws, {"A": 78})

    # ---- Dashboard (KPIs; charts added at the end) ----
    ws = wb.create_sheet("Dashboard")
    ws.cell(row=1, column=1, value="FinancePulse dashboard — Jan–Oct 2026").font = Font(bold=True, size=13)
    ws.cell(row=3, column=1, value="KPI").font = Font(bold=True)
    ws.cell(row=3, column=2, value="Value").font = Font(bold=True)
    ws.cell(row=3, column=4, value="How computed").font = Font(bold=True)
    kpis = [
        ("Total income (Rs)", "=Monthly!C12", MONEY, "Sum of Monthly income"),
        ("Total expenses (Rs)", "=Monthly!D12", MONEY, "Sum of Monthly expenses"),
        ("Net cash flow (Rs)", "=Monthly!E12", MONEY, "Income minus expenses"),
        ("Cash balance, closing (Rs)", "=Monthly!I12", MONEY, "Opening cash + cumulative net"),
        ("Total sales revenue (Rs)", "=Monthly!F12", MONEY, "Sum of Sales revenue"),
        ("Total sales cost (Rs)", "=Monthly!G12", MONEY, "Sum of Sales cost"),
        ("Gross profit (Rs)", "=Monthly!H12", MONEY, "Sales revenue minus sales cost"),
        ("Savings rate", "=Monthly!J12", PCT, "Net cash flow / total income"),
        ("Portfolio value (Rs)", f"=Investments!G{inv_total_row}", MONEY, "Sum of holding values"),
        ("Total invested (Rs)", f"=Investments!F{inv_total_row}", MONEY, "Sum of units x buy price"),
        ("Total gain (Rs)", f"=Investments!H{inv_total_row}", MONEY, "Portfolio value minus invested"),
        ("Gain %", f"=Investments!I{inv_total_row}", PCT, "Total gain / total invested"),
    ]
    for i, (label, formula, fmt, note) in enumerate(kpis, start=4):
        ws.cell(row=i, column=1, value=label)
        c = ws.cell(row=i, column=2, value=formula)
        c.number_format = fmt
        ws.cell(row=i, column=4, value=note)
    ws.cell(row=18, column=1, value="Expense category").font = Font(bold=True)
    ws.cell(row=18, column=2, value="Total (Rs)").font = Font(bold=True)
    for i, cat in enumerate(EXP_CATEGORIES, start=19):
        ws.cell(row=i, column=1, value=cat)
        c = ws.cell(row=i, column=2,
                    value=f"=SUMIF({EXP_CAT},$A{i},{EXP_AMT})")
        c.number_format = MONEY
    widths(ws, {"A": 30, "B": 18, "D": 44})

    # ---- Monthly ----
    ws = wb.create_sheet("Monthly")
    headers = ["MonthKey", "Month", "Income", "Expenses", "Net",
               "Sales revenue", "Sales cost", "Gross profit",
               "Cumulative cash", "Savings rate"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, len(headers))
    for i, (key, label) in enumerate(MONTHS, start=2):
        ws.cell(row=i, column=1, value=key)
        ws.cell(row=i, column=2, value=label)
        ws.cell(row=i, column=3, value=f"=SUMIFS({INC_AMT},{INC_KEY},$A{i})")
        ws.cell(row=i, column=4, value=f"=SUMIFS({EXP_AMT},{EXP_KEY},$A{i})")
        ws.cell(row=i, column=5, value=f"=C{i}-D{i}")
        ws.cell(row=i, column=6, value=f"=SUMIFS({SAL_REV},{SAL_KEY},$A{i})")
        ws.cell(row=i, column=7, value=f"=SUMIFS({SAL_COST},{SAL_KEY},$A{i})")
        ws.cell(row=i, column=8, value=f"=F{i}-G{i}")
        ws.cell(row=i, column=9, value=f"=Assumptions!$B$2+SUM($E$2:E{i})")
        ws.cell(row=i, column=10, value=f"=IF(C{i}=0,0,E{i}/C{i})")
        for col in (3, 4, 5, 6, 7, 8, 9):
            ws.cell(row=i, column=col).number_format = MONEY
        ws.cell(row=i, column=10).number_format = PCT
    ws.cell(row=12, column=2, value="TOTAL")
    for col, letter in ((3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H")):
        c = ws.cell(row=12, column=col, value=f"=SUM({letter}2:{letter}11)")
        c.number_format = MONEY
        c.font = Font(bold=True)
    c = ws.cell(row=12, column=9, value="=I11")
    c.number_format = MONEY
    c.font = Font(bold=True)
    c = ws.cell(row=12, column=10, value="=IF(C12=0,0,E12/C12)")
    c.number_format = PCT
    c.font = Font(bold=True)
    ws.freeze_panes = "A2"
    widths(ws, {"A": 11, "B": 11, "C": 14, "D": 14, "E": 14, "F": 15,
                "G": 14, "H": 14, "I": 16, "J": 13})

    # ---- Income ----
    ws = wb.create_sheet("Income")
    for j, h in enumerate(["Date", "Source", "Category", "Amount", "MonthKey"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 5)
    for i, r in enumerate(income, start=2):
        ws.cell(row=i, column=1, value=dval(r["date"])).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2, value=r["source"])
        ws.cell(row=i, column=3, value=r["category"])
        ws.cell(row=i, column=4, value=int(r["amount"])).number_format = MONEY
    for i in range(2, inc_last + 1):
        ws.cell(row=i, column=5, value=f"=YEAR(A{i})*100+MONTH(A{i})")
    ws.auto_filter.ref = f"A1:E{inc_last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 13, "B": 22, "C": 12, "D": 13, "E": 11})

    # ---- Expenses ----
    ws = wb.create_sheet("Expenses")
    for j, h in enumerate(["Date", "Category", "Description", "Amount", "MonthKey"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 5)
    for i, r in enumerate(expenses, start=2):
        ws.cell(row=i, column=1, value=dval(r["date"])).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2, value=r["category"])
        ws.cell(row=i, column=3, value=r["description"])
        ws.cell(row=i, column=4, value=int(r["amount"])).number_format = MONEY
    for i in range(2, exp_last + 1):
        ws.cell(row=i, column=5, value=f"=YEAR(A{i})*100+MONTH(A{i})")
    ws.auto_filter.ref = f"A1:E{exp_last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 13, "B": 14, "C": 30, "D": 13, "E": 11})

    # ---- Sales ----
    ws = wb.create_sheet("Sales")
    for j, h in enumerate(["Date", "Product", "Qty", "Unit price", "Cost per unit",
                           "Revenue", "Cost", "Profit", "MonthKey"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 9)
    for i, r in enumerate(sales, start=2):
        ws.cell(row=i, column=1, value=dval(r["date"])).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2, value=r["product"])
        ws.cell(row=i, column=3, value=int(r["qty"]))
        ws.cell(row=i, column=4, value=int(r["unit_price"])).number_format = MONEY
        ws.cell(row=i, column=5, value=int(r["cost_per_unit"])).number_format = MONEY
    for i in range(2, sal_last + 1):
        ws.cell(row=i, column=6, value=f"=C{i}*D{i}").number_format = MONEY
        ws.cell(row=i, column=7, value=f"=C{i}*E{i}").number_format = MONEY
        ws.cell(row=i, column=8, value=f"=F{i}-G{i}").number_format = MONEY
        ws.cell(row=i, column=9, value=f"=YEAR(A{i})*100+MONTH(A{i})")
    ws.auto_filter.ref = f"A1:I{sal_last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 13, "B": 20, "C": 7, "D": 12, "E": 14, "F": 12,
                "G": 12, "H": 12, "I": 11})

    # ---- Investments ----
    ws = wb.create_sheet("Investments")
    for j, h in enumerate(["Holding", "Type", "Units", "Buy price", "Current price",
                           "Invested", "Value", "Gain", "Gain %"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 9)
    for i, r in enumerate(investments, start=2):
        ws.cell(row=i, column=1, value=r["holding"])
        ws.cell(row=i, column=2, value=r["type"])
        ws.cell(row=i, column=3, value=float(r["units"])).number_format = MONEY2
        ws.cell(row=i, column=4, value=float(r["buy_price"])).number_format = MONEY2
        ws.cell(row=i, column=5, value=float(r["current_price"])).number_format = MONEY2
        ws.cell(row=i, column=6, value=f"=C{i}*D{i}").number_format = MONEY2
        ws.cell(row=i, column=7, value=f"=C{i}*E{i}").number_format = MONEY2
        ws.cell(row=i, column=8, value=f"=G{i}-F{i}").number_format = MONEY2
        ws.cell(row=i, column=9, value=f"=IF(F{i}=0,0,H{i}/F{i})").number_format = PCT
    ws.cell(row=inv_total_row, column=1, value="TOTAL").font = Font(bold=True)
    for col, letter in ((6, "F"), (7, "G"), (8, "H")):
        c = ws.cell(row=inv_total_row, column=col,
                    value=f"=SUM({letter}2:{letter}{inv_last})")
        c.number_format = MONEY2
        c.font = Font(bold=True)
    c = ws.cell(row=inv_total_row, column=9,
                value=f"=IF(F{inv_total_row}=0,0,H{inv_total_row}/F{inv_total_row})")
    c.number_format = PCT
    c.font = Font(bold=True)
    ws.freeze_panes = "A2"
    widths(ws, {"A": 20, "B": 12, "C": 10, "D": 12, "E": 14, "F": 14,
                "G": 14, "H": 14, "I": 10})

    # ---- Assumptions ----
    ws = wb.create_sheet("Assumptions")
    ws.cell(row=1, column=1, value="Item").font = Font(bold=True)
    ws.cell(row=1, column=2, value="Value").font = Font(bold=True)
    ws.cell(row=2, column=1, value="Opening cash balance, 1 Jan 2026 (Rs)")
    ws.cell(row=2, column=2, value=125000).number_format = MONEY
    ws.cell(row=3, column=1, value="Currency")
    ws.cell(row=3, column=2, value="INR")
    ws.cell(row=4, column=1, value="Period")
    ws.cell(row=4, column=2, value="Jan–Oct 2026")
    ws.cell(row=5, column=1, value="Entry headroom (formula rows per raw sheet)")
    ws.cell(row=5, column=2, value=HEADROOM)
    widths(ws, {"A": 46, "B": 16})

    # ---- Checks ----
    ws = wb.create_sheet("Checks")
    for j, h in enumerate(["Check", "Value", "Expected", "Result"], start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 4)
    checks = [
        ("Dashboard income = sum of Income amounts", "=Dashboard!B4", f"=SUM({INC_AMT})"),
        ("Dashboard expenses = sum of Expense amounts", "=Dashboard!B5", f"=SUM({EXP_AMT})"),
        ("Monthly net total = Dashboard net", "=Monthly!E12", "=Dashboard!B6"),
        ("Cash balance = opening cash + net", "=Dashboard!B7", "=Assumptions!B2+Dashboard!B6"),
        ("Sales revenue tie: Monthly = sum of Sales revenue", "=Monthly!F12", f"=SUM({SAL_REV})"),
        ("Gross profit = sales revenue - sales cost", "=Monthly!H12", "=Monthly!F12-Monthly!G12"),
        ("Portfolio value = sum of holding values", f"=Investments!G{inv_total_row}",
         f"=SUM(Investments!G2:G{inv_last})"),
        ("Total gain = portfolio value - invested", f"=Investments!H{inv_total_row}",
         f"=Investments!G{inv_total_row}-Investments!F{inv_total_row}"),
        ("No negative amounts (minimum entry >= 0)",
         f"=MIN(MIN({INC_AMT}),MIN({EXP_AMT}))", "0"),
        ("Expense mix total = total expenses", "=SUM(Dashboard!B19:B28)", "=Dashboard!B5"),
    ]
    for i, (label, val_f, exp_f) in enumerate(checks, start=2):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=val_f).number_format = MONEY2
        ws.cell(row=i, column=3, value=exp_f).number_format = MONEY2
        ws.cell(row=i, column=4, value=f'=IF(ABS(B{i}-C{i})<0.01,"PASS","FAIL")')
    # Check 9 is a threshold check, not an equality: minimum entry >= 0.
    ws.cell(row=10, column=4, value='=IF(B10>=0,"PASS","FAIL")')
    widths(ws, {"A": 52, "B": 18, "C": 18, "D": 10})

    # ---- Charts (last, after all source sheets exist) ----
    dash = wb["Dashboard"]
    bar = BarChart()
    bar.title = "Income vs expenses by month"
    bar.y_axis.title = "Rs"
    data = Reference(wb["Monthly"], min_col=3, min_row=1, max_row=11, max_col=4)
    cats = Reference(wb["Monthly"], min_col=2, min_row=2, max_row=11)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    dash.add_chart(bar, "A31")
    line = LineChart()
    line.title = "Cumulative cash balance"
    line.y_axis.title = "Rs"
    data2 = Reference(wb["Monthly"], min_col=9, min_row=1, max_row=11)
    line.add_data(data2, titles_from_data=True)
    line.set_categories(cats)
    dash.add_chart(line, "A46")
    pie = PieChart()
    pie.title = "Expense mix"
    labels = Reference(dash, min_col=1, min_row=19, max_row=28)
    piedata = Reference(dash, min_col=2, min_row=18, max_row=28)
    pie.add_data(piedata, titles_from_data=True)
    pie.set_categories(labels)
    dash.add_chart(pie, "A61")

    wb.save(OUT)
    print(f"wrote {OUT} (income {len(income)}, expenses {len(expenses)}, "
          f"sales {len(sales)}, investments {len(investments)})")


if __name__ == "__main__":
    main()
