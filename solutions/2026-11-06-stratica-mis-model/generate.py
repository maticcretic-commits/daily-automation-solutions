#!/usr/bin/env python3
"""StraticaMIS — 3-statement financial model + MIS dashboard for an Indian SME.

Built around a real Freelancer brief ("Stratica OS — Financial Planning,
Budgeting & Investor-Ready Financial Model", project 40804245): develop
financial dashboards and MIS reports plus a 3-statement model with budgets,
forecasts and cash-flow management.

Demo company: Shree Traders, a fictional wholesale trading SME. Demo data is
fictional and deterministic (seeded): monthly inputs Apr 2025 - Mar 2026
(FY 2026-27). Every figure in the workbook is a live Excel formula off the
Inputs sheet — nothing hard-coded except the monthly inputs themselves and
the opening balance sheet at 31 Mar 2025.

Sheets: Guide, Inputs, PL (profit & loss), BS (balance sheet), CF (cash
flow, indirect method), Budget (budget vs actual), Dashboard (10 KPIs +
3 native charts), Checks. No macros, no structured table references, plain
A1 ranges with the column letter repeated on both ends (LibreOffice-safe).

Model rules (all visible in the formulas):
- Receivables close at 50% of monthly revenue (15 days), inventory at 80%
  of monthly COGS (24 days), payables at 45% of monthly COGS.
- Depreciation is a flat Rs 60,000/month; term loan repays Rs 1,00,000/month
  and bears 1% interest per month on the opening balance.
- Tax is 25% of profit before tax when positive, else nil.
- Cash flow uses the indirect method, so the balance sheet balances by
  construction; the Checks sheet verifies it numerically.
- Budget revenue starts at Rs 19,00,000 and grows 4% a month; budget gross
  margin is 45% and budget operating expense is a flat Rs 5,20,000/month.
"""
import csv
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
N = 12  # Apr 2025 - Mar 2026
FIRST, LAST = 2, 2 + N - 1  # data rows (2..13); row 14 = FY total
FY = 14

# Opening balance sheet, 31 Mar 2025 (Rs).
OPEN_CASH = 1_200_000
OPEN_RECV = 1_800_000
OPEN_INV = 1_500_000
OPEN_FA = 4_000_000
OPEN_PAY = 1_000_000
OPEN_LOAN = 2_500_000
OPEN_CAPITAL = 3_000_000
OPEN_RE = 2_000_000
DEPN = 60_000
LOAN_REPAY = 100_000
TAX_RATE = 0.25

NAVY = "1F3864"
TEAL = "0F7B7A"
LIGHT = "D9E2F3"
GREY = "F2F2F2"
MONEY = "#,##0"
PCT1 = "0.0%"
PCT2 = "0.00%"


def build_data():
    rng = random.Random(20261106)
    months = []
    y, m = 2025, 4
    for _ in range(N):
        months.append(f"{y:04d}-{m:02d}")
        m += 1
        if m == 13:
            y, m = y + 1, 1
    seasonal = {9: 1.08, 10: 1.06, 11: 1.03}  # Oct/Nov/Dec festive bump (0-based)
    capex = {2: 250_000, 6: 400_000, 9: 150_000}  # Jun, Oct, Jan
    rows = []
    for i, month in enumerate(months):
        rev = 1_850_000 * (1.035 ** i) * seasonal.get(i, 1.0) * (1 + rng.gauss(0, 0.03))
        rev = round(rev / 1000) * 1000
        cogs = round(rev * min(0.57, max(0.53, rng.gauss(0.55, 0.01))) / 1000) * 1000
        sal = round(260_000 * (1.01 ** i) * (1 + rng.gauss(0, 0.02)) / 1000) * 1000
        rent = 85_000
        mkt = max(80_000, round(140_000 * (1 + rng.gauss(0, 0.15)) / 1000) * 1000)
        adm = round(105_000 * (1 + rng.gauss(0, 0.08)) / 1000) * 1000
        oth = round(60_000 * (1 + rng.gauss(0, 0.10)) / 1000) * 1000
        rows.append((month, rev, cogs, sal, rent, mkt, adm, oth, capex.get(i, 0)))
    return rows


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def money_range(ws, rows, cols, fmt=MONEY):
    for r in rows:
        for c in cols:
            ws.cell(row=r, column=c).number_format = fmt


def build_workbook(rows):
    wb = Workbook()

    # ---- Inputs ----
    ws = wb.active
    ws.title = "Inputs"
    headers = ["Month", "Revenue", "COGS", "Salaries", "Rent", "Marketing",
               "Admin", "Other opex", "Capex"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, row in enumerate(rows):
        r = FIRST + i
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)
    money_range(ws, range(FIRST, LAST + 1), range(2, 10))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{LAST}"
    for i, w in enumerate([10, 15, 15, 15, 12, 15, 15, 15, 15], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- PL ----
    pl = wb.create_sheet("PL")
    pl_headers = ["Month", "Revenue", "COGS", "Gross profit", "Salaries", "Rent",
                  "Marketing", "Admin", "Other opex", "Total opex", "EBITDA",
                  "Depreciation", "Interest", "Profit before tax", "Tax",
                  "Net profit"]
    pl.append(pl_headers)
    style_header(pl, 1, len(pl_headers))
    for i in range(N):
        r = FIRST + i
        pl.cell(row=r, column=1, value=f"=Inputs!A{r}")
        pl.cell(row=r, column=2, value=f"=Inputs!B{r}")
        pl.cell(row=r, column=3, value=f"=Inputs!C{r}")
        pl.cell(row=r, column=4, value=f"=B{r}-C{r}")
        pl.cell(row=r, column=5, value=f"=Inputs!D{r}")
        pl.cell(row=r, column=6, value=f"=Inputs!E{r}")
        pl.cell(row=r, column=7, value=f"=Inputs!F{r}")
        pl.cell(row=r, column=8, value=f"=Inputs!G{r}")
        pl.cell(row=r, column=9, value=f"=Inputs!H{r}")
        pl.cell(row=r, column=10, value=f"=SUM(E{r}:I{r})")
        pl.cell(row=r, column=11, value=f"=D{r}-J{r}")
        pl.cell(row=r, column=12, value=DEPN)
        # Interest: 1% of opening term-loan balance (BS col H, prior row).
        if i == 0:
            pl.cell(row=r, column=13, value=f"=ROUND({OPEN_LOAN}*0.01,0)")
        else:
            pl.cell(row=r, column=13, value=f"=ROUND(BS!H{r-1}*0.01,0)")
        pl.cell(row=r, column=14, value=f"=K{r}-L{r}-M{r}")
        pl.cell(row=r, column=15, value=f"=ROUND({TAX_RATE}*MAX(N{r},0),0)")
        pl.cell(row=r, column=16, value=f"=N{r}-O{r}")
    pl.cell(row=FY, column=1, value="FY 2026-27")
    for c in range(2, 17):
        col = get_column_letter(c)
        pl.cell(row=FY, column=c, value=f"=SUM({col}{FIRST}:{col}{LAST})")
        pl.cell(row=FY, column=c).font = Font(bold=True)
    pl.cell(row=FY, column=1).font = Font(bold=True)
    money_range(pl, list(range(FIRST, LAST + 1)) + [FY], range(2, 17))
    pl.freeze_panes = "A2"
    for i, w in enumerate([10, 14, 14, 14, 14, 12, 14, 14, 14, 14, 14, 14, 14, 16, 14, 14], 1):
        pl.column_dimensions[get_column_letter(i)].width = w

    # ---- BS ----
    bs = wb.create_sheet("BS")
    bs_headers = ["Month", "Cash", "Receivables", "Inventory",
                  "Fixed assets (net)", "Total assets", "Payables", "Term loan",
                  "Share capital", "Retained earnings", "Total liab. + equity",
                  "Balance check", "RE movement check"]
    bs.append(bs_headers)
    style_header(bs, 1, len(bs_headers))
    for i in range(N):
        r = FIRST + i
        bs.cell(row=r, column=1, value=f"=Inputs!A{r}")
        # Cash comes from the cash-flow statement (closing cash).
        bs.cell(row=r, column=2, value=f"=CF!N{r}")
        bs.cell(row=r, column=3, value=f"=ROUND(Inputs!B{r}*0.50,0)")
        bs.cell(row=r, column=4, value=f"=ROUND(Inputs!C{r}*0.80,0)")
        if i == 0:
            bs.cell(row=r, column=5, value=f"={OPEN_FA}+Inputs!I{r}-{DEPN}")
        else:
            bs.cell(row=r, column=5, value=f"=E{r-1}+Inputs!I{r}-{DEPN}")
        bs.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})")
        bs.cell(row=r, column=7, value=f"=ROUND(Inputs!C{r}*0.45,0)")
        if i == 0:
            bs.cell(row=r, column=8, value=f"={OPEN_LOAN}-{LOAN_REPAY}")
        else:
            bs.cell(row=r, column=8, value=f"=H{r-1}-{LOAN_REPAY}")
        bs.cell(row=r, column=9, value=OPEN_CAPITAL)
        if i == 0:
            bs.cell(row=r, column=10, value=f"={OPEN_RE}+PL!P{r}")
        else:
            bs.cell(row=r, column=10, value=f"=J{r-1}+PL!P{r}")
        bs.cell(row=r, column=11, value=f"=SUM(G{r}:J{r})")
        bs.cell(row=r, column=12, value=f"=F{r}-K{r}")
        if i == 0:
            bs.cell(row=r, column=13, value=f"=J{r}-{OPEN_RE}-PL!P{r}")
        else:
            bs.cell(row=r, column=13, value=f"=J{r}-J{r-1}-PL!P{r}")
    money_range(bs, range(FIRST, LAST + 1), range(2, 14))
    bs.freeze_panes = "A2"
    for i, w in enumerate([10, 14, 14, 14, 16, 14, 14, 14, 14, 16, 18, 14, 16], 1):
        bs.column_dimensions[get_column_letter(i)].width = w

    # ---- CF (indirect method) ----
    cf = wb.create_sheet("CF")
    cf_headers = ["Month", "Net profit", "+ Depreciation", "- Change in receivables",
                  "- Change in inventory", "+ Change in payables", "Operating cash flow",
                  "Capex", "Investing cash flow", "Loan repayment",
                  "Financing cash flow", "Net change in cash", "Opening cash",
                  "Closing cash", "Cash tie check"]
    cf.append(cf_headers)
    style_header(cf, 1, len(cf_headers))
    for i in range(N):
        r = FIRST + i
        cf.cell(row=r, column=1, value=f"=Inputs!A{r}")
        cf.cell(row=r, column=2, value=f"=PL!P{r}")
        cf.cell(row=r, column=3, value=DEPN)
        if i == 0:
            cf.cell(row=r, column=4, value=f"=-(BS!C{r}-{OPEN_RECV})")
            cf.cell(row=r, column=5, value=f"=-(BS!D{r}-{OPEN_INV})")
            cf.cell(row=r, column=6, value=f"=BS!G{r}-{OPEN_PAY}")
        else:
            cf.cell(row=r, column=4, value=f"=-(BS!C{r}-BS!C{r-1})")
            cf.cell(row=r, column=5, value=f"=-(BS!D{r}-BS!D{r-1})")
            cf.cell(row=r, column=6, value=f"=BS!G{r}-BS!G{r-1}")
        cf.cell(row=r, column=7, value=f"=SUM(B{r}:F{r})")
        cf.cell(row=r, column=8, value=f"=-Inputs!I{r}")
        cf.cell(row=r, column=9, value=f"=H{r}")
        cf.cell(row=r, column=10, value=f"=-{LOAN_REPAY}")
        cf.cell(row=r, column=11, value=f"=J{r}")
        cf.cell(row=r, column=12, value=f"=G{r}+I{r}+K{r}")
        if i == 0:
            cf.cell(row=r, column=13, value=OPEN_CASH)
        else:
            cf.cell(row=r, column=13, value=f"=N{r-1}")
        cf.cell(row=r, column=14, value=f"=M{r}+L{r}")
        cf.cell(row=r, column=15, value=f"=BS!B{r}-N{r}")
    cf.cell(row=FY, column=1, value="FY 2026-27")
    for c in range(2, 15):
        col = get_column_letter(c)
        if c in (13, 14, 15):
            continue
        cf.cell(row=FY, column=c, value=f"=SUM({col}{FIRST}:{col}{LAST})")
        cf.cell(row=FY, column=c).font = Font(bold=True)
    cf.cell(row=FY, column=1).font = Font(bold=True)
    money_range(cf, list(range(FIRST, LAST + 1)) + [FY], range(2, 16))
    cf.freeze_panes = "A2"
    for i, w in enumerate([10, 14, 14, 18, 18, 16, 18, 13, 16, 14, 16, 16, 14, 14, 14], 1):
        cf.column_dimensions[get_column_letter(i)].width = w

    # ---- Budget vs Actual ----
    bg = wb.create_sheet("Budget")
    bg_headers = ["Month", "Budget revenue", "Actual revenue", "Revenue variance (Rs)",
                  "Revenue variance %", "Budget gross profit", "Actual gross profit",
                  "Budget EBITDA", "Actual EBITDA", "Budget net profit",
                  "Actual net profit", "Net profit variance %"]
    bg.append(bg_headers)
    style_header(bg, 1, len(bg_headers))
    for i in range(N):
        r = FIRST + i
        bg.cell(row=r, column=1, value=f"=Inputs!A{r}")
        bg.cell(row=r, column=2, value=f"=ROUND(1900000*1.04^{i},0)")
        bg.cell(row=r, column=3, value=f"=PL!B{r}")
        bg.cell(row=r, column=4, value=f"=C{r}-B{r}")
        bg.cell(row=r, column=5, value=f"=IF(B{r}=0,0,D{r}/B{r})")
        bg.cell(row=r, column=6, value=f"=ROUND(B{r}*0.45,0)")
        bg.cell(row=r, column=7, value=f"=PL!D{r}")
        bg.cell(row=r, column=8, value=f"=F{r}-520000")
        bg.cell(row=r, column=9, value=f"=PL!K{r}")
        # Budget PBT = budget EBITDA - depreciation - budget interest (Rs 25,000).
        bg.cell(row=r, column=10,
               value=f"=(H{r}-{DEPN}-25000)-ROUND({TAX_RATE}*MAX(H{r}-{DEPN}-25000,0),0)")
        bg.cell(row=r, column=11, value=f"=PL!P{r}")
        bg.cell(row=r, column=12, value=f"=IF(J{r}=0,0,(K{r}-J{r})/J{r})")
    bg.cell(row=FY, column=1, value="FY 2026-27")
    for c in (2, 3, 4, 6, 7, 8, 9, 10, 11):
        col = get_column_letter(c)
        bg.cell(row=FY, column=c, value=f"=SUM({col}{FIRST}:{col}{LAST})")
        bg.cell(row=FY, column=c).font = Font(bold=True)
    bg.cell(row=FY, column=5, value=f"=IF(B{FY}=0,0,D{FY}/B{FY})").font = Font(bold=True)
    bg.cell(row=FY, column=12, value=f"=IF(J{FY}=0,0,(K{FY}-J{FY})/J{FY})").font = Font(bold=True)
    bg.cell(row=FY, column=1).font = Font(bold=True)
    money_range(bg, list(range(FIRST, LAST + 1)) + [FY], [2, 3, 4, 6, 7, 8, 9, 10, 11])
    for r in list(range(FIRST, LAST + 1)) + [FY]:
        for c in (5, 12):
            bg.cell(row=r, column=c).number_format = PCT1
    bg.freeze_panes = "A2"
    for i, w in enumerate([10, 15, 15, 18, 15, 16, 16, 15, 15, 16, 15, 16], 1):
        bg.column_dimensions[get_column_letter(i)].width = w

    # ---- Dashboard ----
    db = wb.create_sheet("Dashboard")
    db.append(["Shree Traders — MIS Dashboard, FY 2026-27 (fictional demo data)"])
    db["A1"].font = Font(bold=True, size=14, color=NAVY)
    db.append(["KPI", "Value"])
    style_header(db, 2, 2)
    kpis = [
        ("FY revenue", f"=PL!B{FY}", MONEY),
        ("FY net profit", f"=PL!P{FY}", MONEY),
        ("Net margin", f"=PL!P{FY}/PL!B{FY}", PCT1),
        ("EBITDA margin", f"=PL!K{FY}/PL!B{FY}", PCT1),
        ("Closing cash (Mar 2026)", f"=BS!B{LAST}", MONEY),
        ("Current ratio (Mar 2026)", f"=(BS!B{LAST}+BS!C{LAST}+BS!D{LAST})/BS!G{LAST}", "0.00"),
        ("Receivables days (Mar 2026)", f"=BS!C{LAST}/PL!B{LAST}*30", "0.0"),
        ("Inventory days (Mar 2026)", f"=BS!D{LAST}/PL!C{LAST}*30", "0.0"),
        ("Revenue budget attainment (FY)", f"=Budget!C{FY}/Budget!B{FY}", PCT1),
        ("FY operating cash flow", f"=CF!G{FY}", MONEY),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        r = 3 + i
        db.cell(row=r, column=1, value=label)
        c = db.cell(row=r, column=2, value=formula)
        c.number_format = fmt
    db.column_dimensions["A"].width = 36
    db.column_dimensions["B"].width = 18

    cats = Reference(bg, min_col=1, min_row=FIRST, max_row=LAST)
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "Revenue: budget vs actual"
    ch1.y_axis.number_format = "#,##0"
    ch1.add_data(Reference(bg, min_col=2, min_row=1, max_row=LAST), titles_from_data=True)
    ch1.add_data(Reference(bg, min_col=3, min_row=1, max_row=LAST), titles_from_data=True)
    ch1.set_categories(cats)
    ch1.height, ch1.width = 8, 16
    db.add_chart(ch1, "D2")
    ch2 = LineChart()
    ch2.title = "EBITDA by month"
    ch2.y_axis.number_format = "#,##0"
    ch2.add_data(Reference(pl, min_col=11, min_row=1, max_row=LAST), titles_from_data=True)
    ch2.set_categories(cats)
    ch2.height, ch2.width = 8, 16
    db.add_chart(ch2, "D20")
    ch3 = LineChart()
    ch3.title = "Closing cash balance"
    ch3.y_axis.number_format = "#,##0"
    ch3.add_data(Reference(cf, min_col=14, min_row=1, max_row=LAST), titles_from_data=True)
    ch3.set_categories(cats)
    ch3.height, ch3.width = 8, 16
    db.add_chart(ch3, "D38")

    # ---- Checks ----
    ck = wb.create_sheet("Checks")
    ck.append(["Check", "Result"])
    style_header(ck, 1, 2)
    checks = [
        ("Balance sheet balances every month (assets = liab. + equity)",
         f"=IF(AND(MAX(BS!L{FIRST}:L{LAST})=0,MIN(BS!L{FIRST}:L{LAST})=0),\"PASS\",\"FAIL\")"),
        ("Cash ties: balance sheet = cash-flow statement, every month",
         f"=IF(AND(MAX(CF!O{FIRST}:O{LAST})=0,MIN(CF!O{FIRST}:O{LAST})=0),\"PASS\",\"FAIL\")"),
        ("Retained earnings move by net profit each month",
         f"=IF(AND(MAX(BS!M{FIRST}:M{LAST})=0,MIN(BS!M{FIRST}:M{LAST})=0),\"PASS\",\"FAIL\")"),
        ("FY net profit = sum of monthly net profit",
         f"=IF(PL!P{FY}-SUM(PL!P{FIRST}:P{LAST})=0,\"PASS\",\"FAIL\")"),
        ("Tax is never negative",
         f"=IF(MIN(PL!O{FIRST}:O{LAST})>=0,\"PASS\",\"FAIL\")"),
        ("Closing cash is positive every month",
         f"=IF(MIN(CF!N{FIRST}:N{LAST})>0,\"PASS\",\"FAIL\")"),
        ("Term loan at Mar 2026 = Rs 13,00,000 (25,00,000 - 12 x 1,00,000)",
         f"=IF(BS!H{LAST}=1300000,\"PASS\",\"FAIL\")"),
        ("Fixed assets at Mar 2026 = opening + capex - 12 x depreciation",
         f"=IF(BS!E{LAST}=({OPEN_FA}+SUM(Inputs!I{FIRST}:I{LAST})-12*{DEPN}),\"PASS\",\"FAIL\")"),
    ]
    for i, (label, formula) in enumerate(checks):
        r = 2 + i
        ck.cell(row=r, column=1, value=label)
        ck.cell(row=r, column=2, value=formula)
    ck.column_dimensions["A"].width = 62
    ck.column_dimensions["B"].width = 12

    # ---- Guide ----
    gd = wb.create_sheet("Guide", 0)
    gd.append(["StraticaMIS — how to use this workbook"])
    gd["A1"].font = Font(bold=True, size=14, color=NAVY)
    lines = [
        "1. Paste your own monthly figures into Inputs (revenue, COGS, salary/rent/",
        "   marketing/admin/other opex, capex). The P&L, balance sheet, cash flow,",
        "   budget comparison, dashboard KPIs and charts all update from those inputs.",
        "2. Opening balances at 31 Mar 2025 sit inside the month-1 formulas on BS/CF",
        "   (cash 12,00,000; receivables 18,00,000; inventory 15,00,000; fixed assets",
        "   40,00,000; payables 10,00,000; term loan 25,00,000; capital 30,00,000;",
        "   retained earnings 20,00,000). Change them there if your opening differs.",
        "3. Model rules: receivables 50% of revenue, inventory 80% of COGS, payables",
        "   45% of COGS; depreciation Rs 60,000/month; loan repays Rs 1,00,000/month",
        "   at 1% monthly interest on the opening balance; tax 25% of positive PBT.",
        "   Cash flow uses the indirect method, so the balance sheet balances.",
        "4. Budget: revenue starts at Rs 19,00,000 growing 4% a month, gross margin",
        "   45%, operating expense Rs 5,20,000/month, budget interest Rs 25,000/month.",
        "   Edit the Budget formulas to match your own plan.",
        "5. Demo data is fictional (seeded, deterministic). Nothing here is financial",
        "   advice; it is a reporting template, not a recommendation. The model carries",
        "   no overdraft or short-term borrowings, so the current ratio runs high.",
    ]
    for i, line in enumerate(lines, 2):
        gd.cell(row=i, column=1, value=line)
    gd.column_dimensions["A"].width = 110
    return wb


def main():
    rows = build_data()
    with open(BASE / "data" / "inputs.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "revenue", "cogs", "salaries", "rent", "marketing",
                    "admin", "other_opex", "capex"])
        w.writerows(rows)
    wb = build_workbook(rows)
    wb.save(BASE / "stratica-mis-model.xlsx")
    print(f"Wrote {len(rows)} months, {BASE / 'stratica-mis-model.xlsx'}")


if __name__ == "__main__":
    main()
