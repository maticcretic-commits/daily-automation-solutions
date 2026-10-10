#!/usr/bin/env python3
"""AX-Style Finance Dashboards.

Turns Dynamics-AX-style GL exports (chart of accounts + GL transactions CSVs)
into a finance dashboard workbook with LIVE Excel formulas - no hard-coded
numbers anywhere in the statements:

  GL_Transactions  raw AX-style export (Entry, Date, Account, Debit, Credit, Net)
  Chart_of_Accounts  account master with groups + cash-flow sections
  Trial_Balance     debit / credit / balance per account, all SUMIFS
  P&L               monthly profit & loss, Apr-26 .. Sep-26, all formulas
  Cash_Flow         indirect-method cash flow for the period, all formulas
  Dashboard         KPIs + 3 native Excel charts (line, bar, column)

Usage:
    python generate.py [--data DIR] [--out FILE]

Sample data in data/ covers Apr-Sep 2026 for fictional 'Desert Rose Trading LLC'.
"""
import argparse
import csv
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

HERE = Path(__file__).resolve().parent
PERIOD_START = dt.date(2026, 4, 1)
N_MONTHS = 6
MONTHS = [dt.date(2026, 4 + i, 1) for i in range(N_MONTHS)]  # Apr..Sep 2026

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
MONEY_FMT = '#,##0'
PCT_FMT = '0.0%'
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------- data load
def load_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ---------------------------------------------------------------- workbook
def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def add_named(wb, name, ref):
    wb.defined_names.add(DefinedName(name, attr_text=ref))


def build(data_dir, out_path):
    coa = load_csv(data_dir / "chart_of_accounts.csv")
    gl = load_csv(data_dir / "gl_transactions.csv")
    acct_name = {r["account_code"]: r["account_name"] for r in coa}

    wb = Workbook()

    # ------------------------------------------------- GL_Transactions
    ws = wb.active
    ws.title = "GL_Transactions"
    headers = ["Entry", "Date", "Account Code", "Account Name",
               "Description", "Debit", "Credit", "Net (Dr-Cr)"]
    ws.append(headers)
    for r in gl:
        ws.append([r["entry_id"], dt.date.fromisoformat(r["date"]), r["account_code"],
                   acct_name[r["account_code"]], r["description"],
                   float(r["debit"]), float(r["credit"]),
                   f"=F{ws.max_row + 1}-G{ws.max_row + 1}"])
    style_header(ws, len(headers))
    n_gl = ws.max_row
    for col, w in zip("ABCDEFGH", [10, 12, 14, 24, 32, 14, 14, 14]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2, max_row=n_gl, min_col=6, max_col=8):
        for c in row:
            c.number_format = MONEY_FMT
    ws.freeze_panes = "A2"
    add_named(wb, "GL_Date", f"GL_Transactions!$B$2:$B${n_gl}")
    add_named(wb, "GL_Acct", f"GL_Transactions!$C$2:$C${n_gl}")
    add_named(wb, "GL_Net", f"GL_Transactions!$H$2:$H${n_gl}")
    add_named(wb, "GL_Debit", f"GL_Transactions!$F$2:$F${n_gl}")
    add_named(wb, "GL_Credit", f"GL_Transactions!$G$2:$G${n_gl}")

    # ------------------------------------------------- Chart_of_Accounts
    ws = wb.create_sheet("Chart_of_Accounts")
    headers = ["Account Code", "Account Name", "Group", "Normal Balance", "Cash-flow Section"]
    ws.append(headers)
    for r in coa:
        ws.append([r["account_code"], r["account_name"], r["account_group"],
                   r["normal_balance"], r["cashflow_section"]])
    style_header(ws, len(headers))
    for col, w in zip("ABCDE", [14, 30, 16, 15, 18]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    # ------------------------------------------------- Trial_Balance
    ws = wb.create_sheet("Trial_Balance")
    ws.append(["Account Code", "Account Name", "Group", "Debit", "Credit", "Balance (Dr-Cr)"])
    for i, r in enumerate(coa, start=2):
        code = r["account_code"]
        ws.cell(row=i, column=1, value=code)
        ws.cell(row=i, column=2, value=r["account_name"])
        ws.cell(row=i, column=3, value=r["account_group"])
        ws.cell(row=i, column=4, value=f'=SUMIFS(GL_Debit,GL_Acct,"{code}")')
        ws.cell(row=i, column=5, value=f'=SUMIFS(GL_Credit,GL_Acct,"{code}")')
        ws.cell(row=i, column=6, value=f"=D{i}-E{i}")
    t = len(coa) + 2
    ws.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=t, column=4, value=f"=SUM(D2:D{t-1})").font = Font(bold=True)
    ws.cell(row=t, column=5, value=f"=SUM(E2:E{t-1})").font = Font(bold=True)
    ws.cell(row=t + 1, column=1, value="Balanced?")
    ws.cell(row=t + 1, column=2,
            value=f'=IF(ABS(D{t}-E{t})<0.01,"OK - balanced","OUT OF BALANCE")')
    style_header(ws, 6)
    for col, w in zip("ABCDEF", [14, 30, 16, 16, 16, 18]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2, max_row=t, min_col=4, max_col=6):
        for c in row:
            c.number_format = MONEY_FMT
    ws.freeze_panes = "A2"
    tb_cash_row = next(i for i, r in enumerate(coa, start=2) if r["account_code"] == "1000")
    tb_cash_ref = f"'Trial_Balance'!F{tb_cash_row}"

    # ------------------------------------------------- P&L
    ws = wb.create_sheet("P&L")
    ws["A1"] = "Profit & Loss (SAR)"
    ws["A1"].font = TITLE_FONT
    headers = ["Line"] + [m.strftime("%b-%y") for m in MONTHS] + ["Total"]
    ws.append([])  # row 2 blank
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=j, value=h)
        if 1 < j <= 7:
            ws.cell(row=2, column=j, value=MONTHS[j - 2]).number_format = "mmm-yy"
    # month start dates live hidden in row 2 (B2:G2), headers in row 3 show text
    style_header(ws, len(headers), row=3)

    def msum(code, row, revenue=False):
        """Monthly SUMIFS formula for one account at P&L row `row`."""
        sign = "-" if revenue else ""
        for j in range(2, 8):  # B..G
            col = get_column_letter(j)
            ws.cell(row=row, column=j,
                    value=f"={sign}SUMIFS(GL_Net,GL_Acct,\"{code}\","
                          f"GL_Date,\">=\"&{col}$2,GL_Date,\"<\"&EDATE({col}$2,1))")
        ws.cell(row=row, column=8, value=f"=SUM(B{row}:G{row})")

    # (line, kind, spec) - spec line numbers are 1-based into this list
    lines = [
        ("Sales Revenue", "rev", "4000"),                              # 1
        ("Service Revenue", "rev", "4100"),                            # 2
        ("Total Revenue", "sum", (1, 2)),                              # 3
        ("Cost of Goods Sold", "exp", "5000"),                         # 4
        ("Gross Profit", "sub", (3, 4)),                               # 5 = 3-4
        ("Salaries & Wages", "exp", "6000"),                           # 6
        ("Office Rent", "exp", "6100"),                                 # 7
        ("Utilities", "exp", "6200"),                                  # 8
        ("Marketing", "exp", "6300"),                                   # 9
        ("Depreciation", "exp", "6400"),                               # 10
        ("Other Operating Expenses", "exp", "6999"),                   # 11
        ("Total Operating Expenses", "sum", (6, 7, 8, 9, 10, 11)),     # 12
        ("Operating Income", "sub", (5, 12)),                          # 13 = 5-12
        ("Other Income", "rev", "7000"),                               # 14
        ("Other Expenses", "exp", "7100"),                             # 15
        ("NET INCOME", "sub2", (13, 14, 15)),                          # 16 = 13+14-15
        ("Total Expenses (for charts)", "sum", (4, 12, 15)),           # 17
    ]
    r0 = 4
    for k, (label, kind, spec) in enumerate(lines):
        r = r0 + k
        ws.cell(row=r, column=1, value=label)
        if kind in ("rev", "exp"):
            msum(spec, r, revenue=(kind == "rev"))
        elif kind == "sum":
            for j in range(2, 9):
                col = get_column_letter(j)
                ws.cell(row=r, column=j,
                        value="=" + "+".join(f"{col}{r0 + s - 1}" for s in spec))
        elif kind == "sub":
            a, b = spec
            for j in range(2, 9):
                col = get_column_letter(j)
                ws.cell(row=r, column=j, value=f"={col}{r0 + a - 1}-{col}{r0 + b - 1}")
        elif kind == "sub2":
            a, b, c = spec
            for j in range(2, 9):
                col = get_column_letter(j)
                ws.cell(row=r, column=j,
                        value=f"={col}{r0 + a - 1}+{col}{r0 + b - 1}-{col}{r0 + c - 1}")
    last = r0 + len(lines) - 1
    for row in ws.iter_rows(min_row=r0, max_row=last, min_col=2, max_col=8):
        for c in row:
            c.number_format = MONEY_FMT
            c.border = BORDER
    for bold_row in (r0 + 2, r0 + 4, r0 + 11, r0 + 12, r0 + 15, r0 + 16):
        for c in ws[bold_row]:
            c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 30
    for j in range(2, 9):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.freeze_panes = "B4"
    ni_row = r0 + 15
    rev_row, exp_row = r0 + 2, r0 + 16
    gp_row = r0 + 4

    # ------------------------------------------------- Cash_Flow
    ws = wb.create_sheet("Cash_Flow")
    ws["A1"] = "Cash Flow Statement - indirect method (SAR), Apr-Sep 2026"
    ws["A1"].font = TITLE_FONT
    ws["K1"] = PERIOD_START
    ws["K2"] = f"=EDATE(K1,{N_MONTHS})"
    ws.append([])
    per = '"<"&$K$1'
    per_ge, per_lt = '">="&$K$1', '"<"&$K$2'

    def bal(code, only_period=False):
        base = f'SUMIFS(GL_Net,GL_Acct,"{code}"'
        if only_period:
            return base + f',GL_Date,{per_ge},GL_Date,{per_lt})'
        return base + ')'

    def bal_open(code):
        return f'SUMIFS(GL_Net,GL_Acct,"{code}",GL_Date,{per})'

    r0_cf = 3  # first cash-flow body row (row 2 holds K1/K2 helpers + col headers)
    rows_cf = [
        ("Operating activities", None, True),
        ("Net income", f"='P&L'!H{ni_row}", False),
        ("Add back: depreciation (non-cash)", f"='P&L'!H{r0 + 9}", False),
        ("Less: increase in trade receivables",
         f"=-(({bal('1100')})-({bal_open('1100')}))", False),
        ("Less: increase in inventory",
         f"=-(({bal('1200')})-({bal_open('1200')}))", False),
        ("Add: increase in trade payables",
         f"=-(({bal('2100')})-({bal_open('2100')}))", False),
        ("Net cash from operating activities", "sum", False),
        ("Investing activities", None, True),
        ("Equipment purchases", f"=-({bal('1500', True)})", False),
        ("Net cash from investing activities", "sum", False),
        ("Financing activities", None, True),
        ("Bank loan (drawdown less repayments)", f"=-({bal('2500', True)})", False),
        ("Net cash from financing activities", "sum", False),
        ("NET CHANGE IN CASH", "sum3", False),
        ("Cash at beginning of period", f"={bal_open('1000')}", False),
        ("CASH AT END OF PERIOD", "sum2", False),
        ("Check: ties to trial balance cash?",
         f'=IF(ABS(B{r0_cf + 15}-{tb_cash_ref})<0.01,"OK - ties to trial balance","DIFF")', False),
    ]
    ws.cell(row=r0_cf - 1, column=1, value="Line").font = Font(bold=True)
    ws.cell(row=r0_cf - 1, column=2, value="SAR").font = Font(bold=True)
    op_rows, inv_rows, fin_rows = [], [], []
    section = None
    for k, (label, formula, is_hdr) in enumerate(rows_cf):
        r = r0_cf + k
        ws.cell(row=r, column=1, value=label)
        if is_hdr:
            ws.cell(row=r, column=1).font = Font(bold=True, color="1F4E78")
            section = label.split()[0]
            continue
        if section == "Operating":
            op_rows.append(r)
        elif section == "Investing":
            inv_rows.append(r)
        elif section == "Financing":
            fin_rows.append(r)
        if formula == "sum":
            members = [m for m in {"Operating": op_rows, "Investing": inv_rows,
                                   "Financing": fin_rows}[section] if m != r]
            ws.cell(row=r, column=2, value="=" + "+".join(f"B{m}" for m in members))
        elif formula == "sum3":
            ws.cell(row=r, column=2, value=f"=B{r0_cf + 6}+B{r0_cf + 9}+B{r0_cf + 12}")
        elif formula == "sum2":
            ws.cell(row=r, column=2, value=f"=B{r0_cf + 13}+B{r0_cf + 14}")
        else:
            ws.cell(row=r, column=2, value=formula)
    for row in ws.iter_rows(min_row=r0_cf, max_row=r0_cf + len(rows_cf) - 1, min_col=2, max_col=2):
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.number_format = MONEY_FMT
    for bold_r in (r0_cf + 6, r0_cf + 9, r0_cf + 12, r0_cf + 13, r0_cf + 15):
        ws.cell(row=bold_r, column=1).font = Font(bold=True)
        ws.cell(row=bold_r, column=2).font = Font(bold=True)
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["K"].hidden = True
    cf_close_ref = f"'Cash_Flow'!B{r0_cf + 15}"
    cf_netchange_ref = f"'Cash_Flow'!B{r0_cf + 13}"

    # ------------------------------------------------- Dashboard
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Finance Dashboard - Desert Rose Trading LLC (SAR)"
    ws["A1"].font = Font(size=16, bold=True, color="1F4E78")
    ws["A2"] = "Apr-Sep 2026  |  Source: AX-style GL export  |  All figures are live formulas"
    ws["A2"].font = Font(italic=True, color="808080")
    kpis = [
        ("Total Revenue (6 mo)", f"='P&L'!H{rev_row}", MONEY_FMT),
        ("Total Expenses (6 mo)", f"='P&L'!H{exp_row}", MONEY_FMT),
        ("Net Income (6 mo)", f"='P&L'!H{ni_row}", MONEY_FMT),
        ("Gross Margin %", f"=IF('P&L'!H{rev_row}=0,0,'P&L'!H{gp_row}/'P&L'!H{rev_row})", PCT_FMT),
        ("Net Cash Change", f"={cf_netchange_ref}", MONEY_FMT),
        ("Closing Cash", f"={cf_close_ref}", MONEY_FMT),
    ]
    ws["A4"] = "Key figures"
    ws["A4"].font = Font(bold=True, size=12, color="1F4E78")
    for k, (label, formula, fmt) in enumerate(kpis):
        r = 5 + k
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = fmt
        c.font = Font(size=12, bold=True)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 20

    pls = wb["P&L"]
    # Chart 1: monthly revenue vs expenses (line)
    c1 = LineChart()
    c1.title = "Monthly Revenue vs Expenses"
    c1.style = 2
    c1.height, c1.width = 7.5, 15
    data = Reference(pls, min_col=2, max_col=7, min_row=rev_row, max_row=exp_row)
    cats = Reference(pls, min_col=2, max_col=7, min_row=3)
    c1.add_data(data, from_rows=True, titles_from_data=False)
    c1.set_categories(cats)
    c1.series[0].title = SeriesLabel(v="Revenue")
    c1.series[1].title = SeriesLabel(v="Expenses")
    ws.add_chart(c1, "D4")
    # Chart 2: opex by category (bar)
    c2 = BarChart()
    c2.title = "Operating Expenses by Category (6 mo total)"
    c2.style = 10
    c2.height, c2.width = 7.5, 15
    data = Reference(pls, min_col=8, min_row=r0 + 5, max_row=r0 + 10)
    cats = Reference(pls, min_col=1, min_row=r0 + 5, max_row=r0 + 10)
    c2.add_data(data, titles_from_data=False)
    c2.set_categories(cats)
    c2.shape = 4
    ws.add_chart(c2, "D20")
    # Chart 3: monthly net income (column)
    c3 = BarChart()
    c3.title = "Monthly Net Income"
    c3.style = 11
    c3.height, c3.width = 7.5, 15
    data = Reference(pls, min_col=2, max_col=7, min_row=ni_row)
    c3.add_data(data, from_rows=True, titles_from_data=False)
    c3.set_categories(Reference(pls, min_col=2, max_col=7, min_row=3))
    c3.series[0].title = SeriesLabel(v="Net Income")
    ws.add_chart(c3, "D36")

    wb.save(out_path)
    print(f"wrote {out_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=str(HERE / "data"))
    ap.add_argument("--out", default=str(HERE / "ax-finance-dashboards.xlsx"))
    a = ap.parse_args()
    build(Path(a.data), Path(a.out))


if __name__ == "__main__":
    main()
