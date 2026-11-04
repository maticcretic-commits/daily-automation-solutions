#!/usr/bin/env python3
"""Build monthly-close-pack.xlsx from data/journal.csv + accounts/budget.

Sheets: Guide, Dashboard, Journal, TrialBalance, IncomeStatement,
BalanceSheet, CashFlow, Monthly, Budget, Assumptions, Checks.
Rules honoured (learned the hard way): plain A1 ranges only, range
literals repeat the column letter on both ends ($D$2:$D$1693), no Table
objects, chart series titles come from header cells (titles_from_data),
dates are real date values, charts are added only after every source
sheet exists. No macros. The Journal sheet carries 150 pre-filled
Month Key formula rows of headroom, so new manual lines pasted into
columns A-E/G-H inside the headroom are picked up without editing any
formula; a blank entry ID yields a blank month key, never an error.
Every statement figure is a live formula off the Journal / TrialBalance
— nothing is typed as a value except source data and budgets.
"""
import csv
import datetime
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
OUT = os.path.join(HERE, "monthly-close-pack.xlsx")
HEADROOM = 150
AS_OF = datetime.date(2026, 11, 4)

NAVY = "1F3864"
MONEY = "#,##0"
PCT = "0.0%"
DATEFMT = "yyyy-mm-dd"

MONTHS = [(202604, "Apr 2026"), (202605, "May 2026"), (202606, "Jun 2026"),
          (202607, "Jul 2026"), (202608, "Aug 2026"), (202609, "Sep 2026"),
          (202610, "Oct 2026")]


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
    accounts = read_csv("accounts.csv")
    budget = read_csv("budget.csv")
    journal = read_csv("journal.csv")
    n = len(journal)
    last = 1 + n + HEADROOM
    nacct = len(accounts)
    tb_last = 1 + nacct

    JA = f"Journal!$A$2:$A${last}"
    JB = f"Journal!$B$2:$B${last}"
    JC = f"Journal!$C$2:$C${last}"
    JD = f"Journal!$D$2:$D${last}"
    JE = f"Journal!$E$2:$E${last}"
    JF = f"Journal!$F$2:$F${last}"
    JG = f"Journal!$G$2:$G${last}"
    TBA = f"TrialBalance!$A$2:$A${tb_last}"
    TBB = f"TrialBalance!$B$2:$B${tb_last}"
    TBH = f"TrialBalance!$H$2:$H${tb_last}"
    TBI = f"TrialBalance!$I$2:$I${tb_last}"

    wb = Workbook()

    # ---- Assumptions (created first so formulas can reference it) ----
    wa = wb.active
    wa.title = "Assumptions"
    wa["A1"] = "Setting"
    wa["B1"] = "Value"
    style_header(wa, 1, 2)
    wa["A2"] = "As-of date (month-end close date for this pack)"
    wa["B2"] = AS_OF
    wa["B2"].number_format = DATEFMT
    wa["A3"] = "Period covered"
    wa["B3"] = "Apr-Oct 2026 (7 months)"
    wa["A4"] = "Business (fictional demo data)"
    wa["B4"] = "Deshmukh Trading Co., Pune — FMCG wholesaler"
    wa["A5"] = "Headroom formula rows on Journal"
    wa["B5"] = HEADROOM
    wa["A6"] = "Currency"
    wa["B6"] = "INR (Rs)"
    wa["A8"] = ("Change B2 to re-date the pack. Do not type over the "
                "Month Key formula column F on Journal.")
    widths(wa, {"A": 46, "B": 44})

    # ---- Guide ----
    wg = wb.create_sheet("Guide")
    guide = [
        ("Monthly Close Pack — Deshmukh Trading Co., Pune", True),
        ("Built for the Monthly Bookkeeping & Reporting brief: a "
         "double-entry journal turned into a trial balance, profit & "
         "loss statement, balance sheet, indirect cash-flow statement, "
         "monthly trend and budget-vs-actual pack — every figure a live "
         "formula, with reconciliation checks. No macros.", False),
        ("How to add entries: paste or type new lines into columns "
         f"A-E and G-H of Journal, on any row up to row {last} — "
         "column F (Month Key) already carries formulas there, so the "
         "trial balance, statements and monthly trend all update. "
         "Every entry must balance: total debits = total credits for "
         "the entry ID, or Checks row 2 fails. Do not paste values "
         "over column F, and do not insert rows above row 2.", False),
        ("How to read it: TrialBalance proves the books balance. "
         "IncomeStatement shows revenue, cost of goods sold, gross "
         "profit and operating expenses. BalanceSheet carries "
         "retained earnings as opening retained earnings plus this "
         "period's net income, less drawings — and must balance to "
         "zero difference. CashFlow (indirect method) must tie "
         "exactly to the change in Cash. Checks must all read PASS.",
         False),
        ("Budgets live on the Budget sheet (typed values) — change "
         "them there and the variances and Dashboard follow. The "
         "Power BI mapping in powerbi/MEASURES.md rebuilds the same "
         "model as a .pbix.", False),
        ("Data in this demo is fictional (seeded, deterministic) — "
         "replace data/journal.csv with the real journal and rebuild, "
         "or paste real lines straight into Journal.", False),
    ]
    for i, (text, bold) in enumerate(guide, start=1):
        c = wg.cell(row=i, column=1, value=text)
        c.font = Font(bold=bold, size=12 if i == 1 else 11)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    wg.column_dimensions["A"].width = 130

    # ---- Journal ----
    wj = wb.create_sheet("Journal")
    headers = ["Entry ID", "Date", "Account", "Debit (Rs)", "Credit (Rs)",
               "Month Key", "Category", "Narration"]
    for col, h in enumerate(headers, start=1):
        wj.cell(row=1, column=col, value=h)
    style_header(wj, 1, len(headers))
    for r, row in enumerate(journal, start=2):
        wj.cell(row=r, column=1, value=row["entry_id"])
        wj.cell(row=r, column=2, value=dval(row["date"])).number_format = DATEFMT
        wj.cell(row=r, column=3, value=row["account"])
        wj.cell(row=r, column=4, value=int(row["debit"])).number_format = MONEY
        wj.cell(row=r, column=5, value=int(row["credit"])).number_format = MONEY
        wj.cell(row=r, column=7, value=row["category"])
        wj.cell(row=r, column=8, value=row["narration"])
    for r in range(2, last + 1):
        wj.cell(row=r, column=6,
                value=f'=IF($A{r}="","",YEAR($B{r})*100+MONTH($B{r}))')
    wj.freeze_panes = "A2"
    wj.auto_filter.ref = f"A1:H{last}"
    widths(wj, {"A": 10, "B": 12, "C": 26, "D": 13, "E": 13, "F": 11,
                "G": 14, "H": 42})

    # ---- TrialBalance ----
    wt = wb.create_sheet("TrialBalance")
    theaders = ["Account", "Type", "Normal", "Opening Dr (Rs)",
                "Opening Cr (Rs)", "Dr Movements (Rs)", "Cr Movements (Rs)",
                "Closing Dr (Rs)", "Closing Cr (Rs)", "Net Movement (Rs)"]
    for col, h in enumerate(theaders, start=1):
        wt.cell(row=1, column=col, value=h)
    style_header(wt, 1, len(theaders))
    for r, a in enumerate(accounts, start=2):
        wt.cell(row=r, column=1, value=a["account"])
        wt.cell(row=r, column=2, value=a["type"])
        wt.cell(row=r, column=3, value=a["normal"])
        wt.cell(row=r, column=4, value=int(a["opening_debit"])).number_format = MONEY
        wt.cell(row=r, column=5, value=int(a["opening_credit"])).number_format = MONEY
        wt.cell(row=r, column=6,
                value=f"=SUMIF({JC},$A{r},{JD})").number_format = MONEY
        wt.cell(row=r, column=7,
                value=f"=SUMIF({JC},$A{r},{JE})").number_format = MONEY
        wt.cell(row=r, column=8,
                value=f"=MAX($D{r}-$E{r}+$F{r}-$G{r},0)").number_format = MONEY
        wt.cell(row=r, column=9,
                value=f"=MAX(-($D{r}-$E{r}+$F{r}-$G{r}),0)").number_format = MONEY
        wt.cell(row=r, column=10,
                value=f"=($H{r}-$I{r})-($D{r}-$E{r})").number_format = MONEY
    tr = tb_last + 1
    wt.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col in range(4, 11):
        letter = chr(64 + col)
        c = wt.cell(row=tr, column=col,
                    value=f"=SUM({letter}2:{letter}{tb_last})")
        c.font = Font(bold=True)
        c.number_format = MONEY
    wt.freeze_panes = "A2"
    widths(wt, {"A": 26, "B": 11, "C": 8, "D": 15, "E": 15, "F": 16,
                "G": 16, "H": 15, "I": 15, "J": 16})

    # ---- IncomeStatement ----
    wi = wb.create_sheet("IncomeStatement")
    wi["A1"] = "Profit & Loss — Apr-Oct 2026 (Rs)"
    wi["A1"].font = Font(bold=True, size=12)
    wi["A2"] = "Sales Revenue"
    wi["A3"] = "Other Income"
    wi["A4"] = "TOTAL REVENUE"
    wi["A5"] = "Cost of Goods Sold"
    wi["A6"] = "GROSS PROFIT"
    wi["A7"] = "Rent Expense"
    wi["A8"] = "Salaries Expense"
    wi["A9"] = "Utilities Expense"
    wi["A10"] = "Depreciation Expense"
    wi["A11"] = "Marketing Expense"
    wi["A12"] = "Bank Charges"
    wi["A13"] = "TOTAL OPERATING EXPENSES"
    wi["A14"] = "NET INCOME"
    rev_rows = [(2, "Sales Revenue", TBI), (3, "Other Income", TBI)]
    for row, _name, colrange in rev_rows:
        wi.cell(row=row, column=2,
                value=f"=SUMIFS({colrange},{TBA},$A{row})").number_format = MONEY
    wi.cell(row=4, column=2, value="=SUM($B$2:$B$3)").number_format = MONEY
    wi.cell(row=5, column=2,
            value=f"=SUMIFS({TBH},{TBA},$A5)").number_format = MONEY
    wi.cell(row=6, column=2, value="=$B$4-$B$5").number_format = MONEY
    for row in range(7, 13):
        wi.cell(row=row, column=2,
                value=f"=SUMIFS({TBH},{TBA},$A{row})").number_format = MONEY
    wi.cell(row=13, column=2, value="=SUM($B$7:$B$12)").number_format = MONEY
    wi.cell(row=14, column=2, value="=$B$6-$B$13").number_format = MONEY
    for row in (4, 6, 13, 14):
        wi.cell(row=row, column=1).font = Font(bold=True)
        wi.cell(row=row, column=2).font = Font(bold=True)
    widths(wi, {"A": 28, "B": 18})

    def tb_cell(account_name, col_letter):
        """Direct TrialBalance cell for an account (row order = CSV order)."""
        idx = next(i for i, a in enumerate(accounts)
                   if a["account"] == account_name)
        return f"TrialBalance!${col_letter}${2 + idx}"

    # ---- BalanceSheet ----
    wb_ = wb.create_sheet("BalanceSheet")
    wb_["A1"] = "Balance Sheet — as at 4 Nov 2026 (Rs)"
    wb_["A1"].font = Font(bold=True, size=12)
    cap_cell = tb_cell("Owner's Capital", "I")
    bs_rows = [
        (2, "Cash", f"={tb_cell('Cash', 'H')}-{tb_cell('Cash', 'I')}"),
        (3, "Accounts Receivable",
         f"={tb_cell('Accounts Receivable', 'H')}-{tb_cell('Accounts Receivable', 'I')}"),
        (4, "Inventory",
         f"={tb_cell('Inventory', 'H')}-{tb_cell('Inventory', 'I')}"),
        (5, "Prepaid Expenses",
         f"={tb_cell('Prepaid Expenses', 'H')}-{tb_cell('Prepaid Expenses', 'I')}"),
        (6, "Equipment (gross)", f"={tb_cell('Equipment', 'H')}"),
        (7, "Less: Accumulated Depreciation",
         f"=-{tb_cell('Accumulated Depreciation', 'I')}"),
        (8, "TOTAL ASSETS", "=SUM($B$2:$B$7)"),
        (10, "Accounts Payable", f"={tb_cell('Accounts Payable', 'I')}"),
        (11, "GST Payable", f"={tb_cell('GST Payable', 'I')}"),
        (12, "Short-term Loan", f"={tb_cell('Short-term Loan', 'I')}"),
        (13, "TOTAL LIABILITIES", "=SUM($B$10:$B$12)"),
        (15, "Owner's Capital", f"={cap_cell}"),
        (16, "Retained Earnings (opening)",
         f"={tb_cell('Retained Earnings', 'I')}"),
        (17, "Add: Net income for the period", "=IncomeStatement!$B$14"),
        (18, "Retained Earnings (closing)", "=$B$16+$B$17"),
        (19, "Less: Drawings", f"=-{tb_cell('Drawings', 'H')}"),
        (20, "TOTAL EQUITY", "=$B$15+$B$18+$B$19"),
        (21, "TOTAL LIABILITIES & EQUITY", "=$B$13+$B$20"),
        (22, "BALANCE CHECK (Assets - L&E, must be 0)", "=$B$8-$B$21"),
    ]
    wb_["A9"] = "LIABILITIES"
    wb_["A9"].font = Font(bold=True)
    wb_["A14"] = "EQUITY"
    wb_["A14"].font = Font(bold=True)
    for row, label, formula in bs_rows:
        wb_.cell(row=row, column=1, value=label)
        wb_.cell(row=row, column=2, value=formula).number_format = MONEY
    for row in (8, 13, 20, 21, 22):
        wb_.cell(row=row, column=1).font = Font(bold=True)
        wb_.cell(row=row, column=2).font = Font(bold=True)
    widths(wb_, {"A": 40, "B": 18})

    def signed_delta(account_name):
        """Closing signed balance minus opening signed balance, TB row."""
        idx = next(i for i, a in enumerate(accounts)
                   if a["account"] == account_name)
        r = 2 + idx
        return (f"((TrialBalance!$H${r}-TrialBalance!$I${r})"
                f"-(TrialBalance!$D${r}-TrialBalance!$E${r}))")

    # ---- CashFlow (indirect) ----
    wc = wb.create_sheet("CashFlow")
    wc["A1"] = "Cash Flow (indirect) — Apr-Oct 2026 (Rs)"
    wc["A1"].font = Font(bold=True, size=12)
    cap_delta = signed_delta("Owner's Capital")
    cf_rows = [
        (2, "Net income", "=IncomeStatement!$B$14"),
        (3, "Add: Depreciation (non-cash)",
         f"={tb_cell('Accumulated Depreciation', 'G')}"),
        (4, "Less: Increase in Accounts Receivable",
         f"=-{signed_delta('Accounts Receivable')}"),
        (5, "Less: Increase in Inventory",
         f"=-{signed_delta('Inventory')}"),
        (6, "Less: Increase in Prepaid Expenses",
         f"=-{signed_delta('Prepaid Expenses')}"),
        (7, "Add: Increase in Accounts Payable",
         f"=-{signed_delta('Accounts Payable')}"),
        (8, "Add: Increase in GST Payable",
         f"=-{signed_delta('GST Payable')}"),
        (9, "NET CASH FROM OPERATIONS", "=SUM($B$2:$B$8)"),
        (10, "Less: Equipment purchases",
         f"=-{signed_delta('Equipment')}"),
        (11, "NET CASH FROM INVESTING", "=$B$10"),
        (12, "Add: Loan movement (drawdown less repayments)",
         f"=-{signed_delta('Short-term Loan')}"),
        (13, "Add: Capital injection",
         f"=-{cap_delta}"),
        (14, "Less: Drawings", f"=-{tb_cell('Drawings', 'H')}"),
        (15, "NET CASH FROM FINANCING", "=SUM($B$12:$B$14)"),
        (16, "NET CHANGE IN CASH", "=$B$9+$B$11+$B$15"),
        (17, "Cash at 1 Apr 2026", f"={tb_cell('Cash', 'D')}"),
        (18, "Cash at 4 Nov 2026",
         f"={tb_cell('Cash', 'H')}-{tb_cell('Cash', 'I')}"),
        (19, "Change in Cash per Trial Balance", "=$B$18-$B$17"),
        (20, "CHECK: statement change less TB change (must be 0)",
         "=$B$16-$B$19"),
    ]
    for row, label, formula in cf_rows:
        wc.cell(row=row, column=1, value=label)
        wc.cell(row=row, column=2, value=formula).number_format = MONEY
    for row in (9, 11, 15, 16, 20):
        wc.cell(row=row, column=1).font = Font(bold=True)
        wc.cell(row=row, column=2).font = Font(bold=True)
    widths(wc, {"A": 52, "B": 18})

    # ---- Monthly ----
    wm = wb.create_sheet("Monthly")
    mheaders = ["Month Key", "Month", "Revenue (Rs)", "COGS (Rs)",
                "Gross Profit (Rs)", "Operating Expenses (Rs)",
                "Net Income (Rs)", "Journal Lines", "Cumulative Net (Rs)"]
    for col, h in enumerate(mheaders, start=1):
        wm.cell(row=1, column=col, value=h)
    style_header(wm, 1, len(mheaders))
    for i, (key, label) in enumerate(MONTHS, start=2):
        wm.cell(row=i, column=1, value=key)
        wm.cell(row=i, column=2, value=label)
        wm.cell(row=i, column=3,
                value=f'=SUMIFS({JE},{JF},$A{i},{JG},"Revenue")'
                ).number_format = MONEY
        wm.cell(row=i, column=4,
                value=f'=SUMIFS({JD},{JF},$A{i},{JG},"COGS")'
                ).number_format = MONEY
        wm.cell(row=i, column=5, value=f"=$C{i}-$D{i}").number_format = MONEY
        wm.cell(row=i, column=6,
                value=f'=SUMIFS({JD},{JF},$A{i},{JG},"Opex")'
                ).number_format = MONEY
        wm.cell(row=i, column=7, value=f"=$E{i}-$F{i}").number_format = MONEY
        wm.cell(row=i, column=8, value=f"=COUNTIF({JF},$A{i})")
        prev = f"$I{i - 1}" if i > 2 else "0"
        wm.cell(row=i, column=9,
                value=f"={prev}+$G{i}").number_format = MONEY
    mtot = 2 + len(MONTHS)
    wm.cell(row=mtot, column=1, value="TOTAL").font = Font(bold=True)
    wm.cell(row=mtot, column=2, value="Apr-Oct 2026").font = Font(bold=True)
    for col, letter in ((3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"),
                        (8, "H")):
        c = wm.cell(row=mtot, column=col,
                    value=f"=SUM({letter}2:{letter}{mtot - 1})")
        c.font = Font(bold=True)
        if col != 8:
            c.number_format = MONEY
    wm.cell(row=mtot, column=9, value=f"=$I{mtot - 1}").number_format = MONEY
    wm.cell(row=mtot, column=9).font = Font(bold=True)
    widths(wm, {"A": 11, "B": 11, "C": 15, "D": 13, "E": 17, "F": 20,
                "G": 15, "H": 13, "I": 18})

    # ---- Budget ----
    wbu = wb.create_sheet("Budget")
    bheaders = ["Month Key", "Month", "Sales Budget (Rs)",
                "Sales Actual (Rs)", "Sales Variance (Rs)",
                "Sales Variance %", "Net Budget (Rs)", "Net Actual (Rs)",
                "Net Variance (Rs)"]
    for col, h in enumerate(bheaders, start=1):
        wbu.cell(row=1, column=col, value=h)
    style_header(wbu, 1, len(bheaders))
    for i, b in enumerate(budget, start=2):
        wbu.cell(row=i, column=1, value=int(b["month_key"]))
        wbu.cell(row=i, column=2, value=b["month"])
        wbu.cell(row=i, column=3, value=int(b["sales_budget"])).number_format = MONEY
        wbu.cell(row=i, column=4, value=f"=Monthly!$C{i}").number_format = MONEY
        wbu.cell(row=i, column=5, value=f"=$D{i}-$C{i}").number_format = MONEY
        wbu.cell(row=i, column=6,
                 value=f"=IF($C{i}=0,0,$E{i}/$C{i})").number_format = PCT
        wbu.cell(row=i, column=7, value=int(b["net_budget"])).number_format = MONEY
        wbu.cell(row=i, column=8, value=f"=Monthly!$G{i}").number_format = MONEY
        wbu.cell(row=i, column=9, value=f"=$H{i}-$G{i}").number_format = MONEY
    btot = 2 + len(budget)
    wbu.cell(row=btot, column=1, value="TOTAL").font = Font(bold=True)
    wbu.cell(row=btot, column=2, value="Apr-Oct 2026").font = Font(bold=True)
    for col, letter in ((3, "C"), (4, "D"), (5, "E"), (7, "G"), (8, "H"),
                        (9, "I")):
        c = wbu.cell(row=btot, column=col,
                     value=f"=SUM({letter}2:{letter}{btot - 1})")
        c.font = Font(bold=True)
        c.number_format = MONEY
    wbu.cell(row=btot, column=6,
             value=f"=IF($C{btot}=0,0,$E{btot}/$C{btot})").number_format = PCT
    wbu.cell(row=btot, column=6).font = Font(bold=True)
    widths(wbu, {"A": 11, "B": 11, "C": 17, "D": 16, "E": 18, "F": 15,
                 "G": 15, "H": 14, "I": 16})

    # ---- Dashboard (created after source sheets; charts added last) ----
    wd = wb.create_sheet("Dashboard")
    wd["A1"] = "Monthly Close Dashboard — Deshmukh Trading Co., Pune"
    wd["A1"].font = Font(bold=True, size=13)
    wd["A2"] = "As of:"
    wd["B2"] = "=Assumptions!$B$2"
    wd["B2"].number_format = DATEFMT
    wd["A3"] = "KPI"
    wd["B3"] = "Value"
    style_header(wd, 3, 2)
    kpis = [
        ("Total revenue (Rs)", "=IncomeStatement!$B$4", MONEY),
        ("Cost of goods sold (Rs)", "=IncomeStatement!$B$5", MONEY),
        ("Gross profit (Rs)", "=IncomeStatement!$B$6", MONEY),
        ("Operating expenses (Rs)", "=IncomeStatement!$B$13", MONEY),
        ("Net income (Rs)", "=IncomeStatement!$B$14", MONEY),
        ("Net margin", "=IF($B$4=0,0,$B$8/$B$4)", PCT),
        ("Total assets (Rs)", "=BalanceSheet!$B$8", MONEY),
        ("Total liabilities (Rs)", "=BalanceSheet!$B$13", MONEY),
        ("Total equity (Rs)", "=BalanceSheet!$B$20", MONEY),
        ("Cash at period end (Rs)", "=CashFlow!$B$18", MONEY),
        ("Net cash from operations (Rs)", "=CashFlow!$B$9", MONEY),
        ("Journal lines", f"=COUNTA({JA})", "0"),
        ("Accounts in trial balance",
         f"=COUNTA(TrialBalance!$A$2:$A${tb_last})", "0"),
    ]
    for i, (label, formula, fmt) in enumerate(kpis, start=4):
        wd.cell(row=i, column=1, value=label)
        c = wd.cell(row=i, column=2, value=formula)
        c.number_format = fmt
    # Expense-mix table for the pie chart (rows 18-25).
    wd["A18"] = "Expense"
    wd["B18"] = "Amount (Rs)"
    style_header(wd, 18, 2)
    for i, is_row in enumerate(range(7, 13), start=19):
        wd.cell(row=i, column=1, value=f"=IncomeStatement!$A{is_row}")
        wd.cell(row=i, column=2,
                value=f"=IncomeStatement!$B{is_row}").number_format = MONEY
    wd["A25"] = "TOTAL"
    wd["A25"].font = Font(bold=True)
    wd["B25"] = "=SUM($B$19:$B$24)"
    wd["B25"].font = Font(bold=True)
    wd["B25"].number_format = MONEY
    widths(wd, {"A": 34, "B": 18, "C": 4, "D": 18, "E": 18})

    # ---- Checks ----
    wk = wb.create_sheet("Checks")
    for col, h in enumerate(["Check", "Difference", "Expected", "Result"],
                            start=1):
        wk.cell(row=1, column=col, value=h)
    style_header(wk, 1, 4)
    checks = [
        ("Journal debits = credits (all lines)",
         f"=SUM({JD})-SUM({JE})"),
        ("Trial balance debits = credits",
         f"=TrialBalance!$H${tr}-TrialBalance!$I${tr}"),
        ("Opening trial balance balances",
         f"=SUM(TrialBalance!$D$2:$D${tb_last})"
         f"-SUM(TrialBalance!$E$2:$E${tb_last})"),
        ("P&L net = trial-balance revenue/COGS/expenses",
         "=IncomeStatement!$B$14-(SUMIFS(TrialBalance!$I$2:$I$22,"
         "TrialBalance!$B$2:$B$22,\"Revenue\")"
         "-SUMIFS(TrialBalance!$H$2:$H$22,"
         "TrialBalance!$B$2:$B$22,\"COGS\")"
         "-SUMIFS(TrialBalance!$H$2:$H$22,"
         "TrialBalance!$B$2:$B$22,\"Expense\"))"),
        ("Balance sheet balances (Assets - L&E)",
         "=BalanceSheet!$B$22"),
        ("Cash flow ties to change in Cash",
         "=CashFlow!$B$20"),
        ("Monthly revenue total = P&L revenue",
         f"=Monthly!$C${mtot}-IncomeStatement!$B$4"),
        ("Monthly net total = P&L net",
         f"=Monthly!$G${mtot}-IncomeStatement!$B$14"),
        ("Budget sales actual total = Monthly revenue",
         f"=Budget!$D${btot}-Monthly!$C${mtot}"),
        ("Closing cash is not negative (1 = negative)",
         "=IF(CashFlow!$B$18<0,1,0)"),
        ("Trial-balance net movements sum to zero",
         f"=TrialBalance!$J${tr}"),
        ("Dashboard net income = P&L net income",
         "=Dashboard!$B$8-IncomeStatement!$B$14"),
    ]
    for i, (label, formula) in enumerate(checks, start=2):
        wk.cell(row=i, column=1, value=label)
        c = wk.cell(row=i, column=2, value=formula)
        c.number_format = MONEY
        wk.cell(row=i, column=3, value=0)
        wk.cell(row=i, column=4,
                value=f'=IF(ABS($B{i})<0.005,"PASS","FAIL")')
    red = PatternFill("solid", fgColor="FFC7CE")
    wk.conditional_formatting.add(
        f"D2:D{1 + len(checks)}",
        CellIsRule(operator="equal", formula=['"FAIL"'], fill=red))
    widths(wk, {"A": 56, "B": 14, "C": 10, "D": 10})

    # ---- Charts (only now that every source sheet exists) ----
    bar = BarChart()
    bar.type = "col"
    bar.title = "Revenue vs COGS by month"
    bar.y_axis.title = "Rs"
    bar_data = Reference(wm, min_col=3, min_row=1, max_row=8, max_col=4)
    bar_cats = Reference(wm, min_col=2, min_row=2, max_row=8)
    bar.add_data(bar_data, titles_from_data=True)
    bar.set_categories(bar_cats)
    bar.height = 7.5
    bar.width = 14
    wd.add_chart(bar, "D3")

    pie = PieChart()
    pie.title = "Operating expenses by type"
    pie_data = Reference(wd, min_col=2, min_row=18, max_row=24)
    pie_cats = Reference(wd, min_col=1, min_row=19, max_row=24)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)
    pie.height = 7.5
    pie.width = 12
    wd.add_chart(pie, "D18")

    line = LineChart()
    line.title = "Cumulative net income"
    line.y_axis.title = "Rs"
    line_data = Reference(wm, min_col=9, min_row=1, max_row=8)
    line.add_data(line_data, titles_from_data=True)
    line.set_categories(bar_cats)
    line.height = 7.5
    line.width = 14
    wd.add_chart(line, "D33")

    # Sensible tab order and active sheet.
    wb._sheets = [wb["Guide"], wb["Dashboard"], wb["Journal"],
                  wb["TrialBalance"], wb["IncomeStatement"],
                  wb["BalanceSheet"], wb["CashFlow"], wb["Monthly"],
                  wb["Budget"], wb["Assumptions"], wb["Checks"]]
    wb.active = wb["Dashboard"]
    wb.save(OUT)
    print(f"wrote {OUT} ({n} journal lines, {nacct} accounts, "
          f"Journal formulas to row {last})")


if __name__ == "__main__":
    main()
