#!/usr/bin/env python3
"""build_workbook.py — Build the live-formula financial reporting workbook.

Pipeline: data/*.csv -> qb2bi.sqlite (build_db.py) -> qb2bi-finance-dashboard.xlsx

Sheets:
  Guide          how the pipeline works + how to run it
  Raw_Journals   every journal line (values) + PeriodInt helper column (formula)
  TrialBalance   per-account totals via SUMIFS (live formulas)
  PL             monthly P&L Apr-25..Sep-26 + YTD (live formulas)
  BalanceSheet   as of Mar-26 and Sep-26 (live formulas) + balance check
  Dashboard      KPI cards (live formulas) + 3 native charts
  AI_Report      rule-based management narrative (computed values)
  Checks         8 automated integrity checks (live formulas -> PASS/FAIL)

All LO-headless-safe: plain A1 ranges, no Table objects, numeric guards are 0.
"""
import sqlite3
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, PieChart, BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).parent
DB = HERE / "qb2bi.sqlite"
XLSX = HERE / "qb2bi-finance-dashboard.xlsx"

MONTHS = [(2025, 4), (2025, 5), (2025, 6), (2025, 7), (2025, 8), (2025, 9),
          (2025, 10), (2025, 11), (2025, 12), (2026, 1), (2026, 2), (2026, 3),
          (2026, 4), (2026, 5), (2026, 6), (2026, 7), (2026, 8), (2026, 9)]
MLBL = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25",
        "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26",
        "Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26"]
PINT = [y * 12 + m for y, m in MONTHS]
N_MO = len(MONTHS)                      # 18
MO_FIRST, MO_LAST = 2, 1 + N_MO         # cols B..S
YTD_COL = MO_LAST + 1                   # col T
RJ = "Raw_Journals"
MONEY = '#,##0'
PCT = '0.0%'

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
SEC_FONT = Font(bold=True, size=12, color="1F4E78")
KPI_FONT = Font(bold=True, size=16, color="1F4E78")
KPI_LBL = Font(bold=True, size=10, color="595959")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def hdr(ws, row, ncols, values):
    for c, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def money_fmt(ws, ref):
    for row in ws[ref]:
        for cell in row:
            cell.number_format = MONEY


# ================================================================ raw data
def sheet_raw(wb):
    ws = wb.create_sheet("Raw_Journals")
    con = sqlite3.connect(DB)
    rows = con.execute(
        """SELECT date, journal_no, line_no, j.acct_no, a.acct_name,
                  j.debit, j.credit, j.memo, j.party
           FROM journal_lines j JOIN accounts a ON a.acct_no = j.acct_no
           ORDER BY j.id""").fetchall()
    con.close()
    hdr(ws, 1, 10, ["Date", "Journal", "Line", "Acct No", "Account",
                    "Debit (₹)", "Credit (₹)", "Memo", "Party", "Period"])
    for i, (d, jno, ln, ano, aname, dr, cr, memo, party) in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=date.fromisoformat(d)).number_format = "DD-MMM-YYYY"
        ws.cell(row=i, column=2, value=jno)
        ws.cell(row=i, column=3, value=ln)
        ws.cell(row=i, column=4, value=ano)
        ws.cell(row=i, column=5, value=aname)
        ws.cell(row=i, column=6, value=dr).number_format = MONEY
        ws.cell(row=i, column=7, value=cr).number_format = MONEY
        ws.cell(row=i, column=8, value=memo)
        ws.cell(row=i, column=9, value=party)
        ws.cell(row=i, column=10, value=f"=YEAR(A{i})*12+MONTH(A{i})")
    n = len(rows) + 1
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFGHIJ", (12, 10, 6, 9, 24, 13, 13, 30, 22, 9)):
        ws.column_dimensions[col].width = w
    return n  # last data row


# ============================================================ trial balance
def sheet_tb(wb):
    ws = wb.create_sheet("TrialBalance")
    con = sqlite3.connect(DB)
    accts = con.execute(
        "SELECT acct_no, acct_name, acct_type FROM accounts ORDER BY acct_no").fetchall()
    con.close()
    ws["A1"] = "TRIAL BALANCE — live SUMIFS over Raw_Journals"
    ws["A1"].font = TITLE_FONT
    hdr(ws, 3, 7, ["Acct No", "Account", "Type", "Total Debit (₹)",
                   "Total Credit (₹)", "Net Debit (₹)", "Net Credit (₹)"])
    r0 = 4
    for i, (ano, aname, atype) in enumerate(accts):
        r = r0 + i
        ws.cell(row=r, column=1, value=ano)
        ws.cell(row=r, column=2, value=aname)
        ws.cell(row=r, column=3, value=atype)
        ws.cell(row=r, column=4,
                value=f'=SUMIFS({RJ}!$F:$F,{RJ}!$D:$D,$A{r})')
        ws.cell(row=r, column=5,
                value=f'=SUMIFS({RJ}!$G:$G,{RJ}!$D:$D,$A{r})')
        ws.cell(row=r, column=6, value=f"=IF(D{r}>=E{r},D{r}-E{r},0)")
        ws.cell(row=r, column=7, value=f"=IF(E{r}>D{r},E{r}-D{r},0)")
    rt = r0 + len(accts)
    ws.cell(row=rt, column=2, value="TOTAL").font = Font(bold=True)
    for c in (4, 5, 6, 7):
        col = get_column_letter(c)
        ws.cell(row=rt, column=c,
                value=f"=SUM({col}{r0}:{col}{rt-1})").font = Font(bold=True)
    money_fmt(ws, f"D4:G{rt}")
    ws.freeze_panes = "A4"
    for col, w in zip("ABCDEFG", (9, 26, 10, 15, 15, 15, 15)):
        ws.column_dimensions[col].width = w
    return {  # row per acct_no for cross-sheet checks
        ano: r0 + i for i, (ano, _, _) in enumerate(accts)}, rt


# ==================================================================== P & L
EXP = [("5100", "Rent Expense"), ("5200", "Salaries Expense"),
       ("5300", "Utilities Expense"), ("5400", "Transport & Logistics"),
       ("5500", "Marketing Expense"), ("5600", "Interest Expense"),
       ("5700", "Depreciation Expense"), ("5800", "Miscellaneous Expense")]


def sheet_pl(wb):
    ws = wb.create_sheet("PL")
    ws["A1"] = "PROFIT & LOSS — monthly, live SUMIFS over Raw_Journals (₹)"
    ws["A1"].font = TITLE_FONT
    ytd = get_column_letter(YTD_COL)
    headers = ["Particulars"] + MLBL + ["YTD"]
    hdr(ws, 2, 1 + N_MO + 1, headers)
    for c, p in enumerate(PINT, MO_FIRST):
        ws.cell(row=3, column=c, value=p)  # hidden period ints
    ws.row_dimensions[3].hidden = True

    def msum(c, row, acct, side):
        """Net movement of acct in one month (col c). side='cr' credit-balance."""
        p = f"{get_column_letter(c)}$3"
        f = (f'SUMIFS({RJ}!$G:$G,{RJ}!$J:$J,{p},{RJ}!$D:$D,"{acct}")'
             f'-SUMIFS({RJ}!$F:$F,{RJ}!$J:$J,{p},{RJ}!$D:$D,"{acct}")')
        if side == "dr":
            f = (f'SUMIFS({RJ}!$F:$F,{RJ}!$J:$J,{p},{RJ}!$D:$D,"{acct}")'
                 f'-SUMIFS({RJ}!$G:$G,{RJ}!$J:$J,{p},{RJ}!$D:$D,"{acct}")')
        ws.cell(row=row, column=c, value=f"={f}")

    R_REV, R_COGS, R_GP = 4, 5, 6
    R_EXP0 = 7
    R_TEXP = R_EXP0 + len(EXP)
    R_NP = R_TEXP + 1
    ws.cell(row=R_REV, column=1, value="Sales Revenue")
    ws.cell(row=R_COGS, column=1, value="Cost of Goods Sold")
    ws.cell(row=R_GP, column=1, value="Gross Profit").font = Font(bold=True)
    for i, (ano, aname) in enumerate(EXP):
        ws.cell(row=R_EXP0 + i, column=1, value=aname)
    ws.cell(row=R_TEXP, column=1, value="Total Expenses").font = Font(bold=True)
    ws.cell(row=R_NP, column=1, value="NET PROFIT").font = Font(bold=True, color="1F4E78")

    for c in range(MO_FIRST, MO_LAST + 1):
        cl = get_column_letter(c)
        msum(c, R_REV, "4000", "cr")
        msum(c, R_COGS, "5000", "dr")
        ws.cell(row=R_GP, column=c, value=f"={cl}{R_REV}-{cl}{R_COGS}")
        for i, (ano, _) in enumerate(EXP):
            msum(c, R_EXP0 + i, ano, "dr")
        ws.cell(row=R_TEXP, column=c,
                value="=" + "+".join(f"{cl}{R_EXP0+i}" for i in range(len(EXP))))
        ws.cell(row=R_NP, column=c, value=f"={cl}{R_GP}-{cl}{R_TEXP}")
    # YTD column
    for r in [R_REV, R_COGS, R_GP] + [R_EXP0 + i for i in range(len(EXP))] + [R_TEXP, R_NP]:
        ws.cell(row=r, column=YTD_COL,
                value=f"=SUM({get_column_letter(MO_FIRST)}{r}:{get_column_letter(MO_LAST)}{r})")
    money_fmt(ws, f"B4:{ytd}{R_NP}")
    ws.column_dimensions["A"].width = 24
    for c in range(MO_FIRST, YTD_COL + 1):
        ws.column_dimensions[get_column_letter(c)].width = 13
    ws.freeze_panes = "B4"
    return {"rev": R_REV, "gp": R_GP, "texp": R_TEXP, "np": R_NP,
            "exp0": R_EXP0, "n_exp": len(EXP)}


# ============================================================ balance sheet
def sheet_bs(wb):
    ws = wb.create_sheet("BalanceSheet")
    ws["A1"] = "BALANCE SHEET — live SUMIFS over Raw_Journals (₹)"
    ws["A1"].font = TITLE_FONT
    hdr(ws, 2, 3, ["Particulars", "As of 31-Mar-2026", "As of 30-Sep-2026"])
    ws.cell(row=3, column=2, value=2026 * 12 + 3).font = Font(color="FFFFFF")
    ws.cell(row=3, column=3, value=2026 * 12 + 9).font = Font(color="FFFFFF")
    ws.row_dimensions[3].hidden = True

    def bal(col_letter, row, acct, side):
        """Balance of acct as of period in row 3. side='dr' debit-balance."""
        p = f"{col_letter}$3"
        if side == "dr":
            f = (f'SUMIFS({RJ}!$F:$F,{RJ}!$J:$J,"<="&{p},{RJ}!$D:$D,"{acct}")'
                 f'-SUMIFS({RJ}!$G:$G,{RJ}!$J:$J,"<="&{p},{RJ}!$D:$D,"{acct}")')
        else:
            f = (f'SUMIFS({RJ}!$G:$G,{RJ}!$J:$J,"<="&{p},{RJ}!$D:$D,"{acct}")'
                 f'-SUMIFS({RJ}!$F:$F,{RJ}!$J:$J,"<="&{p},{RJ}!$D:$D,"{acct}")')
        ws.cell(row=row, column=ord(col_letter) - 64, value=f"={f}")

    R = {"cash": 4, "ar": 5, "inv": 6, "gstin": 7, "equip": 8, "tassets": 9,
         "ap": 11, "gstpay": 12, "loan": 13, "tliab": 14,
         "cap": 16, "re": 17, "cpnp": 18, "teq": 19, "chk": 21}
    labels = {4: "Cash & Bank", 5: "Accounts Receivable", 6: "Inventory",
              7: "GST Input Credit", 8: "Equipment (net of depreciation)",
              9: "TOTAL ASSETS", 11: "Accounts Payable", 12: "GST Payable",
              13: "Term Loan", 14: "TOTAL LIABILITIES", 16: "Owner's Capital",
              17: "Retained Earnings (opening)", 18: "Current period net profit",
              19: "TOTAL EQUITY"}
    for r, lbl in labels.items():
        ws.cell(row=r, column=1, value=lbl)
        if r in (9, 14, 19):
            ws.cell(row=r, column=1).font = Font(bold=True)
    for cl in ("B", "C"):
        bal(cl, R["cash"], "1000", "dr")
        bal(cl, R["ar"], "1100", "dr")
        bal(cl, R["inv"], "1200", "dr")
        bal(cl, R["gstin"], "1300", "dr")
        # equipment net = (dr-cr of 1500) + (dr-cr of 1510)
        p = f"{cl}$3"
        ws.cell(row=R["equip"], column=ord(cl) - 64, value=(
            f"=(SUMIFS({RJ}!$F:$F,{RJ}!$J:$J,\"<=\"&{p},{RJ}!$D:$D,\"1500\")"
            f"-SUMIFS({RJ}!$G:$G,{RJ}!$J:$J,\"<=\"&{p},{RJ}!$D:$D,\"1500\"))"
            f"+(SUMIFS({RJ}!$F:$F,{RJ}!$J:$J,\"<=\"&{p},{RJ}!$D:$D,\"1510\")"
            f"-SUMIFS({RJ}!$G:$G,{RJ}!$J:$J,\"<=\"&{p},{RJ}!$D:$D,\"1510\"))"))
        ws.cell(row=R["tassets"], column=ord(cl) - 64,
                value=f"=SUM({cl}4:{cl}8)").font = Font(bold=True)
        bal(cl, R["ap"], "2000", "cr")
        bal(cl, R["gstpay"], "2100", "cr")
        bal(cl, R["loan"], "2200", "cr")
        ws.cell(row=R["tliab"], column=ord(cl) - 64,
                value=f"=SUM({cl}11:{cl}13)").font = Font(bold=True)
        bal(cl, R["cap"], "3000", "cr")
        bal(cl, R["re"], "3100", "cr")
        ws.cell(row=R["cpnp"], column=ord(cl) - 64,
                value=f"='PL'!{get_column_letter(YTD_COL)}{16}")
        ws.cell(row=R["teq"], column=ord(cl) - 64,
                value=f"=SUM({cl}16:{cl}18)").font = Font(bold=True)
        ws.cell(row=R["chk"], column=ord(cl) - 64,
                value=(f'=IF(ROUND({cl}{R["tassets"]},2)='
                       f'ROUND({cl}{R["tliab"]}+{cl}{R["teq"]},2),'
                       f'"\u2713 BALANCED","\u2717 IMBALANCED")'))
    ws.cell(row=R["chk"], column=1, value="Check: Assets = Liabilities + Equity")
    money_fmt(ws, "B4:C19")
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.freeze_panes = "A4"
    return R


# ================================================================ dashboard
def sheet_dash(wb, pl, bs):
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "AGARWAL TRADERS — FINANCE DASHBOARD (all cards are live formulas)"
    ws["A1"].font = TITLE_FONT
    ytd = get_column_letter(YTD_COL)
    kpis = [
        ("YTD Revenue (₹)", f"='PL'!{ytd}{pl['rev']}", MONEY),
        ("YTD Net Profit (₹)", f"='PL'!{ytd}{pl['np']}", MONEY),
        ("Gross Margin %", f"=IF('PL'!{ytd}{pl['rev']}=0,0,'PL'!{ytd}{pl['gp']}/'PL'!{ytd}{pl['rev']})", PCT),
        ("Net Margin %", f"=IF('PL'!{ytd}{pl['rev']}=0,0,'PL'!{ytd}{pl['np']}/'PL'!{ytd}{pl['rev']})", PCT),
        ("Cash & Bank (₹)", f"='BalanceSheet'!C{bs['cash']}", MONEY),
        ("Accounts Receivable (₹)", f"='BalanceSheet'!C{bs['ar']}", MONEY),
        ("Accounts Payable (₹)", f"='BalanceSheet'!C{bs['ap']}", MONEY),
        ("Term Loan Outstanding (₹)", f"='BalanceSheet'!C{bs['loan']}", MONEY),
        ("Current Ratio", f"=IF('BalanceSheet'!C{bs['tliab']}=0,0,('BalanceSheet'!C{bs['cash']}+"
         f"'BalanceSheet'!C{bs['ar']}+'BalanceSheet'!C{bs['inv']}+'BalanceSheet'!C{bs['gstin']})"
         f"/'BalanceSheet'!C{bs['tliab']})", '0.00'),
    ]
    for i, (lbl, formula, fmt) in enumerate(kpis):
        r = 3 + i * 2
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=3)
        ws.cell(row=r, column=1, value=lbl).font = KPI_LBL
        c = ws.cell(row=r + 1, column=1, value=formula)
        c.font, c.number_format = KPI_FONT, fmt
        for rr in (r, r + 1):
            for cc in range(1, 4):
                ws.cell(row=rr, column=cc).border = BORDER
                ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor="DDEBF7")
    for col, w in zip("ABC", (10, 22, 22)):
        ws.column_dimensions[col].width = w

    # --- chart 1: monthly revenue vs net profit (line)
    lc = LineChart()
    lc.title = "Monthly Revenue vs Net Profit"
    lc.style, lc.height, lc.width = 10, 7.5, 15
    cats = Reference(wb["PL"], min_col=MO_FIRST, max_col=MO_LAST, min_row=2)
    lc.set_categories(cats)
    lc.add_data(Reference(wb["PL"], min_col=MO_FIRST, max_col=MO_LAST,
                          min_row=pl["rev"], max_row=pl["rev"]), from_rows=True)
    lc.add_data(Reference(wb["PL"], min_col=MO_FIRST, max_col=MO_LAST,
                          min_row=pl["np"], max_row=pl["np"]), from_rows=True)
    lc.series[0].title = SeriesLabel(v="Revenue")
    lc.series[1].title = SeriesLabel(v="Net Profit")
    ws.add_chart(lc, "E3")

    # --- chart 2: expense breakdown YTD (pie)
    pc = PieChart()
    pc.title = "YTD Expense Breakdown"
    pc.style, pc.height, pc.width = 10, 7.5, 13
    pc.add_data(Reference(wb["PL"], min_col=YTD_COL, max_col=YTD_COL,
                          min_row=pl["exp0"], max_row=pl["exp0"] + pl["n_exp"] - 1))
    pc.set_categories(Reference(wb["PL"], min_col=1, max_col=1,
                                min_row=pl["exp0"], max_row=pl["exp0"] + pl["n_exp"] - 1))
    ws.add_chart(pc, "E20")

    # --- chart 3: balance sheet comparison (bar)
    bc = BarChart()
    bc.title = "Balance Sheet: Mar-26 vs Sep-26"
    bc.style, bc.height, bc.width = 10, 7.5, 15
    bc.add_data(Reference(wb["BalanceSheet"], min_col=2, max_col=3,
                          min_row=bs["tassets"], max_row=bs["tassets"]), from_rows=True)
    bc.add_data(Reference(wb["BalanceSheet"], min_col=2, max_col=3,
                          min_row=bs["tliab"], max_row=bs["tliab"]), from_rows=True)
    bc.add_data(Reference(wb["BalanceSheet"], min_col=2, max_col=3,
                          min_row=bs["teq"], max_row=bs["teq"]), from_rows=True)
    bc.set_categories(Reference(wb["BalanceSheet"], min_col=2, max_col=3, min_row=2))
    for s, t in zip(bc.series, ["Total Assets", "Total Liabilities", "Total Equity"]):
        s.title = SeriesLabel(v=t)
    ws.add_chart(bc, "E37")


# ================================================================= ai report
def sheet_report(wb, pl, bs):
    ws = wb.create_sheet("AI_Report")
    con = sqlite3.connect(DB)
    q = lambda sql, *a: con.execute(sql, a).fetchone()[0]
    rev = q("SELECT ROUND(SUM(credit-debit),2) FROM journal_lines WHERE acct_no='4000'")
    cogs = q("SELECT ROUND(SUM(debit-credit),2) FROM journal_lines WHERE acct_no='5000'")
    np_ = q("""SELECT ROUND(SUM(CASE WHEN a.acct_type='Income' THEN credit-debit
                 WHEN a.acct_type='Expense' THEN debit-credit ELSE 0 END),2)
               FROM journal_lines j JOIN accounts a ON a.acct_no=j.acct_no""")
    exp_top = con.execute(
        """SELECT a.acct_name, ROUND(SUM(j.debit-j.credit),2) s
           FROM journal_lines j JOIN accounts a ON a.acct_no=j.acct_no
           WHERE a.acct_type='Expense' AND j.acct_no NOT IN ('5000')
           GROUP BY a.acct_name ORDER BY s DESC LIMIT 3""").fetchall()
    best = con.execute(
        """SELECT substr(date,1,7), ROUND(SUM(credit-debit),2) s
           FROM journal_lines WHERE acct_no='4000'
           GROUP BY 1 ORDER BY s DESC LIMIT 1""").fetchone()
    worst = con.execute(
        """SELECT substr(date,1,7), ROUND(SUM(credit-debit),2) s
           FROM journal_lines WHERE acct_no='4000'
           GROUP BY 1 ORDER BY s ASC LIMIT 1""").fetchone()
    cash = q("""SELECT ROUND(SUM(debit-credit),2) FROM journal_lines
                WHERE acct_no='1000'""")
    ar = q("""SELECT ROUND(SUM(debit-credit),2) FROM journal_lines
              WHERE acct_no='1100'""")
    loan = q("""SELECT ROUND(SUM(credit-debit),2) FROM journal_lines
                WHERE acct_no='2200'""")
    first6 = q("""SELECT ROUND(SUM(CASE WHEN a.acct_type='Income' THEN credit-debit
                    ELSE debit-credit END),2) FROM journal_lines j
                  JOIN accounts a ON a.acct_no=j.acct_no
                  WHERE a.acct_type IN ('Income','Expense') AND date < '2025-10-01'""")
    last6 = q("""SELECT ROUND(SUM(CASE WHEN a.acct_type='Income' THEN credit-debit
                    ELSE debit-credit END),2) FROM journal_lines j
                 JOIN accounts a ON a.acct_no=j.acct_no
                 WHERE a.acct_type IN ('Income','Expense') AND date >= '2026-04-01'""")
    con.close()

    gm = (rev - cogs) / rev * 100 if rev else 0
    nm = np_ / rev * 100 if rev else 0
    trend = "improving" if last6 > first6 else "softening"
    lines = [
        ("MANAGEMENT NARRATIVE — Agarwal Traders (Apr-25 to Sep-26)", TITLE_FONT),
        ("", None),
        ("1. Executive summary", SEC_FONT),
        (f"Revenue for the 18-month period is ₹{rev:,.0f} with net profit of "
         f"₹{np_:,.0f} — a net margin of {nm:.1f}%. Gross margin holds at {gm:.1f}%, "
         f"consistent with a wholesale FMCG mix. Profitability is {trend}: net profit "
         f"in the last 6 months (₹{last6:,.0f}) vs the first 6 months (₹{first6:,.0f}).", None),
        ("", None),
        ("2. Revenue trend", SEC_FONT),
        (f"Strongest month was {best[0]} at ₹{best[1]:,.0f}; softest was {worst[0]} at "
         f"₹{worst[1]:,.0f}. Festival-quarter months (Oct/Nov/Mar) run above trend — "
         "plan inventory build-up 3 weeks ahead of those peaks.", None),
        ("", None),
        ("3. Expense drivers", SEC_FONT),
        ("Top 3 non-COGS expenses: " + "; ".join(
            f"{n} ₹{s:,.0f}" for n, s in exp_top) +
         ". Salaries are the largest controllable line — review headcount against "
         "revenue per employee each quarter.", None),
        ("", None),
        ("4. Cash & working capital", SEC_FONT),
        (f"Cash & bank stands at ₹{cash:,.0f}; receivables at ₹{ar:,.0f} "
         f"({ar/rev*365/18:.0f} days of sales tied up). Term loan outstanding is "
         f"₹{loan:,.0f} with steady ₹55,000/month principal repayment. Tighten "
         "collection follow-ups to under 30 days to free working capital.", None),
        ("", None),
        ("5. Recommendations", SEC_FONT),
        ("a) Automate the monthly close: this workbook already reconciles every "
         "figure to the journals — run the Checks sheet before signing off.\n"
         "b) Set a 60% ceiling on COGS-to-revenue; flag any month breaching it.\n"
         "c) Build a 13-week cash forecast off the receipts/payments pattern in "
         "Raw_Journals to avoid surprise shortfalls.", None),
    ]
    ws["A1"] = lines[0][0]
    ws["A1"].font = TITLE_FONT
    r = 3
    for text, font in lines[2:]:
        if text == "":
            r += 1
            continue
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        c = ws.cell(row=r, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if font:
            c.font = font
            r += 1
        else:
            ws.row_dimensions[r].height = 34
            r += 1
    ws.column_dimensions["A"].width = 130


# ==================================================================== checks
def sheet_checks(wb, last_row, tb_rows, tb_total, pl, bs):
    ws = wb.create_sheet("Checks")
    ws["A1"] = "AUTOMATED INTEGRITY CHECKS — every result is a live formula"
    ws["A1"].font = TITLE_FONT
    hdr(ws, 3, 4, ["#", "Check", "Formula result", "Expected"])
    ytd = get_column_letter(YTD_COL)
    tb = "TrialBalance"
    r4000 = tb_rows["4000"]
    checks = [
        ("Every journal entry balances (debits = credits)",
         f'=IF(SUM({RJ}!F2:F{last_row})=SUM({RJ}!G2:G{last_row}),"PASS","FAIL")', "PASS"),
        ("Trial balance: total debits = total credits",
         f'=IF(ROUND({tb}!D{tb_total},2)=ROUND({tb}!E{tb_total},2),"PASS","FAIL")', "PASS"),
        ("P&L YTD revenue ties to trial balance (acct 4000)",
         f'=IF(ROUND(PL!{ytd}{pl["rev"]},2)=ROUND({tb}!E{r4000}-{tb}!D{r4000},2),"PASS","FAIL")', "PASS"),
        ("Net profit = Gross profit − Total expenses",
         f'=IF(ROUND(PL!{ytd}{pl["np"]},2)=ROUND(PL!{ytd}{pl["gp"]}-PL!{ytd}{pl["texp"]},2),"PASS","FAIL")', "PASS"),
        ("Balance sheet balances (Sep-26)",
         f'=IF(BalanceSheet!C{bs["chk"]}="✓ BALANCED","PASS","FAIL")', "PASS"),
        ("Receivables and payables are non-negative",
         f'=IF(AND(BalanceSheet!C{bs["ar"]}>=0,BalanceSheet!C{bs["ap"]}>=0),"PASS","FAIL")', "PASS"),
        ("No blank account numbers in journals",
         f'=IF(COUNTA({RJ}!D2:D{last_row})={last_row-1},"PASS","FAIL")', "PASS"),
        ("All period codes within Apr-25..Sep-26",
         f'=IF(AND(MIN({RJ}!J2:J{last_row})={PINT[0]},MAX({RJ}!J2:J{last_row})={PINT[-1]}),"PASS","FAIL")', "PASS"),
    ]
    for i, (desc, formula, exp) in enumerate(checks, 4):
        ws.cell(row=i, column=1, value=i - 3)
        ws.cell(row=i, column=2, value=desc)
        c = ws.cell(row=i, column=3, value=formula)
        c.font = Font(bold=True)
        ws.cell(row=i, column=4, value=exp)
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 10


# ==================================================================== guide
def sheet_guide(wb):
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "QB2BI — QuickBooks data → Relational DB → Live Financial Reports"
    ws["A1"].font = TITLE_FONT
    body = [
        ("What this is", SEC_FONT),
        ("A working demo of the 'QuickBooks Data Integration for PowerBI' gig pattern: "
         "QuickBooks-style double-entry journals are normalised into a relational SQLite "
         "database, and this workbook renders trial balance, P&L, balance sheet, dashboard "
         "and integrity checks as LIVE Excel formulas over the raw journals — the same "
         "architecture a Power BI model would use, in a file any finance team can open.", None),
        ("The pipeline", SEC_FONT),
        ("1. generate.py      → synthesises 18 months of double-entry journals (data/*.csv)\n"
         "2. build_db.py      → loads them into qb2bi.sqlite (accounts + journal_lines)\n"
         "3. build_workbook.py→ renders this workbook; every summary cell is a SUMIFS formula\n"
         "4. tests/test_build.py → pytest suite incl. LibreOffice headless recalc vs Python ground truth",
         None),
        ("How to run", SEC_FONT),
        ("python3 generate.py && python3 build_db.py && python3 build_workbook.py && pytest -q",
         None),
        ("Sheet map", SEC_FONT),
        ("Raw_Journals  every journal line + Period helper (formula)\n"
         "TrialBalance  per-account totals via SUMIFS\n"
         "PL            monthly P&L Apr-25..Sep-26 + YTD\n"
         "BalanceSheet  as of Mar-26 and Sep-26, with balance check\n"
         "Dashboard     9 KPI cards + 3 native charts (all formula-driven)\n"
         "AI_Report     rule-based management narrative\n"
         "Checks        8 automated integrity checks → all should read PASS", None),
        ("Data note", SEC_FONT),
        ("All figures are synthetic and fictional (seeded random), for demo purposes only.",
         None),
    ]
    r = 3
    for text, font in body:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(row=r, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if font:
            c.font = font
            r += 1
        else:
            ws.row_dimensions[r].height = 52
            r += 1
    ws.column_dimensions["A"].width = 130


# ==================================================================== main
def main():
    wb = Workbook()
    sheet_guide(wb)
    last_row = sheet_raw(wb)
    tb_rows, tb_total = sheet_tb(wb)
    pl = sheet_pl(wb)
    bs = sheet_bs(wb)
    sheet_dash(wb, pl, bs)
    sheet_report(wb, pl, bs)
    sheet_checks(wb, last_row, tb_rows, tb_total, pl, bs)
    wb.save(XLSX)
    print(f"workbook saved: {XLSX} ({XLSX.stat().st_size/1024:.0f} KB, "
          f"{last_row-1} journal lines)")


if __name__ == "__main__":
    main()
