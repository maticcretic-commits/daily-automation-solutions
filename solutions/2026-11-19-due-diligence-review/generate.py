#!/usr/bin/env python3
"""Build due-diligence-review.xlsx — DueDiligence Review Pack.

Built around the Freelancer lead "Pre-Investment Due Diligence Review"
(finance-adjacent reporting wedge): a pre-investment financial
due-diligence review of a fictional Indian private limited company,
Sutlej Agro Foods Pvt. Ltd., Ludhiana, over FY24-FY26.

Sheets:
  Guide      — what the pack is, how to read it, fictional-data note.
  Financials — 3-year P&L + balance sheet (Rs lakh), derived lines live.
  Ratios     — 12 ratios x 3 years + growth block, all live formulas.
  BankRec    — 12-month book-vs-bank reconciliation + open items with
               days outstanding at the as-of date (live).
  Loans      — 5 loan facilities with LT/current split + contingent
               liabilities and % of net worth.
  RedFlags   — 9 scored red flags; every status is a live trigger
               formula reading the other sheets; risk score = 31 (HIGH).
  Dashboard  — target, as-of date, 15 KPIs, severity table, 3 charts.
  Checks     — 10 integrity checks against the independent Python
               ground truth; all must read PASS.

Formula rules (LibreOffice 24.2): plain A1-style ranges, no Table
objects, no structured references; every range literal repeats the
column letter on both ends.
"""
import csv
import datetime as dt
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
XLSX = BASE / "due-diligence-review.xlsx"

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14)
BOLD = Font(bold=True, size=11)
NUM = "#,##0"
NUM2 = "#,##0.00"
RATIO = "0.00"
PCT = "0.0%"
DATE_FMT = "yyyy-mm-dd"

YEARS = ["FY24", "FY25", "FY26"]
# Financials sheet layout: P&L rows 4-14, balance sheet rows 17-31.
PL_ITEMS = ["Revenue", "COGS", "Gross profit", "Operating expenses",
            "EBITDA", "Depreciation", "EBIT", "Interest", "PBT", "Tax",
            "Net income"]
BS_ITEMS = ["Share capital", "Reserves", "Net worth",
            "Long-term borrowings", "Short-term borrowings",
            "Trade payables", "Other current liabilities",
            "Total liabilities & equity", "Fixed assets (net)",
            "Inventory", "Trade receivables", "Cash & bank",
            "Other current assets", "Total assets"]
PL_DERIVED = {6: "={c}4-{c}5", 8: "={c}6-{c}7", 10: "={c}8-{c}9",
              12: "={c}10-{c}11", 14: "={c}12-{c}13"}
BS_DERIVED = {19: "={c}17+{c}18", 24: "=SUM({c}19:{c}23)",
              30: "=SUM({c}25:{c}29)"}

RATIO_DEFS = [
    ("Current ratio",
     "=({f}{c}26+{f}{c}27+{f}{c}28+{f}{c}29)"
     "/({f}{c}21+{f}{c}22+{f}{c}23)", RATIO,
     "Current assets / current liabilities; target >= 1.5"),
    ("Quick ratio",
     "=({f}{c}27+{f}{c}28+{f}{c}29)"
     "/({f}{c}21+{f}{c}22+{f}{c}23)", RATIO,
     "Liquid assets / current liabilities; target >= 1.0"),
    ("Debt / equity",
     "=({f}{c}20+{f}{c}21)/{f}{c}19", RATIO,
     "Total borrowings / net worth; watch above 2.0"),
    ("Interest coverage",
     "={f}{c}10/{f}{c}11", RATIO,
     "EBIT / interest; below 2.0x is thin"),
    ("Gross margin", "={f}{c}6/{f}{c}4", PCT, "Gross profit / revenue"),
    ("EBITDA margin", "={f}{c}8/{f}{c}4", PCT, "EBITDA / revenue"),
    ("Net margin", "={f}{c}14/{f}{c}4", PCT, "Net income / revenue"),
    ("ROCE", "={f}{c}10/({f}{c}19+{f}{c}20)", PCT,
     "EBIT / (net worth + long-term borrowings)"),
    ("ROE", "={f}{c}14/{f}{c}19", PCT, "Net income / net worth"),
    ("Receivable days", "={f}{c}27/{f}{c}4*365", NUM,
     "Trade receivables / revenue x 365"),
    ("Payable days", "={f}{c}22/{f}{c}5*365", NUM,
     "Trade payables / COGS x 365"),
    ("Inventory days", "={f}{c}26/{f}{c}5*365", NUM,
     "Inventory / COGS x 365"),
]


def load_csv(name):
    with open(DATA / name, newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT


def build():
    fin_rows = load_csv("financials.csv")
    fin = {r["item"]: [float(r["fy24"]), float(r["fy25"]),
                       float(r["fy26"])] for r in fin_rows}
    bank_rows = load_csv("bankrec.csv")
    items_rows = load_csv("bank_items.csv")
    loan_rows = load_csv("loans.csv")
    cont_rows = load_csv("contingent.csv")
    flag_rows = load_csv("redflags.csv")
    meta = {r["key"]: r["value"] for r in load_csv("meta.csv")}
    company = meta["company"]
    as_of = dt.date.fromisoformat(meta["as_of"])
    cagr = (fin["Revenue"][2] / fin["Revenue"][0]) ** 0.5 - 1

    wb = openpyxl.Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = ("Pre-Investment Due Diligence Review — "
                "Sutlej Agro Foods Pvt. Ltd.")
    ws["A1"].font = TITLE_FONT
    lines = [
        "A formula-driven pre-investment due-diligence review pack for a",
        f"fictional Indian private limited company ({company}),",
        f"financial years FY24-FY26, as of {as_of.isoformat()}. Built as a",
        "working demo for the Freelancer project 'Pre-Investment Due",
        "Diligence Review' (finance-adjacent reporting wedge).",
        "",
        "ALL DATA IS FICTIONAL — invented for the demo. No real company,",
        "no real filings, no real bank statements.",
        "",
        "WHAT EACH SHEET DOES",
        "  Financials  3-year P&L and balance sheet (Rs lakh). Gross",
        "              profit, EBITDA, EBIT, PBT, net income, net worth",
        "              and both totals are live formulas.",
        "  Ratios      12 ratios x 3 years plus a growth block (revenue",
        "              CAGR, cash growth, contingent liabilities as % of",
        "              net worth, receivables vs revenue growth).",
        "  BankRec     12-month book-vs-bank reconciliation from the",
        "              cash book, plus the open items at the as-of date",
        "              with days outstanding (live).",
        "  Loans       5 loan facilities with long-term / current split,",
        "              plus contingent liabilities and % of net worth.",
        "  RedFlags    9 scored red flags. Every status is a live",
        "              trigger formula reading the other sheets — change",
        "              a number and the flags move. Risk score 31 = HIGH.",
        "  Dashboard   Target, as-of date, 15 KPIs, severity table and",
        "              3 native charts.",
        "  Checks      10 integrity checks vs the independent Python",
        "              ground truth; all must read PASS.",
        "",
        "HEADLINE VERDICT (live formula on Dashboard)",
        f"  Revenue CAGR {cagr:.1%} over FY24-FY26, but receivable days",
        "  climb every year, quick ratio stays under 1.0x, interest",
        "  cover never reaches 2.0x, contingent liabilities run to",
        "  24.8% of net worth, and a 0% director loan sits repayable",
        "  on demand. Risk score 31 — HIGH: proceed only with price",
        "  protection and conditions.",
        "",
        "HOW TO RUN THE TESTS",
        "  python3 tests/test_build.py   (rebuilds, then LibreOffice",
        "  headless recalc vs the Python ground truth in data/)",
    ]
    for i, line in enumerate(lines, start=2):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 78

    # ---------------- Financials ----------------
    ws = wb.create_sheet("Financials")
    ws["A1"] = "Financials — Sutlej Agro Foods Pvt. Ltd. (Rs lakh)"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "Line item (Rs lakh)"
    for j, yr in enumerate(YEARS):
        ws.cell(row=3, column=2 + j, value=yr)
    style_header(ws, 3, 4)
    for i, item in enumerate(PL_ITEMS):
        r = 4 + i
        ws.cell(row=r, column=1, value=item)
        if r in PL_DERIVED:
            for j, col in enumerate("BCD"):
                ws.cell(row=r, column=2 + j,
                       value=PL_DERIVED[r].format(c=col))
        else:
            for j in range(3):
                ws.cell(row=r, column=2 + j, value=fin[item][j])
    ws.cell(row=16, column=1, value="Balance sheet (Rs lakh)").font = BOLD
    for i, item in enumerate(BS_ITEMS):
        r = 17 + i
        ws.cell(row=r, column=1, value=item)
        if r in BS_DERIVED:
            for col in "BCD":
                ws.cell(row=r, column=2 + "BCD".index(col),
                       value=BS_DERIVED[r].format(c=col))
        else:
            for j in range(3):
                ws.cell(row=r, column=2 + j, value=fin[item][j])
    ws.cell(row=31, column=1,
           value="Check: Total assets - Total liabilities & equity")
    for col in "BCD":
        ws.cell(row=31, column=2 + "BCD".index(col),
               value=f"={col}30-{col}24")
    for row in list(range(4, 15)) + list(range(17, 32)):
        for col in range(2, 5):
            ws.cell(row=row, column=col).number_format = NUM
    for r in (6, 8, 10, 12, 14, 19, 24, 30, 31):
        for col in range(1, 5):
            ws.cell(row=r, column=col).font = BOLD
    ws.column_dimensions["A"].width = 42
    for col in "BCD":
        ws.column_dimensions[col].width = 16

    # ---------------- Ratios ----------------
    ws = wb.create_sheet("Ratios")
    ws["A1"] = "Ratios — FY24-FY26"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "Ratio"
    for j, yr in enumerate(YEARS):
        ws.cell(row=3, column=2 + j, value=yr)
    ws.cell(row=3, column=5, value="Reading")
    style_header(ws, 3, 5)
    for i, (label, formula, fmt, reading) in enumerate(RATIO_DEFS):
        r = 4 + i
        ws.cell(row=r, column=1, value=label)
        for j, col in enumerate("BCD"):
            c = ws.cell(row=r, column=2 + j,
                       value=formula.format(f="Financials!", c=col))
            c.number_format = fmt
        ws.cell(row=r, column=5, value=reading)
    ws.cell(row=17, column=1, value="Revenue CAGR FY24->FY26")
    ws.cell(row=17, column=2,
           value="=(Financials!D4/Financials!B4)^(1/2)-1").number_format = PCT
    ws.cell(row=18, column=1, value="Cash growth FY24->FY26")
    ws.cell(row=18, column=2,
           value="=(Financials!D28-Financials!B28)/Financials!B28"
           ).number_format = PCT
    ws.cell(row=19, column=1,
           value="Contingent liabilities % of net worth")
    ws.cell(row=19, column=2, value="=Loans!B17").number_format = PCT
    ws.cell(row=20, column=1, value="Receivables growth FY24->FY26")
    ws.cell(row=20, column=2,
           value="=(Financials!D27-Financials!B27)/Financials!B27"
           ).number_format = PCT
    ws.cell(row=21, column=1, value="Revenue growth FY24->FY26")
    ws.cell(row=21, column=2,
           value="=(Financials!D4-Financials!B4)/Financials!B4"
           ).number_format = PCT
    for r in (17, 18, 19, 20, 21):
        ws.cell(row=r, column=1).font = BOLD
    ws.column_dimensions["A"].width = 42
    for col in "BCD":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["E"].width = 52

    # ---------------- BankRec ----------------
    ws = wb.create_sheet("BankRec")
    ws["A1"] = "Bank reconciliation — Apr 2025 to Mar 2026 (Rs lakh)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "As-of date"
    ws["B2"] = as_of
    ws["B2"].number_format = DATE_FMT
    headers = ["Month", "Receipts", "Payments", "Book balance",
               "Bank balance", "Difference (book - bank)"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, 6)
    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
                   "Sep", "Oct", "Nov", "Dec"]
    for i, br in enumerate(bank_rows):
        r = 4 + i
        y, m = br["month"].split("-")
        ws.cell(row=r, column=1,
               value=f"{month_names[int(m) - 1]} {y}")
        ws.cell(row=r, column=2, value=float(br["receipts"]))
        ws.cell(row=r, column=3, value=float(br["payments"]))
        if i == 0:
            ws.cell(row=r, column=4,
                   value="=Financials!C28+B4-C4")
        else:
            ws.cell(row=r, column=4, value=f"=D{r - 1}+B{r}-C{r}")
        ws.cell(row=r, column=5, value=float(br["bank_balance"]))
        ws.cell(row=r, column=6, value=f"=D{r}-E{r}")
        for col in (2, 3, 4, 5, 6):
            ws.cell(row=r, column=col).number_format = NUM2
    ws.cell(row=16, column=1, value="Total").font = BOLD
    ws.cell(row=16, column=2, value="=SUM(B4:B15)").number_format = NUM2
    ws.cell(row=16, column=3, value="=SUM(C4:C15)").number_format = NUM2
    ws.cell(row=18, column=1,
           value="Open items at 31 Mar 2026").font = BOLD
    for j, h in enumerate(["Item", "Date", "Amount (Rs lakh)",
                           "Effect on book - bank",
                           "Days outstanding"], start=1):
        ws.cell(row=19, column=j, value=h)
    style_header(ws, 19, 5)
    open_items = [r for r in items_rows if not r["cleared"]]
    for i, it in enumerate(open_items):
        r = 20 + i
        ws.cell(row=r, column=1, value=it["item"])
        d = ws.cell(row=r, column=2,
                   value=dt.date.fromisoformat(it["date"]))
        d.number_format = DATE_FMT
        ws.cell(row=r, column=3, value=float(it["amount"])
               ).number_format = NUM2
        ws.cell(row=r, column=4, value=float(it["effect"])
               ).number_format = NUM2
        ws.cell(row=r, column=5, value=f"=$B$2-B{r}").number_format = NUM
    ws.cell(row=24, column=1, value="Total open-item effect").font = BOLD
    ws.cell(row=24, column=4, value="=SUM(D20:D23)").number_format = NUM2
    ws.cell(row=25, column=1,
           value="Check: open items - statement difference").font = BOLD
    ws.cell(row=25, column=2, value="=D24-F15").number_format = NUM2
    ws.column_dimensions["A"].width = 44
    for col in "BCDEF":
        ws.column_dimensions[col].width = 18

    # ---------------- Loans ----------------
    ws = wb.create_sheet("Loans")
    ws["A1"] = "Loan facilities & contingent liabilities (Rs lakh)"
    ws["A1"].font = TITLE_FONT
    headers = ["Facility", "Lender", "Sanctioned", "Outstanding", "Rate",
               "EMI / month", "Overdue", "Security / notes",
               "Long-term portion", "Current portion"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, 10)
    for i, lr in enumerate(loan_rows):
        r = 4 + i
        ws.cell(row=r, column=1, value=lr["facility"])
        ws.cell(row=r, column=2, value=lr["lender"])
        ws.cell(row=r, column=3, value=float(lr["sanctioned"]))
        ws.cell(row=r, column=4, value=float(lr["outstanding"]))
        ws.cell(row=r, column=5, value=float(lr["rate"])
               ).number_format = PCT
        emi = lr["emi"]
        ws.cell(row=r, column=6,
               value=float(emi) if emi else "—").number_format = NUM2
        ws.cell(row=r, column=7, value=int(lr["overdue"]))
        ws.cell(row=r, column=8, value=lr["security"])
        ws.cell(row=r, column=9, value=float(lr["lt_portion"]))
        ws.cell(row=r, column=10, value=float(lr["current_portion"]))
        for col in (3, 4, 9, 10):
            ws.cell(row=r, column=col).number_format = NUM2
    ws.cell(row=9, column=1, value="Total").font = BOLD
    for col, letter in ((3, "C"), (4, "D"), (9, "I"), (10, "J")):
        ws.cell(row=9, column=col,
               value=f"=SUM({letter}4:{letter}8)").number_format = NUM2
    ws.cell(row=11, column=1,
           value="Contingent liabilities (Rs lakh)").font = BOLD
    ws.cell(row=12, column=1, value="Item").font = BOLD
    ws.cell(row=12, column=2, value="Amount").font = BOLD
    for i, cr in enumerate(cont_rows):
        r = 13 + i
        ws.cell(row=r, column=1, value=cr["item"])
        ws.cell(row=r, column=2, value=float(cr["amount"])
               ).number_format = NUM2
    ws.cell(row=16, column=1, value="Total").font = BOLD
    ws.cell(row=16, column=2, value="=SUM(B13:B15)").number_format = NUM2
    ws.cell(row=17, column=1, value="% of FY26 net worth").font = BOLD
    ws.cell(row=17, column=2, value="=B16/Financials!D19"
           ).number_format = PCT
    ws.cell(row=19, column=1, value=(
        "Note: total outstanding must equal FY26 total borrowings "
        "(Financials long-term + short-term).")).font = BOLD
    widths = [30, 22, 13, 13, 9, 12, 9, 30, 16, 15]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    # ---------------- RedFlags ----------------
    ws = wb.create_sheet("RedFlags")
    ws["A1"] = "Red flags — scored (all statuses are live formulas)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Auditor emphasis of matter on GST demand (FY26):"
    ws["B2"] = "Yes"
    headers = ["ID", "Area", "Red flag", "Severity", "Weight",
               "Status (live)", "Points"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, 7)
    triggers = {
        "F1": ('=IF(AND(Ratios!D13>Ratios!C13,'
               'Ratios!C13>Ratios!B13),"TRIGGERED","CLEAR")'),
        "F2": '=IF(Ratios!B19>0.2,"TRIGGERED","CLEAR")',
        "F3": '=IF(AND(Loans!D8>0,Loans!E8=0),"TRIGGERED","CLEAR")',
        "F4": '=IF(MAX(Ratios!B7:D7)<2,"TRIGGERED","CLEAR")',
        "F5": '=IF(MAX(Ratios!B5:D5)<1,"TRIGGERED","CLEAR")',
        "F6": '=IF(MAX(BankRec!E20:E23)>30,"TRIGGERED","CLEAR")',
        "F7": '=IF(Loans!B13>0,"TRIGGERED","CLEAR")',
        "F8": '=IF(Ratios!B20>Ratios!B21,"TRIGGERED","CLEAR")',
        "F9": '=IF($B$2="Yes","TRIGGERED","CLEAR")',
    }
    for i, fr in enumerate(flag_rows):
        r = 4 + i
        ws.cell(row=r, column=1, value=fr["id"])
        ws.cell(row=r, column=2, value=fr["area"])
        ws.cell(row=r, column=3, value=fr["flag"])
        ws.cell(row=r, column=4, value=fr["severity"])
        ws.cell(row=r, column=5, value=int(fr["weight"]))
        ws.cell(row=r, column=6, value=triggers[fr["id"]])
        ws.cell(row=r, column=7, value=f'=IF(F{r}="TRIGGERED",E{r},0)')
    ws.cell(row=13, column=1, value="RISK SCORE").font = BOLD
    ws.cell(row=13, column=7, value="=SUM(G4:G12)").font = BOLD
    ws.cell(row=14, column=1, value=(
        "Verdict bands: 30+ HIGH — proceed only with price protection "
        "and conditions; 15-29 MODERATE — conditions required; "
        "under 15 LOW — no material flags.")).font = BOLD
    widths = [6, 24, 78, 10, 8, 12, 8]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Due Diligence Dashboard — Sutlej Agro Foods Pvt. Ltd."
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "Target company"
    ws["B3"] = company
    ws["A4"] = "As-of date"
    ws.cell(row=4, column=2, value="=BankRec!B2").number_format = DATE_FMT
    ws["A5"] = "Review basis"
    ws["B5"] = ("Audited financials FY24-FY26 + bank statements + loan "
                "schedules (fictional demo data)")
    ws.cell(row=7, column=1, value="KEY FINDINGS").font = BOLD
    kpis = [
        ("Revenue FY26 (Rs lakh)", "=Financials!D4", NUM2),
        ("Revenue CAGR FY24->FY26", "=Ratios!B17", PCT),
        ("EBITDA margin FY26", "=Ratios!D9", PCT),
        ("Current ratio FY26", "=Ratios!D4", RATIO),
        ("Quick ratio FY26", "=Ratios!D5", RATIO),
        ("Debt / equity FY26", "=Ratios!D6", RATIO),
        ("Interest coverage FY26", "=Ratios!D7", RATIO),
        ("Contingent liabilities (Rs lakh)", "=Loans!B16", NUM2),
        ("Contingent liabilities % of net worth", "=Ratios!B19", PCT),
        ("Red flags triggered", '=COUNTIF(RedFlags!F4:F12,"TRIGGERED")',
         NUM),
        ("High-severity flags",
         '=COUNTIFS(RedFlags!D4:D12,"High",RedFlags!F4:F12,"TRIGGERED")',
         NUM),
        ("Medium-severity flags",
         '=COUNTIFS(RedFlags!D4:D12,"Medium",RedFlags!F4:F12,'
         '"TRIGGERED")', NUM),
        ("Low-severity flags",
         '=COUNTIFS(RedFlags!D4:D12,"Low",RedFlags!F4:F12,"TRIGGERED")',
         NUM),
        ("Risk score (max 31)", "=RedFlags!G13", NUM),
        ("Verdict",
         '=IF(B21>=30,"HIGH — proceed only with price protection and '
         'conditions",IF(B21>=15,"MODERATE — conditions required",'
         '"LOW — no material flags"))', None),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        r = 8 + i
        ws.cell(row=r, column=1, value=label)
        c = ws.cell(row=r, column=2, value=formula)
        if fmt:
            c.number_format = fmt
        if r in (21, 22):
            ws.cell(row=r, column=1).font = BOLD
            c.font = BOLD
    ws.cell(row=24, column=1, value="Severity").font = BOLD
    ws.cell(row=24, column=2, value="Triggered").font = BOLD
    ws.cell(row=24, column=3, value="Weight each").font = BOLD
    for i, (sev, ref, w) in enumerate([("High", "=B18", 5),
                                       ("Medium", "=B19", 3),
                                       ("Low", "=B20", 1)]):
        r = 25 + i
        ws.cell(row=r, column=1, value=sev)
        ws.cell(row=r, column=2, value=ref)
        ws.cell(row=r, column=3, value=w)
    ws.cell(row=28, column=1, value="Total").font = BOLD
    ws.cell(row=28, column=2, value="=SUM(B25:B27)").font = BOLD
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 14

    # Charts
    fin_ws = wb["Financials"]
    rat_ws = wb["Ratios"]
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "Revenue vs EBITDA (Rs lakh)"
    ch1.y_axis.title = "Rs lakh"
    ch1.x_axis.title = "Year"
    for row in (4, 8):
        ch1.add_data(Reference(fin_ws, min_col=1, max_col=4,
                               min_row=row, max_row=row),
                     from_rows=True, titles_from_data=True)
    ch1.set_categories(Reference(fin_ws, min_col=2, max_col=4,
                                 min_row=3, max_row=3))
    ch1.height = 7.5
    ch1.width = 14
    ws.add_chart(ch1, "A30")

    ch2 = LineChart()
    ch2.title = "Current ratio vs quick ratio"
    ch2.y_axis.title = "Ratio"
    ch2.x_axis.title = "Year"
    for row in (4, 5):
        ch2.add_data(Reference(rat_ws, min_col=1, max_col=4,
                               min_row=row, max_row=row),
                     from_rows=True, titles_from_data=True)
    ch2.set_categories(Reference(rat_ws, min_col=2, max_col=4,
                                 min_row=3, max_row=3))
    ch2.height = 7.5
    ch2.width = 14
    ws.add_chart(ch2, "A46")

    ch3 = BarChart()
    ch3.type = "col"
    ch3.title = "Red flags triggered by severity"
    ch3.y_axis.title = "Flags"
    ch3.add_data(Reference(ws, min_col=2, min_row=24, max_row=27),
                 titles_from_data=True)
    ch3.set_categories(Reference(ws, min_col=1, min_row=25, max_row=27))
    ch3.height = 7.5
    ch3.width = 14
    ws.add_chart(ch3, "A62")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws["A1"] = "Integrity checks — all must read PASS"
    ws["A1"].font = TITLE_FONT
    for j, h in enumerate(["#", "Check", "Expected", "Actual",
                           "Result"], start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, 5)
    checks = [
        ("Balance sheet balances, FY24", 0, "=Financials!B31"),
        ("Balance sheet balances, FY25", 0, "=Financials!C31"),
        ("Balance sheet balances, FY26", 0, "=Financials!D31"),
        ("Loan schedule total = FY26 borrowings", 950, "=Loans!D9"),
        ("Bank rec difference, Mar 2026", 1.55, "=BankRec!F15"),
        ("Open items reconcile to statement", 0, "=BankRec!B25"),
        ("Contingent liabilities total", 130, "=Loans!B16"),
        ("Revenue CAGR FY24->FY26", cagr, "=Ratios!B17"),
        ("Risk score", 31, "=RedFlags!G13"),
        ("Red flags triggered", 9, "=Dashboard!B17"),
    ]
    for i, (label, expected, actual) in enumerate(checks):
        r = 4 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=expected)
        ws.cell(row=r, column=4, value=actual)
        ws.cell(row=r, column=5,
               value=f'=IF(ABS(C{r}-D{r})<0.01,"PASS","FAIL")')
        ws.cell(row=r, column=3).number_format = NUM2
        ws.cell(row=r, column=4).number_format = NUM2
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 10

    wb.save(XLSX)
    print(f"built {XLSX}")
    print(f"revenue CAGR {cagr:.4%}, risk score 31, flags 9")


if __name__ == "__main__":
    build()
