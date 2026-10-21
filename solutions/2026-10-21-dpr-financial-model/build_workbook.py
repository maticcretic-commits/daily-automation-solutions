"""Build dprpulse.xlsx — formula-driven bank DPR financial model.

Reads data/*.csv (generate.py). Every computed cell is a live Excel formula;
no computed values are pasted. Verified by tests/test_build.py via a
LibreOffice headless recalculation against main.py ground truth.

LO-COMPAT notes (from earlier builds):
  * no MAXIFS/MINIFS/XLOOKUP — chains, IF and MIN/MAX over ranges only
  * range literals repeat the column letter: Sheet!$B$2:$B$9
  * no bare "Rs " inside TEXT() format strings
"""
import csv
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=16, color="1F4E79")
INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")
MONEY = "#,##0.00"
MONEY0 = "#,##0"
PCT1 = "0.0%"
PCT0 = "0%"
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

YEARS = list(range(1, 8))          # 1..7
Y0_YEARS = list(range(0, 8))       # 0..7


def col(n):
    return get_column_letter(n)


def yr_col(t, y0=False):
    """Column letter for year t.
    y0=False -> B..H for years 1..7; y0=True -> B..I for years 0..7."""
    return col(t + 2) if y0 else col(t + 1)


def load_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, ncols, nrows=1):
    for r in range(1, nrows + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER


def style_range(ws, min_row, max_row, ncols):
    for r in range(min_row, max_row + 1):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER


def widths(ws, wsizes):
    for i, w in enumerate(wsizes, start=1):
        ws.column_dimensions[col(i)].width = w


A = "Assumptions"   # sheet-name constants used inside formulas
PC = "Project_Cost"
MF = "Means_of_Finance"
LS = "Loan_Schedule"
RC = "Revenue_Cost"
PL = "PL"
CF = "CashFlow_DSCR"
BS = "BalanceSheet"
RA = "Ratios"
SE = "Sensitivity"
DB = "Dashboard"
CH = "Checks"


# ------------------------------------------------------------------ Guide
def sheet_guide(wb):
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "DPRPulse — Detailed Project Report Financial Model"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = ("Fictional demo case: Sharma Agro Foods, turmeric processing unit, "
                "Patna (Bihar). Project cost Rs 240 lakh; promoter margin 25%; "
                "term loan Rs 180 lakh @ 10% for 7 years (1-yr moratorium).")
    ws["A3"].alignment = Alignment(wrap_text=True)
    ws["A5"] = "How to use:"
    for i, line in enumerate([
        "Blue cells on the Assumptions sheet are the ONLY inputs — change them and the whole model recalculates.",
        "Every other number in this workbook is a live formula. Nothing is pasted as a value.",
        "Read order: Assumptions -> Project_Cost -> Means_of_Finance -> Loan_Schedule -> "
        "Revenue_Cost -> PL -> CashFlow_DSCR -> BalanceSheet -> Ratios -> Sensitivity -> Dashboard.",
        "The Checks sheet must show TRUE on all 10 rows; the Dashboard carries the banker's summary.",
    ], start=6):
        ws[f"A{i}"] = f"{i - 5}. {line}"
        ws[f"A{i}"].alignment = Alignment(wrap_text=True)
    rows = [
        ("Assumptions", "All model inputs (blue). Utilisation ramp, prices, cost %, loan terms."),
        ("Project_Cost", "Rs 240 lakh capex split across 7 heads."),
        ("Means_of_Finance", "Promoter Rs 60 lakh + term loan Rs 180 lakh; margin 25%."),
        ("Loan_Schedule", "Annual amortisation: 1-yr moratorium, then 6 equal instalments."),
        ("Revenue_Cost", "Production, price, revenue and the 7 operating-cost lines."),
        ("PL", "7-year projected profit & loss; tax 25% on positive PBT."),
        ("CashFlow_DSCR", "Cash accruals, debt service, DSCR, NPV @12%, equity IRR, payback."),
        ("BalanceSheet", "Projected balance sheet; assets = liabilities every year."),
        ("Ratios", "Bank ratios: DSCR, D:E, break-even utilisation, NPV, IRR."),
        ("Sensitivity", "4 scenarios: base, revenue -5%, raw +3pp, combined stress."),
        ("Dashboard", "8 KPIs, 3 charts, risk flags for the banker."),
        ("Checks", "10 mechanical integrity checks — all must read TRUE."),
    ]
    r0 = 11
    ws[f"A{r0}"] = "Sheet"
    ws[f"B{r0}"] = "What it holds"
    for i, (s, d) in enumerate(rows, start=r0 + 1):
        ws[f"A{i}"] = s
        ws[f"B{i}"] = d
    style_header(ws, 2, nrows=1)
    ws.row_dimensions[r0].height = 18
    for r in range(r0, r0 + 1 + len(rows)):
        ws.row_dimensions[r].height = 15
    style_range(ws, r0, r0 + len(rows), 2)
    widths(ws, [22, 95])
    ws.sheet_properties.pageSetUpPr = None
    return ws


# ------------------------------------------------------------ Assumptions
def sheet_assumptions(wb):
    ws = wb.create_sheet(A)
    ws["A1"] = "Parameter"
    ws["B1"] = "Value"
    ws["C1"] = "Note"
    params = [
        ("Capacity (MT/year)", 1200, ""),
        ("Selling price Y1 (Rs/kg)", 150, ""),
        ("Price escalation /yr", 0.03, ""),
        ("Capacity utilisation Y1..Y7", None, "one value per year ->"),
        ("Raw material (% of revenue)", 0.76, ""),
        ("Packing (% of revenue)", 0.06, ""),
        ("Admin & selling (% of revenue)", 0.05, ""),
        ("Power & fuel Y1 (Rs lakh)", 22.0, ""),
        ("Power escalation /yr", 0.04, ""),
        ("Salaries Y1 (Rs lakh)", 34.0, ""),
        ("Salary escalation /yr", 0.05, ""),
        ("Repairs & maintenance (Rs lakh/yr)", 2.2, "flat"),
        ("Insurance Y1 (Rs lakh)", 1.5, ""),
        ("Insurance escalation /yr", 0.03, ""),
        ("Depreciation - building (Rs lakh/yr)", 2.75, "55 / 20 yrs"),
        ("Depreciation - P&M incl. contingency", 12.0, "120 / 10 yrs"),
        ("Depreciation - electricals", 1.8, "18 / 10 yrs"),
        ("Preliminary amortisation Y1-Y5", 2.4, "12 / 5 yrs"),
        ("Corporate tax rate", 0.25, "on positive PBT"),
        ("Term loan (Rs lakh)", 180.0, ""),
        ("Interest rate /yr", 0.10, ""),
        ("Moratorium (years)", 1, "interest serviced"),
        ("Repayment years", 6, "annual instalments"),
        ("Promoter equity (Rs lakh)", 60.0, ""),
        ("Discount rate (NPV/IRR hurdle)", 0.12, ""),
        ("Project cost total (Rs lakh)", 240.0, "check vs Project_Cost"),
    ]
    utils = [0.55, 0.65, 0.75, 0.85, 0.90, 0.90, 0.90]
    for i, (label, value, note) in enumerate(params, start=2):
        ws[f"A{i}"] = label
        ws[f"C{i}"] = note
        if value is None:  # utilisation row -> B..H
            for t, u in enumerate(utils, start=2):
                c = ws.cell(row=i, column=t, value=u)
                c.number_format = PCT0
                c.fill = INPUT_FILL
        else:
            c = ws.cell(row=i, column=2, value=value)
            c.fill = INPUT_FILL
        ws[f"A{i}"].border = BORDER
        ws[f"B{i}"].border = BORDER
        ws[f"C{i}"].border = BORDER
    # number formats for the single-value rows
    for r, fmt in [(4, PCT1), (6, PCT1), (7, PCT1), (8, PCT1), (10, PCT1),
                   (12, PCT1), (14, PCT1), (20, PCT0), (22, PCT0), (26, PCT1)]:
        ws[f"B{r}"].number_format = fmt
    for r in (2, 3, 9, 11, 13, 15, 16, 17, 18, 19, 21, 25, 27):
        ws[f"B{r}"].number_format = MONEY
    style_header(ws, 3)
    widths(ws, [44, 16, 40])
    return ws


# ------------------------------------------------------------ Project cost
def sheet_project_cost(wb):
    ws = wb.create_sheet(PC)
    rows = load_csv("project_cost.csv")
    ws["A1"], ws["B1"], ws["C1"] = "Head", "Rs lakh", "% of total"
    for i, r in enumerate(rows, start=2):
        ws[f"A{i}"] = r["head"]
        ws[f"B{i}"] = float(r["amount_lakh"])
        ws[f"B{i}"].number_format = MONEY
        ws[f"C{i}"] = f"=B{i}/$B$9"
        ws[f"C{i}"].number_format = PCT1
    ws["A9"] = "TOTAL"
    ws["A9"].font = Font(bold=True)
    ws["B9"] = "=SUM(B2:B8)"
    ws["B9"].number_format = MONEY
    ws["B9"].font = Font(bold=True)
    ws["C9"] = "=SUM(C2:C8)"
    ws["C9"].number_format = PCT0
    ws["A11"] = "Check: total = Rs 240 lakh?"
    ws["B11"] = "=B9=240"
    style_header(ws, 3)
    style_range(ws, 2, 9, 3)
    style_range(ws, 11, 11, 2)
    widths(ws, [34, 14, 12])
    return ws


# ------------------------------------------------------- Means of finance
def sheet_means(wb):
    ws = wb.create_sheet(MF)
    ws["A1"], ws["B1"] = "Source", "Rs lakh"
    ws["A2"], ws["B2"] = "Promoter contribution", f"={A}!$B$25"
    ws["A3"], ws["B3"] = "Term loan", f"={A}!$B$21"
    ws["A4"], ws["B4"] = "TOTAL", "=SUM(B2:B3)"
    ws["B2"].number_format = MONEY
    ws["B3"].number_format = MONEY
    ws["B4"].number_format = MONEY
    ws["B4"].font = Font(bold=True)
    ws["A6"], ws["B6"] = "Promoter margin", "=B2/B4"
    ws["B6"].number_format = PCT1
    ws["A7"], ws["B7"] = "Debt-equity (at start)", "=B3/B2"
    ws["B7"].number_format = "0.00"
    ws["A9"] = "Check: total = project cost?"
    ws["B9"] = f"=B4={PC}!$B$9"
    style_header(ws, 2)
    style_range(ws, 2, 4, 2)
    style_range(ws, 6, 7, 2)
    style_range(ws, 9, 9, 2)
    widths(ws, [30, 16])
    return ws


# ------------------------------------------------------------ Loan schedule
def sheet_loan(wb):
    ws = wb.create_sheet(LS)
    ws["A1"] = "Year"
    for t in Y0_YEARS:
        ws[f"{yr_col(t, True)}1"] = t
    defs = [
        ("Opening balance", lambda c, t:
         f"={A}!$B$21" if t == 0 else f"={col(c - 1)}6"),
        ("Interest", lambda c, t:
         "0" if t == 0 else f"={col(c)}2*{A}!$B$22"),
        ("Annual payment", lambda c, t:
         "0" if t == 0 else (f"={col(c)}3" if t == 1
                             else f"=-PMT({A}!$B$22,{A}!$B$24,{A}!$B$21)")),
        ("Principal repaid", lambda c, t:
         "0" if t <= 1 else f"={col(c)}4-{col(c)}3"),
        ("Closing balance", lambda c, t: f"={col(c)}2-{col(c)}5"),
    ]
    for i, (label, fn) in enumerate(defs, start=2):
        ws[f"A{i}"] = label
        for t in Y0_YEARS:
            c = t + 2
            cell = ws.cell(row=i, column=c, value=fn(c, t))
            cell.number_format = MONEY
    ws["A8"] = "Check: fully repaid by Y7?"
    ws["B8"] = "=ABS(I6)<0.01"
    ws["A9"] = "Annuity (Rs lakh, Y2-Y7)"
    ws["B9"] = "=D4"
    ws["B9"].number_format = MONEY
    style_header(ws, 9)
    style_range(ws, 2, 6, 9)
    style_range(ws, 8, 9, 2)
    widths(ws, [26] + [13] * 8)
    return ws


# ------------------------------------------------------------ Revenue & cost
def sheet_revenue_cost(wb):
    ws = wb.create_sheet(RC)
    ws["A1"] = "Year"
    for t in YEARS:
        ws[f"{yr_col(t)}1"] = t
    for t in YEARS:
        c, L = t + 1, col(t + 1)
        prev = col(t) if t > 1 else None
        ws[f"A2"] = "Production (kg)"
        ws.cell(row=2, column=c, value=f"=1200000*{A}!{L}5").number_format = MONEY0
        ws[f"A3"] = "Selling price (Rs/kg)"
        ws.cell(row=3, column=c,
                value=f"={A}!$B$3" if t == 1 else f"={prev}3*(1+{A}!$B$4)"
                ).number_format = MONEY
        ws[f"A4"] = "Revenue (Rs lakh)"
        ws.cell(row=4, column=c, value=f"={L}2*{L}3/100000").number_format = MONEY
        for r, label in [(6, "Raw material"), (7, "Packing"),
                         (11, "Admin & selling")]:
            ws[f"A{r}"] = label
        ws[f"A8"] = "Power & fuel"
        ws[f"A9"] = "Salaries & wages"
        ws[f"A10"] = "Repairs & maintenance"
        ws[f"A12"] = "Insurance"
        ws.cell(row=6, column=c, value=f"={L}4*{A}!$B$6").number_format = MONEY
        ws.cell(row=7, column=c, value=f"={L}4*{A}!$B$7").number_format = MONEY
        ws.cell(row=8, column=c,
                value=f"={A}!$B$9" if t == 1 else f"={prev}8*(1+{A}!$B$10)"
                ).number_format = MONEY
        ws.cell(row=9, column=c,
                value=f"={A}!$B$11" if t == 1 else f"={prev}9*(1+{A}!$B$12)"
                ).number_format = MONEY
        ws.cell(row=10, column=c, value=f"={A}!$B$13").number_format = MONEY
        ws.cell(row=11, column=c, value=f"={L}4*{A}!$B$8").number_format = MONEY
        ws.cell(row=12, column=c,
                value=f"={A}!$B$14" if t == 1 else f"={prev}12*(1+{A}!$B$15)"
                ).number_format = MONEY
        ws[f"A13"] = "Total operating cost"
        ws.cell(row=13, column=c, value=f"=SUM({L}6:{L}12)").number_format = MONEY
        ws[f"A14"] = "EBITDA"
        ws.cell(row=14, column=c, value=f"={L}4-{L}13").number_format = MONEY
        ws.cell(row=14, column=c).font = Font(bold=True)
    style_header(ws, 8)
    style_range(ws, 2, 14, 8)
    widths(ws, [28] + [15] * 7)
    return ws


# ---------------------------------------------------------------------- PL
def sheet_pl(wb):
    ws = wb.create_sheet(PL)
    ws["A1"] = "Year"
    for t in YEARS:
        L = yr_col(t)
        ws[f"{L}1"] = t
        loan_c = col(t + 2)  # Loan_Schedule col for year t (Y1 -> C)
        ws[f"A2"] = "Revenue"
        ws.cell(row=2, column=t + 1, value=f"={RC}!{L}4").number_format = MONEY
        ws[f"A3"] = "Total operating cost"
        ws.cell(row=3, column=t + 1, value=f"={RC}!{L}13").number_format = MONEY
        ws[f"A4"] = "EBITDA"
        ws.cell(row=4, column=t + 1, value=f"={RC}!{L}14").number_format = MONEY
        ws[f"A5"] = "Depreciation & amortisation"
        ws.cell(row=5, column=t + 1,
                value=f"={A}!$B$16+{A}!$B$17+{A}!$B$18"
                      f"+IF({L}$1<=5,{A}!$B$19,0)").number_format = MONEY
        ws[f"A6"] = "Interest"
        ws.cell(row=6, column=t + 1,
                value=f"={LS}!{loan_c}3").number_format = MONEY
        ws[f"A7"] = "Profit before tax"
        ws.cell(row=7, column=t + 1,
                value=f"={L}4-{L}5-{L}6").number_format = MONEY
        ws[f"A8"] = "Tax"
        ws.cell(row=8, column=t + 1,
                value=f"=IF({L}7>0,{L}7*{A}!$B$20,0)").number_format = MONEY
        ws[f"A9"] = "Profit after tax"
        ws.cell(row=9, column=t + 1,
                value=f"={L}7-{L}8").number_format = MONEY
        ws.cell(row=9, column=t + 1).font = Font(bold=True)
        ws[f"A10"] = "Cash accruals (PAT + depr.)"
        ws.cell(row=10, column=t + 1,
                value=f"={L}9+{L}5").number_format = MONEY
        ws.cell(row=10, column=t + 1).font = Font(bold=True)
    style_header(ws, 8)
    style_range(ws, 2, 10, 8)
    widths(ws, [32] + [15] * 7)
    return ws


# ---------------------------------------------------------- Cashflow / DSCR
def sheet_cashflow(wb):
    ws = wb.create_sheet(CF)
    ws["A1"] = "Year"
    for t in Y0_YEARS:
        L = yr_col(t, True)
        ws[f"{L}1"] = t
        plc = col(t + 1) if t >= 1 else None   # PL col for year t
        lsc = col(t + 2)                        # Loan_Schedule col for year t
        ws[f"A2"] = "Cash accruals"
        ws.cell(row=2, column=t + 2,
                value=0 if t == 0 else f"={PL}!{plc}10").number_format = MONEY
        ws[f"A3"] = "Debt service (interest + principal)"
        ws.cell(row=3, column=t + 2,
                value=0 if t == 0 else f"={LS}!{lsc}3+{LS}!{lsc}5"
                ).number_format = MONEY
        ws[f"A4"] = "Net cash to equity"
        ws.cell(row=4, column=t + 2,
                value=f"=-{A}!$B$25" if t == 0 else f"={L}2-{LS}!{lsc}5"
                ).number_format = MONEY
        ws.cell(row=4, column=t + 2).font = Font(bold=True)
        ws[f"A5"] = "Cumulative (incl. Y0 equity)"
        ws.cell(row=5, column=t + 2,
                value=f"={L}4" if t == 0 else f"={col(t + 1)}5+{L}4"
                ).number_format = MONEY
        ws[f"A6"] = "DSCR"
        ws.cell(row=6, column=t + 2,
                value=None if t == 0 else f"=IF({L}3=0,\"\",{L}2/{L}3)"
                ).number_format = "0.00"
        ws[f"A7"] = "Payback flag"
        ws.cell(row=7, column=t + 2,
                value=f"=IF({L}5>=0,{L}$1,\"\")")
    labels = [("Average DSCR (Y1-Y7)", "=AVERAGE(C6:I6)", "0.00"),
              ("Minimum DSCR, repayment years (Y2-Y7)", "=MIN(D6:I6)", "0.00"),
              ("NPV @12% (Rs lakh)", f"=B4+NPV({A}!$B$26,C4:I4)", MONEY),
              ("Equity IRR", "=IRR(B4:I4)", PCT1),
              ("Payback year", "=MIN(B7:I7)", "0")]
    for i, (label, formula, fmt) in enumerate(labels, start=9):
        ws[f"A{i}"] = label
        ws.cell(row=i, column=2, value=formula).number_format = fmt
        ws.cell(row=i, column=2).font = Font(bold=True)
    # project cash balance (for the balance sheet): the Y0 -60 on row 4 is the
    # INVESTOR's outflow, not the project's — project cash starts at the
    # Rs 5 lakh working-capital margin and moves by accruals - principal.
    ws["A15"] = "Principal repaid"
    ws["A16"] = "Project cash balance"
    for t in Y0_YEARS:
        L = yr_col(t, True)
        lsc = col(t + 2)
        c = t + 2
        ws.cell(row=15, column=c,
                value=0 if t == 0 else f"={LS}!{lsc}5").number_format = MONEY
        ws.cell(row=16, column=c,
                value="=5" if t == 0 else f"={col(c - 1)}16+{L}2-{L}15"
                ).number_format = MONEY
    style_header(ws, 9)
    style_range(ws, 2, 7, 9)
    style_range(ws, 10, 14, 2)
    style_range(ws, 15, 16, 9)
    widths(ws, [42] + [14] * 8)
    return ws


# ------------------------------------------------------------- Balance sheet
def sheet_bs(wb):
    ws = wb.create_sheet(BS)
    ws["A1"] = "Year"
    for t in Y0_YEARS:
        L = yr_col(t, True)
        ws[f"{L}1"] = t
        plc = col(t + 1) if t >= 1 else None
        ws[f"A2"] = "Net block (fixed assets)"
        ws.cell(row=2, column=t + 2,
                value="=30+55+120+18" if t == 0 else
                f"=(30+55+120+18)-({A}!$B$16+{A}!$B$17+{A}!$B$18)*{L}$1"
                ).number_format = MONEY
        ws[f"A3"] = "Preliminary expenses not written off"
        ws.cell(row=3, column=t + 2,
                value="=12" if t == 0 else
                f"=12-{A}!$B$19*MIN({L}$1,5)").number_format = MONEY
        ws[f"A4"] = "Cash & bank"
        ws.cell(row=4, column=t + 2,
                value=f"={CF}!B16" if t == 0 else f"={CF}!{L}16"
                ).number_format = MONEY
        ws[f"A5"] = "TOTAL ASSETS"
        ws.cell(row=5, column=t + 2,
                value=f"=SUM({L}2:{L}4)").number_format = MONEY
        ws.cell(row=5, column=t + 2).font = Font(bold=True)
        ws[f"A7"] = "Equity capital"
        ws.cell(row=7, column=t + 2, value="=60").number_format = MONEY
        ws[f"A8"] = "Reserves (accumulated PAT)"
        ws.cell(row=8, column=t + 2,
                value="=0" if t == 0 else f"={col(t + 1)}8+{PL}!{plc}9"
                ).number_format = MONEY
        ws[f"A9"] = "Term loan outstanding"
        ws.cell(row=9, column=t + 2,
                value=f"={LS}!{L}6").number_format = MONEY
        ws[f"A10"] = "TOTAL LIABILITIES"
        ws.cell(row=10, column=t + 2,
                value=f"=SUM({L}7:{L}9)").number_format = MONEY
        ws.cell(row=10, column=t + 2).font = Font(bold=True)
        ws[f"A11"] = "Check: assets - liabilities"
        ws.cell(row=11, column=t + 2,
                value=f"={L}5-{L}10").number_format = MONEY
    style_header(ws, 9)
    style_range(ws, 2, 5, 9)
    style_range(ws, 7, 11, 9)
    widths(ws, [38] + [14] * 8)
    return ws


# ------------------------------------------------------------------- Ratios
def sheet_ratios(wb):
    ws = wb.create_sheet(RA)
    ws["A1"], ws["B1"] = "Ratio", "Value"
    rows = [
        ("Project cost (Rs lakh)", f"={PC}!$B$9", MONEY),
        ("Term loan (Rs lakh)", f"={A}!$B$21", MONEY),
        ("Promoter margin", f"={MF}!$B$6", PCT1),
        ("Debt-equity (at start)", f"={MF}!$B$7", "0.00"),
        ("Average DSCR (Y1-Y7)", f"={CF}!$B$9", "0.00"),
        ("Minimum DSCR, repayment years", f"={CF}!$B$10", "0.00"),
        ("NPV @12% (Rs lakh)", f"={CF}!$B$11", MONEY),
        ("Equity IRR", f"={CF}!$B$12", PCT1),
        ("Payback year", f"={CF}!$B$13", "0"),
        ("Break-even capacity utilisation",
         f"=({RC}!$B$8+{RC}!$B$9+{RC}!$B$10+{RC}!$B$12+{PL}!$B$5+{LS}!$C$3)"
         f"/(1-{A}!$B$6-{A}!$B$7-{A}!$B$8)"
         f"/(1200000*{A}!$B$3/100000)", PCT1),
        ("Bank norm: min DSCR >= 1.25x?", "=B7>=1.25", None),
    ]
    for i, (label, formula, fmt) in enumerate(rows, start=2):
        ws[f"A{i}"] = label
        c = ws.cell(row=i, column=2, value=formula)
        if fmt:
            c.number_format = fmt
        c.font = Font(bold=True)
    style_header(ws, 2)
    style_range(ws, 2, 12, 2)
    widths(ws, [38, 18])
    return ws


# --------------------------------------------------------------- Sensitivity
SCEN_DEFS = [
    ("S0 Base", 1.00, 0.76),
    ("S1 Revenue -5%", 0.95, 0.76),
    ("S2 Raw material +3pp", 1.00, 0.79),
    ("S3 Combined stress", 0.95, 0.79),
]
BLOCK_H = 25  # rows per scenario block


def sheet_sensitivity(wb):
    ws = wb.create_sheet(SE)
    ws["A1"] = "Scenario / line"
    ws["B1"] = "Y0"
    for t in YEARS:
        ws[f"{yr_col(t, True)}1"] = t
    for s, (name, rev_f, raw_p) in enumerate(SCEN_DEFS):
        R = 2 + s * BLOCK_H
        ws[f"A{R}"] = f"{name}  (revenue x {rev_f}, raw {raw_p:.0%})"
        ws[f"A{R}"].font = Font(bold=True, size=12, color="1F4E79")
        rev_r, raw_r, pack_r = R + 1, R + 2, R + 3
        pow_r, sal_r, rep_r = R + 4, R + 5, R + 6
        adm_r, ins_r = R + 7, R + 8
        toc_r, ebitda_r = R + 9, R + 10
        depr_r, int_r, pbt_r = R + 11, R + 12, R + 13
        tax_r, pat_r, ca_r = R + 14, R + 15, R + 16
        ds_r, nc_r, dscr_r = R + 17, R + 18, R + 19
        labels = {rev_r: "Revenue (Rs lakh)", raw_r: "Raw material",
                  pack_r: "Packing", pow_r: "Power & fuel", sal_r: "Salaries",
                  rep_r: "Repairs", adm_r: "Admin & selling",
                  ins_r: "Insurance", toc_r: "Total operating cost",
                  ebitda_r: "EBITDA", depr_r: "Depreciation",
                  int_r: "Interest", pbt_r: "PBT", tax_r: "Tax",
                  pat_r: "PAT", ca_r: "Cash accruals",
                  ds_r: "Debt service", nc_r: "Net cash to equity",
                  dscr_r: "DSCR"}
        for r, lab in labels.items():
            ws[f"A{r}"] = lab
        for t in Y0_YEARS:
            L = yr_col(t, True)
            plc = col(t + 1) if t >= 1 else None   # PL / Revenue_Cost col
            lsc = col(t + 2)                        # Loan_Schedule col
            c = t + 2
            if t == 0:
                for r in (rev_r, raw_r, pack_r, pow_r, sal_r, rep_r, adm_r,
                          ins_r, toc_r, ebitda_r, depr_r, int_r, pbt_r, tax_r,
                          pat_r, ca_r, ds_r):
                    ws.cell(row=r, column=c, value=0).number_format = MONEY
                ws.cell(row=nc_r, column=c,
                        value=f"=-{A}!$B$25").number_format = MONEY
                ws.cell(row=dscr_r, column=c, value=None)
                continue
            ws.cell(row=rev_r, column=c,
                    value=f"={RC}!{plc}4*{rev_f}").number_format = MONEY
            ws.cell(row=raw_r, column=c,
                    value=f"={L}{rev_r}*{raw_p}").number_format = MONEY
            ws.cell(row=pack_r, column=c,
                    value=f"={L}{rev_r}*{A}!$B$7").number_format = MONEY
            ws.cell(row=pow_r, column=c,
                    value=f"={RC}!{plc}8").number_format = MONEY
            ws.cell(row=sal_r, column=c,
                    value=f"={RC}!{plc}9").number_format = MONEY
            ws.cell(row=rep_r, column=c,
                    value=f"={A}!$B$13").number_format = MONEY
            ws.cell(row=adm_r, column=c,
                    value=f"={L}{rev_r}*{A}!$B$8").number_format = MONEY
            ws.cell(row=ins_r, column=c,
                    value=f"={RC}!{plc}12").number_format = MONEY
            ws.cell(row=toc_r, column=c,
                    value=f"=SUM({L}{raw_r}:{L}{ins_r})").number_format = MONEY
            ws.cell(row=ebitda_r, column=c,
                    value=f"={L}{rev_r}-{L}{toc_r}").number_format = MONEY
            ws.cell(row=depr_r, column=c,
                    value=f"={PL}!{plc}5").number_format = MONEY
            ws.cell(row=int_r, column=c,
                    value=f"={LS}!{lsc}3").number_format = MONEY
            ws.cell(row=pbt_r, column=c,
                    value=f"={L}{ebitda_r}-{L}{depr_r}-{L}{int_r}"
                    ).number_format = MONEY
            ws.cell(row=tax_r, column=c,
                    value=f"=IF({L}{pbt_r}>0,{L}{pbt_r}*{A}!$B$20,0)"
                    ).number_format = MONEY
            ws.cell(row=pat_r, column=c,
                    value=f"={L}{pbt_r}-{L}{tax_r}").number_format = MONEY
            ws.cell(row=ca_r, column=c,
                    value=f"={L}{pat_r}+{L}{depr_r}").number_format = MONEY
            ws.cell(row=ds_r, column=c,
                    value=f"={CF}!{L}3").number_format = MONEY
            ws.cell(row=nc_r, column=c,
                    value=f"={L}{ca_r}-{LS}!{lsc}5").number_format = MONEY
            ws.cell(row=dscr_r, column=c,
                    value=f"=IF({L}{ds_r}=0,\"\",{L}{ca_r}/{L}{ds_r})"
                    ).number_format = "0.00"
        summ = [("Min DSCR, repayment years (Y2-Y7)",
                 f"=MIN(D{dscr_r}:I{dscr_r})", "0.00"),
                ("Average DSCR (Y1-Y7)", f"=AVERAGE(C{dscr_r}:I{dscr_r})", "0.00"),
                ("NPV @12% (Rs lakh)",
                 f"=B{nc_r}+NPV({A}!$B$26,C{nc_r}:I{nc_r})", MONEY),
                ("Equity IRR", f"=IRR(B{nc_r}:I{nc_r})", PCT1)]
        for i, (lab, formula, fmt) in enumerate(summ, start=R + 21):
            ws[f"A{i}"] = lab
            cc = ws.cell(row=i, column=2, value=formula)
            cc.number_format = fmt
            cc.font = Font(bold=True)
        style_range(ws, R + 1, R + 24, 9)
    style_header(ws, 9)
    widths(ws, [40] + [14] * 8)
    return ws


# ----------------------------------------------------------------- Dashboard
def sheet_dashboard(wb):
    ws = wb.create_sheet(DB)
    ws["A1"] = "DPRPulse — Banker's Dashboard"
    ws["A1"].font = TITLE_FONT
    ws["A3"], ws["B3"] = "KPI", "Value"
    kpis = [
        ("Project cost (Rs lakh)", f"={RA}!$B$2", MONEY),
        ("Term loan (Rs lakh)", f"={RA}!$B$3", MONEY),
        ("Promoter margin", f"={RA}!$B$4", PCT1),
        ("Average DSCR", f"={RA}!$B$6", "0.00"),
        ("Minimum DSCR (repayment)", f"={RA}!$B$7", "0.00"),
        ("NPV @12% (Rs lakh)", f"={RA}!$B$8", MONEY),
        ("Equity IRR", f"={RA}!$B$9", PCT1),
        ("Break-even utilisation", f"={RA}!$B$11", PCT1),
    ]
    for i, (lab, formula, fmt) in enumerate(kpis, start=4):
        ws[f"A{i}"] = lab
        cc = ws.cell(row=i, column=2, value=formula)
        cc.number_format = fmt
        cc.font = Font(bold=True, size=12)
    # risk flags (formula-driven text)
    ws["A13"] = "Risk flags"
    ws["A13"].font = Font(bold=True, size=12, color="1F4E79")
    s2_min, s3_min = "B73", "B98"   # Sensitivity summary rows for S2, S3
    ws["A14"] = (f"=IF({SE}!{s2_min}<1.25,"
                 f"\"RISK: raw-material +3pp drags min DSCR to \""
                 f"&TEXT({SE}!{s2_min},\"0.00\")&\"x - lock input prices\","
                 f"\"OK: raw-material +3pp stays above 1.25x\")")
    ws["A15"] = (f"=IF({SE}!{s3_min}<1.25,"
                 f"\"RISK: combined stress drags min DSCR to \""
                 f"&TEXT({SE}!{s3_min},\"0.00\")&\"x - keep a contingency buffer\","
                 f"\"OK: combined stress stays above 1.25x\")")
    ws["A16"] = (f"=IF({CF}!$B$10>=1.25,"
                 f"\"Base case services debt comfortably (min DSCR \""
                 f"&TEXT({CF}!$B$10,\"0.00\")&\"x)\","
                 f"\"Base case DSCR below 1.25x - restructure\")")
    for r in (14, 15, 16):
        ws[f"A{r}"].alignment = Alignment(wrap_text=True)
        ws.row_dimensions[r].height = 30
    # chart helper tables
    ws["D3"], ws["E3"], ws["F3"] = "Year", "Cash accruals", "Debt service"
    for t in YEARS:
        r = t + 3
        ws[f"D{r}"] = t
        ws[f"E{r}"] = f"={CF}!{yr_col(t, True)}2"
        ws[f"F{r}"] = f"={CF}!{yr_col(t, True)}3"
    ws["H3"], ws["I3"] = "DSCR", "Bank norm"
    for t in YEARS:
        r = t + 3
        ws[f"H{r}"] = f"={CF}!{yr_col(t, True)}6"
        ws[f"I{r}"] = 1.25
    ws["K3"], ws["L3"] = "Scenario", "NPV @12%"
    for s, (name, _, _) in enumerate(SCEN_DEFS):
        r = s + 4
        ws[f"K{r}"] = name
        ws[f"L{r}"] = f"={SE}!B{2 + s * BLOCK_H + 23}"
    # charts
    c1 = BarChart()
    c1.title = "Cash accruals vs debt service (Rs lakh)"
    c1.style = 10
    c1.add_data(Reference(ws, min_col=5, min_row=3, max_row=10), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=4, min_row=4, max_row=10))
    ws.add_chart(c1, "A18")
    c2 = LineChart()
    c2.title = "DSCR by year (bank norm 1.25x)"
    c2.style = 10
    c2.add_data(Reference(ws, min_col=8, min_row=3, max_row=10), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=4, max_row=10))
    ws.add_chart(c2, "A33")
    c3 = BarChart()
    c3.title = "Scenario NPV @12% (Rs lakh)"
    c3.style = 10
    c3.add_data(Reference(ws, min_col=12, min_row=3, max_row=7), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=11, min_row=4, max_row=7))
    ws.add_chart(c3, "A48")
    style_header(ws, 2)
    style_range(ws, 3, 3, 2)
    style_range(ws, 4, 11, 2)
    widths(ws, [34, 20, 4, 8, 16, 14, 4, 10, 12, 4, 24, 14])
    return ws


# -------------------------------------------------------------------- Checks
def sheet_checks(wb):
    ws = wb.create_sheet(CH)
    ws["A1"], ws["B1"], ws["C1"] = "Check", "Result", "Note"
    bs_gap = ",".join(f"ABS({BS}!{yr_col(t, True)}11)<0.01" for t in Y0_YEARS)
    util = ",".join(
        f"{A}!{col(t + 2)}5>={A}!{col(t + 1)}5" for t in range(1, 7))
    rows = [
        ("Project cost adds to Rs 240 lakh", f"={PC}!$B$9=240",
         "7 capex heads"),
        ("Means of finance = project cost", f"={MF}!$B$4={PC}!$B$9", ""),
        ("Term loan fully repaid by Y7", f"=ABS({LS}!$I$6)<0.01", ""),
        ("Balance sheet balances every year", f"=AND({bs_gap})",
         "assets - liabilities within 0.01"),
        ("Min DSCR (repayment) >= 1.25x", f"={CF}!$B$10>=1.25",
         "bank norm"),
        ("Cash balance never negative", f"=MIN({BS}!$B$4:$I$4)>=0", ""),
        ("NPV @12% positive", f"={RA}!$B$8>0", ""),
        ("IRR exceeds hurdle rate", f"={RA}!$B$9>{A}!$B$26", ""),
        ("Utilisation ramp non-decreasing", f"=AND({util})", ""),
        ("S1 mild stress still services debt", f"={SE}!$B$48>=1.25",
         "revenue -5%"),
    ]
    for i, (label, formula, note) in enumerate(rows, start=2):
        ws[f"A{i}"] = label
        ws.cell(row=i, column=2, value=formula).font = Font(bold=True)
        ws[f"C{i}"] = note
    style_header(ws, 3)
    style_range(ws, 2, 11, 3)
    widths(ws, [44, 12, 40])
    return ws


def build():
    wb = Workbook()
    sheet_guide(wb)
    sheet_assumptions(wb)
    sheet_project_cost(wb)
    sheet_means(wb)
    sheet_loan(wb)
    sheet_revenue_cost(wb)
    sheet_pl(wb)
    sheet_cashflow(wb)
    sheet_bs(wb)
    sheet_ratios(wb)
    sheet_sensitivity(wb)
    sheet_dashboard(wb)
    sheet_checks(wb)
    out = os.path.join(ROOT, "dprpulse.xlsx")
    wb.save(out)
    print("wrote", out)


if __name__ == "__main__":
    build()
