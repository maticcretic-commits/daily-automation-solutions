#!/usr/bin/env python3
"""LenderMatch — manufacturing-acquisition lender comparison pack (Excel).

Built around a real Freelancer brief ("Loan Sourcing for Manufacturing
Buyout", project 40804560): a curated list of non-traditional, asset-based
and SBA 7(a) lenders for manufacturing acquisitions, with terms and a
comparison matrix.

What this builds: the comparison matrix itself, as a formula-driven
workbook. The buyer enters one acquisition scenario (price, equity,
revenue, EBITDA, credit score, other debt); every lender is screened for
eligibility, priced with a live PMT payment, scored on a documented
weighted model, and ranked. The Dashboard names the best pick and the
cheapest payment; the Checks sheet proves the arithmetic.

Demo scenario: a fictional $2,400,000 manufacturing acquisition, 20%
buyer equity, $1,920,000 target loan, $6,500,000 target revenue,
$145,000 monthly EBITDA, 720 credit score. All 14 lenders are fictional
and deterministic (data/lenders.csv).

Sheets: Guide, Lenders, Scenario, Comparison, Dashboard (11 KPIs + 3
native charts), Checks. No macros, no structured table references, plain
A1 ranges with the column letter repeated on both ends (LibreOffice-safe).

Model rules (all visible in the formulas):
- A lender is ELIGIBLE when: target loan <= max loan AND target loan <=
  acquisition price x LTV AND target revenue >= min revenue AND buyer
  credit score >= min credit.
- Mid rate = (min rate + max rate) / 2; monthly payment = PMT(mid rate /
  12, tenor x 12, -target loan); total interest = payment x months - loan.
- DSCR = monthly EBITDA / (monthly payment + other monthly debt).
- Weighted score (0-100, eligible lenders only): rate 35, closing speed
  20, LTV 20, origination fee 15, tenor 10. Each component is normalised
  0-1 across all 14 lenders (lower rate / faster close / lower fee score
  higher; higher LTV / longer tenor score higher). Ineligible lenders
  score 0 and rank blank.
- Rates are stored as 0-1 fractions, never 0-100 points.
"""
import csv
import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
AS_OF = datetime.date(2026, 11, 8)
OUT = BASE / "lendermatch.xlsx"

NAVY = "1F3864"
TEAL = "0F7B7A"
LIGHT = "D9E2F3"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color=NAVY, size=14)
SECTION_FONT = Font(bold=True, color=NAVY, size=11)
MONEY = "$#,##0"
MONEY2 = "$#,##0.00"
PCT2 = "0.00%"
PCT0 = "0%"
SCORE = "0.0"
DSCR_FMT = "0.00"
DAYS_FMT = "0"

# Scenario inputs (yellow cells the buyer changes).
SCENARIO = {
    "price": 2_400_000,
    "equity_pct": 0.20,
    "revenue": 6_500_000,
    "ebitda_monthly": 145_000,
    "credit": 720,
    "other_debt_monthly": 12_000,
}


def read_lenders():
    with open(BASE / "data" / "lenders.csv", newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def build():
    lenders = read_lenders()
    n = len(lenders)
    last = n + 1  # last data row (row 1 is the header)
    wb = Workbook()

    # ---------------------------------------------------------------- Guide
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "LenderMatch — Manufacturing Acquisition Lender Comparison Pack"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "What this is"
    ws["A3"].font = SECTION_FONT
    ws["A4"] = ("A comparison matrix for financing a manufacturing acquisition: "
                "14 lenders (SBA 7(a), asset-based, non-traditional, conventional) "
                "screened against one acquisition scenario, priced, scored and ranked. "
                "Built around the Freelancer brief 'Loan Sourcing for Manufacturing "
                "Buyout' (project 40804560) — a curated lender list with terms and a "
                "comparison matrix. All lenders and the scenario are fictional demo "
                "data; replace them with real quotes before deciding.")
    ws["A4"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A4:F4")
    ws.row_dimensions[4].height = 60
    ws["A6"] = "How to use it"
    ws["A6"].font = SECTION_FONT
    steps = [
        "1. Scenario sheet — change the YELLOW cells: acquisition price, buyer equity %, "
        "target revenue, monthly EBITDA, credit score, other monthly debt. Everything "
        "recalculates.",
        "2. Lenders sheet — the lender table. Replace rows with real lender quotes "
        "(keep the column order); the Comparison sheet picks them up automatically.",
        "3. Comparison sheet — eligibility, monthly payment (PMT at the mid rate), total "
        "interest, upfront fees, DSCR and the weighted score + rank for every lender.",
        "4. Dashboard sheet — the best pick, the cheapest payment, the fastest close, "
        "and the rate / payment / score charts.",
        "5. Checks sheet — 10 in-workbook consistency checks; all must read PASS.",
    ]
    for i, s in enumerate(steps):
        ws.cell(row=7 + i, column=1, value=s)
        ws.cell(row=7 + i, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=7 + i, start_column=1, end_row=7 + i, end_column=6)
        ws.row_dimensions[7 + i].height = 30
    ws["A13"] = "Scoring model (weights, 0-100)"
    ws["A13"].font = SECTION_FONT
    ws["A14"] = ("Rate 35  |  Closing speed 20  |  LTV 20  |  Origination fee 15  |  "
                 "Tenor 10. Each component is normalised 0-1 across all lenders in the "
                 "table. Lower rate, faster close and lower fee score higher; higher "
                 "LTV and longer tenor score higher. Ineligible lenders score 0.")
    ws["A14"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A14:F14")
    ws.row_dimensions[14].height = 45
    ws["A16"] = "Assumptions"
    ws["A16"].font = SECTION_FONT
    ws["A17"] = ("Payments use each lender's MID rate (average of min and max) over the "
                 "lender's own tenor; origination fees are shown separately and are NOT "
                 "rolled into the payment. DSCR = monthly EBITDA / (loan payment + other "
                 "monthly debt). Eligibility needs all four tests to pass (max loan, LTV "
                 "headroom, min revenue, min credit). As-of date: 2026-11-08.")
    ws["A17"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A17:F17")
    ws.row_dimensions[17].height = 45
    ws["A19"] = "Sheets: Lenders (source table) · Scenario (inputs) · Comparison (matrix) · Dashboard (KPIs + charts) · Checks (10 PASS/FAIL)"
    ws["A19"].font = Font(italic=True, color="404040")
    ws.merge_cells("A19:F19")
    widths(ws, {"A": 28, "B": 22, "C": 22, "D": 22, "E": 22, "F": 22})

    # -------------------------------------------------------------- Lenders
    ws = wb.create_sheet("Lenders")
    headers = ["Lender", "Type", "Max loan $", "Min rate", "Max rate", "LTV",
               "Tenor (yrs)", "Orig. fee", "Close (days)", "Min revenue $",
               "Min credit", "Collateral", "Prepay penalty"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, len(headers))
    ws.row_dimensions[1].height = 30
    for i, r in enumerate(lenders, start=2):
        ws.cell(row=i, column=1, value=r["lender"])
        ws.cell(row=i, column=2, value=r["type"])
        ws.cell(row=i, column=3, value=int(r["max_loan"])).number_format = MONEY
        ws.cell(row=i, column=4, value=float(r["min_rate"])).number_format = PCT2
        ws.cell(row=i, column=5, value=float(r["max_rate"])).number_format = PCT2
        ws.cell(row=i, column=6, value=float(r["ltv"])).number_format = PCT0
        ws.cell(row=i, column=7, value=int(r["tenor_yrs"]))
        ws.cell(row=i, column=8, value=float(r["orig_fee"])).number_format = PCT2
        ws.cell(row=i, column=9, value=int(r["close_days"]))
        ws.cell(row=i, column=10, value=int(r["min_revenue"])).number_format = MONEY
        ws.cell(row=i, column=11, value=int(r["min_credit"]))
        ws.cell(row=i, column=12, value=r["collateral"])
        ws.cell(row=i, column=13, value=r["prepay_penalty"])
    widths(ws, {"A": 32, "B": 16, "C": 14, "D": 10, "E": 10, "F": 8, "G": 11,
                "H": 10, "I": 12, "J": 15, "K": 11, "L": 34, "M": 14})

    # ------------------------------------------------------------- Scenario
    ws = wb.create_sheet("Scenario")
    ws["A1"] = "Acquisition scenario — change the YELLOW cells"
    ws["A1"].font = TITLE_FONT
    rows = [
        ("Acquisition price $", SCENARIO["price"], MONEY,
         "Purchase price of the target manufacturing business.", True),
        ("Buyer equity %", SCENARIO["equity_pct"], PCT0,
         "Cash the buyer puts in; the rest is the target loan.", True),
        ("Target loan $", "=B3*(1-B4)", MONEY,
         "Loan required = price x (1 - equity %).", False),
        ("Buyer equity $", "=B3*B4", MONEY,
         "Cash required from the buyer.", False),
        ("Target annual revenue $", SCENARIO["revenue"], MONEY,
         "Target business revenue — lenders screen on a minimum.", True),
        ("Monthly EBITDA $", SCENARIO["ebitda_monthly"], MONEY,
         "Target business monthly EBITDA — feeds DSCR.", True),
        ("Buyer credit score", SCENARIO["credit"], "0",
         "Buyer credit score — lenders screen on a minimum.", True),
        ("Other monthly debt $", SCENARIO["other_debt_monthly"], MONEY,
         "Existing monthly debt service — feeds DSCR.", True),
    ]
    for i, (label, val, fmt, note, is_input) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=val)
        c.number_format = fmt
        if is_input:
            c.fill = INPUT_FILL
        ws.cell(row=i, column=3, value=note)
    ws["A12"] = "As-of date"
    ws["A12"].font = Font(bold=True)
    ws.cell(row=12, column=2, value=AS_OF).number_format = "yyyy-mm-dd"
    widths(ws, {"A": 26, "B": 20, "C": 62})

    # ----------------------------------------------------------- Comparison
    ws = wb.create_sheet("Comparison")
    headers = ["Lender", "Type", "Eligible", "Mid rate", "Monthly payment $",
               "Total interest $", "Upfront fees $", "DSCR", "Close (days)",
               "Score (0-100)", "Rank",
               "Rate#", "Pay#", "Int#", "Close#", "LTV#"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, len(headers))
    ws.row_dimensions[1].height = 30
    for i in range(2, last + 1):
        L = f"Lenders!A{i}"
        ws.cell(row=i, column=1, value=f"={L}")
        ws.cell(row=i, column=2, value=f"=Lenders!B{i}")
        # Eligibility: max loan, LTV headroom, min revenue, min credit.
        ws.cell(row=i, column=3, value=(
            f"=IF(AND(Scenario!$B$5<=Lenders!C{i},"
            f"Scenario!$B$5<=Scenario!$B$3*Lenders!F{i},"
            f"Scenario!$B$7>=Lenders!J{i},Scenario!$B$9>=Lenders!K{i}),"
            f"\"Yes\",\"No\")"))
        ws.cell(row=i, column=4, value=f"=(Lenders!D{i}+Lenders!E{i})/2"
               ).number_format = PCT2
        ws.cell(row=i, column=5, value=(
            f"=IF($C{i}=\"Yes\",PMT($D{i}/12,Lenders!G{i}*12,-Scenario!$B$5),\"\")"
        )).number_format = MONEY2
        ws.cell(row=i, column=6, value=(
            f"=IF($C{i}=\"Yes\",$E{i}*Lenders!G{i}*12-Scenario!$B$5,\"\")"
        )).number_format = MONEY
        ws.cell(row=i, column=7, value=(
            f"=IF($C{i}=\"Yes\",Scenario!$B$5*Lenders!H{i},\"\")"
        )).number_format = MONEY
        ws.cell(row=i, column=8, value=(
            f"=IF($C{i}=\"Yes\",Scenario!$B$8/($E{i}+Scenario!$B$10),\"\")"
        )).number_format = DSCR_FMT
        ws.cell(row=i, column=9, value=f"=Lenders!I{i}").number_format = DAYS_FMT
        # Weighted score: rate 35, speed 20, LTV 20, fee 15, tenor 10.
        ws.cell(row=i, column=10, value=(
            f"=IF($C{i}=\"Yes\","
            f"35*(MAX($D$2:$D${last})-$D{i})/(MAX($D$2:$D${last})-MIN($D$2:$D${last}))"
            f"+20*(MAX(Lenders!$I$2:$I${last})-Lenders!I{i})/(MAX(Lenders!$I$2:$I${last})-MIN(Lenders!$I$2:$I${last}))"
            f"+20*(Lenders!F{i}-MIN(Lenders!$F$2:$F${last}))/(MAX(Lenders!$F$2:$F${last})-MIN(Lenders!$F$2:$F${last}))"
            f"+15*(MAX(Lenders!$H$2:$H${last})-Lenders!H{i})/(MAX(Lenders!$H$2:$H${last})-MIN(Lenders!$H$2:$H${last}))"
            f"+10*(Lenders!G{i}-MIN(Lenders!$G$2:$G${last}))/(MAX(Lenders!$G$2:$G${last})-MIN(Lenders!$G$2:$G${last}))"
            f",0)"
        )).number_format = SCORE
        ws.cell(row=i, column=11, value=(
            f"=IF($C{i}=\"Yes\",RANK($J{i},$J$2:$J${last},0),\"\")"))
        # Sentinel helpers for LO-safe MIN/MAX aggregates ("" never feeds maths).
        ws.cell(row=i, column=12, value=f"=IF($C{i}=\"Yes\",$D{i},1)"
               ).number_format = PCT2
        ws.cell(row=i, column=13, value=f"=IF($C{i}=\"Yes\",$E{i},999999)"
               ).number_format = MONEY2
        ws.cell(row=i, column=14, value=f"=IF($C{i}=\"Yes\",$F{i},999999999)"
               ).number_format = MONEY
        ws.cell(row=i, column=15, value=f"=IF($C{i}=\"Yes\",$I{i},999)"
               ).number_format = DAYS_FMT
        ws.cell(row=i, column=16, value=f"=Lenders!F{i}").number_format = PCT0
    # Hide the helper columns from the printed matrix (still live for formulas).
    for col in ("L", "M", "N", "O", "P"):
        ws.column_dimensions[col].hidden = True
    widths(ws, {"A": 32, "B": 16, "C": 10, "D": 10, "E": 17, "F": 16, "G": 15,
                "H": 8, "I": 12, "J": 13, "K": 7})

    # ------------------------------------------------------------ Dashboard
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "LenderMatch — Dashboard"
    ws["A1"].font = TITLE_FONT
    kpis = [
        ("Eligible lenders", f"=COUNTIF(Comparison!$C$2:$C${last},\"Yes\")", "0",
         "of 14 lenders pass all four screens"),
        ("Best weighted score", f"=SUMPRODUCT(MAX((Comparison!$C$2:$C${last}=\"Yes\")*Comparison!$J$2:$J${last}))", SCORE,
         "0-100; weights rate 35 / speed 20 / LTV 20 / fee 15 / tenor 10"),
        ("Best pick", f"=INDEX(Comparison!$A$2:$A${last},MATCH($B$4,Comparison!$J$2:$J${last},0))", "@",
         "highest weighted score among eligible lenders"),
        ("Lowest mid rate", "=MIN(Comparison!$L$2:$L$15)", PCT2,
         "among eligible lenders"),
        ("Lowest monthly payment", "=MIN(Comparison!$M$2:$M$15)", MONEY2,
         "PMT at mid rate over the lender's tenor"),
        ("Fastest close (days)", "=MIN(Comparison!$O$2:$O$15)", DAYS_FMT,
         "among eligible lenders"),
        ("Highest LTV", f"=SUMPRODUCT(MAX((Comparison!$C$2:$C${last}=\"Yes\")*Comparison!$P$2:$P${last}))", PCT0,
         "among eligible lenders"),
        ("Total interest — best pick", f"=INDEX(Comparison!$F$2:$F${last},MATCH($B$4,Comparison!$J$2:$J${last},0))", MONEY,
         "payment x months - loan, at the best pick's terms"),
        ("DSCR — best pick", f"=INDEX(Comparison!$H$2:$H${last},MATCH($B$4,Comparison!$J$2:$J${last},0))", DSCR_FMT,
         "monthly EBITDA / (payment + other debt)"),
        ("Upfront fees — best pick", f"=INDEX(Comparison!$G$2:$G${last},MATCH($B$4,Comparison!$J$2:$J${last},0))", MONEY,
         "origination fee on the target loan"),
        ("Target loan $", "=Scenario!B5", MONEY,
         "acquisition price x (1 - equity %)"),
    ]
    # NOTE: B3 = eligible count, B4 = best score, B5 = best-pick name.
    for i, (label, formula, fmt, note) in enumerate(kpis, start=3):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=formula)
        c.number_format = fmt
        c.font = Font(bold=True, size=12)
        ws.cell(row=i, column=3, value=note)
    widths(ws, {"A": 28, "B": 24, "C": 64})

    # Charts (native openpyxl charts over the Comparison matrix).
    comp = wb["Comparison"]
    cats = Reference(comp, min_col=1, min_row=2, max_row=last)
    ch_score = BarChart()
    ch_score.type = "col"
    ch_score.title = "Weighted score by lender (0-100)"
    ch_score.y_axis.title = "Score"
    ch_score.add_data(Reference(comp, min_col=10, min_row=1, max_row=last), titles_from_data=True)
    ch_score.set_categories(cats)
    ch_score.shape = 4
    ch_score.height = 7.5
    ch_score.width = 15
    ws.add_chart(ch_score, "A15")
    ch_pay = BarChart()
    ch_pay.type = "col"
    ch_pay.title = "Monthly payment by lender ($)"
    ch_pay.y_axis.title = "$ / month"
    ch_pay.add_data(Reference(comp, min_col=5, min_row=1, max_row=last), titles_from_data=True)
    ch_pay.set_categories(cats)
    ch_pay.shape = 4
    ch_pay.height = 7.5
    ch_pay.width = 15
    ws.add_chart(ch_pay, "A31")
    ch_rate = BarChart()
    ch_rate.type = "col"
    ch_rate.title = "Mid rate by lender"
    ch_rate.y_axis.title = "Rate"
    ch_rate.add_data(Reference(comp, min_col=4, min_row=1, max_row=last), titles_from_data=True)
    ch_rate.set_categories(cats)
    ch_rate.shape = 4
    ch_rate.height = 7.5
    ch_rate.width = 15
    ws.add_chart(ch_rate, "A47")

    # --------------------------------------------------------------- Checks
    ws = wb.create_sheet("Checks")
    ws["A1"] = "Checks — all must read PASS"
    ws["A1"].font = TITLE_FONT
    checks = [
        ("Loan = price x (1 - equity %)",
         "=IF(Scenario!B5=Scenario!B3*(1-Scenario!B4),\"PASS\",\"FAIL\")"),
        ("Eligible count = COUNTIF on Comparison",
         f"=IF(Dashboard!B3=COUNTIF(Comparison!$C$2:$C${last},\"Yes\"),\"PASS\",\"FAIL\")"),
        ("Best pick's score = MAX score in the matrix",
         f"=IF(INDEX(Comparison!$J$2:$J${last},MATCH(Dashboard!B5,Comparison!$A$2:$A${last},0))=MAX(Comparison!$J$2:$J${last}),\"PASS\",\"FAIL\")"),
        ("No eligible score above 100 or below 0",
         f"=IF(AND(COUNTIFS(Comparison!$C$2:$C${last},\"Yes\",Comparison!$J$2:$J${last},\">100\")=0,COUNTIFS(Comparison!$C$2:$C${last},\"Yes\",Comparison!$J$2:$J${last},\"<0\")=0),\"PASS\",\"FAIL\")"),
        ("Ineligible lenders score 0",
         f"=IF(COUNTIFS(Comparison!$C$2:$C${last},\"No\",Comparison!$J$2:$J${last},\"<>0\")=0,\"PASS\",\"FAIL\")"),
        ("Ineligible payment sentinels intact (999999 each)",
         f"=IF(SUMPRODUCT((Comparison!$C$2:$C${last}=\"No\")*Comparison!$M$2:$M${last})=999999*COUNTIF(Comparison!$C$2:$C${last},\"No\"),\"PASS\",\"FAIL\")"),
        ("Best pick total interest = payment x months - loan",
         f"=IF(ABS(INDEX(Comparison!$F$2:$F${last},MATCH(Dashboard!B5,Comparison!$A$2:$A${last},0))-(INDEX(Comparison!$E$2:$E${last},MATCH(Dashboard!B5,Comparison!$A$2:$A${last},0))*INDEX(Lenders!$G$2:$G${last},MATCH(Dashboard!B5,Lenders!$A$2:$A${last},0))*12-Scenario!$B$5))<1,\"PASS\",\"FAIL\")"),
        ("Best pick DSCR = EBITDA / (payment + other debt)",
         f"=IF(ABS(INDEX(Comparison!$H$2:$H${last},MATCH(Dashboard!B5,Comparison!$A$2:$A${last},0))-(Scenario!$B$8/(INDEX(Comparison!$E$2:$E${last},MATCH(Dashboard!B5,Comparison!$A$2:$A${last},0))+Scenario!$B$10)))<0.01,\"PASS\",\"FAIL\")"),
        ("All rates stored as 0-1 fractions",
         "=IF(AND(MAX(Lenders!D2:D15)<1,MAX(Lenders!E2:E15)<1,MIN(Lenders!D2:D15)>0,MIN(Lenders!E2:E15)>0),\"PASS\",\"FAIL\")"),
        ("Rank 1 = best pick",
         f"=IF(INDEX(Comparison!$A$2:$A${last},MATCH(1,Comparison!$K$2:$K${last},0))=Dashboard!B5,\"PASS\",\"FAIL\")"),
    ]
    for i, (label, formula) in enumerate(checks, start=3):
        ws.cell(row=i, column=1, value=f"{i - 2}. {label}")
        ws.cell(row=i, column=2, value=formula).font = Font(bold=True)
    widths(ws, {"A": 58, "B": 10})

    wb.save(OUT)
    print(f"Built {OUT} — {n} lenders, {len(checks)} checks")


if __name__ == "__main__":
    build()
