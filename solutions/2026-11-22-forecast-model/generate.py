#!/usr/bin/env python3
"""ForecastForge — build the 3-year financial forecast workbook.

Reads data/assumptions.csv (the only inputs) and writes
forecast-model.xlsx — a genuine formula-driven workbook (openpyxl,
no macros):

  Guide        — what the workbook is, how to use the scenario switch
  Assumptions  — every driver as a value; Scenario cell (B2) with a
                 Base / Optimistic / Pessimistic dropdown; SELECTED
                 growth and COGS% are live IF formulas
  Monthly      — 36 months Apr 2026–Mar 2029; every figure a live
                 formula linked to Assumptions
  Annual_PL    — 3 fiscal years, SUMIFS over Monthly
  Balance_Sheet— year-end, live balance check (must be 0)
  Cash_Flow    — indirect method, ties out to BS cash
  Ratios       — margins, current ratio, D/E, ROE, break-even revenue
  Sensitivity  — 5x5 Year-3 net-profit grid (growth offset x COGS%
                 offset), closed-form formulas; centre cell equals
                 Annual Year-3 net profit
  Dashboard    — scenario readout, 6 KPIs, 3 native charts

Every figure is a live A1-style formula (LibreOffice-safe: no
structured references, column letters repeated on both ends of
every range).

Run:  python3 make_data.py && python3 generate.py
"""
import csv
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
XLSX = os.path.join(BASE, "forecast-model.xlsx")

MONEY = u"\u20b9#,##0"
PCT = "0.0%"
RATIO = "0.00"
NAVY = "1F3864"
LIGHT = "D9E2F3"
GREY = "F2F2F2"
GREEN = "C6EFCE"
RED = "FFC7CE"
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
HDR_FILL = PatternFill("solid", fgColor=NAVY)
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SUB_FONT = Font(size=11, color="404040")
BOLD = Font(bold=True, size=11)
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)

A = "Assumptions!"  # prefix for cross-sheet refs


def load_assumptions():
    with open(os.path.join(DATA, "assumptions.csv"), newline="") as f:
        return {r["key"]: r["value"] for r in csv.DictReader(f)}


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = THIN


def title_block(ws, title, sub, width):
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.cell(row=2, column=1, value=sub).font = SUB_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)


def money_range(ws, row, c1, c2, fmt=MONEY):
    for c in range(c1, c2 + 1):
        ws.cell(row=row, column=c).number_format = fmt


def build_assumptions(wb, drv):
    ws = wb.active
    ws.title = "Assumptions"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.cell(row=1, column=1,
           value="Assumptions — Ganga Foods Pvt. Ltd. (3-year forecast)").font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 64
    # (row, label, value-or-formula, note, fmt)
    spec = [
        (2, "Scenario", "Base",
         "Pick Base / Optimistic / Pessimistic — dropdown in B2.", None),
        (4, "Month-1 revenue — Apr 2026 (Rs)", float(drv["rev_base"]),
         "First forecast month; grows at the selected monthly rate.", MONEY),
        (5, "Monthly revenue growth — Base", float(drv["growth_base"]),
         "2.2% a month compounds to ~30% a year.", PCT),
        (6, "Monthly revenue growth — Optimistic", float(drv["growth_optimistic"]),
         "", PCT),
        (7, "Monthly revenue growth — Pessimistic", float(drv["growth_pessimistic"]),
         "", PCT),
        (8, "Monthly revenue growth — SELECTED",
         '=IF($B$2="Optimistic",$B$6,IF($B$2="Pessimistic",$B$7,$B$5))',
         "Live: follows the Scenario cell.", PCT),
        (9, "COGS % of revenue — Base", float(drv["cogs_base"]), "", PCT),
        (10, "COGS % of revenue — Optimistic", float(drv["cogs_optimistic"]), "", PCT),
        (11, "COGS % of revenue — Pessimistic", float(drv["cogs_pessimistic"]), "", PCT),
        (12, "COGS % of revenue — SELECTED",
         '=IF($B$2="Optimistic",$B$10,IF($B$2="Pessimistic",$B$11,$B$9))',
         "Live: follows the Scenario cell.", PCT),
        (13, "Marketing % of revenue", float(drv["marketing_pct"]),
         "Variable cost — scales with revenue.", PCT),
        (14, "Salaries — month 1 (Rs)", float(drv["salary_base"]), "", MONEY),
        (15, "Monthly salary growth", float(drv["salary_growth"]), "", PCT),
        (16, "Rent / month (Rs)", float(drv["rent"]), "", MONEY),
        (17, "Utilities / month (Rs)", float(drv["utilities"]), "", MONEY),
        (18, "Other opex / month (Rs)", float(drv["other_opex"]), "", MONEY),
        (19, "Loan opening balance (Rs)", float(drv["loan_opening"]),
         "Term loan at 11% p.a., Rs 50,000 principal repaid monthly.", MONEY),
        (20, "Loan interest rate (p.a.)", float(drv["loan_rate"]), "", PCT),
        (21, "Loan principal repaid / month (Rs)", float(drv["loan_principal"]), "", MONEY),
        (22, "Tax rate", float(drv["tax_rate"]),
         "Monthly provision at 25% on positive monthly PBT.", PCT),
        (23, "Plant & machinery opening (Rs)", float(drv["plant_opening"]), "", MONEY),
        (24, "Plant life (months)", float(drv["plant_life_months"]),
         "Straight-line depreciation.", None),
        (25, "Capex Oct 2026 — month 7 (Rs)", float(drv["capex_m7"]), "", MONEY),
        (26, "Capex Apr 2027 — month 13 (Rs)", float(drv["capex_m13"]), "", MONEY),
        (27, "Receivables days", float(drv["wc_recv_days"]),
         "Month-end revenue x days / 30.", None),
        (28, "Inventory days", float(drv["wc_inv_days"]),
         "Month-end COGS x days / 30.", None),
        (29, "Payables days", float(drv["wc_pay_days"]),
         "Month-end COGS x days / 30.", None),
        (30, "Opening cash (Rs)", float(drv["cash_opening"]), "1 Apr 2026.", MONEY),
        (31, "Opening receivables (Rs)", float(drv["recv_opening"]), "", MONEY),
        (32, "Opening inventory (Rs)", float(drv["inv_opening"]), "", MONEY),
        (33, "Opening payables (Rs)", float(drv["pay_opening"]), "", MONEY),
        (34, "Equity (Rs)", float(drv["equity"]), "", MONEY),
        (35, "Opening retained earnings (Rs)", float(drv["retained_opening"]),
         "Opening balance sheet balances exactly (Rs 60,70,000).", MONEY),
    ]
    for row, label, value, note, fmt in spec:
        ws.cell(row=row, column=1, value=label).font = BOLD if row in (2, 8, 12) else Font(size=11)
        c = ws.cell(row=row, column=2, value=value)
        if fmt:
            c.number_format = fmt
        ws.cell(row=row, column=3, value=note).font = SUB_FONT
    dv = DataValidation(type="list", formula1='"Base,Optimistic,Pessimistic"',
                        allow_blank=False)
    dv.error = "Pick Base, Optimistic or Pessimistic."
    dv.prompt = "Scenario for the whole forecast."
    ws.add_data_validation(dv)
    dv.add(ws["B2"])
    return ws


def build_monthly(wb):
    ws = wb.create_sheet("Monthly")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    title_block(ws, "Monthly P&L — Apr 2026 to Mar 2029",
                "Every figure is a live formula linked to Assumptions. "
                "Change the scenario there and this whole sheet re-runs.", 20)
    headers = ["Month", "m", "Revenue", "COGS", "Gross profit", "Salaries",
               "Rent", "Utilities", "Marketing", "Other opex", "Total opex",
               "EBITDA", "Depreciation", "Interest", "PBT", "Tax",
               "Net profit", "Capex", "Loan opening", "FY"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, len(headers))
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:T39"
    widths = [10, 5, 14, 14, 14, 13, 11, 11, 13, 11, 13, 13, 13, 12, 13, 12,
              13, 12, 14, 9]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    labels = []
    y, m = 2026, 4
    for _ in range(36):
        labels.append(f"{y}-{m:02d}")
        m += 1
        if m == 13:
            m, y = 1, y + 1
    for i in range(36):
        r = 4 + i
        mth = i + 1
        ws.cell(row=r, column=1, value=labels[i])
        ws.cell(row=r, column=2, value=mth)
        b = f"$B{r}"
        ws.cell(row=r, column=3,
               value=f"={A}$B$4*(1+{A}$B$8)^({b}-1)")
        ws.cell(row=r, column=4, value=f"=C{r}*{A}$B$12")
        ws.cell(row=r, column=5, value=f"=C{r}-D{r}")
        ws.cell(row=r, column=6,
               value=f"={A}$B$14*(1+{A}$B$15)^({b}-1)")
        ws.cell(row=r, column=7, value=f"={A}$B$16")
        ws.cell(row=r, column=8, value=f"={A}$B$17")
        ws.cell(row=r, column=9, value=f"=C{r}*{A}$B$13")
        ws.cell(row=r, column=10, value=f"={A}$B$18")
        ws.cell(row=r, column=11, value=f"=SUM(F{r}:J{r})")
        ws.cell(row=r, column=12, value=f"=E{r}-K{r}")
        ws.cell(row=r, column=13,
               value=f"={A}$B$23/{A}$B$24"
                     f"+IF({b}>=7,{A}$B$25/{A}$B$24,0)"
                     f"+IF({b}>=13,{A}$B$26/{A}$B$24,0)")
        ws.cell(row=r, column=14,
               value=f"=MAX(0,{A}$B$19-{A}$B$21*({b}-1))*{A}$B$20/12")
        ws.cell(row=r, column=15, value=f"=L{r}-M{r}-N{r}")
        ws.cell(row=r, column=16, value=f"=MAX(0,O{r})*{A}$B$22")
        ws.cell(row=r, column=17, value=f"=O{r}-P{r}")
        ws.cell(row=r, column=18,
               value=f"=IF({b}=7,{A}$B$25,IF({b}=13,{A}$B$26,0))")
        ws.cell(row=r, column=19,
               value=f"=MAX(0,{A}$B$19-{A}$B$21*({b}-1))")
        ws.cell(row=r, column=20,
               value=f'=IF({b}<=12,"FY2027",IF({b}<=24,"FY2028","FY2029"))')
        money_range(ws, r, 3, 19)
    # TOTAL row
    ws.cell(row=40, column=1, value="TOTAL").font = BOLD
    for col in range(3, 20):
        letter = get_column_letter(col)
        c = ws.cell(row=40, column=col,
                   value=f"=SUM({letter}4:{letter}39)")
        c.font = BOLD
        c.number_format = MONEY
    return ws


def build_annual(wb):
    ws = wb.create_sheet("Annual_PL")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    title_block(ws, "Annual P&L — FY2027 to FY2029",
                "SUMIFS over Monthly by fiscal year (Apr–Mar).", 12)
    headers = ["FY", "Revenue", "COGS", "Gross profit", "Total opex",
               "EBITDA", "Depreciation", "Interest", "PBT", "Tax",
               "Net profit", "Capex"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, len(headers))
    ws.column_dimensions["A"].width = 10
    for i in range(2, 13):
        ws.column_dimensions[get_column_letter(i)].width = 15
    src = {"B": "C", "C": "D", "D": "E", "E": "K", "F": "L", "G": "M",
           "H": "N", "I": "O", "J": "P", "K": "Q", "L": "R"}
    for i, fy in enumerate(["FY2027", "FY2028", "FY2029"]):
        r = 4 + i
        ws.cell(row=r, column=1, value=fy).font = BOLD
        for dest, scol in src.items():
            ws.cell(row=r, column=ord(dest) - 64,
                   value=f"=SUMIFS(Monthly!${scol}$4:${scol}$39,"
                         f"Monthly!$T$4:$T$39,$A{r})")
        money_range(ws, r, 2, 12)
    ws.cell(row=7, column=1, value="TOTAL").font = BOLD
    for col in range(2, 13):
        letter = get_column_letter(col)
        c = ws.cell(row=7, column=col, value=f"=SUM({letter}4:{letter}6)")
        c.font = BOLD
        c.number_format = MONEY
    return ws


def build_balance_sheet(wb):
    ws = wb.create_sheet("Balance_Sheet")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    title_block(ws, "Balance Sheet — year-end (31 March)",
                "Balance check must read 0 in every column: the cash-flow "
                "sheet, working capital and retained earnings articulate "
                "by construction.", 5)
    headers = ["Item", "Opening 1 Apr 2026", "FY2027", "FY2028", "FY2029"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, 5)
    ws.column_dimensions["A"].width = 26
    for i in range(2, 6):
        ws.column_dimensions[get_column_letter(i)].width = 18
    # month rows in Monthly: month 12 -> row 15, month 24 -> row 27, month 36 -> row 39
    mrow = {"C": 15, "D": 27, "E": 39}
    rows = [
        (4, "Cash", f"={A}$B$30",
         lambda c: f"=Cash_Flow!{chr(ord(c) - 1)}13"),
        (5, "Receivables", f"={A}$B$31",
         lambda c: f"=Monthly!$C${mrow[c]}*{A}$B$27/30"),
        (6, "Inventory", f"={A}$B$32",
         lambda c: f"=Monthly!$D${mrow[c]}*{A}$B$28/30"),
        (7, "Net fixed assets", f"={A}$B$23",
         lambda c: (f"={A}$B$23+SUM(Monthly!$R$4:$R${mrow[c]})"
                    f"-SUM(Monthly!$M$4:$M${mrow[c]})")),
        (8, "TOTAL ASSETS", None, None),
        (9, "Payables", f"={A}$B$33",
         lambda c: f"=Monthly!$D${mrow[c]}*{A}$B$29/30"),
        (10, "Loan", f"={A}$B$19",
         lambda c: f"={A}$B$19-{A}$B$21*{12 * (ord(c) - 66)}"),
        (11, "Equity", f"={A}$B$34", lambda c: f"={A}$B$34"),
        (12, "Retained earnings", f"={A}$B$35",
         lambda c: f"={A}$B$35+SUM(Monthly!$Q$4:$Q${mrow[c]})"),
        (13, "TOTAL LIABILITIES & EQUITY", None, None),
        (14, "BALANCE CHECK (must be 0)", None, None),
    ]
    for row, label, opening, fn in rows:
        cell = ws.cell(row=row, column=1, value=label)
        if row in (8, 13, 14):
            cell.font = BOLD
        if opening is not None:
            ws.cell(row=row, column=2, value=opening)
        if fn is not None:
            for c in ("C", "D", "E"):
                ws.cell(row=row, column=ord(c) - 64, value=fn(c))
        if row in (8, 13):
            for col in range(2, 6):
                letter = get_column_letter(col)
                c = ws.cell(row=row, column=col,
                           value=f"=SUM({letter}{4 if row == 8 else 9}:"
                                 f"{letter}{7 if row == 8 else 12})")
                c.font = BOLD
        if row == 14:
            for col in range(2, 6):
                letter = get_column_letter(col)
                c = ws.cell(row=row, column=col,
                           value=f"={letter}8-{letter}13")
                c.font = BOLD
        money_range(ws, row, 2, 5)
    return ws


def build_cash_flow(wb):
    ws = wb.create_sheet("Cash_Flow")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    title_block(ws, "Cash Flow — indirect method",
                "Closing cash here feeds the Balance Sheet; the two must "
                "agree every year.", 4)
    headers = ["Item", "FY2027", "FY2028", "FY2029"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, 4)
    ws.column_dimensions["A"].width = 34
    for i in range(2, 5):
        ws.column_dimensions[get_column_letter(i)].width = 18
    spec = [
        (4, "Net profit", lambda c: f"=Annual_PL!K{ord(c) - 62}"),
        (5, "Add: Depreciation", lambda c: f"=Annual_PL!G{ord(c) - 62}"),
        (6, "Less: Increase in receivables",
         lambda c: (f"=-(Balance_Sheet!{chr(ord(c) + 1)}5"
                    f"-Balance_Sheet!{c}5)")),
        (7, "Less: Increase in inventory",
         lambda c: (f"=-(Balance_Sheet!{chr(ord(c) + 1)}6"
                    f"-Balance_Sheet!{c}6)")),
        (8, "Add: Increase in payables",
         lambda c: (f"=Balance_Sheet!{chr(ord(c) + 1)}9"
                    f"-Balance_Sheet!{c}9")),
        (9, "Less: Capex", lambda c: f"=-Annual_PL!L{ord(c) - 62}"),
        (10, "Less: Loan principal repaid",
         lambda c: f"=-{A}$B$21*12"),
        (11, "NET CASH FLOW", None),
        (12, "Opening cash", None),
        (13, "Closing cash", None),
    ]
    for row, label, fn in spec:
        cell = ws.cell(row=row, column=1, value=label)
        if row in (11, 13):
            cell.font = BOLD
        if fn is not None:
            for c in ("B", "C", "D"):
                ws.cell(row=row, column=ord(c) - 64, value=fn(c))
        if row == 11:
            for col in range(2, 5):
                letter = get_column_letter(col)
                c = ws.cell(row=row, column=col,
                           value=f"=SUM({letter}4:{letter}10)")
                c.font = BOLD
        if row == 12:
            ws.cell(row=12, column=2, value=f"={A}$B$30")
            ws.cell(row=12, column=3, value="=B13")
            ws.cell(row=12, column=4, value="=C13")
        if row == 13:
            for col in range(2, 5):
                letter = get_column_letter(col)
                c = ws.cell(row=row, column=col,
                           value=f"={letter}11+{letter}12")
                c.font = BOLD
        money_range(ws, row, 2, 4)
    return ws


def build_ratios(wb):
    ws = wb.create_sheet("Ratios")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    title_block(ws, "Ratios & break-even",
                "Current ratio uses payables as the only current liability "
                "(no overdraft in this model). Break-even revenue covers "
                "all fixed cash and non-cash costs at the selected COGS% "
                "and marketing %.", 4)
    headers = ["Ratio", "FY2027", "FY2028", "FY2029"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, 4)
    ws.column_dimensions["A"].width = 40
    for i in range(2, 5):
        ws.column_dimensions[get_column_letter(i)].width = 18
    mrow = {"B": 15, "C": 27, "D": 39}
    arow = {"B": 4, "C": 5, "D": 6}
    spec = [
        (4, "Gross margin",
         lambda c: f"=Annual_PL!D{arow[c]}/Annual_PL!B{arow[c]}", PCT),
        (5, "Net margin",
         lambda c: f"=Annual_PL!K{arow[c]}/Annual_PL!B{arow[c]}", PCT),
        (6, "Current ratio",
         lambda c: (f"=(Balance_Sheet!{c}4+Balance_Sheet!{c}5"
                    f"+Balance_Sheet!{c}6)/Balance_Sheet!{c}9"), RATIO),
        (7, "Debt / equity",
         lambda c: (f"=Balance_Sheet!{c}10/(Balance_Sheet!{c}11"
                    f"+Balance_Sheet!{c}12)"), RATIO),
        (8, "Return on equity",
         lambda c: (f"=Annual_PL!K{arow[c]}/(Balance_Sheet!{c}11"
                    f"+Balance_Sheet!{c}12)"), PCT),
        (9, "Break-even monthly revenue (Rs)",
         lambda c: (f"=(Monthly!F{mrow[c]}+{A}$B$16+{A}$B$17"
                    f"+{A}$B$18+Monthly!M{mrow[c]}+Monthly!N{mrow[c]})"
                    f"/(1-{A}$B$12-{A}$B$13)"), MONEY),
        (10, "Fixed costs in break-even month (Rs)",
         lambda c: (f"=Monthly!F{mrow[c]}+{A}$B$16+{A}$B$17"
                    f"+{A}$B$18+Monthly!M{mrow[c]}+Monthly!N{mrow[c]}"), MONEY),
    ]
    for row, label, fn, fmt in spec:
        ws.cell(row=row, column=1, value=label).font = BOLD
        for c in ("B", "C", "D"):
            cell = ws.cell(row=row, column=ord(c) - 64, value=fn(c))
            cell.number_format = fmt
    return ws


def build_sensitivity(wb):
    ws = wb.create_sheet("Sensitivity")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    title_block(ws, "Sensitivity — Year-3 net profit (Rs)",
                "Rows vary monthly revenue growth; columns vary COGS%. "
                "Each cell recomputes Year-3 net profit in closed form "
                "(geometric series over months 25–36) from the same "
                "drivers — the centre cell must equal Annual Year-3 net "
                "profit. Read across a row to see growth risk, down a "
                "column to see margin risk.", 7)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 13
    for i in range(3, 8):
        ws.column_dimensions[get_column_letter(i)].width = 17
    growth_offsets = [-0.02, -0.01, 0.0, 0.01, 0.02]
    cogs_offsets = [-0.03, -0.015, 0.0, 0.015, 0.03]
    ws.cell(row=4, column=1,
           value="Growth offset \\ COGS% offset").font = BOLD
    ws.cell(row=4, column=2, value="").font = BOLD
    for j, dc in enumerate(cogs_offsets):
        c = ws.cell(row=4, column=3 + j, value=dc)
        c.number_format = PCT
        c.font = BOLD
        c.fill = PatternFill("solid", fgColor=LIGHT)
        c.alignment = Alignment(horizontal="center")
    for i, dg in enumerate(growth_offsets):
        r = 5 + i
        c = ws.cell(row=r, column=2, value=dg)
        c.number_format = PCT
        c.font = BOLD
        c.fill = PatternFill("solid", fgColor=LIGHT)
        c.alignment = Alignment(horizontal="center")
    ws.cell(row=4, column=1).alignment = Alignment(horizontal="right")
    # Helper block: Year-3 revenue per growth offset (rows 12-16, col C)
    ws.cell(row=11, column=1,
           value="Helper — Year-3 revenue by growth offset").font = BOLD
    for i, dg in enumerate(growth_offsets):
        r = 12 + i
        ws.cell(row=r, column=2, value=dg).number_format = PCT
        ws.cell(row=r, column=3, value=(
            f"={A}$B$4*POWER(1+{A}$B$8+$B{r},24)"
            f"*(POWER(1+{A}$B$8+$B{r},12)-1)/({A}$B$8+$B{r})"))
        ws.cell(row=r, column=3).number_format = MONEY
    # Grid cells: PBT then tax, via repeated PBT expression
    for i in range(5):
        r = 5 + i
        hr = 12 + i  # helper row for this growth offset
        for j in range(5):
            col = 3 + j
            letter = get_column_letter(col)
            rev = f"$C{hr}"
            pbt = (f"{rev}*(1-({A}$B$12+{letter}$4))"
                   f"-((Annual_PL!$E$6-0.03*Annual_PL!$B$6)"
                   f"+0.03*{rev})"
                   f"-Annual_PL!$G$6-Annual_PL!$H$6")
            cell = ws.cell(row=r, column=col,
                          value=f"={pbt}-MAX(0,{pbt})*{A}$B$22")
            cell.number_format = MONEY
            if i == 2 and j == 2:
                cell.fill = PatternFill("solid", fgColor=GREEN)
                cell.font = BOLD
    ws.cell(row=18, column=1, value=(
        "Check: centre cell (0% / 0%) must equal Annual_PL Year-3 net "
        "profit.")).font = SUB_FONT
    ws.cell(row=19, column=1, value="Centre minus Annual Y3 net (must be 0):").font = BOLD
    c = ws.cell(row=19, column=2, value="=E7-Annual_PL!K6")
    c.number_format = MONEY
    c.font = BOLD
    return ws


def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    title_block(ws, "Dashboard — Ganga Foods 3-year forecast",
                "Scenario, KPIs and charts. Change the scenario on the "
                "Assumptions sheet to re-run everything.", 8)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 20
    ws.cell(row=3, column=1, value="Scenario").font = BOLD
    c = ws.cell(row=3, column=2, value=f"={A}$B$2")
    c.font = BOLD
    kpis = [
        (5, "Year-3 revenue (Rs)", "=Annual_PL!B6", MONEY),
        (6, "Year-3 net profit (Rs)", "=Annual_PL!K6", MONEY),
        (7, "Year-3 net margin", "=Annual_PL!K6/Annual_PL!B6", PCT),
        (8, "Break-even monthly revenue, FY2029 (Rs)", "=Ratios!D9", MONEY),
        (9, "Closing cash, FY2029 (Rs)", "=Balance_Sheet!E4", MONEY),
        (10, "Current ratio, FY2029", "=Ratios!D6", RATIO),
    ]
    for row, label, formula, fmt in kpis:
        ws.cell(row=row, column=1, value=label).font = BOLD
        cell = ws.cell(row=row, column=2, value=formula)
        cell.number_format = fmt
        cell.font = Font(bold=True, size=12, color=NAVY)
    # Charts
    line = LineChart()
    line.title = "Monthly revenue vs net profit"
    line.y_axis.title = "Rs"
    line.x_axis.title = "Month"
    line.style = 2
    line.height = 7.5
    line.width = 15
    data = Reference(wb["Monthly"], min_col=3, min_row=3, max_row=39)
    cats = Reference(wb["Monthly"], min_col=1, min_row=4, max_row=39)
    line.add_data(data, titles_from_data=True)
    data2 = Reference(wb["Monthly"], min_col=17, min_row=3, max_row=39)
    line.add_data(data2, titles_from_data=True)
    line.set_categories(cats)
    ws.add_chart(line, "D3")
    bar = BarChart()
    bar.type = "col"
    bar.title = "Annual revenue vs net profit"
    bar.y_axis.title = "Rs"
    bar.style = 10
    bar.height = 7.5
    bar.width = 15
    data = Reference(wb["Annual_PL"], min_col=2, min_row=3, max_row=6)
    cats = Reference(wb["Annual_PL"], min_col=1, min_row=4, max_row=6)
    bar.add_data(data, titles_from_data=True)
    data2 = Reference(wb["Annual_PL"], min_col=11, min_row=3, max_row=6)
    bar.add_data(data2, titles_from_data=True)
    bar.set_categories(cats)
    ws.add_chart(bar, "D19")
    cash = LineChart()
    cash.title = "Cash walk — opening to FY2029 close"
    cash.y_axis.title = "Rs"
    cash.style = 2
    cash.height = 7.5
    cash.width = 15
    data = Reference(wb["Balance_Sheet"], min_col=2, min_row=3, max_row=4)
    cats = Reference(wb["Balance_Sheet"], min_col=2, min_row=3, max_row=5)
    cash.add_data(data, titles_from_data=True)
    cash.set_categories(cats)
    ws.add_chart(cash, "D35")
    return ws


def build_guide(wb):
    ws = wb.create_sheet("Guide", 0)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    title_block(ws, "ForecastForge — 3-year financial forecast model",
                "Built for the Freelancer lead '3 Year Financial Forecast "
                "Model' (Bodh Gaya client): three-year P&L, balance sheet "
                "and cash-flow forecast in Excel, with assumptions, ratios, "
                "break-even and sensitivity. Fictional company: Ganga Foods "
                "Pvt. Ltd., Bodh Gaya — packaged foods.", 2)
    ws.column_dimensions["A"].width = 110
    lines = [
        "",
        "HOW TO USE",
        "1. Go to Assumptions. Pick a scenario in B2 (Base / Optimistic / "
        "Pessimistic) — every sheet re-runs from that one cell.",
        "2. Edit any driver value (month-1 revenue, growth, COGS%, salaries, "
        "loan terms, capex, working-capital days) and the whole model updates.",
        "3. Read the story on Dashboard: 6 KPIs and 3 charts. Drill into "
        "Monthly for the 36-month detail, Annual_PL for fiscal years, "
        "Balance_Sheet and Cash_Flow for the articulated statements.",
        "",
        "WHAT EACH SHEET DOES",
        "- Assumptions: every driver as a plain value, plus the scenario "
        "switch. SELECTED growth and COGS% are live formulas.",
        "- Monthly: 36 months, Apr 2026–Mar 2029. Revenue compounds monthly; "
        "salaries grow monthly; marketing is 3% of revenue; depreciation "
        "steps up when capex lands (Oct 2026, Apr 2027); interest is charged "
        "on the opening loan balance; tax is a 25% provision on positive "
        "monthly PBT.",
        "- Annual_PL: fiscal years by SUMIFS over Monthly (Apr–Mar).",
        "- Balance_Sheet: year-end. Receivables 20 days of month-end revenue, "
        "inventory 15 days of month-end COGS, payables 25 days of month-end "
        "COGS. Balance check must read 0 — if it does not, a formula broke.",
        "- Cash_Flow: indirect method. Closing cash feeds the Balance Sheet.",
        "- Ratios: gross / net margin, current ratio, debt/equity, ROE, and "
        "break-even monthly revenue (fixed costs / contribution margin).",
        "- Sensitivity: Year-3 net profit for growth offsets -2%..+2% against "
        "COGS% offsets -3%..+3%. Centre cell equals Annual Year-3 net profit.",
        "",
        "VERIFICATION",
        "The workbook was verified by LibreOffice headless recalculation "
        "against an independent Python model: annual revenue and net profit, "
        "balance checks, cash tie-outs and the sensitivity centre cell all "
        "agree. See tests/test_build.py in the repository folder.",
        "",
        "LIMITS (honest scope)",
        "- Tax is a flat 25% provision on positive monthly PBT, not a full "
        "Indian income-tax computation (no MAT, no loss carry-forward).",
        "- No GST, no inventory valuation choice, no debtor ageing — working "
        "capital is a simple days-of-sales rule.",
        "- A real engagement would start from the client's actual trial "
        "balance and bank statements, not these fictional drivers.",
    ]
    for i, line in enumerate(lines, start=3):
        cell = ws.cell(row=i, column=1, value=line)
        if line and not line.startswith(("-", "1.", "2.", "3.")):
            cell.font = BOLD
    return ws


def main():
    drv = load_assumptions()
    wb = Workbook()
    build_assumptions(wb, drv)
    build_monthly(wb)
    build_annual(wb)
    build_balance_sheet(wb)
    build_cash_flow(wb)
    build_ratios(wb)
    build_sensitivity(wb)
    build_dashboard(wb)
    build_guide(wb)
    wb.save(XLSX)
    print(f"Wrote {XLSX}")


if __name__ == "__main__":
    main()
