#!/usr/bin/env python3
"""Build forecast-model.xlsx — ForecastForge, a 3-year financial forecast
model for a fictional wholesale distributor (Kosi Traders Pvt. Ltd., Patna).

Built around the Freelancer lead "3 Year Financial Forecast Model"
(ID 40821252): a flexible 3-year Excel forecast model — revenue, EBITDA,
net income, cash flow, assumptions sheet, scenario toggles — plus a
concise dashboard with charts.

Sheets:
  Guide        — what the pack is and how to use it.
  Assumptions  — FY26 actuals + opening position, scenario table
                 (Base / Optimistic / Pessimistic), and the active
                 scenario block driven by the Dashboard selector.
  P&L          — FY26 actual + FY27-29 forecast: Revenue, COGS,
                 Gross Profit, opex lines, EBITDA, Depreciation,
                 Interest, PBT, Tax, Net Income. Every forecast cell
                 is a live formula off Assumptions.
  CashFlow     — FY27-29: Net Income + Depreciation -/+ working
                 capital moves - Capex - debt repayment = net cash
                 flow; opening / closing cash.
  BalanceSheet — FY26-29: Cash, Receivables, Inventory, Net Fixed
                 Assets / Payables, Debt, Share Capital, Retained
                 Earnings, with a live balance check (must be 0).
  Dashboard    — scenario selector (B3) driving the whole model,
                 FY29 KPIs, a 4-year summary table, 2 native charts.
  Checks       — 10 integrity checks; all must read PASS.

Articulation (verified by the Checks sheet): retained earnings =
opening + cumulative net income; debt falls by the annual repayment;
cash ties between CashFlow and BalanceSheet; the balance sheet
balances in every year and every scenario.

Formula rules (LibreOffice 24.2): plain A1-style ranges, no Table
objects, no structured references; every range literal repeats the
column letter on both ends.
"""
import csv
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
XLSX = BASE / "forecast-model.xlsx"

YEARS = ["FY26", "FY27", "FY28", "FY29"]
SCENARIOS = ["Base", "Optimistic", "Pessimistic"]

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14)
BOLD = Font(bold=True, size=11)
MONEY = "#,##0"
PCT = "0.0%"

# (csv key, label, is_percent)
BASE_INPUTS = [
    ("base_revenue_fy26", "Base revenue FY26 (Rs)", False),
    ("cogs_fy26", "COGS FY26 (Rs)", False),
    ("salaries_fy26", "Salaries FY26 (Rs)", False),
    ("rent_fy26", "Rent FY26 (Rs)", False),
    ("marketing_fy26", "Marketing FY26 (Rs)", False),
    ("admin_fy26", "Admin FY26 (Rs)", False),
    ("depreciation_fy26", "Depreciation FY26 (Rs)", False),
    ("interest_fy26", "Interest FY26 (Rs)", False),
    ("opening_cash", "Opening cash (Rs)", False),
    ("opening_receivables", "Opening receivables (Rs)", False),
    ("opening_inventory", "Opening inventory (Rs)", False),
    ("opening_net_fixed_assets", "Opening net fixed assets (Rs)", False),
    ("opening_payables", "Opening payables (Rs)", False),
    ("opening_debt", "Opening debt (Rs)", False),
    ("share_capital", "Share capital (Rs)", False),
    ("opening_retained_earnings", "Opening retained earnings (Rs)", False),
    ("annual_capex", "Annual capex (Rs)", False),
    ("annual_debt_repayment", "Annual debt repayment (Rs)", False),
    ("interest_rate", "Interest rate (on opening debt)", True),
    ("depreciation_rate", "Depreciation rate", True),
    ("tax_rate", "Tax rate", True),
    ("receivables_pct_revenue", "Receivables % of revenue", True),
    ("inventory_pct_cogs", "Inventory % of COGS", True),
    ("payables_pct_cogs", "Payables % of COGS", True),
]
# Assumptions value cells: rows 3..26 in column B, in BASE_INPUTS order.
A = {key: f"B{3 + i}" for i, (key, _, _) in enumerate(BASE_INPUTS)}


def load_inputs():
    with open(DATA / "base_inputs.csv", newline="") as f:
        vals = {r["key"]: float(r["value"]) for r in csv.DictReader(f)}
    with open(DATA / "scenarios.csv", newline="") as f:
        scenarios = list(csv.DictReader(f))
    assert set(vals) == {k for k, _, _ in BASE_INPUTS}, set(vals) ^ {
        k for k, _, _ in BASE_INPUTS}
    assert [s["scenario"] for s in scenarios] == SCENARIOS
    return vals, scenarios


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT


def money_cells(ws, cells):
    for ref in cells:
        ws[ref].number_format = MONEY


def build():
    vals, scenarios = load_inputs()
    wb = openpyxl.Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = ("3-Year Financial Forecast Model — "
                "Kosi Traders Pvt. Ltd., Patna")
    ws["A1"].font = TITLE_FONT
    lines = [
        "A formula-driven 3-year forecast model for a wholesale FMCG",
        "distributor (FY26 actuals, FY27-FY29 forecast). Built as a",
        "working demo for the Freelancer project '3 Year Financial",
        "Forecast Model' (ID 40821252).",
        "",
        "HOW TO USE",
        "1. Pick a scenario on the Dashboard (cell B3 — Base /",
        "   Optimistic / Pessimistic). Every sheet updates from live",
        "   formulas; nothing in the forecast years is hard-coded.",
        "2. Change any assumption on the Assumptions sheet and the",
        "   P&L, cash flow, balance sheet and dashboard follow.",
        "3. Read the Checks sheet after any change — all 10 checks",
        "   must read PASS (balance sheet balances, cash ties out,",
        "   retained earnings and debt articulate).",
        "",
        "WHAT THE PESSIMISTIC SCENARIO SHOWS",
        "Under Pessimistic the model goes cash-negative in FY28:",
        "growth of 8/6/5% at a 28% margin does not cover capex plus",
        "debt repayment. That is the model doing its job — it flags",
        "a funding gap of roughly Rs 3.5 lakh by FY28, not a bug.",
        "",
        "SHEETS",
        "Assumptions  — FY26 actuals, opening position, scenario",
        "               table and the active-scenario block.",
        "P&L          — revenue, COGS, gross profit, opex, EBITDA,",
        "               depreciation, interest, PBT, tax, net income.",
        "CashFlow     — net income + depreciation, working-capital",
        "               moves, capex, debt repayment, closing cash.",
        "BalanceSheet — assets, liabilities, equity, balance check.",
        "Dashboard    — scenario selector, FY29 KPIs, 4-year",
        "               summary table, 2 native charts.",
        "Checks       — 10 integrity checks; all must read PASS.",
        "",
        "DATA — all fictional. Kosi Traders is invented; FY26 figures",
        "are plausible inputs for the demo, not a real company's books.",
    ]
    for i, line in enumerate(lines, start=3):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 78

    # ---------------- Assumptions ----------------
    ws = wb.create_sheet("Assumptions")
    ws["A1"] = "Forecast assumptions — Kosi Traders Pvt. Ltd."
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "FY26 actuals & opening position (Rs unless noted)"
    ws["A2"].font = BOLD
    for i, (key, label, is_pct) in enumerate(BASE_INPUTS):
        r = 3 + i
        ws.cell(row=r, column=1, value=label)
        c = ws.cell(row=r, column=2, value=vals[key])
        c.number_format = PCT if is_pct else MONEY
    ws["A28"] = "Scenario table"
    ws["A28"].font = BOLD
    ws.append([])  # row 29 is the header row below (written explicitly)
    for col, h in enumerate(["Scenario", "Rev growth Y1", "Rev growth Y2",
                             "Rev growth Y3", "Gross margin",
                             "Opex growth"], start=1):
        ws.cell(row=29, column=col, value=h)
    style_header(ws, 29, 6)
    for i, s in enumerate(scenarios):
        r = 30 + i
        ws.cell(row=r, column=1, value=s["scenario"])
        for j, k in enumerate(["growth_y1", "growth_y2", "growth_y3",
                               "gross_margin", "opex_growth"], start=2):
            c = ws.cell(row=r, column=j, value=float(s[k]))
            c.number_format = PCT
    ws["A34"] = "Active scenario (driven by Dashboard!B3)"
    ws["A34"].font = BOLD
    ws["A35"] = "Scenario"
    ws["B35"] = "=Dashboard!B3"
    ws["B35"].font = BOLD
    active = [
        ("Active rev growth Y1", "B"),
        ("Active rev growth Y2", "C"),
        ("Active rev growth Y3", "D"),
        ("Active gross margin", "E"),
        ("Active opex growth", "F"),
    ]
    for i, (label, col) in enumerate(active):
        r = 36 + i
        ws.cell(row=r, column=1, value=label)
        c = ws.cell(
            row=r, column=2,
            value=(f"=INDEX({col}$30:{col}$32,"
                   f"MATCH($B$35,$A$30:$A$32,0))"))
        c.number_format = PCT
    ws.column_dimensions["A"].width = 34
    for col in "BCDEF":
        ws.column_dimensions[col].width = 16
    ws.freeze_panes = "A3"

    # Shorthand: active-scenario cells on Assumptions.
    G1, G2, G3 = "Assumptions!$B$36", "Assumptions!$B$37", "Assumptions!$B$38"
    GM, GO = "Assumptions!$B$39", "Assumptions!$B$40"

    # ---------------- P&L ----------------
    ws = wb.create_sheet("P&L")
    ws["A1"] = "Profit & Loss — FY26 actual, FY27-FY29 forecast (Rs)"
    ws["A1"].font = TITLE_FONT
    ws.append([])
    ws.append([])
    for col, h in enumerate(["Line", "FY26 Actual", "FY27", "FY28",
                             "FY29"], start=1):
        ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 5)
    pl = [
        ("Revenue", f"=Assumptions!{A['base_revenue_fy26']}",
         f"=B5*(1+{G1})", f"=C5*(1+{G2})", f"=D5*(1+{G3})"),
        ("COGS", f"=Assumptions!{A['cogs_fy26']}",
         f"=C5*(1-{GM})", f"=D5*(1-{GM})", f"=E5*(1-{GM})"),
        ("Gross Profit", "=B5-B6", "=C5-C6", "=D5-D6", "=E5-E6"),
        ("Salaries", f"=Assumptions!{A['salaries_fy26']}",
         f"=B8*(1+{GO})", f"=C8*(1+{GO})", f"=D8*(1+{GO})"),
        ("Rent", f"=Assumptions!{A['rent_fy26']}",
         f"=B9*(1+{GO})", f"=C9*(1+{GO})", f"=D9*(1+{GO})"),
        ("Marketing", f"=Assumptions!{A['marketing_fy26']}",
         f"=B10*(1+{GO})", f"=C10*(1+{GO})", f"=D10*(1+{GO})"),
        ("Admin", f"=Assumptions!{A['admin_fy26']}",
         f"=B11*(1+{GO})", f"=C11*(1+{GO})", f"=D11*(1+{GO})"),
        ("Total Opex", "=SUM(B8:B11)", "=SUM(C8:C11)", "=SUM(D8:D11)",
         "=SUM(E8:E11)"),
        ("EBITDA", "=B7-B12", "=C7-C12", "=D7-D12", "=E7-E12"),
        ("Depreciation", f"=Assumptions!{A['depreciation_fy26']}",
         f"=Assumptions!{A['depreciation_rate']}*"
         f"(BalanceSheet!B8+Assumptions!{A['annual_capex']})",
         f"=Assumptions!{A['depreciation_rate']}*"
         f"(BalanceSheet!C8+Assumptions!{A['annual_capex']})",
         f"=Assumptions!{A['depreciation_rate']}*"
         f"(BalanceSheet!D8+Assumptions!{A['annual_capex']})"),
        ("Interest", f"=Assumptions!{A['interest_fy26']}",
         f"=Assumptions!{A['interest_rate']}*BalanceSheet!B11",
         f"=Assumptions!{A['interest_rate']}*BalanceSheet!C11",
         f"=Assumptions!{A['interest_rate']}*BalanceSheet!D11"),
        ("Profit Before Tax", "=B13-B14-B15", "=C13-C14-C15",
         "=D13-D14-D15", "=E13-E14-E15"),
        ("Tax", f"=MAX(0,B16)*Assumptions!{A['tax_rate']}",
         f"=MAX(0,C16)*Assumptions!{A['tax_rate']}",
         f"=MAX(0,D16)*Assumptions!{A['tax_rate']}",
         f"=MAX(0,E16)*Assumptions!{A['tax_rate']}"),
        ("Net Income", "=B16-B17", "=C16-C17", "=D16-D17", "=E16-E17"),
    ]
    for i, (label, b, c, d, e) in enumerate(pl):
        r = 5 + i
        ws.cell(row=r, column=1, value=label).font = (
            BOLD if label in ("Gross Profit", "Total Opex", "EBITDA",
                              "Profit Before Tax", "Net Income") else Font())
        for col, formula in zip("BCDE", (b, c, d, e)):
            cell = ws.cell(row=r, column="BCDE".index(col) + 2,
                          value=formula)
            cell.number_format = MONEY
    ws.column_dimensions["A"].width = 20
    for col in "BCDE":
        ws.column_dimensions[col].width = 17
    ws.freeze_panes = "A5"

    # ---------------- CashFlow ----------------
    ws = wb.create_sheet("CashFlow")
    ws["A1"] = "Cash Flow — FY27-FY29 forecast (Rs)"
    ws["A1"].font = TITLE_FONT
    ws.append([])
    ws.append([])
    for col, h in enumerate(["Line", "FY27", "FY28", "FY29"], start=1):
        ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 4)
    cf = [
        ("Net Income", "='P&L'!C18", "='P&L'!D18", "='P&L'!E18"),
        ("+ Depreciation", "='P&L'!C14", "='P&L'!D14", "='P&L'!E14"),
        ("- Increase in Receivables",
         "=-(BalanceSheet!C6-BalanceSheet!B6)",
         "=-(BalanceSheet!D6-BalanceSheet!C6)",
         "=-(BalanceSheet!E6-BalanceSheet!D6)"),
        ("- Increase in Inventory",
         "=-(BalanceSheet!C7-BalanceSheet!B7)",
         "=-(BalanceSheet!D7-BalanceSheet!C7)",
         "=-(BalanceSheet!E7-BalanceSheet!D7)"),
        ("+ Increase in Payables",
         "=BalanceSheet!C10-BalanceSheet!B10",
         "=BalanceSheet!D10-BalanceSheet!C10",
         "=BalanceSheet!E10-BalanceSheet!D10"),
        ("- Capex", f"=-Assumptions!{A['annual_capex']}",
         f"=-Assumptions!{A['annual_capex']}",
         f"=-Assumptions!{A['annual_capex']}"),
        ("- Debt repayment",
         f"=-Assumptions!{A['annual_debt_repayment']}",
         f"=-Assumptions!{A['annual_debt_repayment']}",
         f"=-Assumptions!{A['annual_debt_repayment']}"),
        ("Net Cash Flow", "=SUM(B5:B11)", "=SUM(C5:C11)",
         "=SUM(D5:D11)"),
        ("Opening Cash", f"=Assumptions!{A['opening_cash']}",
         "=B14", "=C14"),
        ("Closing Cash", "=B12+B13", "=C12+C13", "=D12+D13"),
    ]
    for i, (label, b, c, d) in enumerate(cf):
        r = 5 + i
        ws.cell(row=r, column=1, value=label).font = (
            BOLD if label in ("Net Cash Flow", "Closing Cash")
            else Font())
        for col, formula in zip("BCD", (b, c, d)):
            cell = ws.cell(row=r, column="BCD".index(col) + 2,
                          value=formula)
            cell.number_format = MONEY
    ws.column_dimensions["A"].width = 30
    for col in "BCD":
        ws.column_dimensions[col].width = 17
    ws.freeze_panes = "A5"

    # ---------------- BalanceSheet ----------------
    ws = wb.create_sheet("BalanceSheet")
    ws["A1"] = "Balance Sheet — FY26 actual, FY27-FY29 forecast (Rs)"
    ws["A1"].font = TITLE_FONT
    ws.append([])
    ws.append([])
    for col, h in enumerate(["Line", "FY26 Actual", "FY27", "FY28",
                             "FY29"], start=1):
        ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 5)
    bs = [
        ("Cash", f"=Assumptions!{A['opening_cash']}",
         "=CashFlow!B14", "=CashFlow!C14", "=CashFlow!D14"),
        ("Receivables", f"=Assumptions!{A['opening_receivables']}",
         f"=Assumptions!{A['receivables_pct_revenue']}*'P&L'!C5",
         f"=Assumptions!{A['receivables_pct_revenue']}*'P&L'!D5",
         f"=Assumptions!{A['receivables_pct_revenue']}*'P&L'!E5"),
        ("Inventory", f"=Assumptions!{A['opening_inventory']}",
         f"=Assumptions!{A['inventory_pct_cogs']}*'P&L'!C6",
         f"=Assumptions!{A['inventory_pct_cogs']}*'P&L'!D6",
         f"=Assumptions!{A['inventory_pct_cogs']}*'P&L'!E6"),
        ("Net Fixed Assets",
         f"=Assumptions!{A['opening_net_fixed_assets']}",
         f"=B8+Assumptions!{A['annual_capex']}-'P&L'!C14",
         f"=C8+Assumptions!{A['annual_capex']}-'P&L'!D14",
         f"=D8+Assumptions!{A['annual_capex']}-'P&L'!E14"),
        ("Total Assets", "=SUM(B5:B8)", "=SUM(C5:C8)", "=SUM(D5:D8)",
         "=SUM(E5:E8)"),
        ("Payables", f"=Assumptions!{A['opening_payables']}",
         f"=Assumptions!{A['payables_pct_cogs']}*'P&L'!C6",
         f"=Assumptions!{A['payables_pct_cogs']}*'P&L'!D6",
         f"=Assumptions!{A['payables_pct_cogs']}*'P&L'!E6"),
        ("Debt", f"=Assumptions!{A['opening_debt']}",
         f"=B11-Assumptions!{A['annual_debt_repayment']}",
         f"=C11-Assumptions!{A['annual_debt_repayment']}",
         f"=D11-Assumptions!{A['annual_debt_repayment']}"),
        ("Share Capital", f"=Assumptions!{A['share_capital']}",
         "=B12", "=C12", "=D12"),
        ("Retained Earnings",
         f"=Assumptions!{A['opening_retained_earnings']}",
         "=B13+'P&L'!C18", "=C13+'P&L'!D18", "=D13+'P&L'!E18"),
        ("Total Liabilities & Equity", "=SUM(B10:B13)",
         "=SUM(C10:C13)", "=SUM(D10:D13)", "=SUM(E10:E13)"),
        ("Balance Check (must be 0)", "=B9-B14", "=C9-C14",
         "=D9-D14", "=E9-E14"),
    ]
    for i, (label, b, c, d, e) in enumerate(bs):
        r = 5 + i
        ws.cell(row=r, column=1, value=label).font = (
            BOLD if label in ("Total Assets",
                              "Total Liabilities & Equity",
                              "Balance Check (must be 0)") else Font())
        for col, formula in zip("BCDE", (b, c, d, e)):
            cell = ws.cell(row=r, column="BCDE".index(col) + 2,
                          value=formula)
            cell.number_format = MONEY
    ws.column_dimensions["A"].width = 28
    for col in "BCDE":
        ws.column_dimensions[col].width = 17
    ws.freeze_panes = "A5"

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Forecast Dashboard — Kosi Traders Pvt. Ltd."
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Change the scenario in B3 — every KPI, table and chart "
                "below is live formulas.")
    ws["A3"] = "Scenario"
    ws["B3"] = "Base"
    ws["A4"] = "Active scenario"
    ws["B4"] = "=Assumptions!B35"
    for rr in (3, 4):
        ws.cell(row=rr, column=1).font = BOLD
        ws.cell(row=rr, column=2).font = BOLD
    dv = DataValidation(type="list",
                        formula1='"Base,Optimistic,Pessimistic"',
                        allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B3")

    ws["A6"] = "Key figures (Rs unless noted)"
    ws["A6"].font = BOLD
    kpis = [
        ("Revenue FY29", "='P&L'!E5", MONEY),
        ("EBITDA FY29", "='P&L'!E13", MONEY),
        ("Net Income FY29", "='P&L'!E18", MONEY),
        ("Closing Cash FY29", "=CashFlow!D14", MONEY),
        ("EBITDA margin FY29",
         "=IF('P&L'!E5=0,0,'P&L'!E13/'P&L'!E5)", PCT),
        ("Revenue CAGR FY26-FY29",
         "=IF('P&L'!B5=0,0,('P&L'!E5/'P&L'!B5)^(1/3)-1)", PCT),
        ("Total Net Income FY27-FY29", "=SUM('P&L'!C18:E18)", MONEY),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        r = 7 + i
        ws.cell(row=r, column=1, value=label).font = BOLD
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = fmt

    ws["A14"] = "Summary by year"
    ws["A14"].font = BOLD
    for col, h in enumerate(["Year", "Revenue", "EBITDA", "Net Income",
                             "Closing Cash"], start=1):
        ws.cell(row=15, column=col, value=h)
    style_header(ws, 15, 5)
    summary = [
        ("FY26", "='P&L'!B5", "='P&L'!B13", "='P&L'!B18",
         f"=Assumptions!{A['opening_cash']}"),
        ("FY27", "='P&L'!C5", "='P&L'!C13", "='P&L'!C18",
         "=CashFlow!B14"),
        ("FY28", "='P&L'!D5", "='P&L'!D13", "='P&L'!D18",
         "=CashFlow!C14"),
        ("FY29", "='P&L'!E5", "='P&L'!E13", "='P&L'!E18",
         "=CashFlow!D14"),
    ]
    for i, (year, rev, ebitda, ni, cash) in enumerate(summary):
        r = 16 + i
        ws.cell(row=r, column=1, value=year)
        for col, formula in zip("BCDE", (rev, ebitda, ni, cash)):
            cell = ws.cell(row=r, column="BCDE".index(col) + 2,
                          value=formula)
            cell.number_format = MONEY
    ws.column_dimensions["A"].width = 30
    for col in "BCDE":
        ws.column_dimensions[col].width = 17

    ch1 = BarChart()
    ch1.type = "col"
    ch1.style = 10
    ch1.title = "Revenue vs EBITDA by year (Rs)"
    ch1.y_axis.title = "Rs"
    ch1.height = 7.5
    ch1.width = 15
    data_ref = Reference(ws, min_col=2, min_row=15, max_row=19, max_col=3)
    cats_ref = Reference(ws, min_col=1, min_row=16, max_row=19)
    ch1.add_data(data_ref, titles_from_data=True)
    ch1.set_categories(cats_ref)
    ch1.shape = 4
    ws.add_chart(ch1, "G6")

    ch2 = LineChart()
    ch2.style = 10
    ch2.title = "Closing cash by year (Rs)"
    ch2.y_axis.title = "Rs"
    ch2.height = 7.5
    ch2.width = 15
    data_ref2 = Reference(ws, min_col=5, min_row=15, max_row=19)
    ch2.add_data(data_ref2, titles_from_data=True)
    ch2.set_categories(cats_ref)
    ws.add_chart(ch2, "G23")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws["A1"] = "Integrity checks — all must read PASS"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Checks hold in every scenario — switch Dashboard!B3 "
                "and re-check.")
    ws.append([])
    ws.append(["#", "Check", "Expected", "Actual", "Result"])
    style_header(ws, 4, 5)
    checks = [
        ("Opening balance sheet balances",
         "=BalanceSheet!B15"),
        ("Balance check FY27 = 0",
         "=BalanceSheet!C15"),
        ("Balance check FY28 = 0",
         "=BalanceSheet!D15"),
        ("Balance check FY29 = 0",
         "=BalanceSheet!E15"),
        ("Cash tie-out FY27 (BS vs CashFlow)",
         "=BalanceSheet!C5-CashFlow!B14"),
        ("Cash tie-out FY28 (BS vs CashFlow)",
         "=BalanceSheet!D5-CashFlow!C14"),
        ("Cash tie-out FY29 (BS vs CashFlow)",
         "=BalanceSheet!E5-CashFlow!D14"),
        ("Retained earnings FY29 = opening + 3yr net income",
         "=BalanceSheet!E13-(Assumptions!"
         f"{A['opening_retained_earnings']}+SUM('P&L'!C18:E18))"),
        ("Debt FY29 = opening debt - 3 x repayment",
         "=BalanceSheet!E11-(Assumptions!"
         f"{A['opening_debt']}-3*Assumptions!"
         f"{A['annual_debt_repayment']})"),
        ("Dashboard FY29 revenue = P&L FY29 revenue",
         "=Dashboard!B7-'P&L'!E5"),
    ]
    for i, (label, actual) in enumerate(checks):
        r = 5 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=0)  # numeric, never text
        ws.cell(row=r, column=4, value=actual)
        ws.cell(row=r, column=5,
               value=f'=IF(ABS(D{r})<0.01,"PASS","FAIL")').font = BOLD
    for col, w in zip("ABCDE", [5, 52, 12, 18, 9]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    wb.save(XLSX)
    print(f"built {XLSX}")


if __name__ == "__main__":
    build()
