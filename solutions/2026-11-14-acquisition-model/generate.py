#!/usr/bin/env python3
"""AcquisitionModel — financial model for a business acquisition.

Built around a real Freelancer brief ("Comprehensive Financial Model
for Business Acquisition", posted ~7h before the 2026-11-14 scout,
19 proposals, avg bid $366 USD, Cupertino US): an Excel financial
model to assess buying / expanding a business, with cash-flow
projections, scenario analysis, defined assumptions and a user-friendly
interface.

Fictional target: Sharma Precision Components Pvt. Ltd., Patna —
precision auto-parts manufacturer. Asking price $2,400,000.
All figures are invented for the demo; no real business data is used.

What generate.py builds from data/*.csv (see make_data.py):
- assumptions.csv  -> Assumptions sheet (deal + valuation inputs)
- historical.csv   -> Historical sheet (2023-2025 actuals, live formulas)
- scenarios.csv    -> Scenarios sheet (Base / Bull / Bear drivers)
- acquisition-model.xlsx — formula-driven workbook:
  Guide, Assumptions, Historical, Scenarios, Projections (2025 actuals
  + 2026-2030, scenario selector drives growth / margin / opex),
  Valuation (DCF: FCF, discount factors, terminal value, enterprise
  value, equity value, implied price/share vs asking price, verdict,
  plus a WACC x terminal-growth sensitivity grid), Dashboard (KPIs +
  3 native charts), Checks (10 integrity checks).

Honest outcome of the demo numbers: base-case equity value comes out
well below the $2.4M asking price, so the verdict reads PASS — the
model is the point, not a flattering answer.

Conventions (LibreOffice-safe, as proven on earlier builds): plain A1
ranges with the column letter repeated on both ends, real date values
(never text dates), percentages stored as 0-1 fractions, numeric
columns guarded to 0 instead of "", LO-safe lookups via INDEX/MATCH
(never MAXIFS, never SMALL(IF(...))). Assumption lookups use
INDEX/MATCH on the key column so row order never matters.
"""
import csv
import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT_XLSX = BASE / "acquisition-model.xlsx"

NAVY = "1F3864"
TEAL = "2E75B6"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color=NAVY, size=14)
MONEY_FMT = '#,##0'
MONEY2_FMT = '#,##0.00'
PCT_FMT = '0.0%'
PCT2_FMT = '0.00%'
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")

YEARS_HIST = [2023, 2024, 2025]
YEARS_PROJ = [2026, 2027, 2028, 2029, 2030]


def read_csv(name):
    with open(DATA / name, newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    ws.row_dimensions[row].height = 30


def title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    ws.cell(row=2, column=1, value=subtitle).font = Font(italic=True,
                                                         color="595959")
    return 4


def assume(key):
    """INDEX/MATCH lookup into the Assumptions key/value table (rows 5-14)."""
    return (f'INDEX(Assumptions!$B$5:$B$14,'
            f'MATCH("{key}",Assumptions!$A$5:$A$14,0))')


def build_workbook(assumptions, historical, scenarios):
    asm = {r["key"]: float(r["value"]) for r in assumptions}
    hist = {int(r["year"]): r for r in historical}
    scen = {r["scenario"]: r for r in scenarios}

    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 100
    r = title_block(ws, "AcquisitionModel — Guide",
                    "Sharma Precision Components Pvt. Ltd., Patna  •  "
                    "Financial model for a business acquisition  •  "
                    "Built from the Freelancer brief 'Comprehensive Financial "
                    "Model for Business Acquisition'", 2)
    guide = [
        ("What this is", "A working demo of the brief: a 5-year financial model "
         "to assess buying a business — historical actuals, projected P&L "
         "and free cash flow, three scenarios, a DCF valuation with a "
         "sensitivity grid, and a verdict against the asking price."),
        ("Scenario selector", "Pick Base / Bull / Bear in Projections!B3 — "
         "revenue growth, gross margin and opex growth all switch, and "
         "every projected figure, the DCF and the dashboard follow."),
        ("How to refresh", "1) Edit data/assumptions.csv, data/historical.csv "
         "or data/scenarios.csv. 2) Re-run: python3 make_data.py (only if "
         "you changed the generator inputs) then python3 generate.py. "
         "3) Every projected figure is a live formula, so the model "
         "rebuilds from the new inputs."),
        ("Valuation", "Enterprise value = present value of 2026-2030 free "
         "cash flow at WACC + present value of terminal value "
         "(Gordon growth). Equity value = enterprise value − net debt. "
         "Implied price/share = equity value / shares outstanding."),
        ("Honest outcome", "On the demo numbers the base-case equity value "
         "comes out well below the $2.4M asking price, so the verdict "
         "reads PASS. The model is the deliverable — change the "
         "assumptions and the verdict follows."),
        ("Sheets", "Assumptions — deal + valuation inputs. Historical — "
         "2023-2025 actuals. Scenarios — Base/Bull/Bear drivers. "
         "Projections — 2025 actuals + 2026-2030 P&L and free cash flow. "
         "Valuation — DCF + sensitivity grid. Dashboard — KPIs + charts. "
         "Checks — 10 integrity checks, all must read PASS."),
        ("Fictional data", "Target company, people and all figures are "
         "invented for the demo. No real business data is used."),
    ]
    for label, text in guide:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 42
        r += 1

    # ---------------- Assumptions ----------------
    ws = wb.create_sheet("Assumptions")
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 60
    hr = title_block(ws, "Assumptions — deal + valuation inputs",
                     "Edit the values in column B and re-run generate.py; "
                     "every sheet reads them via INDEX/MATCH on the key.", 3)
    for c, h in enumerate(["key", "value", "what it means"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 3)
    labels = {
        "asking_price_usd": "Seller's asking price for the business",
        "net_debt_usd": "Debt minus cash assumed at close",
        "shares_outstanding": "Shares outstanding (for price/share)",
        "tax_rate": "Corporate tax rate on positive earnings",
        "wacc": "Discount rate for the DCF (WACC)",
        "terminal_growth": "Perpetual growth after 2030 (Gordon growth)",
        "depreciation_usd_per_year": "Annual depreciation (non-cash)",
        "capex_usd_per_year": "Annual maintenance capital expenditure",
        "working_capital_pct_revenue": "Working capital as % of revenue",
        "interest_usd_per_year": "Annual interest on assumed debt",
    }
    pct_keys = {"tax_rate", "wacc", "terminal_growth",
                "working_capital_pct_revenue"}
    for i, (key, val) in enumerate(asm.items()):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=key)
        c = ws.cell(row=r, column=2, value=val)
        c.number_format = PCT_FMT if key in pct_keys else MONEY_FMT
        ws.cell(row=r, column=3, value=labels.get(key, ""))

    # ---------------- Historical ----------------
    ws = wb.create_sheet("Historical")
    ws.sheet_properties.tabColor = TEAL
    for col, w in zip("ABCD", [22, 16, 16, 16]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Historical actuals — 2023-2025",
                     "Revenue, COGS, opex and other items are inputs from "
                     "data/historical.csv; every derived line is a live "
                     "formula. 2025 is the base year for Projections.", 4)
    for c, h in enumerate(["line item", 2023, 2024, 2025], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 4)
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    hist_rows = ["Revenue", "COGS", "Gross profit", "Gross margin %",
                 "Opex", "Other income", "Other expense", "EBITDA",
                 "Depreciation", "EBIT", "Interest", "EBT", "Tax",
                 "Net income"]
    for i, label in enumerate(hist_rows):
        ws.cell(row=hr + 1 + i, column=1, value=label)
    H = hr + 1  # first data row
    for j, y in enumerate(YEARS_HIST):
        col = get_column_letter(2 + j)
        row = hist[y]
        ws.cell(row=H, column=2 + j, value=int(row["revenue"])).number_format = MONEY_FMT
        ws.cell(row=H + 1, column=2 + j, value=int(row["cogs"])).number_format = MONEY_FMT
        c = ws.cell(row=H + 2, column=2 + j, value=f"={col}{H}-{col}{H+1}")
        c.number_format = MONEY_FMT
        c = ws.cell(row=H + 3, column=2 + j,
                   value=f"=IF({col}{H}=0,0,{col}{H+2}/{col}{H})")
        c.number_format = PCT_FMT
        ws.cell(row=H + 4, column=2 + j, value=int(row["opex"])).number_format = MONEY_FMT
        ws.cell(row=H + 5, column=2 + j, value=int(row["other_income"])).number_format = MONEY_FMT
        ws.cell(row=H + 6, column=2 + j, value=int(row["other_expense"])).number_format = MONEY_FMT
        c = ws.cell(row=H + 7, column=2 + j,
                   value=f"={col}{H+2}-{col}{H+4}+{col}{H+5}-{col}{H+6}")
        c.number_format = MONEY_FMT
        c = ws.cell(row=H + 8, column=2 + j, value=f"={assume('depreciation_usd_per_year')}")
        c.number_format = MONEY_FMT
        c = ws.cell(row=H + 9, column=2 + j, value=f"={col}{H+7}-{col}{H+8}")
        c.number_format = MONEY_FMT
        c = ws.cell(row=H + 10, column=2 + j, value=f"={assume('interest_usd_per_year')}")
        c.number_format = MONEY_FMT
        c = ws.cell(row=H + 11, column=2 + j, value=f"={col}{H+9}-{col}{H+10}")
        c.number_format = MONEY_FMT
        c = ws.cell(row=H + 12, column=2 + j,
                   value=f"=IF({col}{H+11}>0,{col}{H+11}*{assume('tax_rate')},0)")
        c.number_format = MONEY_FMT
        c = ws.cell(row=H + 13, column=2 + j, value=f"={col}{H+11}-{col}{H+12}")
        c.number_format = MONEY_FMT
    for i in (0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13):
        ws.cell(row=H + i, column=1).font = Font(
            bold=True) if i in (2, 7, 9, 11, 13) else Font()
    hist_rev = f"Historical!$B${H}:$D${H}"
    hist_2025_rev = f"Historical!$D${H}"
    hist_2025_opex = f"Historical!$D${H+4}"

    # ---------------- Scenarios ----------------
    ws = wb.create_sheet("Scenarios")
    ws.sheet_properties.tabColor = TEAL
    for col, w in zip("ABCD", [14, 18, 16, 16]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Scenarios — Base / Bull / Bear drivers",
                     "Projections!B3 picks one row; its growth, margin and "
                     "opex growth drive every projected year.", 4)
    for c, h in enumerate(["scenario", "revenue_growth", "gross_margin",
                           "opex_growth"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 4)
    for i, s in enumerate(scenarios):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=s["scenario"])
        for j, key in enumerate(("revenue_growth", "gross_margin",
                                 "opex_growth")):
            c = ws.cell(row=r, column=2 + j, value=float(s[key]))
            c.number_format = PCT_FMT
    scen_rng = f"Scenarios!$A${hr+1}:$A${hr+3}"
    scen_growth = f"Scenarios!$B${hr+1}:$B${hr+3}"
    scen_margin = f"Scenarios!$C${hr+1}:$C${hr+3}"
    scen_opex = f"Scenarios!$D${hr+1}:$D${hr+3}"

    # ---------------- Projections ----------------
    ws = wb.create_sheet("Projections")
    ws.sheet_properties.tabColor = NAVY
    for col, w in zip("ABCDEFGHIJ", [34, 16, 16, 16, 16, 16, 16, 14, 16, 14]):
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:J1")
    ws.cell(row=1, column=1,
           value="Projections — 2025 actuals + 2026-2030").font = TITLE_FONT
    ws.merge_cells("A2:J2")
    ws.cell(row=2, column=1,
           value="Pick a scenario in B3; revenue growth, gross margin and "
                 "opex growth switch and every year follows.").font = \
        Font(italic=True, color="595959")
    ws.cell(row=3, column=1, value="Scenario (pick one):").font = Font(bold=True)
    ws.cell(row=3, column=2, value="Base")
    ws.cell(row=3, column=5, value="Revenue growth").font = Font(bold=True)
    ws.cell(row=3, column=6,
           value=f"=INDEX({scen_growth},MATCH($B$3,{scen_rng},0))").number_format = PCT_FMT
    ws.cell(row=3, column=7, value="Gross margin").font = Font(bold=True)
    ws.cell(row=3, column=8,
           value=f"=INDEX({scen_margin},MATCH($B$3,{scen_rng},0))").number_format = PCT_FMT
    ws.cell(row=3, column=9, value="Opex growth").font = Font(bold=True)
    ws.cell(row=3, column=10,
           value=f"=INDEX({scen_opex},MATCH($B$3,{scen_rng},0))").number_format = PCT_FMT
    dv = DataValidation(type="list", formula1=scen_rng, allow_blank=False)
    dv.error = "Pick Base, Bull or Bear"
    ws.add_data_validation(dv)
    dv.add(ws["B3"])
    hr = 5
    for c, h in enumerate(["line item", 2025] + YEARS_PROJ, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 7)
    ws.freeze_panes = ws.cell(row=hr + 1, column=2)
    proj_rows = ["Revenue", "COGS", "Gross profit", "Gross margin %",
                 "Opex", "EBITDA", "Depreciation", "EBIT", "Interest",
                 "EBT", "Tax", "Net income", "NOPAT", "Add: Depreciation",
                 "Less: Capex", "Less: Change in working capital",
                 "Free cash flow", "Working capital",
                 "Change in working capital"]
    P = hr + 1  # first data row (6)
    for i, label in enumerate(proj_rows):
        ws.cell(row=P + i, column=1, value=label)
    # 2025 column B — linked to Historical 2025 (column D).
    # Explicit row map: projections row -> historical row.
    explicit_2025 = {0: H, 1: H + 1, 4: H + 4, 6: H + 8, 8: H + 10}
    for i, hrow in explicit_2025.items():
        c = ws.cell(row=P + i, column=2, value=f"=Historical!D{hrow}")
        c.number_format = MONEY_FMT
    # Column B computed rows:
    ws.cell(row=P + 2, column=2, value=f"=B{P}-B{P+1}").number_format = MONEY_FMT
    ws.cell(row=P + 3, column=2, value=f"=IF(B{P}=0,0,B{P+2}/B{P})").number_format = PCT_FMT
    ws.cell(row=P + 5, column=2, value=f"=B{P+2}-B{P+4}").number_format = MONEY_FMT
    ws.cell(row=P + 7, column=2, value=f"=B{P+5}-B{P+6}").number_format = MONEY_FMT
    ws.cell(row=P + 9, column=2, value=f"=B{P+7}-B{P+8}").number_format = MONEY_FMT
    ws.cell(row=P + 10, column=2,
           value=f"=IF(B{P+9}>0,B{P+9}*{assume('tax_rate')},0)").number_format = MONEY_FMT
    ws.cell(row=P + 11, column=2, value=f"=B{P+9}-B{P+10}").number_format = MONEY_FMT
    ws.cell(row=P + 12, column=2,
           value=f"=B{P+7}*(1-{assume('tax_rate')})").number_format = MONEY_FMT
    ws.cell(row=P + 13, column=2, value=f"=B{P+6}").number_format = MONEY_FMT
    ws.cell(row=P + 14, column=2, value=f"={assume('capex_usd_per_year')}").number_format = MONEY_FMT
    ws.cell(row=P + 15, column=2, value=0).number_format = MONEY_FMT  # base year: no change
    ws.cell(row=P + 16, column=2,
           value=f"=B{P+12}+B{P+13}-B{P+14}-B{P+15}").number_format = MONEY_FMT
    ws.cell(row=P + 17, column=2,
           value=f"=B{P}*{assume('working_capital_pct_revenue')}").number_format = MONEY_FMT
    ws.cell(row=P + 18, column=2, value=0).number_format = MONEY_FMT
    # Year columns C..G (2026..2030)
    for j in range(5):
        col = get_column_letter(3 + j)
        prev = get_column_letter(2 + j)
        ws.cell(row=P, column=3 + j,
               value=f"={prev}{P}*(1+$F$3)").number_format = MONEY_FMT
        ws.cell(row=P + 1, column=3 + j,
               value=f"={col}{P}*(1-$H$3)").number_format = MONEY_FMT
        ws.cell(row=P + 2, column=3 + j,
               value=f"={col}{P}-{col}{P+1}").number_format = MONEY_FMT
        ws.cell(row=P + 3, column=3 + j,
               value=f"=IF({col}{P}=0,0,{col}{P+2}/{col}{P})").number_format = PCT_FMT
        ws.cell(row=P + 4, column=3 + j,
               value=f"={prev}{P+4}*(1+$J$3)").number_format = MONEY_FMT
        ws.cell(row=P + 5, column=3 + j,
               value=f"={col}{P+2}-{col}{P+4}").number_format = MONEY_FMT
        ws.cell(row=P + 6, column=3 + j,
               value=f"={assume('depreciation_usd_per_year')}").number_format = MONEY_FMT
        ws.cell(row=P + 7, column=3 + j,
               value=f"={col}{P+5}-{col}{P+6}").number_format = MONEY_FMT
        ws.cell(row=P + 8, column=3 + j,
               value=f"={assume('interest_usd_per_year')}").number_format = MONEY_FMT
        ws.cell(row=P + 9, column=3 + j,
               value=f"={col}{P+7}-{col}{P+8}").number_format = MONEY_FMT
        ws.cell(row=P + 10, column=3 + j,
               value=f"=IF({col}{P+9}>0,{col}{P+9}*{assume('tax_rate')},0)").number_format = MONEY_FMT
        ws.cell(row=P + 11, column=3 + j,
               value=f"={col}{P+9}-{col}{P+10}").number_format = MONEY_FMT
        ws.cell(row=P + 12, column=3 + j,
               value=f"={col}{P+7}*(1-{assume('tax_rate')})").number_format = MONEY_FMT
        ws.cell(row=P + 13, column=3 + j,
               value=f"={col}{P+6}").number_format = MONEY_FMT
        ws.cell(row=P + 14, column=3 + j,
               value=f"={assume('capex_usd_per_year')}").number_format = MONEY_FMT
        ws.cell(row=P + 15, column=3 + j,
               value=f"={col}{P+18}").number_format = MONEY_FMT
        ws.cell(row=P + 16, column=3 + j,
               value=f"={col}{P+12}+{col}{P+13}-{col}{P+14}-{col}{P+15}").number_format = MONEY_FMT
        ws.cell(row=P + 17, column=3 + j,
               value=f"={col}{P}*{assume('working_capital_pct_revenue')}").number_format = MONEY_FMT
        ws.cell(row=P + 18, column=3 + j,
               value=f"={col}{P+17}-{prev}{P+17}").number_format = MONEY_FMT
    for i in (2, 5, 7, 9, 11, 16):
        ws.cell(row=P + i, column=1).font = Font(bold=True)
    ws.conditional_formatting.add(
        f"B{P+16}:G{P+16}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
    proj_rev = f"Projections!$B${P}:$G${P}"
    proj_ebitda = f"Projections!$B${P+5}:$G${P+5}"
    proj_fcf = f"Projections!$C${P+16}:$G${P+16}"
    proj_years = f"Projections!$B${hr}:$G${hr}"
    fcf_2030 = f"Projections!$G${P+16}"
    rev_2030 = f"Projections!$G${P}"
    ebitda_2030 = f"Projections!$G${P+5}"

    # ---------------- Valuation ----------------
    ws = wb.create_sheet("Valuation")
    ws.sheet_properties.tabColor = NAVY
    for col, w in zip("ABCDEFGHIJKL", [34, 16, 16, 16, 16, 16, 6, 14, 14, 14, 14, 14]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Valuation — discounted cash flow",
                     "Free cash flow discounted at WACC; terminal value by "
                     "Gordon growth. Scenario is the one picked in "
                     "Projections!B3.", 6)
    for c, h in enumerate(["line item", 2026, 2027, 2028, 2029, 2030], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 6)
    V = hr + 1  # first data row
    ws.cell(row=V, column=1, value="Free cash flow")
    for j in range(5):
        col = get_column_letter(2 + j)
        c = ws.cell(row=V, column=2 + j,
                   value=f"=Projections!{get_column_letter(3 + j)}{P+16}")
        c.number_format = MONEY_FMT
    ws.cell(row=V + 1, column=1, value="Discount factor")
    for j in range(5):
        c = ws.cell(row=V + 1, column=2 + j,
                   value=f"=1/(1+{assume('wacc')})^{j+1}")
        c.number_format = '0.0000'
    ws.cell(row=V + 2, column=1, value="Present value of FCF")
    for j in range(5):
        col = get_column_letter(2 + j)
        c = ws.cell(row=V + 2, column=2 + j, value=f"={col}{V}*{col}{V+1}")
        c.number_format = MONEY_FMT
    ws.cell(row=V + 4, column=1, value="Terminal value (end 2030)")
    c = ws.cell(row=V + 4, column=2,
               value=f"=F{V}*(1+{assume('terminal_growth')})"
                     f"/({assume('wacc')}-{assume('terminal_growth')})")
    c.number_format = MONEY_FMT
    ws.cell(row=V + 5, column=1, value="Present value of terminal value")
    c = ws.cell(row=V + 5, column=2, value=f"=B{V+4}*F{V+1}")
    c.number_format = MONEY_FMT
    ws.cell(row=V + 6, column=1, value="Enterprise value").font = Font(bold=True)
    c = ws.cell(row=V + 6, column=2,
               value=f"=SUM(B{V+2}:F{V+2})+B{V+5}")
    c.number_format = MONEY_FMT
    c.font = Font(bold=True)
    ws.cell(row=V + 7, column=1, value="Less: Net debt")
    c = ws.cell(row=V + 7, column=2, value=f"={assume('net_debt_usd')}")
    c.number_format = MONEY_FMT
    ws.cell(row=V + 8, column=1, value="Equity value").font = Font(bold=True)
    c = ws.cell(row=V + 8, column=2, value=f"=B{V+6}-B{V+7}")
    c.number_format = MONEY_FMT
    c.font = Font(bold=True)
    ws.cell(row=V + 9, column=1, value="Shares outstanding")
    c = ws.cell(row=V + 9, column=2, value=f"={assume('shares_outstanding')}")
    c.number_format = MONEY_FMT
    ws.cell(row=V + 10, column=1, value="Implied price per share").font = Font(bold=True)
    c = ws.cell(row=V + 10, column=2,
               value=f"=IF(B{V+9}=0,0,B{V+8}/B{V+9})")
    c.number_format = MONEY2_FMT
    c.font = Font(bold=True)
    ws.cell(row=V + 11, column=1, value="Asking price")
    c = ws.cell(row=V + 11, column=2, value=f"={assume('asking_price_usd')}")
    c.number_format = MONEY_FMT
    ws.cell(row=V + 12, column=1, value="Asking price per share")
    c = ws.cell(row=V + 12, column=2,
               value=f"=IF(B{V+9}=0,0,B{V+11}/B{V+9})")
    c.number_format = MONEY2_FMT
    ws.cell(row=V + 13, column=1, value="Premium / (discount) vs asking")
    c = ws.cell(row=V + 13, column=2,
               value=f"=IF(B{V+11}=0,0,(B{V+8}-B{V+11})/B{V+11})")
    c.number_format = PCT_FMT
    ws.cell(row=V + 14, column=1, value="Verdict").font = Font(bold=True)
    c = ws.cell(row=V + 14, column=2,
               value=f'=IF(B{V+8}>=B{V+11},'
                     f'"ACQUIRE — model value covers the asking price",'
                     f'"PASS — model value is below the asking price")')
    c.font = Font(bold=True)
    # Sensitivity grid
    S = V + 16
    ws.merge_cells(start_row=S, start_column=1, end_row=S, end_column=6)
    ws.cell(row=S, column=1,
           value="Sensitivity — enterprise value (USD) by WACC × terminal "
                 "growth").font = Font(bold=True, color=NAVY, size=12)
    ws.cell(row=S + 1, column=1, value="WACC \\ g").font = HDR_FONT
    ws.cell(row=S + 1, column=1).fill = HDR_FILL
    g_vals = [0.015, 0.02, 0.025, 0.03, 0.035]
    wacc_vals = [0.09, 0.10, 0.11, 0.12, 0.13]
    for j, g in enumerate(g_vals):
        c = ws.cell(row=S + 1, column=2 + j, value=g)
        c.number_format = PCT_FMT
        c.font = HDR_FONT
        c.fill = HDR_FILL
    for j in range(5):
        ws.cell(row=S + 1, column=8 + j,
               value=f"PV FCF {YEARS_PROJ[j]} (helper)").font = Font(
                   italic=True, color="595959")
    for i, w in enumerate(wacc_vals):
        r = S + 2 + i
        c = ws.cell(row=r, column=1, value=w)
        c.number_format = PCT_FMT
        c.font = HDR_FONT
        c.fill = HDR_FILL
        for j in range(5):
            fcol = get_column_letter(2 + j)
            c = ws.cell(row=r, column=8 + j,
                       value=f"=${fcol}${V}/(1+$A{r})^{j+1}")
            c.number_format = MONEY_FMT
        for j, g in enumerate(g_vals):
            gcol = get_column_letter(2 + j)
            c = ws.cell(row=r, column=2 + j,
                       value=f"=SUM($H{r}:$L{r})+"
                             f"($F${V}*(1+{gcol}${S+1})/($A{r}-{gcol}${S+1}))"
                             f"/(1+$A{r})^5")
            c.number_format = MONEY_FMT
    ws.conditional_formatting.add(
        f"B{V+13}:B{V+13}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
    ev_cell = f"Valuation!$B${V+6}"
    eq_cell = f"Valuation!$B${V+8}"
    pps_cell = f"Valuation!$B${V+10}"
    ask_cell = f"Valuation!$B${V+11}"
    prem_cell = f"Valuation!$B${V+13}"
    verdict_cell = f"Valuation!$B${V+14}"
    sens_base = f"Valuation!$D${S+4}"  # WACC 11% row, g 2.5% column
    val_fcf = f"Valuation!$B${V}:$F${V}"
    val_pv = f"Valuation!$B${V+2}:$F${V+2}"
    val_years = f"Valuation!$B${hr}:$F${hr}"

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 60
    r = title_block(ws, "Acquisition dashboard — Sharma Precision Components",
                    "Base scenario unless Projections!B3 says otherwise  •  "
                    "Asking price $2,400,000  •  All KPIs are live formulas.",
                    2)
    ws.cell(row=r, column=1, value="KPI").font = HDR_FONT
    ws.cell(row=r, column=1).fill = HDR_FILL
    ws.cell(row=r, column=2, value="Value").font = HDR_FONT
    ws.cell(row=r, column=2).fill = HDR_FILL
    kpis = [
        ("Scenario in force", "=Projections!$B$3", None),
        ("Asking price", f"={ask_cell}", MONEY_FMT),
        ("Enterprise value (DCF)", f"={ev_cell}", MONEY_FMT),
        ("Equity value", f"={eq_cell}", MONEY_FMT),
        ("Implied price per share", f"={pps_cell}", MONEY2_FMT),
        ("Asking price per share", f"={ask_cell}/"
         f"{assume('shares_outstanding')}", MONEY2_FMT),
        ("Premium / (discount) vs asking", f"={prem_cell}", PCT_FMT),
        ("Verdict", f"={verdict_cell}", None),
        ("2030 revenue", f"={rev_2030}", MONEY_FMT),
        ("2030 EBITDA", f"={ebitda_2030}", MONEY_FMT),
        ("2030 EBITDA margin",
         f"=IF({rev_2030}=0,0,{ebitda_2030}/{rev_2030})", PCT_FMT),
        ("Average annual FCF 2026-2030", f"=AVERAGE({val_fcf})", MONEY_FMT),
        ("Cumulative FCF 2026-2030", f"=SUM({val_fcf})", MONEY_FMT),
    ]
    kpi_first = r + 1
    for i, (label, formula, fmt) in enumerate(kpis):
        rr = kpi_first + i
        ws.cell(row=rr, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=rr, column=2, value=formula)
        if fmt:
            c.number_format = fmt
    # Scenario 2030-revenue table for chart 3
    t = kpi_first + len(kpis) + 2
    ws.cell(row=t, column=1, value="2030 revenue by scenario").font = \
        Font(bold=True, color=NAVY, size=11)
    ws.cell(row=t + 1, column=1, value="Scenario").font = HDR_FONT
    ws.cell(row=t + 1, column=1).fill = HDR_FILL
    ws.cell(row=t + 1, column=2, value="2030 revenue").font = HDR_FONT
    ws.cell(row=t + 1, column=2).fill = HDR_FILL
    for i, name in enumerate(("Base", "Bull", "Bear")):
        rr = t + 2 + i
        ws.cell(row=rr, column=1, value=name)
        c = ws.cell(row=rr, column=2,
                   value=f"={hist_2025_rev}*(1+INDEX({scen_growth},"
                         f'MATCH("{name}",{scen_rng},0)))^5')
        c.number_format = MONEY_FMT
    scen_tbl = f"Dashboard!$A${t+2}:$B${t+4}"
    scen_cats = f"Dashboard!$A${t+2}:$A${t+4}"
    scen_vals = f"Dashboard!$B${t+2}:$B${t+4}"
    # Charts (data lives in rows: one series per row, years across columns)
    col_chart = BarChart()
    col_chart.type = "col"
    col_chart.title = "Revenue vs EBITDA (2025-2030)"
    col_chart.y_axis.title = "USD"
    col_chart.x_axis.title = "Year"
    col_chart.style = 10
    col_chart.add_data(Reference(wb["Projections"], min_col=1, min_row=P,
                                 max_row=P), from_rows=True,
                       titles_from_data=True)
    col_chart.add_data(Reference(wb["Projections"], min_col=1, min_row=P + 5,
                                 max_row=P + 5), from_rows=True,
                       titles_from_data=True)
    col_chart.set_categories(
        Reference(wb["Projections"], min_col=2, min_row=hr, max_col=7))
    col_chart.shape = 4
    ws.add_chart(col_chart, "A28")
    line_chart = LineChart()
    line_chart.title = "Free cash flow 2026-2030"
    line_chart.y_axis.title = "USD"
    line_chart.style = 10
    line_chart.add_data(Reference(wb["Valuation"], min_col=1, min_row=V,
                                  max_row=V), from_rows=True,
                        titles_from_data=True)
    line_chart.set_categories(
        Reference(wb["Valuation"], min_col=2, min_row=hr, max_col=6))
    ws.add_chart(line_chart, "A43")
    bar_chart = BarChart()
    bar_chart.type = "col"
    bar_chart.title = "2030 revenue by scenario"
    bar_chart.y_axis.title = "USD"
    bar_chart.style = 10
    bar_chart.add_data(Reference(ws, min_col=2, min_row=t + 1,
                                 max_row=t + 4), titles_from_data=True)
    bar_chart.set_categories(
        Reference(ws, min_col=1, min_row=t + 2, max_row=t + 4))
    bar_chart.shape = 4
    ws.add_chart(bar_chart, "A58")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws.sheet_properties.tabColor = "C00000"
    for col, w in zip("ABCD", [46, 18, 18, 10]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Integrity checks — all must read PASS",
                     "Expected is computed in Python; Actual is a live formula. "
                     "Money comparisons allow 0.01 tolerance.", 4)
    for c, h in enumerate(["check", "expected", "actual", "result"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 4)
    exp_rev26 = float(hist[2025]["revenue"]) * (1 + float(scen["Base"]["revenue_growth"]))
    checks = [
        ("Historical years loaded", 3,
         f"=COUNT(Historical!$B${H}:$D${H})"),
        ("2025 revenue", 4200000, f"=Historical!$D${H}"),
        ("2026 revenue = 2025 x (1+growth)", round(exp_rev26, 2),
         f"=Projections!$C${P}"),
        ("2026 gross profit = revenue - COGS", 0,
         f"=Projections!C{P+2}-(Projections!C{P}-Projections!C{P+1})"),
        ("2026 EBITDA = gross profit - opex", 0,
         f"=Projections!C{P+5}-(Projections!C{P+2}-Projections!C{P+4})"),
        ("2026 FCF = NOPAT + dep - capex - dWC", 0,
         f"=Projections!C{P+16}-(Projections!C{P+12}+Projections!C{P+13}"
         f"-Projections!C{P+14}-Projections!C{P+15})"),
        ("Enterprise value = PV FCF + PV TV", 0,
         f"=Valuation!B{V+6}-(SUM(Valuation!B{V+2}:F{V+2})"
         f"+Valuation!B{V+5})"),
        ("Equity value = EV - net debt", 0,
         f"=Valuation!B{V+8}-(Valuation!B{V+6}-Valuation!B{V+7})"),
        ("Sensitivity base cell (11% x 2.5%) = EV", 0,
         f"={sens_base}-Valuation!B{V+6}"),
        ("2030 working capital = 12% x 2030 revenue", 0,
         f"=Projections!G{P+17}-(Projections!G{P}"
         f"*{assume('working_capital_pct_revenue')})"),
    ]
    for i, (label, expected, actual) in enumerate(checks):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=label)
        c = ws.cell(row=r, column=2, value=expected)
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=3, value=actual)
        c.number_format = MONEY_FMT
        ws.cell(row=r, column=4,
               value=f'=IF(ABS(B{r}-C{r})<0.01,"PASS","FAIL")')
    return wb


def main():
    assumptions = read_csv("assumptions.csv")
    historical = read_csv("historical.csv")
    scenarios = read_csv("scenarios.csv")
    wb = build_workbook(assumptions, historical, scenarios)
    wb.save(OUT_XLSX)
    rev25 = int([r for r in historical if r["year"] == "2025"][0]["revenue"])
    print(f"built {OUT_XLSX.name}: 2025 revenue {rev25:,}, "
          f"asking ${int(assumptions[0]['value']):,}")


if __name__ == "__main__":
    main()
