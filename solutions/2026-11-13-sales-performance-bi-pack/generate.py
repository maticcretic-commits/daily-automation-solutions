#!/usr/bin/env python3
"""SalesPerf BI Pack — interactive sales-performance dashboard factory.

Built around a real Freelancer brief ("Interactive Sales Performance BI
Dashboard", Freelancer ID 40815164, posted ~12h before the 2026-11-13
scout, 1 proposal, avg INR 25,000): build a brand-new real-time
sales/business-performance dashboard from scratch — revenue, monthly /
quarterly targets, % to goal, products, orders and customer purchasing,
with date/category filters and CSV export; Power BI, Tableau or Looker
Studio, with refresh instructions.

What this builds, from fictional Agarwal Traders, Patna data (FMCG
wholesaler, 2026-01-01 .. 2026-10-31, deterministic data/*.csv, seed
20261113):
1. data/monthly_export.csv + data/category_export.csv — the brief's CSV
   export: the Monthly and By_Category summaries as flat files.
2. sales-performance-bi-pack.xlsx — a formula-driven workbook:
   Guide, Orders (live Amount / Category / Month / Quarter formulas),
   Products + Customers masters, Targets, Monthly (COUNTIFS/SUMIFS vs
   target + % to goal with red/amber/green conditional formatting),
   Quarterly (SUMIFS over Monthly + months-on-target), By_Product /
   By_Customer / By_Region / By_Category (SUMIFS + share), Dashboard
   (working month + category filter cells driving SUMIFS KPIs, 13 KPIs,
   3 native charts), Checks (10 integrity checks).

Demo targets: data/targets.csv sets each month's target as a fixed
percentage of that month's actual revenue (factors 0.92–1.30, see
make_data.py) so every attainment band shows on the dashboard; in
production the client supplies real targets. The method is stated in
make_data.py and the README — targets are an input file, never computed
by the workbook.

Conventions (LibreOffice-safe, as proven on earlier builds): plain A1
ranges with the column letter repeated on both ends, real date values
(never text dates), percentages stored as 0-1 fractions, numeric columns
guarded to 0 instead of "", LO-safe lookups via INDEX/MATCH (never
MAXIFS, never SMALL(IF(...))).
"""
import csv
import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule  # noqa: F401
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT_XLSX = BASE / "sales-performance-bi-pack.xlsx"
OUT_MONTHLY_EXPORT = DATA / "monthly_export.csv"
OUT_CATEGORY_EXPORT = DATA / "category_export.csv"

NAVY = "1F3864"
TEAL = "2E75B6"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color=NAVY, size=14)
MONEY_FMT = '#,##0.00'
PCT_FMT = '0.0%'
DATE_FMT = 'yyyy-mm-dd'
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
AMBER_FILL = PatternFill("solid", fgColor="FFEB9C")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")


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
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = Font(italic=True, color="595959")
    return 4  # first data/header row


def build_workbook(orders, products, customers, targets):
    n = len(orders)
    months = sorted(targets)
    prod_names = [p["product"] for p in products]
    cust_names = [c["customer"] for c in customers]
    regions = sorted({c["region"] for c in customers})
    categories = sorted({p["category"] for p in products})
    quarters = ["Q1 2026", "Q2 2026", "Q3 2026", "Q4 2026"]
    # Orders rows: header 4, data 5..4+n ; Products 5..4+len ; etc.
    o_first, o_last = 5, 4 + n
    p_first, p_last = 5, 4 + len(products)
    c_first, c_last = 5, 4 + len(customers)
    t_first, t_last = 5, 4 + len(months)
    mo_first, mo_last = 5, 4 + len(months)
    mo_total = mo_last + 1
    q_first, q_last = 5, 4 + len(quarters)
    q_total = q_last + 1
    bp_first, bp_last = 5, 4 + len(prod_names)
    bp_total = bp_last + 1
    bc_first, bc_last = 5, 4 + len(cust_names)
    bc_total = bc_last + 1
    br_first, br_last = 5, 4 + len(regions)
    br_total = br_last + 1
    bcat_first, bcat_last = 5, 4 + len(categories)
    bcat_total = bcat_last + 1

    ord_amt = f"Orders!$J${o_first}:$J${o_last}"
    ord_mon = f"Orders!$K${o_first}:$K${o_last}"
    ord_qtr = f"Orders!$L${o_first}:$L${o_last}"
    ord_prod = f"Orders!$F${o_first}:$F${o_last}"
    ord_cust = f"Orders!$C${o_first}:$C${o_last}"
    ord_reg = f"Orders!$D${o_first}:$D${o_last}"
    ord_cat = f"Orders!$G${o_first}:$G${o_last}"
    ord_qty = f"Orders!$H${o_first}:$H${o_last}"
    prod_rng = f"Products!$A${p_first}:$A${p_last}"
    prod_cat = f"Products!$B${p_first}:$B${p_last}"
    cust_rng = f"Customers!$A${c_first}:$A${c_last}"
    cust_reg = f"Customers!$B${c_first}:$B${c_last}"
    tgt_mon = f"Targets!$A${t_first}:$A${t_last}"
    tgt_val = f"Targets!$B${t_first}:$B${t_last}"

    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 95
    r = title_block(ws, "SalesPerf BI Pack — Guide",
                    "Agarwal Traders, Patna (FMCG wholesaler)  •  Jan – Oct 2026  •  "
                    "Built from the Freelancer brief 'Interactive Sales Performance "
                    "BI Dashboard'", 2)
    guide = [
        ("What this is", "A working demo of the brief: a real-time-style sales performance "
         "dashboard — revenue, monthly/quarterly targets, % to goal, products, "
         "orders and customer purchasing, with working date/category filters "
         "and CSV export of the Monthly and By_Category summaries."),
        ("Filters (Dashboard)", "Pick a Focus month and a Focus category on the Dashboard — "
         "the three filtered-revenue KPIs are live SUMIFS over Orders, so they "
         "follow the selection. Every other sheet has AutoFilter for ad-hoc "
         "date/category slicing."),
        ("How to refresh", "1) Append the new month's rows to data/orders.csv and its target "
         "to data/targets.csv. 2) Re-run: python3 generate.py — it rebuilds this "
         "workbook and re-exports the two CSVs. 3) Every figure on Monthly / "
         "Quarterly / By_* / Dashboard is a live formula, so the dashboard "
         "updates from the rebuilt Orders sheet."),
        ("Attainment bands", "Monthly % to goal is conditionally formatted: green >= 100%, "
         "amber 80–99%, red < 80%. Demo targets are set as a fixed percentage of "
         "each month's actual revenue (see make_data.py) so all three bands show; "
         "in production the client supplies real targets."),
        ("CSV export", "data/monthly_export.csv (month, orders, revenue, target, % to goal, "
         "avg order, quarter) and data/category_export.csv (category, orders, "
         "revenue, share) are written by generate.py from the same aggregates "
         "the workbook formulas compute — the test asserts they agree."),
        ("Sheets", "Orders — raw rows (live Amount / Category / Month / Quarter formulas). "
         "Products / Customers — masters. Targets — monthly revenue targets. "
         "Monthly / Quarterly — SUMIFS vs target. By_Product / By_Customer / "
         "By_Region / By_Category — SUMIFS + share. Dashboard — filters, KPIs, "
         "3 charts. Checks — 10 integrity checks, all must read PASS."),
        ("Fictional data", "All customers, products and figures are invented for the demo. "
         "No real business data is used."),
    ]
    for label, text in guide:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 42
        r += 1

    # ---------------- Orders ----------------
    ws = wb.create_sheet("Orders")
    ws.sheet_properties.tabColor = "C00000"
    headers = ["order_id", "date", "customer", "region", "channel", "product",
               "category", "qty", "unit_price", "amount", "month", "quarter"]
    for col, w in zip("ABCDEFGHIJKL",
                      [12, 13, 30, 13, 14, 22, 12, 10, 11, 14, 10, 10]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Orders — raw sales rows",
                     f"{n} rows, Jan – Oct 2026. Category, Amount, Month and Quarter "
                     "are live formulas; paste new rows at the bottom and re-run "
                     "generate.py.", 12)
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 12)
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:L{hr + n}"
    for i, row in enumerate(orders):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=row["order_id"])
        d = datetime.date.fromisoformat(row["date"])
        c = ws.cell(row=r, column=2, value=d)
        c.number_format = DATE_FMT
        ws.cell(row=r, column=3, value=row["customer"])
        ws.cell(row=r, column=4, value=row["region"])
        ws.cell(row=r, column=5, value=row["channel"])
        ws.cell(row=r, column=6, value=row["product"])
        ws.cell(row=r, column=7,
               value=f"=INDEX({prod_cat},MATCH($F{r},{prod_rng},0))")
        ws.cell(row=r, column=8, value=int(row["qty"]))
        c = ws.cell(row=r, column=9, value=float(row["unit_price"]))
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=10, value=f"=H{r}*I{r}")
        c.number_format = MONEY_FMT
        ws.cell(row=r, column=11, value=f'=TEXT(B{r},"yyyy-mm")')
        ws.cell(row=r, column=12,
               value=f'="Q"&INT((MONTH(B{r})-1)/3)+1&" "&YEAR(B{r})')

    # ---------------- Products ----------------
    ws = wb.create_sheet("Products")
    ws.sheet_properties.tabColor = TEAL
    for col, w in zip("ABC", [22, 12, 12]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Products — master list",
                     "Category is looked up from here by Orders (INDEX/MATCH).", 3)
    for c, h in enumerate(["product", "category", "base_price"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 3)
    for i, p in enumerate(products):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=p["product"])
        ws.cell(row=r, column=2, value=p["category"])
        c = ws.cell(row=r, column=3, value=float(p["base_price"]))
        c.number_format = MONEY_FMT

    # ---------------- Customers ----------------
    ws = wb.create_sheet("Customers")
    ws.sheet_properties.tabColor = TEAL
    for col, w in zip("ABC", [32, 13, 14]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Customers — master list",
                     "Region is looked up from here by By_Customer (INDEX/MATCH).", 3)
    for c, h in enumerate(["customer", "region", "channel"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 3)
    for i, cu in enumerate(customers):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=cu["customer"])
        ws.cell(row=r, column=2, value=cu["region"])
        ws.cell(row=r, column=3, value=cu["channel"])

    # ---------------- Targets ----------------
    ws = wb.create_sheet("Targets")
    ws.sheet_properties.tabColor = NAVY
    for col, w in zip("AB", [12, 16]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Targets — monthly revenue targets",
                     "Demo targets: a fixed percentage of each month's actual revenue "
                     "(see make_data.py). In production the client supplies these.", 2)
    for c, h in enumerate(["month", "target"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 2)
    for i, m in enumerate(months):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=m)
        c = ws.cell(row=r, column=2, value=targets[m])
        c.number_format = MONEY_FMT

    # ---------------- Monthly ----------------
    ws = wb.create_sheet("Monthly")
    ws.sheet_properties.tabColor = NAVY
    for col, w in zip("ABCDEFG", [12, 10, 16, 14, 12, 12, 10]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Monthly sales vs target",
                     "Orders and revenue are COUNTIFS / SUMIFS over Orders. "
                     "Attainment = revenue / target.", 7)
    for c, h in enumerate(["month", "orders", "revenue", "target",
                           "pct_to_goal", "avg_order", "quarter"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 7)
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    for i, m in enumerate(months):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=m)
        ws.cell(row=r, column=2, value=f"=COUNTIFS({ord_mon},$A{r})")
        c = ws.cell(row=r, column=3, value=f"=SUMIFS({ord_amt},{ord_mon},$A{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=4,
                   value=f"=INDEX({tgt_val},MATCH($A{r},{tgt_mon},0))")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=5, value=f"=IF($D{r}=0,0,$C{r}/$D{r})")
        c.number_format = PCT_FMT
        c = ws.cell(row=r, column=6, value=f"=IF($B{r}=0,0,$C{r}/$B{r})")
        c.number_format = MONEY_FMT
        ws.cell(row=r, column=7,
               value=f'="Q"&INT((VALUE(MID($A{r},6,2))-1)/3)+1&" "&LEFT($A{r},4)')
    tr = mo_total
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, fmt in ((2, None), (3, MONEY_FMT), (4, MONEY_FMT)):
        c = ws.cell(row=tr, column=col,
                   value=f"=SUM({get_column_letter(col)}{mo_first}:"
                         f"{get_column_letter(col)}{mo_last})")
        c.font = Font(bold=True)
        if fmt:
            c.number_format = fmt
    c = ws.cell(row=tr, column=5, value=f"=IF($D{tr}=0,0,$C{tr}/$D{tr})")
    c.font = Font(bold=True)
    c.number_format = PCT_FMT
    c = ws.cell(row=tr, column=6, value=f"=IF($B{tr}=0,0,$C{tr}/$B{tr})")
    c.font = Font(bold=True)
    c.number_format = MONEY_FMT
    ws.conditional_formatting.add(
        f"E{mo_first}:E{mo_last}",
        CellIsRule(operator="greaterThanOrEqual", formula=["1"],
                   fill=GREEN_FILL))
    ws.conditional_formatting.add(
        f"E{mo_first}:E{mo_last}",
        CellIsRule(operator="between", formula=["0.8", "0.9999"],
                   fill=AMBER_FILL))
    ws.conditional_formatting.add(
        f"E{mo_first}:E{mo_last}",
        CellIsRule(operator="lessThan", formula=["0.8"], fill=RED_FILL))
    mo = f"Monthly!$A${mo_first}:$A${mo_last}"
    mo_ord = f"Monthly!$B${mo_first}:$B${mo_last}"
    mo_rev = f"Monthly!$C${mo_first}:$C${mo_last}"
    mo_tgt = f"Monthly!$D${mo_first}:$D${mo_last}"
    mo_pct = f"Monthly!$E${mo_first}:$E${mo_last}"
    mo_qtr = f"Monthly!$G${mo_first}:$G${mo_last}"

    # ---------------- Quarterly ----------------
    ws = wb.create_sheet("Quarterly")
    ws.sheet_properties.tabColor = NAVY
    for col, w in zip("ABCDEF", [12, 10, 16, 14, 12, 16]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Quarterly sales vs target",
                     "SUMIFS over Monthly; months-on-target counts Monthly rows "
                     "with attainment >= 100%.", 6)
    for c, h in enumerate(["quarter", "orders", "revenue", "target",
                           "pct_to_goal", "months_on_target"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 6)
    for i, q in enumerate(quarters):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=q)
        ws.cell(row=r, column=2, value=f"=SUMIFS({mo_ord},{mo_qtr},$A{r})")
        c = ws.cell(row=r, column=3, value=f"=SUMIFS({mo_rev},{mo_qtr},$A{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=4, value=f"=SUMIFS({mo_tgt},{mo_qtr},$A{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=5, value=f"=IF($D{r}=0,0,$C{r}/$D{r})")
        c.number_format = PCT_FMT
        ws.cell(row=r, column=6,
               value=f'=COUNTIFS({mo_qtr},$A{r},{mo_pct},">=1")')
    tr = q_total
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, fmt in ((2, None), (3, MONEY_FMT), (4, MONEY_FMT)):
        c = ws.cell(row=tr, column=col,
                   value=f"=SUM({get_column_letter(col)}{q_first}:"
                         f"{get_column_letter(col)}{q_last})")
        c.font = Font(bold=True)
        if fmt:
            c.number_format = fmt
    c = ws.cell(row=tr, column=5, value=f"=IF($D{tr}=0,0,$C{tr}/$D{tr})")
    c.font = Font(bold=True)
    c.number_format = PCT_FMT
    c = ws.cell(row=tr, column=6, value=f"=SUM(F{q_first}:F{q_last})")
    c.font = Font(bold=True)

    # ---------------- By_Product ----------------
    ws = wb.create_sheet("By_Product")
    ws.sheet_properties.tabColor = TEAL
    for col, w in zip("ABCDEFG", [22, 12, 10, 12, 16, 12, 10]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Sales by product",
                     "Orders, quantity and revenue are COUNTIFS / SUMIFS over Orders; "
                     "share is revenue / column total.", 7)
    for c, h in enumerate(["product", "category", "orders", "qty", "revenue",
                           "avg_price", "share"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 7)
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:G{hr + len(prod_names)}"
    for i, p in enumerate(prod_names):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=p)
        ws.cell(row=r, column=2,
               value=f"=INDEX({prod_cat},MATCH($A{r},{prod_rng},0))")
        ws.cell(row=r, column=3, value=f"=COUNTIFS({ord_prod},$A{r})")
        ws.cell(row=r, column=4, value=f"=SUMIFS({ord_qty},{ord_prod},$A{r})")
        c = ws.cell(row=r, column=5, value=f"=SUMIFS({ord_amt},{ord_prod},$A{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=6, value=f"=IF($D{r}=0,0,$E{r}/$D{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=7,
                   value=f"=IF($E${bp_total}=0,0,$E{r}/$E${bp_total})")
        c.number_format = PCT_FMT
    tr = bp_total
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, fmt in ((3, None), (4, None), (5, MONEY_FMT)):
        c = ws.cell(row=tr, column=col,
                   value=f"=SUM({get_column_letter(col)}{bp_first}:"
                         f"{get_column_letter(col)}{bp_last})")
        c.font = Font(bold=True)
        if fmt:
            c.number_format = fmt
    c = ws.cell(row=tr, column=6, value=f"=IF($D{tr}=0,0,$E{tr}/$D{tr})")
    c.font = Font(bold=True)
    c.number_format = MONEY_FMT
    c = ws.cell(row=tr, column=7, value=1)
    c.font = Font(bold=True)
    c.number_format = PCT_FMT
    bp_prod = f"By_Product!$A${bp_first}:$A${bp_last}"
    bp_rev = f"By_Product!$E${bp_first}:$E${bp_last}"

    # ---------------- By_Customer ----------------
    ws = wb.create_sheet("By_Customer")
    ws.sheet_properties.tabColor = TEAL
    for col, w in zip("ABCDEF", [32, 13, 10, 16, 12, 10]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Sales by customer",
                     "Region is looked up from Customers; orders and revenue are "
                     "COUNTIFS / SUMIFS over Orders; share is revenue / column total.",
                     6)
    for c, h in enumerate(["customer", "region", "orders", "revenue",
                           "avg_order", "share"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 6)
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:F{hr + len(cust_names)}"
    for i, cu in enumerate(cust_names):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=cu)
        ws.cell(row=r, column=2,
               value=f"=INDEX({cust_reg},MATCH($A{r},{cust_rng},0))")
        ws.cell(row=r, column=3, value=f"=COUNTIFS({ord_cust},$A{r})")
        c = ws.cell(row=r, column=4, value=f"=SUMIFS({ord_amt},{ord_cust},$A{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=5, value=f"=IF($C{r}=0,0,$D{r}/$C{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=6,
                   value=f"=IF($D${bc_total}=0,0,$D{r}/$D${bc_total})")
        c.number_format = PCT_FMT
    tr = bc_total
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, fmt in ((3, None), (4, MONEY_FMT)):
        c = ws.cell(row=tr, column=col,
                   value=f"=SUM({get_column_letter(col)}{bc_first}:"
                         f"{get_column_letter(col)}{bc_last})")
        c.font = Font(bold=True)
        if fmt:
            c.number_format = fmt
    c = ws.cell(row=tr, column=5, value=f"=IF($C{tr}=0,0,$D{tr}/$C{tr})")
    c.font = Font(bold=True)
    c.number_format = MONEY_FMT
    c = ws.cell(row=tr, column=6, value=1)
    c.font = Font(bold=True)
    c.number_format = PCT_FMT
    bc_cust = f"By_Customer!$A${bc_first}:$A${bc_last}"
    bc_rev = f"By_Customer!$D${bc_first}:$D${bc_last}"

    # ---------------- By_Region ----------------
    ws = wb.create_sheet("By_Region")
    ws.sheet_properties.tabColor = TEAL
    for col, w in zip("ABCD", [14, 10, 16, 10]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Sales by region",
                     "Orders and revenue are COUNTIFS / SUMIFS over Orders; "
                     "share is revenue / column total.", 4)
    for c, h in enumerate(["region", "orders", "revenue", "share"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 4)
    for i, rg in enumerate(regions):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=rg)
        ws.cell(row=r, column=2, value=f"=COUNTIFS({ord_reg},$A{r})")
        c = ws.cell(row=r, column=3, value=f"=SUMIFS({ord_amt},{ord_reg},$A{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=4,
                   value=f"=IF($C${br_total}=0,0,$C{r}/$C${br_total})")
        c.number_format = PCT_FMT
    tr = br_total
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, fmt in ((2, None), (3, MONEY_FMT)):
        c = ws.cell(row=tr, column=col,
                   value=f"=SUM({get_column_letter(col)}{br_first}:"
                         f"{get_column_letter(col)}{br_last})")
        c.font = Font(bold=True)
        if fmt:
            c.number_format = fmt
    c = ws.cell(row=tr, column=4, value=1)
    c.font = Font(bold=True)
    c.number_format = PCT_FMT
    br_reg = f"By_Region!$A${br_first}:$A${br_last}"
    br_rev = f"By_Region!$C${br_first}:$C${br_last}"

    # ---------------- By_Category ----------------
    ws = wb.create_sheet("By_Category")
    ws.sheet_properties.tabColor = TEAL
    for col, w in zip("ABCD", [14, 10, 16, 10]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Sales by category",
                     "Orders and revenue are COUNTIFS / SUMIFS over Orders; "
                     "share is revenue / column total.", 4)
    for c, h in enumerate(["category", "orders", "revenue", "share"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 4)
    for i, cat in enumerate(categories):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=f"=COUNTIFS({ord_cat},$A{r})")
        c = ws.cell(row=r, column=3, value=f"=SUMIFS({ord_amt},{ord_cat},$A{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=4,
                   value=f"=IF($C${bcat_total}=0,0,$C{r}/$C${bcat_total})")
        c.number_format = PCT_FMT
    tr = bcat_total
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, fmt in ((2, None), (3, MONEY_FMT)):
        c = ws.cell(row=tr, column=col,
                   value=f"=SUM({get_column_letter(col)}{bcat_first}:"
                         f"{get_column_letter(col)}{bcat_last})")
        c.font = Font(bold=True)
        if fmt:
            c.number_format = fmt
    c = ws.cell(row=tr, column=4, value=1)
    c.font = Font(bold=True)
    c.number_format = PCT_FMT
    bcat_cat = f"By_Category!$A${bcat_first}:$A${bcat_last}"
    bcat_rev = f"By_Category!$C${bcat_first}:$C${bcat_last}"

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 60
    r = title_block(ws, "Sales performance dashboard — Agarwal Traders",
                    "Jan – Oct 2026  •  Revenue vs target, % to goal, products, orders "
                    "and customers. Filters below are live.", 2)
    # Filters (rows 4-8)
    ws.cell(row=4, column=1, value="FILTERS — pick a month and a category").font = \
        Font(bold=True, color=NAVY, size=11)
    ws.cell(row=5, column=1, value="Focus month").font = Font(bold=True)
    ws.cell(row=5, column=2, value=months[-1])
    ws.cell(row=6, column=1, value="Focus category").font = Font(bold=True)
    ws.cell(row=6, column=2, value=categories[0])
    dv_month = DataValidation(type="list", formula1=f"{tgt_mon}",
                              allow_blank=False)
    dv_month.error = "Pick a month from Targets"
    dv_month.prompt = "Pick a month"
    ws.add_data_validation(dv_month)
    dv_month.add(ws["B5"])
    dv_cat = DataValidation(type="list", formula1=f"{bcat_cat}",
                            allow_blank=False)
    dv_cat.error = "Pick a category from By_Category"
    dv_cat.prompt = "Pick a category"
    ws.add_data_validation(dv_cat)
    dv_cat.add(ws["B6"])
    ws.cell(row=7, column=1,
           value="Filtered revenue (month × category)").font = Font(bold=True)
    c = ws.cell(row=7, column=2,
               value=f"=SUMIFS({ord_amt},{ord_mon},$B$5,{ord_cat},$B$6)")
    c.number_format = MONEY_FMT
    ws.cell(row=8, column=1, value="Filtered revenue (month only)").font = \
        Font(bold=True)
    c = ws.cell(row=8, column=2, value=f"=SUMIFS({ord_amt},{ord_mon},$B$5)")
    c.number_format = MONEY_FMT
    ws.cell(row=9, column=1, value="Filtered revenue (category only)").font = \
        Font(bold=True)
    c = ws.cell(row=9, column=2, value=f"=SUMIFS({ord_amt},{ord_cat},$B$6)")
    c.number_format = MONEY_FMT
    # KPIs (header row 11, values 12..)
    ws.cell(row=11, column=1, value="KPI").font = HDR_FONT
    ws.cell(row=11, column=1).fill = HDR_FILL
    ws.cell(row=11, column=2, value="Value").font = HDR_FONT
    ws.cell(row=11, column=2).fill = HDR_FILL
    kpis = [
        ("Total revenue", f"=SUM({mo_rev})", MONEY_FMT),
        ("Total target", f"=SUM({mo_tgt})", MONEY_FMT),
        ("YTD pct to goal", None, PCT_FMT),  # filled below (refs KPI rows)
        ("Total orders", f"=SUM({mo_ord})", None),
        ("Average order value", None, MONEY_FMT),
        ("Months on target (of 10)", f'=COUNTIF({mo_pct},">=1")', None),
        ("Best month by revenue",
         f"=INDEX({mo},MATCH(MAX({mo_rev}),{mo_rev},0))", None),
        ("Best month revenue", f"=MAX({mo_rev})", MONEY_FMT),
        ("Top product by revenue",
         f"=INDEX({bp_prod},MATCH(MAX({bp_rev}),{bp_rev},0))", None),
        ("Top customer by revenue",
         f"=INDEX({bc_cust},MATCH(MAX({bc_rev}),{bc_rev},0))", None),
        ("Top region by revenue",
         f"=INDEX({br_reg},MATCH(MAX({br_rev}),{br_rev},0))", None),
        ("Top category by revenue",
         f"=INDEX({bcat_cat},MATCH(MAX({bcat_rev}),{bcat_rev},0))", None),
        ("Revenue vs target gap", None, MONEY_FMT),
    ]
    kpi_first = 12
    for i, (label, formula, fmt) in enumerate(kpis):
        r = kpi_first + i
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        if label == "YTD pct to goal":
            formula = f"=IF(B{kpi_first + 1}=0,0,B{kpi_first}/B{kpi_first + 1})"
        elif label == "Average order value":
            formula = (f"=IF(B{kpi_first + 3}=0,0,"
                       f"B{kpi_first}/B{kpi_first + 3})")
        elif label == "Revenue vs target gap":
            formula = f"=B{kpi_first}-B{kpi_first + 1}"
        c = ws.cell(row=r, column=2, value=formula)
        if fmt:
            c.number_format = fmt
    # Charts
    col_chart = BarChart()
    col_chart.type = "col"
    col_chart.title = "Monthly revenue vs target"
    col_chart.y_axis.title = "INR"
    col_chart.x_axis.title = "Month"
    col_chart.style = 10
    data = Reference(wb["Monthly"], min_col=3, min_row=4, max_row=mo_last)
    cats = Reference(wb["Monthly"], min_col=1, min_row=mo_first, max_row=mo_last)
    col_chart.add_data(data, titles_from_data=True)
    col_chart.add_data(Reference(wb["Monthly"], min_col=4, min_row=4,
                                 max_row=mo_last), titles_from_data=True)
    col_chart.set_categories(cats)
    col_chart.shape = 4
    ws.add_chart(col_chart, "A27")
    line_chart = LineChart()
    line_chart.title = "% to goal by month"
    line_chart.y_axis.title = "% to goal"
    line_chart.style = 10
    line_chart.add_data(Reference(wb["Monthly"], min_col=5, min_row=4,
                                  max_row=mo_last), titles_from_data=True)
    line_chart.set_categories(cats)
    ws.add_chart(line_chart, "A42")
    pie_chart = PieChart()
    pie_chart.title = "Revenue share by category"
    pie_chart.style = 10
    pie_chart.add_data(Reference(wb["By_Category"], min_col=3, min_row=4,
                                 max_row=bcat_last), titles_from_data=True)
    pie_chart.set_categories(
        Reference(wb["By_Category"], min_col=1, min_row=bcat_first,
                  max_row=bcat_last))
    pie_chart.dataLabels = None
    ws.add_chart(pie_chart, "A57")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws.sheet_properties.tabColor = "C00000"
    for col, w in zip("ABCD", [42, 18, 18, 10]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Integrity checks — all must read PASS",
                     "Expected is computed in Python; Actual is a live formula. "
                     "Money comparisons allow 0.01 tolerance.", 4)
    for c, h in enumerate(["check", "expected", "actual", "result"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 4)
    total_rev = sum(float(o["amount"]) for o in orders)
    total_tgt = sum(targets.values())
    checks = [
        ("Order rows in Orders sheet", n,
         f"=COUNTA(Orders!$A${o_first}:$A${o_last})"),
        ("Monthly revenue = Orders revenue", round(total_rev, 2),
         f"=SUM({mo_rev})"),
        ("Quarterly revenue = Monthly revenue", round(total_rev, 2),
         f"=SUM(Quarterly!$C${q_first}:$C${q_last})"),
        ("By_Product revenue = Monthly revenue", round(total_rev, 2),
         f"=SUM({bp_rev})"),
        ("By_Customer revenue = Monthly revenue", round(total_rev, 2),
         f"=SUM({bc_rev})"),
        ("By_Region revenue = Monthly revenue", round(total_rev, 2),
         f"=SUM({br_rev})"),
        ("By_Category revenue = Monthly revenue", round(total_rev, 2),
         f"=SUM({bcat_rev})"),
        ("Monthly target total = Targets total", total_tgt,
         f"=SUM({mo_tgt})"),
        ("Orders with amount <= 0", 0,
         f'=COUNTIFS({ord_amt},"<=0")'),
        ("Orders with blank category", 0,
         f"=COUNTBLANK({ord_cat})"),
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


def write_exports(orders, targets):
    monthly = {}
    for o in orders:
        m = o["date"][:7]
        monthly.setdefault(m, {"orders": 0, "revenue": 0.0})
        monthly[m]["orders"] += 1
        monthly[m]["revenue"] += float(o["amount"])
    with open(OUT_MONTHLY_EXPORT, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "orders", "revenue", "target", "pct_to_goal",
                    "avg_order", "quarter"])
        for m in sorted(monthly):
            rev = monthly[m]["revenue"]
            tgt = targets[m]
            q = f"Q{(int(m[5:7]) - 1) // 3 + 1} {m[:4]}"
            w.writerow([m, monthly[m]["orders"], f"{rev:.2f}", tgt,
                        f"{rev / tgt:.4f}" if tgt else "0",
                        f"{rev / monthly[m]['orders']:.2f}", q])
    by_cat = {}
    total = sum(float(o["amount"]) for o in orders)
    for o in orders:
        c = o["category"]
        by_cat.setdefault(c, {"orders": 0, "revenue": 0.0})
        by_cat[c]["orders"] += 1
        by_cat[c]["revenue"] += float(o["amount"])
    with open(OUT_CATEGORY_EXPORT, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["category", "orders", "revenue", "share"])
        for c in sorted(by_cat):
            w.writerow([c, by_cat[c]["orders"],
                        f"{by_cat[c]['revenue']:.2f}",
                        f"{by_cat[c]['revenue'] / total:.4f}"])


def main():
    orders = read_csv("orders.csv")
    products = read_csv("products.csv")
    customers = read_csv("customers.csv")
    targets = {r["month"]: int(r["target"])
               for r in read_csv("targets.csv")}
    wb = build_workbook(orders, products, customers, targets)
    wb.save(OUT_XLSX)
    write_exports(orders, targets)
    total = sum(float(o["amount"]) for o in orders)
    print(f"built {OUT_XLSX.name}: {len(orders)} orders, "
          f"total revenue {total:,.2f}, target {sum(targets.values()):,}")


if __name__ == "__main__":
    main()
