#!/usr/bin/env python3
"""WarehousePulse — build the formula-driven warehouse tracker workbook.

Reads data/*.csv (see generate.py) and writes warehouse-tracker-dashboard.xlsx:
SKU_Master, Raw_Receipts, Raw_Issues, Raw_Adjustments, Stock_Ledger,
Bin_Map, Entry (data-entry form), Dashboard (KPIs + 3 charts + category
drill-through), AI_Report, Checks, PowerBI_Model, Guide.
Every computed cell is a live Excel formula: paste new rows into the Raw
sheets and the ledger, bins, dashboard, report and checks all recalculate.
"""
import csv
import os
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "warehouse-tracker-dashboard.xlsx")

PERIOD_START = date(2026, 10, 1)
PERIOD_DAYS = 18  # Oct 1..18

HDR_FILL = PatternFill("solid", fgColor="1F4E5F")
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color="1F4E5F")
SUB_FONT = Font(bold=True, size=11, color="1F4E5F")
MONEY_FMT = '₹#,##0.00'
QTY_FMT = '#,##0'
PCT_FMT = '0.0%'
DATE_FMT = 'dd-mmm-yy'

STATUS_FILL = {
    "OK": PatternFill("solid", fgColor="C6EFCE"),
    "REORDER": PatternFill("solid", fgColor="FFEB9C"),
    "STOCKOUT": PatternFill("solid", fgColor="FFC7CE"),
    "OVERSTOCK": PatternFill("solid", fgColor="DDEBF7"),
}


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)


def build():
    skus = read_csv("skus.csv")
    opening = {r["sku"]: int(r["opening_qty"]) for r in read_csv("opening.csv")}
    receipts = read_csv("receipts.csv")
    issues = read_csv("issues.csv")
    adjustments = read_csv("adjustments.csv")
    bins = read_csv("bins.csv")

    n_sku = len(skus)
    n_rc, n_is, n_ad = len(receipts), len(issues), len(adjustments)
    L0, L1 = 3, 2 + n_sku          # ledger data rows
    R0, R1 = 2, 1 + n_rc           # receipts rows
    I0, I1 = 2, 1 + n_is           # issues rows
    A0, A1 = 2, 1 + n_ad           # adjustments rows
    B0, B1 = 3, 2 + len(bins)      # bin rows
    cats = sorted({s["category"] for s in skus})

    wb = Workbook()

    # ================= Guide =================
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "WarehousePulse — warehouse stock tracker + Power BI-ready model"
    ws["A1"].font = TITLE_FONT
    lines = [
        "Fictional demo data: 'Sharma Wholesale Traders, Patna' (FMCG wholesale).",
        "Opening stock as of 30-Sep-2026; movements 01-Oct-2026 .. 18-Oct-2026.",
        "",
        "DAILY ROUTINE",
        "1. Key today's receipts / issues / adjustments into the Entry sheet (or paste",
        "   rows directly into Raw_Receipts / Raw_Issues / Raw_Adjustments — keep column order).",
        "2. Everything recalculates: Stock_Ledger (closing, WAC, value, days-of-cover,",
        "   status), Bin_Map utilisation, Dashboard KPIs/charts/drill-through,",
        "   AI_Report narrative and the Checks sheet.",
        "3. SKUs flagged REORDER or STOCKOUT on Stock_Ledger are your purchase list;",
        "   the Dashboard drill-through shows the top-5 SKUs by value per category.",
        "",
        "POWER BI",
        "4. Point Power BI at the powerbi/*.csv files (or export the workbook sheets):",
        "   dim_sku + dim_bin + dim_date feed fact_inventory_daily (one row per SKU per day).",
        "   Ready-made DAX measures live in powerbi/MEASURES.md and on PowerBI_Model.",
        "5. The Dashboard's 3 charts and the category drill-through are the Excel",
        "   prototype of the Power BI page — same KPIs, same grain.",
        "",
        "FORMULAS, NOT VALUES",
        "Every computed cell is a live formula (SUMIFS/VLOOKUP/IF/LARGE …). Nothing",
        "is pasted as values, so the workbook never goes stale. Checks!B must all",
        "read TRUE after any recalculation.",
    ]
    for i, line in enumerate(lines, start=3):
        ws[f"A{i}"] = line
    ws.column_dimensions["A"].width = 92

    # ================= SKU_Master =================
    sm = wb.create_sheet("SKU_Master")
    sm_headers = ["sku", "name", "category", "uom", "reorder_level",
                  "std_cost", "bin", "opening_qty"]
    for c, h in enumerate(sm_headers, start=1):
        sm.cell(row=1, column=c, value=h)
    style_header(sm, 1, len(sm_headers))
    for r, s in enumerate(skus, start=2):
        vals = [s["sku"], s["name"], s["category"], s["uom"],
                int(s["reorder_level"]), float(s["std_cost"]), s["bin"],
                opening[s["sku"]]]
        for c, v in enumerate(vals, start=1):
            cell = sm.cell(row=r, column=c, value=v)
            if c in (5, 8):
                cell.number_format = QTY_FMT
            if c == 6:
                cell.number_format = MONEY_FMT
    sm.freeze_panes = "A2"
    sm.auto_filter.ref = f"A1:H{n_sku + 1}"
    for c, w in zip("ABCDEFGH", [9, 30, 14, 7, 13, 11, 7, 12]):
        sm.column_dimensions[c].width = w

    # ================= Raw_Receipts =================
    rr = wb.create_sheet("Raw_Receipts")
    for c, h in enumerate(["date", "grn", "sku", "qty", "unit_cost", "bin",
                           "value"], start=1):
        rr.cell(row=1, column=c, value=h)
    style_header(rr, 1, 7)
    for r, x in enumerate(receipts, start=R0):
        rr.cell(row=r, column=1, value=date.fromisoformat(x["date"])).number_format = DATE_FMT
        rr.cell(row=r, column=2, value=x["grn"])
        rr.cell(row=r, column=3, value=x["sku"])
        rr.cell(row=r, column=4, value=int(x["qty"])).number_format = QTY_FMT
        rr.cell(row=r, column=5, value=float(x["unit_cost"])).number_format = MONEY_FMT
        rr.cell(row=r, column=6, value=x["bin"])
        rr.cell(row=r, column=7, value=f"=D{r}*E{r}").number_format = MONEY_FMT
    rr.freeze_panes = "A2"
    rr.auto_filter.ref = f"A1:G{R1}"
    for c, w in zip("ABCDEFG", [12, 10, 9, 9, 11, 7, 13]):
        rr.column_dimensions[c].width = w

    # ================= Raw_Issues =================
    ri = wb.create_sheet("Raw_Issues")
    for c, h in enumerate(["date", "do_no", "sku", "qty", "bin"], start=1):
        ri.cell(row=1, column=c, value=h)
    style_header(ri, 1, 5)
    for r, x in enumerate(issues, start=I0):
        ri.cell(row=r, column=1, value=date.fromisoformat(x["date"])).number_format = DATE_FMT
        ri.cell(row=r, column=2, value=x["do_no"])
        ri.cell(row=r, column=3, value=x["sku"])
        ri.cell(row=r, column=4, value=int(x["qty"])).number_format = QTY_FMT
        ri.cell(row=r, column=5, value=x["bin"])
    ri.freeze_panes = "A2"
    ri.auto_filter.ref = f"A1:E{I1}"
    for c, w in zip("ABCDE", [12, 10, 9, 9, 7]):
        ri.column_dimensions[c].width = w

    # ================= Raw_Adjustments =================
    ra = wb.create_sheet("Raw_Adjustments")
    for c, h in enumerate(["date", "adj_id", "sku", "qty", "reason"], start=1):
        ra.cell(row=1, column=c, value=h)
    style_header(ra, 1, 5)
    for r, x in enumerate(adjustments, start=A0):
        ra.cell(row=r, column=1, value=date.fromisoformat(x["date"])).number_format = DATE_FMT
        ra.cell(row=r, column=2, value=x["adj_id"])
        ra.cell(row=r, column=3, value=x["sku"])
        ra.cell(row=r, column=4, value=int(x["qty"])).number_format = QTY_FMT
        ra.cell(row=r, column=5, value=x["reason"])
    ra.freeze_panes = "A2"
    ra.auto_filter.ref = f"A1:E{A1}"
    for c, w in zip("ABCDE", [12, 10, 9, 9, 26]):
        ra.column_dimensions[c].width = w

    # ================= Stock_Ledger =================
    lg = wb.create_sheet("Stock_Ledger")
    lg_headers = ["sku", "name", "category", "bin", "reorder_level",
                  "opening", "receipts_qty", "receipts_value", "issues_qty",
                  "adj_qty", "closing", "wac", "stock_value",
                  "avg_daily_issue", "days_of_cover", "status", "reorder_qty"]
    for c, h in enumerate(lg_headers, start=1):
        lg.cell(row=2, column=c, value=h)
    style_header(lg, 2, len(lg_headers))
    lg["A1"] = "Stock_Ledger — every figure a live formula (paste new rows into Raw_*)"
    lg["A1"].font = TITLE_FONT
    std_cost_v = 'VLOOKUP($A{r},SKU_Master!$A:$G,6,FALSE)'
    for r, s in zip(range(L0, L1 + 1), skus):
        rr_ = r
        lg.cell(row=r, column=1, value=s["sku"])
        lg.cell(row=r, column=2, value=s["name"])
        lg.cell(row=r, column=3, value=s["category"])
        lg.cell(row=r, column=4, value=s["bin"])
        lg.cell(row=r, column=5,
                value=f"=VLOOKUP($A{rr_},SKU_Master!$A:$G,5,FALSE)").number_format = QTY_FMT
        lg.cell(row=r, column=6,
                value=f"=VLOOKUP($A{rr_},SKU_Master!$A:$H,8,FALSE)").number_format = QTY_FMT
        lg.cell(row=r, column=7,
                value=f"=SUMIFS(Raw_Receipts!$D:$D,Raw_Receipts!$C:$C,$A{rr_})").number_format = QTY_FMT
        lg.cell(row=r, column=8,
                value=f"=SUMIFS(Raw_Receipts!$G:$G,Raw_Receipts!$C:$C,$A{rr_})").number_format = MONEY_FMT
        lg.cell(row=r, column=9,
                value=f"=SUMIFS(Raw_Issues!$D:$D,Raw_Issues!$C:$C,$A{rr_})").number_format = QTY_FMT
        lg.cell(row=r, column=10,
                value=f"=SUMIFS(Raw_Adjustments!$D:$D,Raw_Adjustments!$C:$C,$A{rr_})").number_format = QTY_FMT
        lg.cell(row=r, column=11,
                value=f"=F{rr_}+G{rr_}-I{rr_}+J{rr_}").number_format = QTY_FMT
        lg.cell(row=r, column=12, value=(
            f"=IF((F{rr_}+G{rr_})=0,{std_cost_v.format(r=rr_)},"
            f"(F{rr_}*{std_cost_v.format(r=rr_)}+H{rr_})/(F{rr_}+G{rr_}))")
            ).number_format = MONEY_FMT
        lg.cell(row=r, column=13,
                value=f"=K{rr_}*L{rr_}").number_format = MONEY_FMT
        lg.cell(row=r, column=14,
                value=f"=I{rr_}/{PERIOD_DAYS}").number_format = "0.0"
        lg.cell(row=r, column=15,
                value=f"=IF(N{rr_}=0,999,K{rr_}/N{rr_})").number_format = "0.0"
        lg.cell(row=r, column=16, value=(
            f'=IF(K{rr_}=0,"STOCKOUT",IF(K{rr_}<=E{rr_},"REORDER",'
            f'IF(O{rr_}>60,"OVERSTOCK","OK")))'))
        lg.cell(row=r, column=17,
                value=f'=IF(P{rr_}="REORDER",MAX(0,E{rr_}*2-K{rr_}),0)'
                ).number_format = QTY_FMT
    lg.freeze_panes = "A3"
    lg.auto_filter.ref = f"A2:Q{L1}"
    widths = [9, 30, 14, 7, 13, 10, 12, 14, 10, 9, 10, 11, 13, 13, 12, 11, 11]
    for c, w in zip("ABCDEFGHIJKLMNOPQ", widths):
        lg.column_dimensions[c].width = w

    # ================= Bin_Map =================
    bm = wb.create_sheet("Bin_Map")
    for c, h in enumerate(["bin", "capacity", "qty_on_hand", "#_skus",
                           "utilisation", "flag"], start=1):
        bm.cell(row=2, column=c, value=h)
    style_header(bm, 2, 6)
    bm["A1"] = "Bin_Map — per-bin stock and utilisation (formula-driven)"
    bm["A1"].font = TITLE_FONT
    for r, b in zip(range(B0, B1 + 1), bins):
        bm.cell(row=r, column=1, value=b["bin"])
        bm.cell(row=r, column=2, value=int(b["capacity"])).number_format = QTY_FMT
        bm.cell(row=r, column=3, value=(
            f"=SUMIFS(Stock_Ledger!$K:$K,Stock_Ledger!$D:$D,$A{r})")
            ).number_format = QTY_FMT
        bm.cell(row=r, column=4, value=(
            f"=COUNTIFS(Stock_Ledger!$D:$D,$A{r})")).number_format = QTY_FMT
        bm.cell(row=r, column=5,
                value=f"=IF(B{r}=0,0,C{r}/B{r})").number_format = PCT_FMT
        bm.cell(row=r, column=6, value=(
            f'=IF(E{r}>0.9,"FULL",IF(C{r}=0,"EMPTY","OK"))'))
    bm.freeze_panes = "A3"
    for c, w in zip("ABCDEF", [8, 10, 13, 8, 12, 8]):
        bm.column_dimensions[c].width = w

    # ================= Entry =================
    en = wb.create_sheet("Entry")
    en["A1"] = "Entry — key today's movements here, then copy the row into the matching Raw_* sheet"
    en["A1"].font = TITLE_FONT
    for c, h in enumerate(["date (dd-mmm-yy)", "type", "sku", "qty",
                           "unit_cost (receipts only)", "sku name (auto)",
                           "bin (auto)"], start=1):
        en.cell(row=3, column=c, value=h)
    style_header(en, 3, 7)
    for r in range(4, 24):
        en.cell(row=r, column=1, value=date(2026, 10, 19)).number_format = DATE_FMT
        en.cell(row=r, column=6,
                value=f'=IF(C{r}="","",VLOOKUP(C{r},SKU_Master!$A:$B,2,FALSE))')
        en.cell(row=r, column=7,
                value=f'=IF(C{r}="","",VLOOKUP(C{r},SKU_Master!$A:$G,7,FALSE))')
    dv_type = DataValidation(type="list",
                             formula1='"Receipt,Issue,Adjustment"',
                             allow_blank=True)
    en.add_data_validation(dv_type)
    dv_type.add(f"B4:B23")
    sku_list = ",".join(s["sku"] for s in skus)
    dv_sku = DataValidation(type="list",
                            formula1=f"SKU_Master!$A$2:$A${n_sku + 1}",
                            allow_blank=True)
    en.add_data_validation(dv_sku)
    dv_sku.add("C4:C23")
    en["A26"] = ("How it works: pick type + SKU from the dropdowns, enter qty; "
                 "the name/bin fill in automatically. Then copy the row into "
                 "Raw_Receipts (needs unit_cost), Raw_Issues or Raw_Adjustments "
                 "(qty may be negative). Use negative qty for damaged/expired stock.")
    en["A26"].alignment = Alignment(wrap_text=True)
    en["A27"] = ("Automation note: these Raw_* tables are plain ranges with "
                 "AutoFilter (no Table objects) so the workbook recalculates in "
                 "Excel, LibreOffice and Google Sheets. A 2-minute Power Query "
                 "'append new rows' routine is described on PowerBI_Model.")
    en["A27"].alignment = Alignment(wrap_text=True)
    for c, w in zip("ABCDEFG", [15, 12, 9, 9, 22, 30, 8]):
        en.column_dimensions[c].width = w

    # ================= Dashboard =================
    dash = wb.create_sheet("Dashboard")
    dash["A1"] = "Dashboard — KPIs (all formula-linked)"
    dash["A1"].font = TITLE_FONT
    kpis = [
        ("Total SKUs", f"=COUNTA(Stock_Ledger!$A${L0}:$A${L1})", QTY_FMT),
        ("Total stock value", f"=SUM(Stock_Ledger!$M${L0}:$M${L1})", MONEY_FMT),
        ("SKUs on REORDER", f'=COUNTIF(Stock_Ledger!$P${L0}:$P${L1},"REORDER")', QTY_FMT),
        ("STOCKOUT SKUs", f'=COUNTIF(Stock_Ledger!$P${L0}:$P${L1},"STOCKOUT")', QTY_FMT),
        ("OVERSTOCK SKUs", f'=COUNTIF(Stock_Ledger!$P${L0}:$P${L1},"OVERSTOCK")', QTY_FMT),
        ("Avg days of cover", f'=AVERAGEIF(Stock_Ledger!$O${L0}:$O${L1},"<900")', "0.0"),
        ("Receipts value (period)", "=SUM(Raw_Receipts!$G:$G)", MONEY_FMT),
        ("Issues qty (period)", "=SUM(Raw_Issues!$D:$D)", QTY_FMT),
    ]
    for i, (label, formula, fmt) in enumerate(kpis, start=3):
        dash[f"A{i}"] = label
        dash[f"A{i}"].font = Font(bold=True)
        dash[f"B{i}"] = formula
        dash[f"B{i}"].number_format = fmt
        dash[f"B{i}"].font = Font(size=13, color="1F4E5F")
    dash.column_dimensions["A"].width = 26
    dash.column_dimensions["B"].width = 18

    # Category summary (column chart source)
    dash["A13"] = "Stock value by category"
    dash["A13"].font = SUB_FONT
    dash["A14"], dash["B14"], dash["C14"] = "Category", "Stock value", "SKUs"
    for c in "ABC":
        dash[f"{c}14"].fill = HDR_FILL
        dash[f"{c}14"].font = HDR_FONT
    cat_rows = []
    for i, cat in enumerate(cats, start=15):
        dash[f"A{i}"] = cat
        dash[f"B{i}"] = (f"=SUMIFS(Stock_Ledger!$M:$M,"
                         f"Stock_Ledger!$C:$C,$A{i})")
        dash[f"B{i}"].number_format = MONEY_FMT
        dash[f"C{i}"] = (f"=COUNTIFS(Stock_Ledger!$C:$C,$A{i})")
        dash[f"C{i}"].number_format = QTY_FMT
        cat_rows.append(i)
    c0, c1 = cat_rows[0], cat_rows[-1]

    cat_bar = BarChart()
    cat_bar.title = "Stock value by category"
    cat_bar.style = 2
    cat_bar.height, cat_bar.width = 7.5, 14
    cat_bar.add_data(Reference(dash, min_col=2, max_col=2, min_row=c0,
                               max_row=c1), titles_from_data=False)
    cat_bar.set_categories(Reference(dash, min_col=1, max_col=1, min_row=c0,
                                     max_row=c1))
    cat_bar.series[0].title = SeriesLabel(v="Stock value")
    dash.add_chart(cat_bar, "D2")

    # Daily movement (line chart source)
    d0 = c1 + 3
    dash[f"A{d0}"] = "Daily receipts vs issues (qty)"
    dash[f"A{d0}"].font = SUB_FONT
    dash[f"A{d0+1}"], dash[f"B{d0+1}"], dash[f"C{d0+1}"] = (
        "Date", "Receipts qty", "Issues qty")
    for c in "ABC":
        dash[f"{c}{d0+1}"].fill = HDR_FILL
        dash[f"{c}{d0+1}"].font = HDR_FONT
    day_rows = []
    for i in range(PERIOD_DAYS):
        r = d0 + 2 + i
        day = PERIOD_START + timedelta(days=i)
        dash.cell(row=r, column=1, value=day).number_format = DATE_FMT
        dash[f"B{r}"] = (f"=SUMIFS(Raw_Receipts!$D:$D,"
                         f"Raw_Receipts!$A:$A,$A{r})")
        dash[f"B{r}"].number_format = QTY_FMT
        dash[f"C{r}"] = (f"=SUMIFS(Raw_Issues!$D:$D,"
                         f"Raw_Issues!$A:$A,$A{r})")
        dash[f"C{r}"].number_format = QTY_FMT
        day_rows.append(r)
    dd0, dd1 = day_rows[0], day_rows[-1]

    mv_line = LineChart()
    mv_line.title = "Daily receipts vs issues (qty)"
    mv_line.style = 2
    mv_line.height, mv_line.width = 7.5, 14
    mv_line.add_data(Reference(dash, min_col=2, max_col=3, min_row=dd0 - 1,
                               max_row=dd1), from_rows=False,
                     titles_from_data=True)
    mv_line.set_categories(Reference(dash, min_col=1, max_col=1, min_row=dd0,
                                     max_row=dd1))
    dash.add_chart(mv_line, "D18")

    # Status mix (pie chart source)
    s0 = dd1 + 3
    dash[f"A{s0}"] = "Stock status mix"
    dash[f"A{s0}"].font = SUB_FONT
    dash[f"A{s0+1}"], dash[f"B{s0+1}"] = "Status", "SKUs"
    for c in "AB":
        dash[f"{c}{s0+1}"].fill = HDR_FILL
        dash[f"{c}{s0+1}"].font = HDR_FONT
    statuses = ["OK", "REORDER", "STOCKOUT", "OVERSTOCK"]
    st_rows = []
    for i, st in enumerate(statuses, start=s0 + 2):
        dash[f"A{i}"] = st
        dash[f"B{i}"] = (f'=COUNTIF(Stock_Ledger!$P${L0}:$P${L1},"{st}")')
        dash[f"B{i}"].number_format = QTY_FMT
        st_rows.append(i)
    st0, st1 = st_rows[0], st_rows[-1]

    pie = PieChart()
    pie.title = "SKUs by stock status"
    pie.style = 2
    pie.height, pie.width = 7.5, 14
    pie.add_data(Reference(dash, min_col=2, max_col=2, min_row=st0,
                           max_row=st1), titles_from_data=False)
    pie.set_categories(Reference(dash, min_col=1, max_col=1, min_row=st0,
                                 max_row=st1))
    dash.add_chart(pie, "D34")

    # Category drill-through (top-5 SKUs by stock value)
    t0 = st1 + 3
    dash[f"A{t0}"] = "Drill-through — top 5 SKUs by stock value in category:"
    dash[f"A{t0}"].font = SUB_FONT
    dash[f"B{t0}"] = cats[0]
    dash[f"B{t0}"].font = Font(bold=True, size=12, color="1F4E5F")
    dv_cat = DataValidation(type="list",
                            formula1='"' + ",".join(cats) + '"',
                            allow_blank=False)
    dash.add_data_validation(dv_cat)
    dv_cat.add(f"B{t0}")
    for c, h in enumerate(["Rank", "SKU", "Name", "Closing qty",
                           "Stock value"], start=1):
        dash.cell(row=t0 + 1, column=c, value=h)
    for c in "ABCDE":
        dash[f"{c}{t0+1}"].fill = HDR_FILL
        dash[f"{c}{t0+1}"].font = HDR_FONT
    # hidden helper column AA: stock value if SKU in selected category
    for r in range(L0, L1 + 1):
        dash[f"AA{r}"] = (f"=IF(Stock_Ledger!$C{r}=$B${t0},"
                          f"Stock_Ledger!$M{r},0)")
        dash[f"AA{r}"].number_format = MONEY_FMT
    dash.column_dimensions["AA"].hidden = True
    for k in range(1, 6):
        r = t0 + 1 + k
        dash[f"A{r}"] = k
        dash[f"B{r}"] = (  # value first (col C), then SKU lookup
            f"=IF(COUNTIF($AA${L0}:$AA${L1},\">0\")>={k},"
            f"LARGE($AA${L0}:$AA${L1},{k}),0)")
        dash[f"B{r}"].number_format = MONEY_FMT
        # NOTE: col B holds the value; SKU in C via MATCH on helper
        dash[f"C{r}"] = (f"=IF($B{r}=0,\"\",INDEX(Stock_Ledger!$A${L0}:"
                         f"$A${L1},MATCH($B{r},$AA${L0}:$AA${L1},0)))")
        dash[f"D{r}"] = (f"=IF($C{r}=\"\",\"\",VLOOKUP($C{r},"
                         f"Stock_Ledger!$A:$B,2,FALSE))")
        dash[f"E{r}"] = (f"=IF($C{r}=\"\",\"\",VLOOKUP($C{r},"
                         f"Stock_Ledger!$A:$K,11,FALSE))")
        dash[f"E{r}"].number_format = QTY_FMT
        dash[f"F{r}"] = f"=$B{r}"
        dash[f"F{r}"].number_format = MONEY_FMT
    # relabel: B=value helper display; keep header row honest
    dash[f"B{t0+1}"] = "Stock value"
    dash[f"C{t0+1}"] = "SKU"
    dash[f"D{t0+1}"] = "Name"
    dash[f"E{t0+1}"] = "Closing qty"
    dash[f"F{t0+1}"] = "Value check"
    for c in "BCDEF":
        dash[f"{c}{t0+1}"].fill = HDR_FILL
        dash[f"{c}{t0+1}"].font = HDR_FONT
        dash.column_dimensions[c].width = 16
    dash.column_dimensions["A"].width = 26
    dash.sheet_properties.pageSetUpPr = None

    # ================= AI_Report =================
    ai = wb.create_sheet("AI_Report")
    ai["A1"] = "AI_Report — rule-based narrative (recalculates with the data)"
    ai["A1"].font = TITLE_FONT
    ai["A3"] = "Key findings"
    ai["A3"].font = SUB_FONT
    ai["B4"] = "=Dashboard!B4"    # total stock value
    ai["B5"] = "=Dashboard!B5"    # reorder count
    ai["B6"] = "=Dashboard!B6"    # stockout count
    ai["B7"] = ('=SUMIFS(Stock_Ledger!$M:$M,Stock_Ledger!$P:$P,"OVERSTOCK")')
    ai["B8"] = "=Dashboard!B9"    # receipts value
    ai["B9"] = "=Dashboard!B10"   # issues qty
    ai["B10"] = ('=IF(SUMIFS(Raw_Issues!$D:$D,Raw_Issues!$A:$A,"<"&'
                 'DATE(2026,10,10))=0,0,SUMIFS(Raw_Issues!$D:$D,'
                 'Raw_Issues!$A:$A,">="&DATE(2026,10,10))/'
                 'SUMIFS(Raw_Issues!$D:$D,Raw_Issues!$A:$A,"<"&'
                 'DATE(2026,10,10))-1)')
    ai["B11"] = f"=MAX(Dashboard!B{c0}:B{c1})/Dashboard!B4"
    ai["B12"] = '=COUNTIF(Bin_Map!$F:$F,"FULL")'
    ai["B13"] = ('=IFERROR(INDEX(Stock_Ledger!$A:$A,MATCH("STOCKOUT",'
                 'Stock_Ledger!$P:$P,0)),"none")')
    ai["B14"] = "=SUM(Stock_Ledger!$Q:$Q)"
    for r in range(4, 15):
        ai[f"B{r}"].number_format = (
            MONEY_FMT if r in (4, 7, 8) else
            (PCT_FMT if r in (10, 11) else ("#,##0" if r != 13 else "General")))
    ai["A5"] = ('="Stock on hand is "&TEXT(B4,"₹#,##0")&" across "&'
                'TEXT(Dashboard!B3,"0")&" SKUs. "&TEXT(B5,"0")&" SKUs need '
                'reorder and "&TEXT(B6,"0")&" are stocked out."')
    ai["A5"].alignment = Alignment(wrap_text=True)
    flags = [
        (7, '=IF(B6>0,"⚠ "&B13&" is STOCKED OUT — expedite purchase or transfer",'
         '"✓ No stockouts")'),
        (8, '=IF(B5>3,"⚠ "&TEXT(B5,"0")&" SKUs at/below reorder level — raise POs ' 
         'for "&TEXT(B14,"#,##0")&" units total","✓ Reorder list is short")'),
        (9, '=IF(B7/B4>0.2,"⚠ "&TEXT(B7,"₹#,##0")&" ("&TEXT(B7/B4,"0%")&") sits in '
         'OVERSTOCK — pause buying there","✓ No material overstock")'),
        (10, '=IF(B10>0.2,"⚠ Dispatches up "&TEXT(B10,"0%")&" in the second half of '
         'the period — watch fast movers","✓ Dispatch pace is steady")'),
        (11, '=IF(B11>0.4,"⚠ One category holds "&TEXT(B11,"0%")&" of stock value — '
         'concentration risk","✓ Stock value spread across categories")'),
        (12, '=IF(B12>0,"⚠ "&TEXT(B12,"0")&" bins are over 90% full — rebalance",'
         '"✓ Bin utilisation is healthy")'),
        (13, '=IF(SUMPRODUCT((Stock_Ledger!$K$3:$K$38<0)+0)>0,"⚠ Negative closing '
         'stock exists — data error","✓ No negative stock")'),
        (14, '"ℹ Next: paste today&apos;s rows into Raw_* and the ledger, bins, '
         'dashboard and this report all refresh."'),
    ]
    ai["A6"] = "Flags"
    ai["A6"].font = SUB_FONT
    for r, f in flags:
        ai[f"A{r}"] = f
        ai[f"A{r}"].alignment = Alignment(wrap_text=True)
    ai.column_dimensions["A"].width = 95
    ai.column_dimensions["B"].width = 18
    for r in range(5, 15):
        ai.row_dimensions[r].height = 24

    # ================= Checks =================
    ch = wb.create_sheet("Checks")
    ch["A1"] = "Checks — all must be TRUE after recalculation"
    ch["A1"].font = TITLE_FONT
    ch["A2"], ch["B2"] = "Check", "Result"
    style_header(ch, 2, 2)
    checks = [
        "Ledger closing = opening + receipts − issues + adjustments (all SKUs)",
        "No negative closing stock",
        "Stock value = SUM(closing × WAC)",
        "Bin quantities tie to ledger closings",
        "Receipts value ties raw ↔ ledger",
        "Issues qty ties raw ↔ ledger",
        "Status counts add up to total SKUs",
        "No orphan SKUs in raw tables (all exist in SKU_Master)",
    ]
    formulas = [
        (f"=SUMPRODUCT(ABS(Stock_Ledger!$K${L0}:$K${L1}-"
         f"(Stock_Ledger!$F${L0}:$F${L1}+Stock_Ledger!$G${L0}:$G${L1}-"
         f"Stock_Ledger!$I${L0}:$I${L1}+Stock_Ledger!$J${L0}:$J${L1})))<0.01"),
        f"=COUNTIF(Stock_Ledger!$K${L0}:$K${L1},\"<0\")=0",
        (f"=ABS(SUM(Stock_Ledger!$M${L0}:$M${L1})-SUMPRODUCT("
         f"Stock_Ledger!$K${L0}:$K${L1},Stock_Ledger!$L${L0}:$L${L1}))<1"),
        (f"=ABS(SUM(Bin_Map!$C${B0}:$C${B1})-SUM(Stock_Ledger!$K${L0}:"
         f"$K${L1}))<0.01"),
        (f"=ABS(SUM(Raw_Receipts!$G:$G)-SUM(Stock_Ledger!$H${L0}:"
         f"$H${L1}))<0.01"),
        (f"=SUM(Raw_Issues!$D:$D)=SUM(Stock_Ledger!$I${L0}:$I${L1})"),
        (f"=(COUNTIF(Stock_Ledger!$P${L0}:$P${L1},\"OK\")+COUNTIF("
         f"Stock_Ledger!$P${L0}:$P${L1},\"REORDER\")+COUNTIF("
         f"Stock_Ledger!$P${L0}:$P${L1},\"STOCKOUT\")+COUNTIF("
         f"Stock_Ledger!$P${L0}:$P${L1},\"OVERSTOCK\"))=COUNTA("
         f"Stock_Ledger!$A${L0}:$A${L1})"),
        (f"=SUMPRODUCT((COUNTIF(SKU_Master!$A:$A,Raw_Receipts!$C$2:$C${R1})"
         f"=0)+0)+SUMPRODUCT((COUNTIF(SKU_Master!$A:$A,Raw_Issues!$C$2:$C${I1})"
         f"=0)+0)+SUMPRODUCT((COUNTIF(SKU_Master!$A:$A,Raw_Adjustments!$C$2:"
         f"$C${A1})=0)+0)=0"),
    ]
    for i, (label, f) in enumerate(zip(checks, formulas), start=3):
        ch[f"A{i}"] = label
        ch[f"B{i}"] = f
        ch[f"B{i}"].font = Font(bold=True)
    ch.column_dimensions["A"].width = 58
    ch.column_dimensions["B"].width = 12

    # ================= PowerBI_Model =================
    pb = wb.create_sheet("PowerBI_Model")
    pb["A1"] = "PowerBI_Model — star schema + DAX for the Power BI dashboard"
    pb["A1"].font = TITLE_FONT
    model = [
        "GRAIN: one row per SKU per day (fact_inventory_daily.csv).",
        "Dimensions: dim_sku (36 SKUs, category, bin), dim_bin (12 bins, capacity),",
        "dim_date (2026-10-01 .. 2026-12-31). All CSVs live in powerbi/.",
        "Relationships: fact[sku] → dim_sku[sku]; fact[bin] → dim_bin[bin];",
        "fact[date] → dim_date[date] (all single-direction, many-to-one).",
        "",
        "DAX MEASURES (also in powerbi/MEASURES.md):",
        "  Closing Qty = SUM(fact[closing_qty])",
        "  Stock Value = SUMX(fact, fact[closing_qty] * RELATED(dim_sku[wac]))",
        "  Days of Cover = DIVIDE([Closing Qty], DIVIDE(SUM(fact[issues_qty]), 18))",
        "  Reorder SKUs = COUNTROWS(FILTER(VALUES(dim_sku[sku]), [Closing Qty] <=",
        "      MAX(dim_sku[reorder_level])))",
        "  Stockout SKUs = COUNTROWS(FILTER(VALUES(dim_sku[sku]), [Closing Qty] = 0))",
        "  Receipts Value = SUM(fact[receipts_value])",
        "  Bin Utilisation % = DIVIDE(SUM(fact[closing_qty]), SUM(dim_bin[capacity]))",
        "",
        "POWER QUERY REFRESH (2 minutes, no VBA):",
        "  Data → Get Data → From Text/CSV → point at data/receipts.csv etc. →",
        "  Append new daily exports with the same column order → Refresh All.",
        "  The Excel workbook mirrors this: paste new rows into Raw_* and every",
        "  formula sheet (ledger, bins, dashboard, report, checks) recalculates.",
        "",
        "DRILL-THROUGH MAPPING (Excel prototype → Power BI page):",
        "  Dashboard KPIs → KPI cards; category column chart → bar visual with",
        "  category axis; daily line chart → line visual by dim_date[date];",
        "  status pie → donut visual; category selector + top-5 table →",
        "  Power BI drill-through page filtered by dim_sku[category].",
    ]
    for i, line in enumerate(model, start=3):
        pb[f"A{i}"] = line
    pb.column_dimensions["A"].width = 95

    wb.save(OUT)
    print(f"wrote {OUT}: {len(wb.sheetnames)} sheets, {n_sku} SKUs, "
          f"{n_rc} receipts, {n_is} issues, {n_ad} adjustments")


if __name__ == "__main__":
    build()
