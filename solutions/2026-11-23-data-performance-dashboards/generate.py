#!/usr/bin/env python3
"""Build performance-dashboards.xlsx for PerfDash.

Reads data/performance_daily.csv + data/targets.csv + data/ground_truth.json
and writes a genuine formula-driven workbook:

  Guide     — walkthrough: what this is, how to refresh it, sheet map
  Raw       — daily channel rows + live derived columns (Month, CTR, CPC,
              ROAS, AOV) + TOTAL row
  Monthly   — SUMIFS by month + target vs actual + variance (+/- %)
  Channel   — SUMIFS by channel + revenue share
  Campaign  — SUMIFS by campaign + orders share
  Targets   — monthly revenue targets (values, from targets.csv)
  Dashboard — focus-month dropdown (data validation) driving 7 KPIs +
              monthly table + 3 native charts
  Checks    — 10 integrity checks vs the Python ground truth

Every computed figure is a live A1-style formula (no hard-coded results,
no Table objects, no structured references). Ranges use full column
letters on both ends so LibreOffice recalculates them.
"""
import csv
import datetime as dt
import json
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "performance-dashboards.xlsx")

MONTHS = ["2026-04", "2026-05", "2026-06", "2026-07",
          "2026-08", "2026-09", "2026-10"]
FOCUS = "2026-10"
RAW_HEADROOM = 600  # SUMIFS ranges extend past the data for pasted rows

MONEY = "\u20b9#,##0.00"
COUNT = "#,##0"
PCT = "0.00%"
PCT1 = "0.0%"
TITLE_FONT = Font(bold=True, size=14)
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
TOTAL_FONT = Font(bold=True)
TOTAL_FILL = PatternFill("solid", fgColor="D9E2F3")
RED_FONT = Font(color="C00000")
WRAP = Alignment(wrap_text=True, vertical="top")


def load():
    with open(os.path.join(DATA, "performance_daily.csv")) as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["date"] = dt.date.fromisoformat(r["date"])
        for k in ("impressions", "clicks", "orders"):
            r[k] = int(r[k])
        for k in ("revenue", "ad_spend"):
            r[k] = float(r[k])
    with open(os.path.join(DATA, "targets.csv")) as f:
        targets = {r["month"]: float(r["target_revenue"])
                   for r in csv.DictReader(f)}
    with open(os.path.join(DATA, "ground_truth.json")) as f:
        gt = json.load(f)
    return rows, targets, gt


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL


def style_total(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = TOTAL_FONT
        cell.fill = TOTAL_FILL


def build_guide(wb):
    ws = wb.active
    ws.title = "Guide"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    lines = [
        ("PerfDash — Data Performance Dashboards", TITLE_FONT),
        ("Working demo built around the Freelancer brief 'Excel Data Performance "
         "Dashboards' (Excel/CSV files into KPI dashboards with PivotTables, "
         "Power Query, Power Pivot, Power BI or Tableau, fully editable, with a "
         "walkthrough). Fictional data: Ganga Retail & Services Pvt. Ltd., Patna "
         "(FMCG wholesaler/retailer), Apr–Oct 2026.", None),
        ("HOW TO USE — WALKTHROUGH", Font(bold=True, size=12)),
        ("1. Paste new daily rows at the bottom of Raw (columns A–I). The Month, "
         "CTR, CPC, ROAS and AOV columns are formulas — copy them down with the row.", None),
        ("2. If you paste more than 600 rows, extend the SUMIFS ranges on Monthly / "
         "Channel / Campaign (they currently read Raw rows 2–600).", None),
        ("3. Change the focus month on Dashboard (cell B3, dropdown) — every KPI, "
         "and the monthly table, refresh for that month.", None),
        ("4. Update Targets any month you revise the plan; Monthly variance and "
         "variance % follow automatically.", None),
        ("5. Read the story on Dashboard: revenue trend, which channel earns the "
         "most, which campaign returns the most per rupee spent.", None),
        ("SHEET MAP", Font(bold=True, size=12)),
        ("Raw — one row per channel-day, exactly as an export would arrive, plus "
         "live derived columns. Monthly — revenue/orders/spend by month vs target. "
         "Channel — the same cut by sales channel, with revenue share. Campaign — "
         "the same cut by campaign, with orders share. Targets — monthly revenue "
         "targets. Dashboard — focus-month KPIs + 3 charts. Checks — 10 integrity "
         "checks; all must read PASS.", None),
        ("DEFINITIONS", Font(bold=True, size=12)),
        ("CTR = Clicks / Impressions. CPC = Ad Spend / Clicks (ad channels) — for "
         "WhatsApp and Walk-in the spend is a per-order cost, so CPC there reads "
         "Spend / Clicks on the same basis. ROAS = Revenue / Ad Spend. "
         "AOV = Revenue / Orders. Variance % = (Revenue − Target) / Target.", None),
        ("WHAT THE DATA SAYS (Apr–Oct 2026)", Font(bold=True, size=12)),
        ("Walk-in is the biggest channel (about a third of revenue) and needs almost "
         "no ad spend — footfall is the profit engine. Paid Campaigns return the "
         "least per rupee (ROAS ~2.5) and deserve the first budget review. October "
         "is the festive peak (Diwali Early-Bird), but it still missed target — "
         "the miss is a demand-planning note, not a dashboard error.", None),
        ("NOTE", Font(bold=True, size=12)),
        ("All figures are fictional demo data. Every computed cell is a live Excel "
         "formula — change any raw value and the whole workbook follows.", None),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        c.alignment = WRAP
        if font:
            c.font = font
    ws.column_dimensions["A"].width = 150
    for i in (2, 4, 5, 6, 7, 8, 10, 12, 14):
        ws.row_dimensions[i].height = 30


def build_raw(wb, rows):
    ws = wb.create_sheet("Raw")
    headers = ["Date", "Channel", "Product", "Campaign", "Impressions",
               "Clicks", "Orders", "Revenue (Rs)", "Ad Spend (Rs)", "Month",
               "CTR", "CPC", "ROAS", "AOV (Rs)"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, r in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=r["date"]).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2, value=r["channel"])
        ws.cell(row=i, column=3, value=r["product"])
        ws.cell(row=i, column=4, value=r["campaign"])
        ws.cell(row=i, column=5, value=r["impressions"]).number_format = COUNT
        ws.cell(row=i, column=6, value=r["clicks"]).number_format = COUNT
        ws.cell(row=i, column=7, value=r["orders"]).number_format = COUNT
        ws.cell(row=i, column=8, value=r["revenue"]).number_format = MONEY
        ws.cell(row=i, column=9, value=r["ad_spend"]).number_format = MONEY
        ws.cell(row=i, column=10, value=f'=TEXT(A{i},"yyyy-mm")')
        ws.cell(row=i, column=11, value=f"=IF(E{i}=0,0,F{i}/E{i})").number_format = PCT
        ws.cell(row=i, column=12, value=f"=IF(F{i}=0,0,I{i}/F{i})").number_format = MONEY
        ws.cell(row=i, column=13, value=f"=IF(I{i}=0,0,H{i}/I{i})").number_format = "0.00"
        ws.cell(row=i, column=14, value=f"=IF(G{i}=0,0,H{i}/G{i})").number_format = MONEY
    total = len(rows) + 2
    ws.cell(row=total, column=1, value="TOTAL")
    for col, letter in ((5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I")):
        ws.cell(row=total, column=col,
               value=f"=SUM({letter}2:{letter}{len(rows) + 1})")
    ws.cell(row=total, column=11,
           value=f"=IF(E{total}=0,0,F{total}/E{total})").number_format = PCT
    ws.cell(row=total, column=12,
           value=f"=IF(F{total}=0,0,I{total}/F{total})").number_format = MONEY
    ws.cell(row=total, column=13,
           value=f"=IF(I{total}=0,0,H{total}/I{total})").number_format = "0.00"
    ws.cell(row=total, column=14,
           value=f"=IF(G{total}=0,0,H{total}/G{total})").number_format = MONEY
    for col in (5, 6, 7):
        ws.cell(row=total, column=col).number_format = COUNT
    for col in (8, 9):
        ws.cell(row=total, column=col).number_format = MONEY
    style_total(ws, total, len(headers))
    ws.auto_filter.ref = f"A1:N{len(rows) + 1}"
    ws.freeze_panes = "A2"
    widths = [12, 17, 20, 18, 13, 10, 9, 14, 13, 9, 8, 10, 8, 11]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    return total


def sumifs(col_letter, crit_range, crit, headroom=RAW_HEADROOM):
    return (f"=SUMIFS(Raw!${col_letter}$2:${col_letter}${headroom},"
            f"Raw!${crit_range}$2:${crit_range}${headroom},{crit})")


def build_monthly(wb):
    ws = wb.create_sheet("Monthly")
    ws["A1"] = "Monthly Performance — revenue, orders and spend vs target"
    ws["A1"].font = TITLE_FONT
    headers = ["Month", "Impressions", "Clicks", "Orders", "Revenue (Rs)",
               "Ad Spend (Rs)", "CTR", "CPC", "ROAS", "AOV (Rs)",
               "Target Revenue (Rs)", "Variance (Rs)", "Variance %"]
    ws.append([])  # row 2 blank
    for j, h in enumerate(headers, start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, len(headers))
    for i, m in enumerate(MONTHS, start=4):
        ws.cell(row=i, column=1, value=m)
        ws.cell(row=i, column=2, value=sumifs("E", "J", f"$A{i}"))
        ws.cell(row=i, column=3, value=sumifs("F", "J", f"$A{i}"))
        ws.cell(row=i, column=4, value=sumifs("G", "J", f"$A{i}"))
        ws.cell(row=i, column=5, value=sumifs("H", "J", f"$A{i}"))
        ws.cell(row=i, column=6, value=sumifs("I", "J", f"$A{i}"))
        ws.cell(row=i, column=7, value=f"=IF(B{i}=0,0,C{i}/B{i})")
        ws.cell(row=i, column=8, value=f"=IF(C{i}=0,0,F{i}/C{i})")
        ws.cell(row=i, column=9, value=f"=IF(F{i}=0,0,E{i}/F{i})")
        ws.cell(row=i, column=10, value=f"=IF(D{i}=0,0,E{i}/D{i})")
        ws.cell(row=i, column=11,
               value=f"=SUMIFS(Targets!$B$2:$B$8,Targets!$A$2:$A$8,$A{i})")
        ws.cell(row=i, column=12, value=f"=E{i}-K{i}")
        ws.cell(row=i, column=13, value=f"=IF(K{i}=0,0,L{i}/K{i})")
    total = 11
    ws.cell(row=total, column=1, value="TOTAL")
    for col, letter in ((2, "B"), (3, "C"), (4, "D"), (5, "E"), (6, "F"),
                        (11, "K"), (12, "L")):
        ws.cell(row=total, column=col,
               value=f"=SUM({letter}4:{letter}10)")
    ws.cell(row=total, column=7, value=f"=IF(B{total}=0,0,C{total}/B{total})")
    ws.cell(row=total, column=8, value=f"=IF(C{total}=0,0,F{total}/C{total})")
    ws.cell(row=total, column=9, value=f"=IF(F{total}=0,0,E{total}/F{total})")
    ws.cell(row=total, column=10, value=f"=IF(D{total}=0,0,E{total}/D{total})")
    ws.cell(row=total, column=13, value=f"=IF(K{total}=0,0,L{total}/K{total})")
    style_total(ws, total, len(headers))
    for r in range(4, 12):
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).number_format = COUNT
        for col in (5, 6, 11, 12):
            ws.cell(row=r, column=col).number_format = MONEY
        ws.cell(row=r, column=7).number_format = PCT
        ws.cell(row=r, column=8).number_format = MONEY
        ws.cell(row=r, column=9).number_format = "0.00"
        ws.cell(row=r, column=10).number_format = MONEY
        ws.cell(row=r, column=13).number_format = PCT
    ws.conditional_formatting.add("M4:M10",
                                  CellIsRule(operator="lessThan", formula=["0"],
                                             font=RED_FONT))
    ws.freeze_panes = "A4"
    widths = [10, 13, 10, 9, 14, 13, 8, 10, 8, 11, 16, 13, 11]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w


def build_channel(wb, gt):
    ws = wb.create_sheet("Channel")
    ws["A1"] = "Performance by Channel"
    ws["A1"].font = TITLE_FONT
    headers = ["Channel", "Impressions", "Clicks", "Orders", "Revenue (Rs)",
               "Ad Spend (Rs)", "CTR", "CPC", "ROAS", "AOV (Rs)",
               "Revenue Share"]
    ws.append([])
    for j, h in enumerate(headers, start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, len(headers))
    channels = [c["channel"] for c in gt["channels"]]
    for i, ch in enumerate(channels, start=4):
        ws.cell(row=i, column=1, value=ch)
        ws.cell(row=i, column=2, value=sumifs("E", "B", f"$A{i}"))
        ws.cell(row=i, column=3, value=sumifs("F", "B", f"$A{i}"))
        ws.cell(row=i, column=4, value=sumifs("G", "B", f"$A{i}"))
        ws.cell(row=i, column=5, value=sumifs("H", "B", f"$A{i}"))
        ws.cell(row=i, column=6, value=sumifs("I", "B", f"$A{i}"))
        ws.cell(row=i, column=7, value=f"=IF(B{i}=0,0,C{i}/B{i})")
        ws.cell(row=i, column=8, value=f"=IF(C{i}=0,0,F{i}/C{i})")
        ws.cell(row=i, column=9, value=f"=IF(F{i}=0,0,E{i}/F{i})")
        ws.cell(row=i, column=10, value=f"=IF(D{i}=0,0,E{i}/D{i})")
        ws.cell(row=i, column=11, value=f"=IF($E$9=0,0,E{i}/$E$9)")
    total = 9
    ws.cell(row=total, column=1, value="TOTAL")
    for col, letter in ((2, "B"), (3, "C"), (4, "D"), (5, "E"), (6, "F")):
        ws.cell(row=total, column=col, value=f"=SUM({letter}4:{letter}8)")
    ws.cell(row=total, column=7, value=f"=IF(B{total}=0,0,C{total}/B{total})")
    ws.cell(row=total, column=8, value=f"=IF(C{total}=0,0,F{total}/C{total})")
    ws.cell(row=total, column=9, value=f"=IF(F{total}=0,0,E{total}/F{total})")
    ws.cell(row=total, column=10, value=f"=IF(D{total}=0,0,E{total}/D{total})")
    ws.cell(row=total, column=11, value=f"=IF(E{total}=0,0,E{total}/E{total})")
    style_total(ws, total, len(headers))
    for r in range(4, 10):
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).number_format = COUNT
        for col in (5, 6):
            ws.cell(row=r, column=col).number_format = MONEY
        ws.cell(row=r, column=7).number_format = PCT
        ws.cell(row=r, column=8).number_format = MONEY
        ws.cell(row=r, column=9).number_format = "0.00"
        ws.cell(row=r, column=10).number_format = MONEY
        ws.cell(row=r, column=11).number_format = PCT1
    ws.freeze_panes = "A4"
    widths = [17, 13, 10, 9, 14, 13, 8, 10, 8, 11, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w


def build_campaign(wb, gt):
    ws = wb.create_sheet("Campaign")
    ws["A1"] = "Performance by Campaign"
    ws["A1"].font = TITLE_FONT
    headers = ["Campaign", "Impressions", "Clicks", "Orders", "Revenue (Rs)",
               "Ad Spend (Rs)", "CTR", "CPC", "ROAS", "AOV (Rs)",
               "Orders Share"]
    ws.append([])
    for j, h in enumerate(headers, start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, len(headers))
    campaigns = [c["campaign"] for c in gt["campaigns"]]
    for i, cp in enumerate(campaigns, start=4):
        ws.cell(row=i, column=1, value=cp)
        ws.cell(row=i, column=2, value=sumifs("E", "D", f"$A{i}"))
        ws.cell(row=i, column=3, value=sumifs("F", "D", f"$A{i}"))
        ws.cell(row=i, column=4, value=sumifs("G", "D", f"$A{i}"))
        ws.cell(row=i, column=5, value=sumifs("H", "D", f"$A{i}"))
        ws.cell(row=i, column=6, value=sumifs("I", "D", f"$A{i}"))
        ws.cell(row=i, column=7, value=f"=IF(B{i}=0,0,C{i}/B{i})")
        ws.cell(row=i, column=8, value=f"=IF(C{i}=0,0,F{i}/C{i})")
        ws.cell(row=i, column=9, value=f"=IF(F{i}=0,0,E{i}/F{i})")
        ws.cell(row=i, column=10, value=f"=IF(D{i}=0,0,E{i}/D{i})")
        ws.cell(row=i, column=11, value=f"=IF($D$10=0,0,D{i}/$D$10)")
    total = 10
    ws.cell(row=total, column=1, value="TOTAL")
    for col, letter in ((2, "B"), (3, "C"), (4, "D"), (5, "E"), (6, "F")):
        ws.cell(row=total, column=col, value=f"=SUM({letter}4:{letter}9)")
    ws.cell(row=total, column=7, value=f"=IF(B{total}=0,0,C{total}/B{total})")
    ws.cell(row=total, column=8, value=f"=IF(C{total}=0,0,F{total}/C{total})")
    ws.cell(row=total, column=9, value=f"=IF(F{total}=0,0,E{total}/F{total})")
    ws.cell(row=total, column=10, value=f"=IF(D{total}=0,0,E{total}/D{total})")
    ws.cell(row=total, column=11, value=f"=IF(D{total}=0,0,D{total}/D{total})")
    style_total(ws, total, len(headers))
    for r in range(4, 11):
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).number_format = COUNT
        for col in (5, 6):
            ws.cell(row=r, column=col).number_format = MONEY
        ws.cell(row=r, column=7).number_format = PCT
        ws.cell(row=r, column=8).number_format = MONEY
        ws.cell(row=r, column=9).number_format = "0.00"
        ws.cell(row=r, column=10).number_format = MONEY
        ws.cell(row=r, column=11).number_format = PCT1
    ws.conditional_formatting.add("I4:I9",
                                  CellIsRule(operator="lessThan", formula=["3"],
                                             font=RED_FONT))
    ws.freeze_panes = "A4"
    widths = [18, 13, 10, 9, 14, 13, 8, 10, 8, 11, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w


def build_targets(wb, targets):
    ws = wb.create_sheet("Targets")
    ws.append(["Month", "Target Revenue (Rs)"])
    style_header(ws, 1, 2)
    for i, m in enumerate(MONTHS, start=2):
        ws.cell(row=i, column=1, value=m)
        ws.cell(row=i, column=2, value=targets[m]).number_format = MONEY
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 20


def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Performance Dashboard — Ganga Retail & Services Pvt. Ltd."
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "Focus month:"
    ws["A3"].font = Font(bold=True)
    ws["B3"] = FOCUS
    ws["B3"].font = Font(bold=True, size=12)
    ws["C3"] = "Change the month (dropdown) — every KPI below refreshes."
    dv = DataValidation(type="list",
                        formula1='"' + ",".join(MONTHS) + '"',
                        allow_blank=False)
    dv.error = "Pick a month from Apr–Oct 2026."
    dv.prompt = "Choose the focus month."
    ws.add_data_validation(dv)
    dv.add("B3")
    ws["A5"] = "KPI (focus month)"
    ws["B5"] = "Value"
    style_header(ws, 5, 2)
    kpis = [
        ("Revenue (Rs)", f"=SUMIFS(Raw!$H$2:$H${RAW_HEADROOM},"
         f"Raw!$J$2:$J${RAW_HEADROOM},$B$3)", MONEY),
        ("Orders", f"=SUMIFS(Raw!$G$2:$G${RAW_HEADROOM},"
         f"Raw!$J$2:$J${RAW_HEADROOM},$B$3)", COUNT),
        ("Ad Spend (Rs)", f"=SUMIFS(Raw!$I$2:$I${RAW_HEADROOM},"
         f"Raw!$J$2:$J${RAW_HEADROOM},$B$3)", MONEY),
        ("ROAS", "=IF(B8=0,0,B6/B8)", "0.00"),
        ("CTR", f"=IF(SUMIFS(Raw!$E$2:$E${RAW_HEADROOM},Raw!$J$2:$J${RAW_HEADROOM},$B$3)=0,0,"
         f"SUMIFS(Raw!$F$2:$F${RAW_HEADROOM},Raw!$J$2:$J${RAW_HEADROOM},$B$3)/"
         f"SUMIFS(Raw!$E$2:$E${RAW_HEADROOM},Raw!$J$2:$J${RAW_HEADROOM},$B$3))", PCT),
        ("AOV (Rs)", "=IF(B7=0,0,B6/B7)", MONEY),
        ("Best Channel (all-time)",
         "=INDEX(Channel!$A$4:$A$8,MATCH(MAX(Channel!$E$4:$E$8),"
         "Channel!$E$4:$E$8,0))", None),
    ]
    for i, (label, formula, fmt) in enumerate(kpis, start=6):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=formula)
        if fmt:
            c.number_format = fmt
    ws["A14"] = "Month"
    ws["B14"] = "Revenue (Rs)"
    ws["C14"] = "Orders"
    style_header(ws, 14, 3)
    for i, m in enumerate(MONTHS, start=15):
        mr = 4 + (i - 15)
        ws.cell(row=i, column=1, value=f"=Monthly!A{mr}")
        ws.cell(row=i, column=2, value=f"=Monthly!E{mr}").number_format = MONEY
        ws.cell(row=i, column=3, value=f"=Monthly!D{mr}").number_format = COUNT
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 40

    rev_chart = BarChart()
    rev_chart.type = "col"
    rev_chart.title = "Monthly Revenue (Rs)"
    rev_chart.y_axis.title = "Rs"
    rev_chart.x_axis.title = "Month"
    rev_chart.height = 7.5
    rev_chart.width = 14
    rev_chart.add_data(Reference(wb["Monthly"], min_col=5, min_row=3,
                                 max_row=10), titles_from_data=True)
    rev_chart.set_categories(Reference(wb["Monthly"], min_col=1, min_row=4,
                                       max_row=10))
    ws.add_chart(rev_chart, "A23")

    pie = PieChart()
    pie.title = "Revenue Share by Channel"
    pie.height = 7.5
    pie.width = 14
    pie.add_data(Reference(wb["Channel"], min_col=5, min_row=3, max_row=8),
                 titles_from_data=True)
    pie.set_categories(Reference(wb["Channel"], min_col=1, min_row=4,
                                 max_row=8))
    pie.dataLabels = None
    ws.add_chart(pie, "A39")

    roas_chart = BarChart()
    roas_chart.type = "col"
    roas_chart.title = "ROAS by Campaign"
    roas_chart.y_axis.title = "ROAS (x)"
    roas_chart.height = 7.5
    roas_chart.width = 14
    roas_chart.add_data(Reference(wb["Campaign"], min_col=9, min_row=3,
                                  max_row=9), titles_from_data=True)
    roas_chart.set_categories(Reference(wb["Campaign"], min_col=1, min_row=4,
                                        max_row=9))
    ws.add_chart(roas_chart, "A55")


def build_checks(wb, rows, gt):
    ws = wb.create_sheet("Checks")
    ws["A1"] = "Integrity Checks — all must read PASS"
    ws["A1"].font = TITLE_FONT
    for j, h in enumerate(["Check", "Expected", "Actual", "Result"], start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, 4)
    n = len(rows)
    total_row = n + 2
    fm = gt["focus_month"]
    checks = [
        ("Raw row count", gt["rows"], f"=COUNT(Raw!A2:A{n + 1})", 0.5),
        ("Raw revenue total (Rs)", gt["totals"]["revenue"],
         f"=SUM(Raw!H2:H{n + 1})", 1.0),
        ("Monthly revenue total (Rs)", gt["totals"]["revenue"],
         "=Monthly!E11", 1.0),
        ("Channel revenue total (Rs)", gt["totals"]["revenue"],
         "=Channel!E9", 1.0),
        ("Campaign revenue total (Rs)", gt["totals"]["revenue"],
         "=Campaign!E10", 1.0),
        ("Focus-month revenue (Rs)", fm["revenue"], "=Dashboard!B6", 1.0),
        ("Focus-month ROAS", fm["roas"], "=Dashboard!B9", 0.01),
        ("October ROAS (Monthly)", fm["roas"], "=Monthly!I10", 0.01),
        ("Focus-month orders", fm["orders"], "=Dashboard!B7", 0.5),
    ]
    for i, (label, expected, formula, tol) in enumerate(checks, start=4):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=expected)
        ws.cell(row=i, column=3, value=formula)
        ws.cell(row=i, column=4,
               value=f'=IF(ABS(B{i}-C{i})<{tol},"PASS","FAIL")')
    i = 13
    ws.cell(row=i, column=1, value="Best channel (all-time)")
    ws.cell(row=i, column=2, value=gt["best_channel"])
    ws.cell(row=i, column=3, value="=Dashboard!B12")
    ws.cell(row=i, column=4, value=f'=IF(B{i}=C{i},"PASS","FAIL")')
    for r in range(4, 14):
        ws.cell(row=r, column=2).number_format = "#,##0.00"
        ws.cell(row=r, column=3).number_format = "#,##0.00"
    ws.cell(row=10, column=2).number_format = "0.00"
    ws.cell(row=10, column=3).number_format = "0.00"
    ws.cell(row=11, column=2).number_format = "0.00"
    ws.cell(row=11, column=3).number_format = "0.00"
    ws.cell(row=4, column=2).number_format = COUNT
    ws.cell(row=4, column=3).number_format = COUNT
    ws.cell(row=12, column=2).number_format = COUNT
    ws.cell(row=12, column=3).number_format = COUNT
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 10


def main():
    rows, targets, gt = load()
    wb = Workbook()
    build_guide(wb)
    build_raw(wb, rows)
    build_monthly(wb)
    build_channel(wb, gt)
    build_campaign(wb, gt)
    build_targets(wb, targets)
    build_dashboard(wb)
    build_checks(wb, rows, gt)
    wb.save(OUT)
    print(f"wrote {OUT} rows={len(rows)} "
          f"total_revenue={gt['totals']['revenue']:,.2f}")


if __name__ == "__main__":
    main()
