#!/usr/bin/env python3
"""Build renewable-mis.xlsx from data/*.csv — every figure a live formula.

Sheets: Guide, Dashboard, Monthly, State, Technology, Sites, Raw_Plan,
Raw_Actual, Checks.
Rules honoured (learned the hard way): plain A1 ranges only, range
literals repeat the column letter on both ends ($F$2:$F$120), no Table
objects, chart series titles come from header cells (titles_from_data),
never plain strings. Raw sheets hold values; every summary cell is a
SUMIFS/COUNTIF formula over the raw ranges, so next month's rows flow
through by extending the ranges.
"""
import csv
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
OUT = os.path.join(HERE, "renewable-mis.xlsx")

NAVY = "1F3864"
TEAL = "2E75B6"
LIGHT = "D9E2F3"
RED = "C00000"

MONTHS = ["2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
DAYS = [30, 31, 30, 31, 31, 30, 31]
STATES = ["Rajasthan", "Gujarat", "Maharashtra", "Karnataka", "Tamil Nadu",
          "Telangana", "Andhra Pradesh", "Madhya Pradesh", "Uttar Pradesh",
          "UAE", "Singapore"]
TECHS = ["Solar", "Wind", "Hybrid", "Rooftop"]


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.reader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def main():
    sites = read_csv("sites.csv")
    plan = read_csv("plan.csv")
    actual = read_csv("actuals.csv")
    site_header, site_rows = sites[0], sites[1:]
    plan_header, plan_rows = plan[0], plan[1:]
    act_header, act_rows = actual[0], actual[1:]
    n_plan = len(plan_rows)
    n_act = len(act_rows)
    plan_last = n_plan + 1
    act_last = n_act + 1
    site_last = len(site_rows) + 1

    wb = Workbook()

    # ---- Guide ----
    ws = wb.active
    ws.title = "Guide"
    guide_lines = [
        ("RenewableMIS — Monthly MIS: Actual vs Annual Operating Plan (AOP)", True),
        ("Demo workbook built around a renewable-energy MIS automation brief: a monthly",
         False),
        ("management report comparing generation and revenue actually achieved against",
         False),
        ("the Annual Operating Plan baseline, by site, state, technology and month.", False),
        ("", False),
        ("HOW TO REFRESH (the automation this demo proves)", True),
        ("1. Drop next month's site generation/revenue export into Raw_Plan / Raw_Actual", False),
        ("   (same columns, one row per site-month). Extend the SUMIFS ranges below.", False),
        ("2. Monthly, State, Technology, Dashboard and Checks all recalculate — no", False),
        ("   hard-coded totals anywhere in the summary sheets.", False),
        ("3. In Power BI, point the model at the same two raw tables; see", False),
        ("   powerbi/MEASURES.md for the DAX equivalents of every KPI.", False),
        ("", False),
        ("SHEET MAP", True),
        ("Dashboard — 10 headline KPIs + 3 charts (technology mix, monthly trend).", False),
        ("Monthly — planned vs actual by month, with availability from outage hours.", False),
        ("State — planned vs actual by state (9 Indian states + UAE + Singapore).", False),
        ("Technology — planned vs actual by Solar / Wind / Hybrid / Rooftop.", False),
        ("Sites — the 17-site portfolio: capacity (MW) and tariff (Rs/kWh).", False),
        ("Raw_Plan / Raw_Actual — one row per site-month; derived columns are formulas.", False),
        ("Checks — 8 integrity checks; every Status must read PASS.", False),
        ("", False),
        ("DATA NOTE", True),
        ("All figures are fictional, generated deterministically (seed 20261029) for", False),
        ("this demo. Two documented events are baked in: heavy curtailment at Kutch", False),
        ("Wind Farm in Jul 2026 (18% of planned) and a 120-hour outage at Kamuthi", False),
        ("Solar Park in Sep 2026. Revenue is in Rs lakh (MWh x tariff / 100).", False),
        ("Availability % = 1 - outage hours / (sites x days x 24).", False),
    ]
    for i, (text, bold) in enumerate(guide_lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        c.font = Font(bold=bold, size=12 if i == 1 else 11, color=NAVY if i == 1 else "000000")
    ws.column_dimensions["A"].width = 110

    # ---- Sites ----
    ws = wb.create_sheet("Sites")
    ws.append(site_header)
    for r in site_rows:
        ws.append([r[0], r[1], r[2], r[3], float(r[4]), float(r[5])])
    style_header(ws, 1, 6)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{site_last}"
    widths(ws, {"A": 10, "B": 26, "C": 16, "D": 12, "E": 13, "F": 16})

    # ---- Raw_Plan ----
    ws = wb.create_sheet("Raw_Plan")
    ws.append(plan_header)
    for r in plan_rows:
        ws.append([r[0], r[1], r[2], r[3], r[4], int(r[5]), float(r[6])])
    style_header(ws, 1, 7)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{plan_last}"
    for row in ws.iter_rows(min_row=2, max_row=plan_last, min_col=6, max_col=6):
        for c in row:
            c.number_format = "#,##0"
    for row in ws.iter_rows(min_row=2, max_row=plan_last, min_col=7, max_col=7):
        for c in row:
            c.number_format = "#,##0.00"
    widths(ws, {"A": 10, "B": 26, "C": 16, "D": 12, "E": 10, "F": 14, "G": 19})

    # ---- Raw_Actual (derived columns are formulas) ----
    ws = wb.create_sheet("Raw_Actual")
    ws.append(act_header + ["VarianceMWh", "AchievementPct", "CurtailmentPct"])
    for i, r in enumerate(act_rows, start=2):
        planned = (f"SUMIFS(Raw_Plan!$F$2:$F${plan_last},"
                   f"Raw_Plan!$A$2:$A${plan_last},$A{i},"
                   f"Raw_Plan!$E$2:$E${plan_last},$E{i})")
        ws.append([r[0], r[1], r[2], r[3], r[4], int(r[5]), float(r[6]),
                   int(r[7]), int(r[8]),
                   f"=$F{i}-{planned}",
                   f"=IF({planned}=0,\"\",$F{i}/{planned})",
                   f"=IF({planned}=0,\"\",$H{i}/{planned})"])
    style_header(ws, 1, 12)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{act_last}"
    for row in ws.iter_rows(min_row=2, max_row=act_last, min_col=6, max_col=10):
        for c in row:
            c.number_format = "#,##0"
    for row in ws.iter_rows(min_row=2, max_row=act_last, min_col=11, max_col=12):
        for c in row:
            c.number_format = "0.0%"
    ws["G2"].number_format = "#,##0.00"
    for row in ws.iter_rows(min_row=2, max_row=act_last, min_col=7, max_col=7):
        for c in row:
            c.number_format = "#,##0.00"
    widths(ws, {"A": 10, "B": 26, "C": 16, "D": 12, "E": 10, "F": 12, "G": 19,
                "H": 15, "I": 12, "J": 13, "K": 14, "L": 14})

    # Range shorthands for summary formulas.
    PF = f"Raw_Plan!$F$2:$F${plan_last}"
    PG = f"Raw_Plan!$G$2:$G${plan_last}"
    PA = f"Raw_Plan!$A$2:$A${plan_last}"
    PE = f"Raw_Plan!$E$2:$E${plan_last}"
    PC = f"Raw_Plan!$C$2:$C${plan_last}"
    PD = f"Raw_Plan!$D$2:$D${plan_last}"
    AF = f"Raw_Actual!$F$2:$F${act_last}"
    AG = f"Raw_Actual!$G$2:$G${act_last}"
    AH = f"Raw_Actual!$H$2:$H${act_last}"
    AI = f"Raw_Actual!$I$2:$I${act_last}"
    AA = f"Raw_Actual!$A$2:$A${act_last}"
    AE = f"Raw_Actual!$E$2:$E${act_last}"
    AC = f"Raw_Actual!$C$2:$C${act_last}"
    AD = f"Raw_Actual!$D$2:$D${act_last}"

    # ---- Monthly ----
    ws = wb.create_sheet("Monthly")
    ws.append(["Month", "Days", "PlannedMWh", "ActualMWh", "VarianceMWh",
               "AchievementPct", "PlannedRevenueLakh", "ActualRevenueLakh",
               "RevenueAchievementPct", "CurtailmentMWh", "OutageHours", "AvailabilityPct"])
    for i, (month, days) in enumerate(zip(MONTHS, DAYS), start=2):
        ws.append([
            month, days,
            f"=SUMIFS({PF},{PE},$A{i})",
            f"=SUMIFS({AF},{AE},$A{i})",
            f"=$D{i}-$C{i}",
            f"=IF($C{i}=0,\"\",$D{i}/$C{i})",
            f"=SUMIFS({PG},{PE},$A{i})",
            f"=SUMIFS({AG},{AE},$A{i})",
            f"=IF($G{i}=0,\"\",$H{i}/$G{i})",
            f"=SUMIFS({AH},{AE},$A{i})",
            f"=SUMIFS({AI},{AE},$A{i})",
            f"=1-$K{i}/(COUNTA(Sites!$A$2:$A${site_last})*$B{i}*24)",
        ])
    # TOTAL row
    ws.append(["TOTAL", "=SUM($B$2:$B$8)", "=SUM($C$2:$C$8)", "=SUM($D$2:$D$8)",
               "=SUM($E$2:$E$8)", "=IF($C$9=0,\"\",$D$9/$C$9)",
               "=SUM($G$2:$G$8)", "=SUM($H$2:$H$8)", "=IF($G$9=0,\"\",$H$9/$G$9)",
               "=SUM($J$2:$J$8)", "=SUM($K$2:$K$8)",
               "=1-$K$9/(COUNTA(Sites!$A$2:$A$" + str(site_last) + ")*$B$9*24)"])
    style_header(ws, 1, 12)
    for c in range(1, 13):
        cell = ws.cell(row=9, column=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
    ws.freeze_panes = "A2"
    for col in ("C", "D", "E", "J", "K"):
        for row in ws.iter_rows(min_row=2, max_row=9, min_col=ord(col) - 64, max_col=ord(col) - 64):
            for c in row:
                c.number_format = "#,##0"
    for col in ("F", "I", "L"):
        for row in ws.iter_rows(min_row=2, max_row=9, min_col=ord(col) - 64, max_col=ord(col) - 64):
            for c in row:
                c.number_format = "0.0%"
    for col in ("G", "H"):
        for row in ws.iter_rows(min_row=2, max_row=9, min_col=ord(col) - 64, max_col=ord(col) - 64):
            for c in row:
                c.number_format = "#,##0.00"
    widths(ws, {"A": 10, "B": 7, "C": 13, "D": 12, "E": 13, "F": 14, "G": 19,
                "H": 18, "I": 20, "J": 15, "K": 12, "L": 14})

    # ---- State ----
    ws = wb.create_sheet("State")
    ws.append(["State", "Sites", "PlannedMWh", "ActualMWh", "AchievementPct",
               "PlannedRevenueLakh", "ActualRevenueLakh", "RevenueAchievementPct",
               "CurtailmentMWh", "OutageHours"])
    for i, state in enumerate(STATES, start=2):
        ws.append([
            state,
            f"=COUNTIF(Sites!$C$2:$C${site_last},$A{i})",
            f"=SUMIFS({PF},{PC},$A{i})",
            f"=SUMIFS({AF},{AC},$A{i})",
            f"=IF($C{i}=0,\"\",$D{i}/$C{i})",
            f"=SUMIFS({PG},{PC},$A{i})",
            f"=SUMIFS({AG},{AC},$A{i})",
            f"=IF($F{i}=0,\"\",$G{i}/$F{i})",
            f"=SUMIFS({AH},{AC},$A{i})",
            f"=SUMIFS({AI},{AC},$A{i})",
        ])
    ws.append(["TOTAL", f"=SUM($B$2:$B${len(STATES) + 1})", f"=SUM($C$2:$C${len(STATES) + 1})",
               f"=SUM($D$2:$D${len(STATES) + 1})", "", f"=SUM($F$2:$F${len(STATES) + 1})",
               f"=SUM($G$2:$G${len(STATES) + 1})", "",
               f"=SUM($I$2:$I${len(STATES) + 1})", f"=SUM($J$2:$J${len(STATES) + 1})"])
    style_header(ws, 1, 10)
    tot = len(STATES) + 2
    for c in range(1, 11):
        cell = ws.cell(row=tot, column=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
    ws.cell(row=tot, column=5, value=f"=IF($C{tot}=0,\"\",$D{tot}/$C{tot})").number_format = "0.0%"
    ws.cell(row=tot, column=8, value=f"=IF($F{tot}=0,\"\",$G{tot}/$F{tot})").number_format = "0.0%"
    ws.freeze_panes = "A2"
    for r in range(2, tot + 1):
        for col, fmt in (("C", "#,##0"), ("D", "#,##0"), ("E", "0.0%"),
                         ("F", "#,##0.00"), ("G", "#,##0.00"), ("H", "0.0%"),
                         ("I", "#,##0"), ("J", "#,##0")):
            ws[f"{col}{r}"].number_format = fmt
    widths(ws, {"A": 16, "B": 7, "C": 13, "D": 12, "E": 14, "F": 19, "G": 18,
                "H": 20, "I": 15, "J": 12})

    # ---- Technology ----
    ws = wb.create_sheet("Technology")
    ws.append(["Technology", "Sites", "PlannedMWh", "ActualMWh", "AchievementPct",
               "PlannedRevenueLakh", "ActualRevenueLakh", "RevenueAchievementPct",
               "CurtailmentMWh", "OutageHours"])
    for i, tech in enumerate(TECHS, start=2):
        ws.append([
            tech,
            f"=COUNTIF(Sites!$D$2:$D${site_last},$A{i})",
            f"=SUMIFS({PF},{PD},$A{i})",
            f"=SUMIFS({AF},{AD},$A{i})",
            f"=IF($C{i}=0,\"\",$D{i}/$C{i})",
            f"=SUMIFS({PG},{PD},$A{i})",
            f"=SUMIFS({AG},{AD},$A{i})",
            f"=IF($F{i}=0,\"\",$G{i}/$F{i})",
            f"=SUMIFS({AH},{AD},$A{i})",
            f"=SUMIFS({AI},{AD},$A{i})",
        ])
    ws.append(["TOTAL", "=SUM($B$2:$B$5)", "=SUM($C$2:$C$5)", "=SUM($D$2:$D$5)", "",
               "=SUM($F$2:$F$5)", "=SUM($G$2:$G$5)", "",
               "=SUM($I$2:$I$5)", "=SUM($J$2:$J$5)"])
    style_header(ws, 1, 10)
    for c in range(1, 11):
        cell = ws.cell(row=6, column=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
    ws.cell(row=6, column=5, value="=IF($C$6=0,\"\",$D$6/$C$6)").number_format = "0.0%"
    ws.cell(row=6, column=8, value="=IF($F$6=0,\"\",$G$6/$F$6)").number_format = "0.0%"
    ws.freeze_panes = "A2"
    for r in range(2, 7):
        for col, fmt in (("C", "#,##0"), ("D", "#,##0"), ("E", "0.0%"),
                         ("F", "#,##0.00"), ("G", "#,##0.00"), ("H", "0.0%"),
                         ("I", "#,##0"), ("J", "#,##0")):
            ws[f"{col}{r}"].number_format = fmt
    widths(ws, {"A": 12, "B": 7, "C": 13, "D": 12, "E": 14, "F": 19, "G": 18,
                "H": 20, "I": 15, "J": 12})

    # ---- Checks ----
    ws = wb.create_sheet("Checks")
    ws.append(["Check", "Expected", "Got", "Diff", "Status"])
    checks = [
        ("Raw_Plan row count (sites x months)", 119, f"=COUNTA(Raw_Plan!$A$2:$A${plan_last})"),
        ("Raw_Actual row count (sites x months)", 119, f"=COUNTA(Raw_Actual!$A$2:$A${act_last})"),
        ("Monthly planned total ties to Raw_Plan", f"=SUM({PF})", "=SUM(Monthly!$C$2:$C$8)"),
        ("Monthly actual total ties to Raw_Actual", f"=SUM({AF})", "=SUM(Monthly!$D$2:$D$8)"),
        ("State planned total ties to Raw_Plan", f"=SUM({PF})", f"=SUM(State!$C$2:$C${len(STATES) + 1})"),
        ("Technology actual total ties to Raw_Actual", f"=SUM({AF})", "=SUM(Technology!$D$2:$D$5)"),
        ("Rows with negative actual generation", 0, f"=COUNTIF({AF},\"<0\")"),
        ("Achievement % outside 0-150%", 0,
         f"=COUNTIF(Raw_Actual!$K$2:$K${act_last},\">1.5\")+COUNTIF(Raw_Actual!$K$2:$K${act_last},\"<0\")"),
    ]
    for i, (label, expected, got) in enumerate(checks, start=2):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=expected)
        ws.cell(row=i, column=3, value=got)
        ws.cell(row=i, column=4, value=f"=$C{i}-$B{i}")
        ws.cell(row=i, column=5, value=f"=IF($D{i}=0,\"PASS\",\"FAIL\")")
    style_header(ws, 1, 5)
    ws.freeze_panes = "A2"
    for r in range(2, 10):
        for col in ("B", "C", "D"):
            ws[f"{col}{r}"].number_format = "#,##0"
    widths(ws, {"A": 46, "B": 14, "C": 14, "D": 10, "E": 10})

    # ---- Dashboard (built last so it can reference the other sheets) ----
    ws = wb.create_sheet("Dashboard", 1)
    ws.cell(row=1, column=1, value="RenewableMIS — Portfolio Dashboard (Apr-Oct 2026)").font = \
        Font(bold=True, size=14, color=NAVY)
    ws.cell(row=2, column=1, value="All values are live formulas over Raw_Plan / Raw_Actual.").font = \
        Font(italic=True, color="595959")
    kpis = [
        ("Sites in portfolio", f"=COUNTA(Sites!$A$2:$A${site_last})", "#,##0"),
        ("Total Planned Generation (MWh)", f"=SUM({PF})", "#,##0"),
        ("Total Actual Generation (MWh)", f"=SUM({AF})", "#,##0"),
        ("Generation Achievement %", "=IF($B$4=0,\"\",$B$5/$B$4)", "0.0%"),
        ("Planned Revenue (Rs lakh)", f"=SUM({PG})", "#,##0.00"),
        ("Actual Revenue (Rs lakh)", f"=SUM({AG})", "#,##0.00"),
        ("Revenue Achievement %", "=IF($B$7=0,\"\",$B$8/$B$7)", "0.0%"),
        ("Curtailment (MWh)", f"=SUM({AH})", "#,##0"),
        ("Curtailment % of Planned", "=IF($B$4=0,\"\",$B$10/$B$4)", "0.0%"),
        ("Total Outage Hours", f"=SUM({AI})", "#,##0"),
        ("Weighted Availability %",
         f"=1-$B$12/(COUNTA(Sites!$A$2:$A${site_last})*SUM(Monthly!$B$2:$B$8)*24)", "0.0%"),
    ]
    # KPI rows start at row 3; value cells are B3..B13. Fix cross-references:
    # row 3 = sites, 4 = planned MWh, 5 = actual MWh, 6 = gen achievement,
    # 7 = planned rev, 8 = actual rev, 9 = rev achievement, 10 = curtailment,
    # 11 = curtailment %, 12 = outage, 13 = availability.
    kpis[3] = ("Generation Achievement %", "=IF($B$4=0,\"\",$B$5/$B$4)", "0.0%")
    kpis[6] = ("Revenue Achievement %", "=IF($B$7=0,\"\",$B$8/$B$7)", "0.0%")
    kpis[8] = ("Curtailment % of Planned", "=IF($B$4=0,\"\",$B$10/$B$4)", "0.0%")
    kpis[10] = ("Weighted Availability %",
                f"=1-$B$12/(COUNTA(Sites!$A$2:$A${site_last})*SUM(Monthly!$B$2:$B$8)*24)", "0.0%")
    for i, (label, formula, fmt) in enumerate(kpis, start=3):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=formula)
        c.number_format = fmt
        c.font = Font(bold=True, size=12)
    ws.cell(row=14, column=1, value="Charts below are wired to Technology / Monthly — they move with the raw data.").font = \
        Font(italic=True, color="595959")
    widths(ws, {"A": 34, "B": 18, "C": 20, "D": 20, "E": 20, "F": 20})

    col_chart = BarChart()
    col_chart.type = "col"
    col_chart.title = "Planned vs Actual Generation by Technology (MWh)"
    col_chart.y_axis.title = "MWh"
    col_chart.style = 10
    data = Reference(wb["Technology"], min_col=3, min_row=1, max_row=5, max_col=4)
    cats = Reference(wb["Technology"], min_col=1, min_row=2, max_row=5)
    col_chart.add_data(data, titles_from_data=True)
    col_chart.set_categories(cats)
    col_chart.height = 7.5
    col_chart.width = 15
    ws.add_chart(col_chart, "A16")

    pie = PieChart()
    pie.title = "Actual Generation Mix by Technology"
    pie.style = 10
    labels = Reference(wb["Technology"], min_col=1, min_row=2, max_row=5)
    pdata = Reference(wb["Technology"], min_col=4, min_row=1, max_row=5)
    pie.add_data(pdata, titles_from_data=True)
    pie.set_categories(labels)
    pie.height = 7.5
    pie.width = 15
    ws.add_chart(pie, "A31")

    line = LineChart()
    line.title = "Monthly Generation Achievement %"
    line.y_axis.title = "Achievement"
    line.y_axis.number_format = "0%"
    line.style = 10
    ldata = Reference(wb["Monthly"], min_col=6, min_row=1, max_row=8)
    lcats = Reference(wb["Monthly"], min_col=1, min_row=2, max_row=8)
    line.add_data(ldata, titles_from_data=True)
    line.set_categories(lcats)
    line.height = 7.5
    line.width = 15
    ws.add_chart(line, "A46")

    wb.save(OUT)
    print(f"wrote {OUT} sheets={wb.sheetnames}")


if __name__ == "__main__":
    main()
