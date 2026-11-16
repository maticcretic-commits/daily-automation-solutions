#!/usr/bin/env python3
"""Build sales-performance-dashboard.xlsx from data/orders.csv + targets.csv.

Sheets:
  Guide     — what the pack is and how to use it.
  Data      — every order (Revenue = Qty x UnitPrice, Month = LEFT(Date,7)).
  Targets   — monthly revenue target per category.
  Monthly   — per-category x month SUMIFS revenue / target / % to goal.
  Dashboard — focus-month + focus-category dropdowns driving 7 KPIs,
              a monthly revenue-vs-target table, a by-category table,
              a by-product table, and 2 native charts.
  Checks    — 10 integrity checks (all must read PASS; they assume the
              Dashboard filters are at their defaults, All / All).

Formula rules learned from earlier packs (LibreOffice 24.2):
  * "All"-fallback inside SUMPRODUCT is ((range=crit)+(crit_cell="All"))
    — parenthesised, and safe because no data row ever equals "All",
    so the factor is exactly 0 or 1, never 2.
  * SUMIFS/COUNTIFS with direct cell criteria (Monthly sheet) are fine.
  * SUMPRODUCT ranges are bounded ($5:$last), never whole columns.
  * COUNT counts numbers only; COUNTA is used for text keys.
"""
import csv
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
XLSX = BASE / "sales-performance-dashboard.xlsx"

CATEGORIES = ["Staples", "Beverages", "Personal Care", "Home Care",
              "Snacks"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 11)]
PRODUCTS = ["Atta 10kg", "Basmati Rice 5kg", "Sugar 1kg", "Tea 500g",
            "Instant Coffee 200g", "Bath Soap (pack of 4)",
            "Shampoo 650ml", "Detergent Powder 3kg",
            "Floor Cleaner 1L", "Biscuits 400g", "Namkeen 500g"]

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14)
BOLD = Font(bold=True, size=11)
MONEY = "#,##0"
MONEY2 = "#,##0.00"
PCT = "0.0%"


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT


def build():
    with open(DATA / "orders.csv", newline="") as f:
        orders = list(csv.DictReader(f))
    with open(DATA / "targets.csv", newline="") as f:
        targets = list(csv.DictReader(f))
    n = len(orders)
    ld = 4 + n  # last Data row (header row 4, data from row 5)
    assert len(targets) == 50, len(targets)

    wb = openpyxl.Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "Sales Performance Dashboard — Ganga Traders, Patna"
    ws["A1"].font = TITLE_FONT
    lines = [
        "A formula-driven Excel sales dashboard for a wholesale FMCG",
        "distributor (Jan–Oct 2026). Built as a working demo for the",
        "Freelancer project 'Interactive Sales Performance BI Dashboard'.",
        "",
        "HOW TO USE",
        "1. Pick a focus month and a focus category on the Dashboard",
        "   (cells B3:B4 — 'All' means no filter). Every KPI, table and",
        "   chart updates from live formulas; nothing is hard-coded.",
        "2. Paste new CSV exports into the Data sheet (same columns) and",
        "   extend the formulas down; Monthly and Dashboard follow.",
        "3. Targets live on the Targets sheet — one row per category x",
        "   month. Change a target and % to goal updates everywhere.",
        "",
        "SHEETS",
        "Data      — one row per order; Revenue = Qty x UnitPrice.",
        "Targets   — monthly revenue target per category.",
        "Monthly   — revenue / target / % to goal / orders per",
        "            category x month (SUMIFS over Data and Targets).",
        "Dashboard — KPIs + monthly, category and product tables,",
        "            with 2 native charts.",
        "Checks    — 10 integrity checks; all must read PASS.",
        "            (Checks assume Dashboard filters are All / All.)",
        "",
        "DATA — all fictional, generated deterministically by",
        "make_data.py (seed 20261116). No real client data.",
    ]
    for i, line in enumerate(lines, start=3):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 78

    # ---------------- Data ----------------
    ws = wb.create_sheet("Data")
    headers = ["Date", "OrderID", "Category", "Product", "Region",
               "Channel", "Qty", "UnitPrice", "Revenue", "Month"]
    ws.append([])  # row 1
    ws.append(["Orders — Ganga Traders, Patna (Jan–Oct 2026)"])  # row 2
    ws["A2"].font = BOLD
    ws.append([])  # row 3
    ws.append(headers)  # row 4
    style_header(ws, 4, len(headers))
    for i, o in enumerate(orders):
        r = 5 + i
        ws.cell(row=r, column=1, value=o["Date"])
        ws.cell(row=r, column=2, value=o["OrderID"])
        ws.cell(row=r, column=3, value=o["Category"])
        ws.cell(row=r, column=4, value=o["Product"])
        ws.cell(row=r, column=5, value=o["Region"])
        ws.cell(row=r, column=6, value=o["Channel"])
        ws.cell(row=r, column=7, value=int(o["Qty"]))
        ws.cell(row=r, column=8, value=float(o["UnitPrice"]))
        ws.cell(row=r, column=9, value=f"=G{r}*H{r}")
        ws.cell(row=r, column=10, value=f"=LEFT(A{r},7)")
    tr = ld + 1
    ws.cell(row=tr, column=1, value="TOTAL").font = BOLD
    ws.cell(row=tr, column=7, value=f"=SUM(G5:G{ld})").font = BOLD
    ws.cell(row=tr, column=9, value=f"=SUM(I5:I{ld})").font = BOLD
    for r in range(5, ld + 1):
        ws.cell(row=r, column=8).number_format = MONEY2
        ws.cell(row=r, column=9).number_format = MONEY2
    ws.cell(row=tr, column=9).number_format = MONEY2
    widths = [12, 13, 15, 22, 14, 11, 8, 11, 13, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:J{ld}"

    # ---------------- Targets ----------------
    ws = wb.create_sheet("Targets")
    ws.append([])
    ws.append(["Monthly revenue target per category (Rs)"])
    ws["A2"].font = BOLD
    ws.append([])
    ws.append(["Category", "Month", "Target"])
    style_header(ws, 4, 3)
    for i, t in enumerate(targets):
        r = 5 + i
        ws.cell(row=r, column=1, value=t["Category"])
        ws.cell(row=r, column=2, value=t["Month"])
        ws.cell(row=r, column=3, value=int(t["Target"]))
        ws.cell(row=r, column=3).number_format = MONEY
    ws.cell(row=55, column=1, value="TOTAL").font = BOLD
    ws.cell(row=55, column=3, value="=SUM(C5:C54)").font = BOLD
    ws.cell(row=55, column=3).number_format = MONEY
    for col, w in zip("ABC", [16, 10, 14]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    # ---------------- Monthly ----------------
    ws = wb.create_sheet("Monthly")
    ws.append([])
    ws.append(["Revenue / target / % to goal / orders per category x month"])
    ws["A2"].font = BOLD
    ws.append([])
    ws.append(["Category", "Month", "Revenue", "Target", "% to Goal",
               "Orders"])
    style_header(ws, 4, 6)
    r = 5
    for cat in CATEGORIES:
        for month in MONTHS:
            ws.cell(row=r, column=1, value=cat)
            ws.cell(row=r, column=2, value=month)
            ws.cell(row=r, column=3, value=(
                f"=SUMIFS(Data!$I$5:$I${ld},Data!$C$5:$C${ld},$A{r},"
                f"Data!$J$5:$J${ld},$B{r})"))
            ws.cell(row=r, column=4, value=(
                "=SUMIFS(Targets!$C$5:$C$54,Targets!$A$5:$A$54,$A"
                f"{r},Targets!$B$5:$B$54,$B{r})"))
            ws.cell(row=r, column=5, value=f'=IF(D{r}=0,"—",C{r}/D{r})')
            ws.cell(row=r, column=6, value=(
                f"=COUNTIFS(Data!$C$5:$C${ld},$A{r},"
                f"Data!$J$5:$J${ld},$B{r})"))
            ws.cell(row=r, column=3).number_format = MONEY
            ws.cell(row=r, column=4).number_format = MONEY
            ws.cell(row=r, column=5).number_format = PCT
            r += 1
    ws.cell(row=55, column=1, value="TOTAL").font = BOLD
    for col, letter in [(3, "C"), (4, "D"), (6, "F")]:
        c = ws.cell(row=55, column=col,
                   value=f"=SUM({letter}5:{letter}54)")
        c.font = BOLD
        c.number_format = MONEY if col != 6 else "0"
    ws.cell(row=55, column=5,
           value='=IF(D55=0,"—",C55/D55)').number_format = PCT
    for col, w in zip("ABCDEF", [16, 10, 14, 14, 10, 9]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Sales Performance Dashboard — Ganga Traders"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Revenue vs target, % to goal, orders and top product. "
                "Change B3 / B4 — everything below is live formulas.")
    ws["A3"] = "Focus month"
    ws["B3"] = "All"
    ws["A4"] = "Focus category"
    ws["B4"] = "All"
    for rr in (3, 4):
        ws.cell(row=rr, column=1).font = BOLD
        ws.cell(row=rr, column=2).font = BOLD
    dv_month = DataValidation(
        type="list", formula1='"All,' + ",".join(MONTHS) + '"',
        allow_blank=False)
    dv_cat = DataValidation(
        type="list", formula1='"All,' + ",".join(CATEGORIES) + '"',
        allow_blank=False)
    ws.add_data_validation(dv_month)
    ws.add_data_validation(dv_cat)
    dv_month.add("B3")
    dv_cat.add("B4")

    d = f"Data!$J$5:$J${ld}"
    dc = f"Data!$C$5:$C${ld}"
    dp = f"Data!$D$5:$D${ld}"
    di = f"Data!$I$5:$I${ld}"
    mf = f"(({d}=$B$3)+($B$3=\"All\"))"      # month factor (All-safe)
    cf = f"(({dc}=$B$4)+($B$4=\"All\"))"     # category factor (All-safe)
    tf_m = "((Targets!$B$5:$B$54=$B$3)+($B$3=\"All\"))"
    tf_c = "((Targets!$A$5:$A$54=$B$4)+($B$4=\"All\"))"

    ws["A6"] = "Key figures (Rs unless noted)"
    ws["A6"].font = BOLD
    kpis = [
        ("Revenue", f"=SUMPRODUCT({mf}*{cf}*{di})", MONEY),
        ("Target", "=SUMPRODUCT(" + tf_m + "*" + tf_c +
         "*Targets!$C$5:$C$54)", MONEY),
        ("% to Goal", '=IF(B8=0,"—",B7/B8)', PCT),
        ("Orders", f"=SUMPRODUCT({mf}*{cf})", "0"),
        ("Avg Order Value", "=IF(B10=0,0,B7/B10)", MONEY2),
        ("Top Product Revenue", "=MAX(B35:B46)", MONEY),
        ("Top Product", "=INDEX(A35:A46,MATCH(B12,B35:B46,0))", None),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        r = 7 + i
        ws.cell(row=r, column=1, value=label).font = BOLD
        c = ws.cell(row=r, column=2, value=formula)
        if fmt:
            c.number_format = fmt

    # Monthly table (row 14 title, 15 header, 16-25 data)
    ws["A14"] = "Revenue vs target by month"
    ws["A14"].font = BOLD
    headers = ["Month", "Revenue", "Target", "% to Goal"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=15, column=col, value=h)
    style_header(ws, 15, 4)
    for i, month in enumerate(MONTHS):
        r = 16 + i
        ws.cell(row=r, column=1, value=month)
        ws.cell(row=r, column=2, value=(
            f"=SUMPRODUCT(({d}=${'A'}{r})*{cf}*{di})"))
        ws.cell(row=r, column=3, value=(
            "=SUMPRODUCT((Targets!$B$5:$B$54=$A" + str(r) + ")*" +
            tf_c + "*Targets!$C$5:$C$54)"))
        ws.cell(row=r, column=4,
               value=f'=IF(C{r}=0,"—",B{r}/C{r})')
        ws.cell(row=r, column=2).number_format = MONEY
        ws.cell(row=r, column=3).number_format = MONEY
        ws.cell(row=r, column=4).number_format = PCT

    # Category table (row 26 title, 27 header, 28-32 data)
    ws["A26"] = "Revenue vs target by category"
    ws["A26"].font = BOLD
    for col, h in enumerate(headers, start=1):
        ws.cell(row=27, column=col,
               value=h if col > 1 else "Category")
    style_header(ws, 27, 4)
    for i, cat in enumerate(CATEGORIES):
        r = 28 + i
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=(
            f"=SUMPRODUCT({mf}*(Data!$C$5:$C${ld}=$A{r})*{di})"))
        ws.cell(row=r, column=3, value=(
            "=SUMPRODUCT(" + tf_m +
            f"*(Targets!$A$5:$A$54=$A{r})*Targets!$C$5:$C$54)"))
        ws.cell(row=r, column=4,
               value=f'=IF(C{r}=0,"—",B{r}/C{r})')
        ws.cell(row=r, column=2).number_format = MONEY
        ws.cell(row=r, column=3).number_format = MONEY
        ws.cell(row=r, column=4).number_format = PCT

    # Product table (row 33 title, 34 header, 35-46 data)
    ws["A33"] = "Revenue by product (drives Top Product above)"
    ws["A33"].font = BOLD
    for col, h in enumerate(["Product", "Revenue", "Orders"], start=1):
        ws.cell(row=34, column=col, value=h)
    style_header(ws, 34, 3)
    for i, prod in enumerate(PRODUCTS):
        r = 35 + i
        ws.cell(row=r, column=1, value=prod)
        ws.cell(row=r, column=2, value=(
            f"=SUMPRODUCT({mf}*{cf}*({dp}=$A{r})*{di})"))
        ws.cell(row=r, column=3, value=(
            f"=SUMPRODUCT({mf}*{cf}*({dp}=$A{r}))"))
        ws.cell(row=r, column=2).number_format = MONEY
        ws.cell(row=r, column=3).number_format = "0"

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 3

    # Charts
    ch1 = BarChart()
    ch1.type = "col"
    ch1.style = 10
    ch1.title = "Monthly revenue vs target (Rs)"
    ch1.y_axis.title = "Rs"
    ch1.x_axis.title = "Month"
    ch1.height = 7.5
    ch1.width = 15
    data_ref = Reference(ws, min_col=2, min_row=15, max_row=25, max_col=3)
    cats_ref = Reference(ws, min_col=1, min_row=16, max_row=25)
    ch1.add_data(data_ref, titles_from_data=True)
    ch1.set_categories(cats_ref)
    ch1.shape = 4
    ws.add_chart(ch1, "F6")

    ch2 = BarChart()
    ch2.type = "col"
    ch2.style = 10
    ch2.title = "Revenue by category (Rs)"
    ch2.y_axis.title = "Rs"
    ch2.height = 7.5
    ch2.width = 15
    data_ref2 = Reference(ws, min_col=2, min_row=27, max_row=32)
    cats_ref2 = Reference(ws, min_col=1, min_row=28, max_row=32)
    ch2.add_data(data_ref2, titles_from_data=True)
    ch2.set_categories(cats_ref2)
    ch2.shape = 4
    ws.add_chart(ch2, "F23")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws["A1"] = "Integrity checks — all must read PASS"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Checks assume the Dashboard filters are at their "
                "defaults (All / All).")
    ws.append([])
    ws.append(["#", "Check", "Expected", "Actual", "Result"])
    style_header(ws, 4, 5)
    checks = [
        ("Orders in Data = generated rows",
         n, f"=COUNTA(Data!A5:A{ld})", None),
        ("Data revenue = Monthly revenue total",
         0, f"=SUM(Data!I5:I{ld})-SUM(Monthly!C5:C54)", "tol"),
        ("Order count = Monthly orders total",
         0, f"=COUNTA(Data!A5:A{ld})-SUM(Monthly!F5:F54)", "tol"),
        ("Targets total = Monthly target total",
         0, "=SUM(Targets!C5:C54)-SUM(Monthly!D5:D54)", "tol"),
        ("Dashboard revenue (All/All) = Monthly revenue",
         0, "=Dashboard!B7-SUM(Monthly!C5:C54)", "tol"),
        ("Overall % to goal within 80–120%",
         "0.8–1.2", "=Dashboard!B9", "band"),
        ("No blank Category in Data",
         0, f"=COUNTBLANK(Data!C5:C{ld})", None),
        ("No non-positive revenue in Data",
         0, f'=COUNTIF(Data!I5:I{ld},"<=0")', None),
        ("Top product = MAX(product table)",
         0, "=Dashboard!B12-MAX(Dashboard!B35:B46)", "tol"),
        ("Category table sums to Dashboard revenue",
         0, "=SUM(Dashboard!B28:B32)-Dashboard!B7", "tol"),
    ]
    for i, (label, expected, actual, kind) in enumerate(checks):
        r = 5 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=expected)
        ws.cell(row=r, column=4, value=actual)
        if kind == "tol":
            result = f'=IF(ABS(D{r})<0.01,"PASS","FAIL")'
        elif kind == "band":
            result = f'=IF(AND(D{r}>=0.8,D{r}<=1.2),"PASS","FAIL")'
        else:
            result = f'=IF(D{r}=C{r},"PASS","FAIL")'
        ws.cell(row=r, column=5, value=result)
        ws.cell(row=r, column=5).font = BOLD
    for col, w in zip("ABCDE", [5, 48, 12, 16, 9]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    wb.save(XLSX)
    print(f"built {XLSX} ({n} orders)")


if __name__ == "__main__":
    build()
