#!/usr/bin/env python3
"""DataOps pipeline: CSVs -> SQLite (SQL layer) -> Excel workbook.

Steps:
1. Load data/*.csv into an in-memory SQLite database.
2. Run the auditable queries in queries.sql (monthly aggregates, category
   and channel breakdowns) and write the results to data/sql_summary.csv.
3. Export a Power BI-ready star schema: data/fact_sales.csv and
   data/dim_products.csv.
4. Build dataops-automation-system.xlsx with openpyxl: raw sheets with
   live derived columns, a formula-driven Monthly sheet that carries a
   SQL cross-check column (Excel SUMIFS vs SQL aggregates, diff must be
   zero), a Dashboard driven by a focus-month dropdown, and a Checks
   sheet with 10 integrity checks.

Every money figure in the workbook is a live formula; the only values
written by the generator are raw inputs and the SQL cross-check numbers.
"""
import csv
import sqlite3
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
XLSX = BASE / "dataops-automation-system.xlsx"

NAVY = "1F3864"
NAVY_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SUB_FONT = Font(bold=True, size=11, color=NAVY)
MONEY = '#,##0.00'
MONEY0 = '#,##0'
PCT = '0.0%'
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def load_csv(name):
    with open(DATA / name, newline="") as f:
        return list(csv.DictReader(f))


def read_queries():
    """Parse queries.sql into {name: sql} on '-- QUERY: name' markers."""
    text = (BASE / "queries.sql").read_text()
    queries = {}
    name, buf = None, []
    for line in text.splitlines():
        if line.startswith("-- QUERY:"):
            if name:
                queries[name] = "\n".join(buf)
            name = line.split(":", 1)[1].strip()
            buf = []
        else:
            buf.append(line)
    if name:
        queries[name] = "\n".join(buf)
    return queries


def build_sqlite():
    con = sqlite3.connect(":memory:")
    cur = con.cursor()
    cur.execute("CREATE TABLE sales (order_id TEXT, date TEXT, product TEXT,"
                " category TEXT, qty INTEGER, unit_price REAL,"
                " unit_cost REAL, channel TEXT, region TEXT)")
    cur.execute("CREATE TABLE purchases (po_id TEXT, date TEXT,"
                " supplier TEXT, product TEXT, category TEXT, qty INTEGER,"
                " unit_cost REAL)")
    cur.execute("CREATE TABLE expenses (date TEXT, category TEXT,"
                " amount REAL, note TEXT)")
    cur.execute("CREATE TABLE products (product TEXT, category TEXT,"
                " unit_cost REAL, unit_price REAL)")
    sales = load_csv("sales.csv")
    cur.executemany(
        "INSERT INTO sales VALUES (?,?,?,?,?,?,?,?,?)",
        [(r["order_id"], r["date"], r["product"], r["category"],
          int(r["qty"]), float(r["unit_price"]), float(r["unit_cost"]),
          r["channel"], r["region"]) for r in sales])
    purch = load_csv("purchases.csv")
    cur.executemany(
        "INSERT INTO purchases VALUES (?,?,?,?,?,?,?)",
        [(r["po_id"], r["date"], r["supplier"], r["product"],
          r["category"], int(r["qty"]), float(r["unit_cost"]))
         for r in purch])
    exp = load_csv("expenses.csv")
    cur.executemany(
        "INSERT INTO expenses VALUES (?,?,?,?)",
        [(r["date"], r["category"], float(r["amount"]), r["note"])
         for r in exp])
    prod = load_csv("products.csv")
    cur.executemany(
        "INSERT INTO products VALUES (?,?,?,?)",
        [(r["product"], r["category"], float(r["unit_cost"]),
          float(r["unit_price"])) for r in prod])
    con.commit()
    return con, sales, purch, exp, prod


def run_pipeline():
    con, sales, purch, exp, prod = build_sqlite()
    queries = read_queries()
    cur = con.cursor()

    monthly = cur.execute(queries["monthly"]).fetchall()
    # monthly rows: (month, revenue, cogs, gross_profit, orders)
    by_category = cur.execute(queries["by_category"]).fetchall()
    by_channel = cur.execute(queries["by_channel"]).fetchall()
    exp_monthly = {r[0]: r[1] for r in
                   cur.execute(queries["expenses_monthly"]).fetchall()}
    purch_monthly = {r[0]: r[1] for r in
                     cur.execute(queries["purchases_monthly"]).fetchall()}

    # SQL summary export (the cross-check source for the workbook).
    with open(DATA / "sql_summary.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "sql_revenue", "sql_cogs", "sql_gross_profit",
                    "sql_orders", "sql_purchases", "sql_expenses"])
        for m, rev, cogs, gp, orders in monthly:
            w.writerow([m, f"{rev:.2f}", f"{cogs:.2f}", f"{gp:.2f}",
                        orders, f"{purch_monthly.get(m, 0):.2f}",
                        f"{exp_monthly.get(m, 0):.2f}"])

    # Power BI-ready star schema.
    with open(DATA / "fact_sales.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["order_id", "date", "month", "product", "category",
                    "channel", "region", "qty", "unit_price", "unit_cost",
                    "revenue", "cost", "margin"])
        for r in sales:
            qty = int(r["qty"])
            rev = qty * float(r["unit_price"])
            cost = qty * float(r["unit_cost"])
            w.writerow([r["order_id"], r["date"], r["date"][:7],
                        r["product"], r["category"], r["channel"],
                        r["region"], qty, r["unit_price"], r["unit_cost"],
                        f"{rev:.2f}", f"{cost:.2f}", f"{rev - cost:.2f}"])
    with open(DATA / "dim_products.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["product", "category", "unit_cost", "unit_price",
                    "margin_per_unit"])
        for r in prod:
            w.writerow([r["product"], r["category"], r["unit_cost"],
                        r["unit_price"],
                        float(r["unit_price"]) - float(r["unit_cost"])])

    build_workbook(sales, purch, exp, monthly, by_category, by_channel,
                   exp_monthly, purch_monthly)
    print(f"monthly rows={len(monthly)} categories={len(by_category)} "
          f"channels={len(by_channel)}")
    return monthly, by_category, by_channel


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = NAVY_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = BOX


def build_workbook(sales, purch, exp, monthly, by_category, by_channel,
                   exp_monthly, purch_monthly):
    wb = openpyxl.Workbook()
    months = [m[0] for m in monthly]
    sql_rev = {m[0]: m[1] for m in monthly}
    sql_gp = {m[0]: m[3] for m in monthly}
    n_sales = len(sales)
    n_purch = len(purch)
    n_exp = len(exp)
    s_last = n_sales + 1   # last data row on Sales (header is row 1)
    p_last = n_purch + 1
    e_last = n_exp + 1

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws["A1"] = "DataOps Automation System"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Narmada Traders, Indore — wholesale FMCG distributor. "
                "Apr–Oct 2026. Fictional demo data (seed 20261120).")
    ws["A2"].font = Font(size=11)
    lines = [
        ("How it works",
         "1) make_data.py generates the source CSVs. 2) pipeline.py loads "
         "them into SQLite and runs the auditable queries in queries.sql. "
         "3) The same script builds this workbook: raw sheets carry live "
         "derived columns, Monthly aggregates with SUMIFS and cross-checks "
         "every figure against the SQL aggregates (diff must be 0), and "
         "Dashboard is driven by the focus-month dropdown in B2."),
        ("Sheets",
         "Sales / Purchases / Expenses — raw data + live derived columns. "
         "Monthly — one row per month, all figures live formulas. "
         "Dashboard — pick a month in B2; KPIs, monthly trend and "
         "category mix update. Checks — 10 integrity checks, all PASS."),
        ("Power BI",
         "data/fact_sales.csv + data/dim_products.csv are a star schema: "
         "load both into Power BI, relate on product, and the measures "
         "Revenue = SUM(fact_sales[revenue]) and Margin % reproduce the "
         "Dashboard KPIs. data/sql_summary.csv is the SQL audit trail."),
        ("Refresh",
         "Paste new rows at the bottom of Sales / Purchases / Expenses, "
         "extend the SUMIFS ranges on Monthly, and re-run pipeline.py — "
         "or point the loader at the live CSV export and re-run."),
        ("Note",
         "This workbook is a demonstration built from fictional data. "
         "It shows the pipeline, not a client's real books."),
    ]
    r = 4
    for title, body in lines:
        ws.cell(row=r, column=1, value=title).font = SUB_FONT
        ws.cell(row=r + 1, column=1, value=body).alignment = \
            Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r + 1].height = 45
        r += 3
    ws.column_dimensions["A"].width = 130

    # ---------------- Sales ----------------
    ws = wb.create_sheet("Sales")
    headers = ["Order", "Date", "Product", "Category", "Qty", "Unit Price",
               "Unit Cost", "Channel", "Region", "Revenue", "Cost",
               "Margin", "Month"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, row in enumerate(sales, start=2):
        ws.cell(row=i, column=1, value=row["order_id"])
        ws.cell(row=i, column=2, value=row["date"])
        ws.cell(row=i, column=3, value=row["product"])
        ws.cell(row=i, column=4, value=row["category"])
        ws.cell(row=i, column=5, value=int(row["qty"]))
        ws.cell(row=i, column=6, value=float(row["unit_price"]))
        ws.cell(row=i, column=7, value=float(row["unit_cost"]))
        ws.cell(row=i, column=8, value=row["channel"])
        ws.cell(row=i, column=9, value=row["region"])
        ws.cell(row=i, column=10, value=f"=E{i}*F{i}")
        ws.cell(row=i, column=11, value=f"=E{i}*G{i}")
        ws.cell(row=i, column=12, value=f"=J{i}-K{i}")
        ws.cell(row=i, column=13, value=f"=LEFT(B{i},7)")
    for col in "JKLM":
        for i in range(2, s_last + 1):
            ws[f"{col}{i}"].number_format = MONEY
    ws.auto_filter.ref = f"A1:M{s_last}"
    ws.freeze_panes = "A2"
    widths = [10, 12, 20, 14, 7, 11, 11, 11, 10, 13, 13, 13, 9]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Purchases ----------------
    ws = wb.create_sheet("Purchases")
    headers = ["PO", "Date", "Supplier", "Product", "Category", "Qty",
               "Unit Cost", "Amount", "Month"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, row in enumerate(purch, start=2):
        ws.cell(row=i, column=1, value=row["po_id"])
        ws.cell(row=i, column=2, value=row["date"])
        ws.cell(row=i, column=3, value=row["supplier"])
        ws.cell(row=i, column=4, value=row["product"])
        ws.cell(row=i, column=5, value=row["category"])
        ws.cell(row=i, column=6, value=int(row["qty"]))
        ws.cell(row=i, column=7, value=float(row["unit_cost"]))
        ws.cell(row=i, column=8, value=f"=F{i}*G{i}")
        ws.cell(row=i, column=9, value=f"=LEFT(B{i},7)")
    for i in range(2, p_last + 1):
        ws[f"H{i}"].number_format = MONEY
    ws.auto_filter.ref = f"A1:I{p_last}"
    ws.freeze_panes = "A2"
    for i, w in enumerate([10, 12, 22, 20, 14, 7, 11, 13, 9], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Expenses ----------------
    ws = wb.create_sheet("Expenses")
    headers = ["Date", "Category", "Amount", "Note", "Month"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, row in enumerate(exp, start=2):
        ws.cell(row=i, column=1, value=row["date"])
        ws.cell(row=i, column=2, value=row["category"])
        ws.cell(row=i, column=3, value=float(row["amount"]))
        ws.cell(row=i, column=4, value=row["note"])
        ws.cell(row=i, column=5, value=f"=LEFT(A{i},7)")
    for i in range(2, e_last + 1):
        ws[f"C{i}"].number_format = MONEY
    ws.auto_filter.ref = f"A1:E{e_last}"
    ws.freeze_panes = "A2"
    for i, w in enumerate([12, 15, 13, 30, 9], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Monthly ----------------
    ws = wb.create_sheet("Monthly")
    headers = ["Month", "Sales Revenue", "COGS", "Gross Profit",
               "Purchases", "Expenses", "Net Profit", "Orders",
               "SQL Revenue", "SQL Gross Profit", "Revenue Diff",
               "GP Diff"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for j, m in enumerate(months, start=2):
        ws.cell(row=j, column=1, value=m)
        ws.cell(row=j, column=2,
               value=f"=SUMIFS(Sales!$J$2:$J${s_last},"
                     f"Sales!$M$2:$M${s_last},$A{j})")
        ws.cell(row=j, column=3,
               value=f"=SUMIFS(Sales!$K$2:$K${s_last},"
                     f"Sales!$M$2:$M${s_last},$A{j})")
        ws.cell(row=j, column=4, value=f"=B{j}-C{j}")
        ws.cell(row=j, column=5,
               value=f"=SUMIFS(Purchases!$H$2:$H${p_last},"
                     f"Purchases!$I$2:$I${p_last},$A{j})")
        ws.cell(row=j, column=6,
               value=f"=SUMIFS(Expenses!$C$2:$C${e_last},"
                     f"Expenses!$E$2:$E${e_last},$A{j})")
        ws.cell(row=j, column=7, value=f"=D{j}-F{j}")
        ws.cell(row=j, column=8,
               value=f"=COUNTIFS(Sales!$M$2:$M${s_last},$A{j})")
        ws.cell(row=j, column=9, value=round(sql_rev[m], 2))
        ws.cell(row=j, column=10, value=round(sql_gp[m], 2))
        ws.cell(row=j, column=11, value=f"=B{j}-I{j}")
        ws.cell(row=j, column=12, value=f"=D{j}-J{j}")
    tot = len(months) + 2
    ws.cell(row=tot, column=1, value="TOTAL").font = Font(bold=True)
    for col in "BCDEFGHIJKL":
        c = ws.cell(row=tot, column=ord(col) - 64)
        if col == "H":
            c.value = f"=SUM(H2:H{tot - 1})"
        else:
            c.value = f"=SUM({col}2:{col}{tot - 1})"
        c.font = Font(bold=True)
    for row in range(2, tot + 1):
        for col in "BCDEFGIJKL":
            ws[f"{col}{row}"].number_format = MONEY
        ws[f"H{row}"].number_format = MONEY0
    ws.freeze_panes = "A2"
    for i, w in enumerate([10, 14, 13, 13, 13, 13, 13, 9, 13, 15, 13, 11],
                          start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Dashboard — Narmada Traders"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Focus month:"
    ws["A2"].font = SUB_FONT
    ws["B2"] = months[-1]
    ws["B2"].font = Font(bold=True, size=12)
    ws["B2"].fill = PatternFill("solid", fgColor="FFF2CC")
    dv = DataValidation(type="list",
                        formula1='"' + ",".join(months) + '"',
                        allow_blank=False)
    dv.add("B2")
    ws.add_data_validation(dv)
    kpis = [
        ("Revenue", f"=SUMIFS(Sales!$J$2:$J${s_last},"
                    f"Sales!$M$2:$M${s_last},$B$2)", MONEY),
        ("COGS", f"=SUMIFS(Sales!$K$2:$K${s_last},"
                 f"Sales!$M$2:$M${s_last},$B$2)", MONEY),
        ("Gross Profit", f"=SUMIFS(Sales!$L$2:$L${s_last},"
                        f"Sales!$M$2:$M${s_last},$B$2)", MONEY),
        ("Gross Margin %", "=IF(B4=0,0,B6/B4)", PCT),
        ("Expenses", f"=SUMIFS(Expenses!$C$2:$C${e_last},"
                    f"Expenses!$E$2:$E${e_last},$B$2)", MONEY),
        ("Net Profit", "=B6-B8", MONEY),
        ("Orders", f"=COUNTIFS(Sales!$M$2:$M${s_last},$B$2)", MONEY0),
    ]
    ws["A3"] = "KPI (focus month)"
    ws["A3"].font = SUB_FONT
    for i, (label, formula, fmt) in enumerate(kpis, start=4):
        ws.cell(row=i, column=1, value=label)
        c = ws.cell(row=i, column=2, value=formula)
        c.number_format = fmt
        c.font = Font(bold=True)
    # Monthly net-profit table (drives the charts).
    ws["A12"] = "Monthly trend"
    ws["A12"].font = SUB_FONT
    ws["A13"] = "Month"
    ws["B13"] = "Revenue"
    ws["C13"] = "Net Profit"
    for c in ("A13", "B13", "C13"):
        ws[c].font = HEADER_FONT
        ws[c].fill = NAVY_FILL
    for j, m in enumerate(months, start=14):
        mr = j - 12  # Monthly sheet row for this month
        ws.cell(row=j, column=1, value=m)
        ws.cell(row=j, column=2, value=f"=Monthly!B{mr}")
        ws.cell(row=j, column=3, value=f"=Monthly!G{mr}")
        ws.cell(row=j, column=2).number_format = MONEY
        ws.cell(row=j, column=3).number_format = MONEY
    # Category table.
    crow = 14 + len(months) + 2
    ws.cell(row=crow - 1, column=1, value="Revenue by category").font = SUB_FONT
    ws.cell(row=crow, column=1, value="Category").font = HEADER_FONT
    ws.cell(row=crow, column=1).fill = NAVY_FILL
    ws.cell(row=crow, column=2, value="Revenue").font = HEADER_FONT
    ws.cell(row=crow, column=2).fill = NAVY_FILL
    cats = sorted({r["category"] for r in sales})
    for i, cat in enumerate(cats):
        r = crow + 1 + i
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2,
               value=f"=SUMIFS(Sales!$J$2:$J${s_last},"
                     f"Sales!$D$2:$D${s_last},$A{r})")
        ws.cell(row=r, column=2).number_format = MONEY
    cat_first = crow + 1
    cat_last = crow + len(cats)

    # Charts (data in rows: one series per row needs from_rows=True).
    bar = BarChart()
    bar.type = "col"
    bar.title = "Monthly revenue"
    bar.y_axis.title = "INR"
    bar.x_axis.title = "Month"
    data = Reference(ws, min_col=2, min_row=13, max_row=13 + len(months))
    cats_ref = Reference(ws, min_col=1, min_row=14,
                         max_row=13 + len(months))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats_ref)
    bar.shape = 4
    ws.add_chart(bar, "D3")

    line = LineChart()
    line.title = "Monthly net profit"
    line.y_axis.title = "INR"
    data = Reference(ws, min_col=3, min_row=13, max_row=13 + len(months))
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats_ref)
    ws.add_chart(line, "D18")

    pie = PieChart()
    pie.title = "Revenue by category"
    labels = Reference(ws, min_col=1, min_row=cat_first, max_row=cat_last)
    data = Reference(ws, min_col=2, min_row=crow, max_row=cat_last)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = None
    ws.add_chart(pie, f"D{crow + 1}")

    for i, w in enumerate([22, 16, 16, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws.append(["#", "Check", "Actual", "Expected", "Result"])
    style_header(ws, 1, 5)
    total_rev = sum(m[1] for m in monthly)
    total_cogs = sum(m[2] for m in monthly)
    total_gp = sum(m[3] for m in monthly)
    total_purch = sum(purch_monthly.values())
    total_exp = sum(exp_monthly.values())
    oct_row = months.index(months[-1]) + 2  # Monthly row for Oct
    checks = [
        ("Sales rows loaded", f"=COUNTA(Sales!A2:A{s_last})", n_sales),
        ("Purchase rows loaded", f"=COUNTA(Purchases!A2:A{p_last})",
         n_purch),
        ("Expense rows loaded", f"=COUNTA(Expenses!A2:A{e_last})", n_exp),
        ("Monthly revenue total = sales revenue",
         f"=Monthly!B{tot}", round(total_rev, 2)),
        ("Monthly COGS total = sales cost",
         f"=Monthly!C{tot}", round(total_cogs, 2)),
        ("Monthly purchases total", f"=Monthly!E{tot}",
         round(total_purch, 2)),
        ("Monthly expenses total", f"=Monthly!F{tot}",
         round(total_exp, 2)),
        ("SQL vs Excel revenue diff = 0",
         f"=SUMPRODUCT(ABS(Monthly!K2:K{tot - 1}))", 0),
        ("SQL vs Excel gross-profit diff = 0",
         f"=SUMPRODUCT(ABS(Monthly!L2:L{tot - 1}))", 0),
        ("Dashboard Oct revenue = Monthly Oct revenue",
         "=Dashboard!B4", round(sql_rev[months[-1]], 2)),
    ]
    for i, (label, formula, expected) in enumerate(checks, start=2):
        ws.cell(row=i, column=1, value=i - 1)
        ws.cell(row=i, column=2, value=label)
        c = ws.cell(row=i, column=3, value=formula)
        c.number_format = MONEY
        ws.cell(row=i, column=4, value=expected).number_format = MONEY
        ws.cell(row=i, column=5,
               value=f'=IF(ABS(C{i}-D{i})<0.01,"PASS","FAIL")')
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 10
    ws.freeze_panes = "A2"

    wb.save(XLSX)


if __name__ == "__main__":
    run_pipeline()
