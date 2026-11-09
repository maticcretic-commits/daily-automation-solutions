#!/usr/bin/env python3
"""SalesDrill — interactive sales dashboard factory (Excel).

Built around a real Freelancer brief ("Interactive Sales Dashboard
Builder", project 40808588, Delhi): turn Excel/CSV order history into a
single interactive sales/revenue dashboard with drill-down by
product/category/customer, period comparison and a refresh workflow.

What this builds: that dashboard itself, as a formula-driven workbook.
A fictional Delhi wholesaler ("Delhi Traders", Apr-Sep 2026, 646 orders,
deterministic data/orders.csv) is consolidated into: Orders (every line a
live Amount/Cost/Margin formula), Products master, Monthly rollup with
prior-month comparison, Category drill-down, Customer drill-down (top 30
by revenue), Dashboard (10 formula-wired KPIs + 3 native charts), Checks
(8 integrity checks). No macros, no structured table references, plain
A1 ranges with the column letter repeated on both ends (LibreOffice-safe),
real date values (never text dates), percentages stored as 0-1 fractions.

Refresh workflow (documented on the Guide sheet): paste new order rows
into Orders, copy the Amount/Cost/Margin/Month formulas down, extend the
range ends in Monthly/Category/Customers/Checks to the new last row.
"""
import csv
import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
OUT = BASE / "sales-drilldown-dashboard.xlsx"

NAVY = "1F3864"
TEAL = "0F7B7A"
LIGHT = "D9E2F3"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color=NAVY, size=14)
SECTION_FONT = Font(bold=True, color=NAVY, size=11)
INR = "\u20b9#,##0"
INR2 = "\u20b9#,##0.00"
PCT1 = "0.0%"
DATEFMT = "yyyy-mm-dd"

MONTHS = [("Apr 2026", datetime.date(2026, 4, 1)),
          ("May 2026", datetime.date(2026, 5, 1)),
          ("Jun 2026", datetime.date(2026, 6, 1)),
          ("Jul 2026", datetime.date(2026, 7, 1)),
          ("Aug 2026", datetime.date(2026, 8, 1)),
          ("Sep 2026", datetime.date(2026, 9, 1))]


def read_csv(name):
    with open(BASE / "data" / name, newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def build():
    orders = read_csv("orders.csv")
    products = read_csv("products.csv")
    customers = read_csv("customers.csv")
    n = len(orders)
    last = n + 1  # last data row in Orders (row 1 is the header)

    # Revenue per customer, for the drill-down sort order (static order;
    # the values themselves are live formulas).
    rev_by_cust = {}
    for r in orders:
        rev_by_cust[r["customer"]] = rev_by_cust.get(r["customer"], 0) \
            + int(r["qty"]) * int(r["unit_price"])
    cust_order = sorted(customers, key=lambda c: -rev_by_cust.get(c["customer"], 0))
    categories = sorted({p["category"] for p in products})

    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    widths(ws, {"A": 28, "B": 110})
    ws["A1"] = "SalesDrill — Interactive Sales Dashboard"
    ws["A1"].font = TITLE_FONT
    guide = [
        ("What this is", "A formula-driven sales dashboard built from order "
         "history (CSV). Every figure on Monthly / Category / Customers / "
         "Dashboard is a live Excel formula over the Orders sheet — change "
         "an order and the whole pack updates."),
        ("The business", "Delhi Traders, a fictional Delhi wholesaler. "
         f"{n} orders, Apr-Sep 2026, {len(products)} products in "
         f"{len(categories)} categories, {len(customers)} customers. "
         "All data is fictional and deterministic (data/*.csv)."),
        ("Sheets", "Orders — raw order lines (Amount/Cost/Margin/Month are "
         "formulas). Products — price & cost master. Monthly — revenue, "
         "orders, units, margin + prior-month comparison. Category — "
         "drill-down by category. Customers — drill-down by customer, "
         "ranked by revenue. Dashboard — 10 KPIs + 3 charts. Checks — "
         "8 integrity checks, all must read PASS."),
        ("How to refresh", "1) Paste new order rows at the bottom of "
         "Orders (columns A-H). 2) Copy the formulas in I, J, K, L down "
         "to the new rows. 3) Extend the range ends in Monthly, Category, "
         "Customers and Checks from row "
         f"{last} to the new last row. 4) Re-check the Checks sheet."),
        ("Reading the KPIs", "Best Month / Top Category / Top Customer are "
         "INDEX+MATCH over the live rollups. Repeat-customer % counts "
         "customers with more than one order. Margin uses the Products "
         "cost master via VLOOKUP — update a cost there and margins move."),
        ("Built from", "Freelancer project 40808588, 'Interactive Sales "
         "Dashboard Builder' (Delhi, INR 100-400/hour, avg bid INR 250). "
         "This workbook is a demo reference implementation, not client work."),
    ]
    r = 3
    for title, text in guide:
        ws.cell(row=r, column=1, value=title).font = SECTION_FONT
        c = ws.cell(row=r, column=2, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 42
        r += 1

    # ---------------- Products ----------------
    ws = wb.create_sheet("Products")
    ws.append(["Product", "Category", "Unit Price (INR)", "Cost Price (INR)"])
    style_header(ws, 1, 4)
    for p in products:
        ws.append([p["product"], p["category"], int(p["unit_price"]),
                   int(p["cost_price"])])
    for row in ws.iter_rows(min_row=2, max_row=len(products) + 1,
                            min_col=3, max_col=4):
        for cell in row:
            cell.number_format = INR
    widths(ws, {"A": 24, "B": 18, "C": 18, "D": 18})

    # ---------------- Orders ----------------
    ws = wb.create_sheet("Orders")
    headers = ["Order ID", "Date", "Customer", "Region", "Category",
               "Product", "Qty", "Unit Price (INR)", "Amount (INR)",
               "Cost (INR)", "Margin (INR)", "Month"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    prod_last = len(products) + 1
    for i, o in enumerate(orders):
        row = i + 2
        d = datetime.date.fromisoformat(o["date"])
        ws.cell(row=row, column=1, value=o["order_id"])
        c = ws.cell(row=row, column=2, value=d)
        c.number_format = DATEFMT
        ws.cell(row=row, column=3, value=o["customer"])
        ws.cell(row=row, column=4, value=o["region"])
        ws.cell(row=row, column=5, value=o["category"])
        ws.cell(row=row, column=6, value=o["product"])
        ws.cell(row=row, column=7, value=int(o["qty"]))
        u = ws.cell(row=row, column=8, value=int(o["unit_price"]))
        u.number_format = INR
        a = ws.cell(row=row, column=9)
        a.value = f"=G{row}*H{row}"
        a.number_format = INR
        co = ws.cell(row=row, column=10)
        co.value = (f"=G{row}*VLOOKUP(F{row},Products!$A$2:$D${prod_last},4,FALSE)")
        co.number_format = INR
        m = ws.cell(row=row, column=11)
        m.value = f"=I{row}-J{row}"
        m.number_format = INR
        mo = ws.cell(row=row, column=12)
        mo.value = f'=TEXT(B{row},"mmm yyyy")'
    ws.auto_filter.ref = f"A1:L{last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 12, "B": 12, "C": 22, "D": 14, "E": 16, "F": 22,
                "G": 7, "H": 16, "I": 15, "J": 14, "K": 14, "L": 11})

    O = f"Orders!$L$2:$L${last}"   # month helper range
    OI = f"Orders!$I$2:$I${last}"  # amount range
    OG = f"Orders!$G$2:$G${last}"  # qty range
    OK = f"Orders!$K$2:$K${last}"  # margin range
    OE = f"Orders!$E$2:$E${last}"  # category range
    OC = f"Orders!$C$2:$C${last}"  # customer range

    # ---------------- Monthly ----------------
    ws = wb.create_sheet("Monthly")
    ws.append(["Month", "Revenue (INR)", "Orders", "Units", "Avg Order (INR)",
               "Margin (INR)", "Margin %", "Prior Month Revenue (INR)",
               "MoM Change"])
    style_header(ws, 1, 9)
    for i, (label, _d) in enumerate(MONTHS):
        row = i + 2
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2,
               value=f'=SUMIFS({OI},{O},$A{row})').number_format = INR
        ws.cell(row=row, column=3,
               value=f'=COUNTIFS({O},$A{row})')
        ws.cell(row=row, column=4,
               value=f'=SUMIFS({OG},{O},$A{row})')
        e = ws.cell(row=row, column=5, value=f'=IF(C{row}=0,0,B{row}/C{row})')
        e.number_format = INR
        ws.cell(row=row, column=6,
               value=f'=SUMIFS({OK},{O},$A{row})').number_format = INR
        g = ws.cell(row=row, column=7, value=f'=IF(B{row}=0,0,F{row}/B{row})')
        g.number_format = PCT1
        if i == 0:
            ws.cell(row=row, column=8, value=0).number_format = INR
            ws.cell(row=row, column=9, value=0).number_format = PCT1
        else:
            ws.cell(row=row, column=8,
                   value=f"=B{row - 1}").number_format = INR
            ch = ws.cell(row=row, column=9,
                        value=f'=IF(H{row}=0,0,(B{row}-H{row})/H{row})')
            ch.number_format = PCT1
    t = len(MONTHS) + 2
    ws.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
    for col, fmt in [(2, INR), (3, None), (4, None), (6, INR)]:
        c = ws.cell(row=t, column=col,
                   value=f"=SUM({get_column_letter(col)}2:{get_column_letter(col)}{t - 1})")
        c.font = Font(bold=True)
        if fmt:
            c.number_format = fmt
    c = ws.cell(row=t, column=5, value=f"=IF(C{t}=0,0,B{t}/C{t})")
    c.font = Font(bold=True)
    c.number_format = INR
    c = ws.cell(row=t, column=7, value=f"=IF(B{t}=0,0,F{t}/B{t})")
    c.font = Font(bold=True)
    c.number_format = PCT1
    widths(ws, {"A": 12, "B": 17, "C": 9, "D": 9, "E": 16, "F": 16,
                "G": 10, "H": 22, "I": 12})
    m_last = t - 1  # last month data row

    # ---------------- Category ----------------
    ws = wb.create_sheet("Category")
    ws.append(["Category", "Revenue (INR)", "Orders", "Units",
               "Margin (INR)", "Share of Revenue", "Avg Order (INR)"])
    style_header(ws, 1, 7)
    for i, cat in enumerate(categories):
        row = i + 2
        ws.cell(row=row, column=1, value=cat)
        ws.cell(row=row, column=2,
               value=f'=SUMIFS({OI},{OE},$A{row})').number_format = INR
        ws.cell(row=row, column=3,
               value=f'=COUNTIFS({OE},$A{row})')
        ws.cell(row=row, column=4,
               value=f'=SUMIFS({OG},{OE},$A{row})')
        ws.cell(row=row, column=5,
               value=f'=SUMIFS({OK},{OE},$A{row})').number_format = INR
        f = ws.cell(row=row, column=6,
                   value=f"=IF(Dashboard!$B$2=0,0,B{row}/Dashboard!$B$2)")
        f.number_format = PCT1
        g = ws.cell(row=row, column=7, value=f"=IF(C{row}=0,0,B{row}/C{row})")
        g.number_format = INR
    widths(ws, {"A": 18, "B": 17, "C": 9, "D": 9, "E": 16, "F": 17, "G": 16})
    cat_last = len(categories) + 1

    # ---------------- Customers ----------------
    ws = wb.create_sheet("Customers")
    ws.append(["Customer", "Region", "Orders", "Revenue (INR)",
               "Share of Revenue", "Avg Order (INR)"])
    style_header(ws, 1, 6)
    region_of = {c["customer"]: c["region"] for c in customers}
    for i, c in enumerate(cust_order):
        row = i + 2
        ws.cell(row=row, column=1, value=c["customer"])
        ws.cell(row=row, column=2, value=region_of[c["customer"]])
        ws.cell(row=row, column=3,
               value=f'=COUNTIFS({OC},$A{row})')
        ws.cell(row=row, column=4,
               value=f'=SUMIFS({OI},{OC},$A{row})').number_format = INR
        e = ws.cell(row=row, column=5,
                   value=f"=IF(Dashboard!$B$2=0,0,D{row}/Dashboard!$B$2)")
        e.number_format = PCT1
        f = ws.cell(row=row, column=6, value=f"=IF(C{row}=0,0,D{row}/C{row})")
        f.number_format = INR
    widths(ws, {"A": 24, "B": 14, "C": 9, "D": 17, "E": 17, "F": 16})
    cust_last = len(cust_order) + 1

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "SalesDrill — Sales Dashboard"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Delhi Traders  •  Apr-Sep 2026  •  all figures live formulas over Orders"
    kpis = [
        ("Total Revenue", f"=SUM(Monthly!B2:B{m_last})", INR),
        ("Total Orders", f"=SUM(Monthly!C2:C{m_last})", None),
        ("Avg Order Value", "=IF(B3=0,0,B2/B3)", INR),
        ("Total Margin", f"=SUM(Monthly!F2:F{m_last})", INR),
        ("Margin %", "=IF(B2=0,0,B5/B2)", PCT1),
        ("Best Month", f"=INDEX(Monthly!A2:A{m_last},MATCH(MAX(Monthly!B2:B{m_last}),Monthly!B2:B{m_last},0))", None),
        ("Best Month Revenue", f"=MAX(Monthly!B2:B{m_last})", INR),
        ("Top Category", f"=INDEX(Category!A2:A{cat_last},MATCH(MAX(Category!B2:B{cat_last}),Category!B2:B{cat_last},0))", None),
        ("Top Customer", f"=INDEX(Customers!A2:A{cust_last},MATCH(MAX(Customers!D2:D{cust_last}),Customers!D2:D{cust_last},0))", None),
        ("Repeat-Customer %", f'=COUNTIF(Customers!C2:C{cust_last},">1")/COUNTA(Customers!A2:A{cust_last})', PCT1),
    ]
    # KPI rows start at row 4 (row 3 is the header row for the block).
    ws["A3"] = "KPI"
    ws["B3"] = "Value"
    style_header(ws, 3, 2)
    for i, (label, formula, fmt) in enumerate(kpis):
        row = 4 + i
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=row, column=2, value=formula)
        if fmt:
            c.number_format = fmt
    widths(ws, {"A": 22, "B": 34, "C": 4, "D": 30, "E": 30})
    # The KPI value cells are B4..B13; fix the formulas that referenced B2/B3/B5
    # (written above against the Dashboard layout) by rewriting them now.
    ws["B4"] = f"=SUM(Monthly!B2:B{m_last})"            # Total Revenue
    ws["B5"] = f"=SUM(Monthly!C2:C{m_last})"            # Total Orders
    ws["B6"] = "=IF(B5=0,0,B4/B5)"                       # Avg Order Value
    ws["B7"] = f"=SUM(Monthly!F2:F{m_last})"            # Total Margin
    ws["B8"] = "=IF(B4=0,0,B7/B4)"                       # Margin %
    ws["B9"] = (f"=INDEX(Monthly!A2:A{m_last},MATCH(MAX(Monthly!B2:B{m_last}),"
                f"Monthly!B2:B{m_last},0))")             # Best Month
    ws["B10"] = f"=MAX(Monthly!B2:B{m_last})"           # Best Month Revenue
    ws["B11"] = (f"=INDEX(Category!A2:A{cat_last},MATCH(MAX(Category!B2:B{cat_last}),"
                 f"Category!B2:B{cat_last},0))")         # Top Category
    ws["B12"] = (f"=INDEX(Customers!A2:A{cust_last},MATCH(MAX(Customers!D2:D{cust_last}),"
                 f"Customers!D2:D{cust_last},0))")       # Top Customer
    ws["B13"] = (f'=COUNTIF(Customers!C2:C{cust_last},">1")'
                 f"/COUNTA(Customers!A2:A{cust_last})")  # Repeat %
    for row, fmt in [(4, INR), (6, INR), (7, INR), (8, PCT1), (10, INR), (13, PCT1)]:
        ws.cell(row=row, column=2).number_format = fmt
    # Category/Customer share formulas reference Dashboard!$B$2 — point them
    # at the real Total Revenue cell B4.
    for sh in ("Category", "Customers"):
        w2 = wb[sh]
        col = 6 if sh == "Category" else 5
        for row in range(2, (cat_last if sh == "Category" else cust_last) + 1):
            valcol = "B" if sh == "Category" else "D"
            w2.cell(row=row, column=col).value = (
                f"=IF(Dashboard!$B$4=0,0,{valcol}{row}/Dashboard!$B$4)")

    # Charts (native openpyxl charts over the live rollups).
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "Monthly Revenue (INR)"
    ch1.y_axis.title = "INR"
    ch1.x_axis.title = "Month"
    ch1.style = 10
    ch1.height = 7.5
    ch1.width = 15
    ch1.add_data(Reference(wb["Monthly"], min_col=2, min_row=1,
                           max_row=m_last), titles_from_data=True)
    ch1.set_categories(Reference(wb["Monthly"], min_col=1, min_row=2,
                                 max_row=m_last))
    ch1.shape = 4
    ws.add_chart(ch1, "D3")

    ch2 = PieChart()
    ch2.title = "Revenue Share by Category"
    ch2.style = 10
    ch2.height = 7.5
    ch2.width = 15
    labels = Reference(wb["Category"], min_col=1, min_row=2, max_row=cat_last)
    data = Reference(wb["Category"], min_col=2, min_row=1, max_row=cat_last)
    ch2.add_data(data, titles_from_data=True)
    ch2.set_categories(labels)
    ch2.dataLabels = None
    ws.add_chart(ch2, "D19")

    ch3 = BarChart()
    ch3.type = "bar"
    ch3.title = "Top 10 Customers by Revenue (INR)"
    ch3.style = 10
    ch3.height = 7.5
    ch3.width = 15
    ch3.add_data(Reference(wb["Customers"], min_col=4, min_row=1,
                           max_row=11), titles_from_data=True)
    ch3.set_categories(Reference(wb["Customers"], min_col=1, min_row=2,
                                 max_row=11))
    ws.add_chart(ch3, "D35")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws.append(["#", "Check", "Expected", "Actual", "Result"])
    style_header(ws, 1, 5)
    checks = [
        ("Orders Amount total equals Monthly revenue total",
         f"=SUM({OI})", f"=Monthly!B{t}"),
        ("Orders count equals Monthly orders total",
         f"=COUNTA(Orders!A2:A{last})", f"=Monthly!C{t}"),
        ("Orders Margin total equals Monthly margin total",
         f"=SUM({OK})", f"=Monthly!F{t}"),
        ("Category revenue total equals Total Revenue",
         f"=SUM(Category!B2:B{cat_last})", "=Dashboard!B4"),
        ("Customer revenue total equals Total Revenue",
         f"=SUM(Customers!D2:D{cust_last})", "=Dashboard!B4"),
        ("Orders with Amount <= 0 (must be 0)", 0,
         f"=SUMPRODUCT(--({OI}<=0))"),
        ("Orders dated outside Apr-Sep 2026 (must be 0)", 0,
         f'=COUNTIFS(Orders!$B$2:$B${last},"<"&DATE(2026,4,1))'
         f'+COUNTIFS(Orders!$B$2:$B${last},">"&DATE(2026,9,30))'),
        ("Blank Product or Customer cells (must be 0)", 0,
         f"=COUNTBLANK(Orders!F2:F{last})+COUNTBLANK(Orders!C2:C{last})"),
    ]
    for i, (label, expected, actual) in enumerate(checks):
        row = i + 2
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=2, value=label)
        e = ws.cell(row=row, column=3, value=expected)
        a = ws.cell(row=row, column=4, value=actual)
        if isinstance(expected, (int, float)):
            e.number_format = INR if expected else "0"
        else:
            e.number_format = INR
        a.number_format = INR
        ws.cell(row=row, column=5,
               value=f'=IF(C{row}=D{row},"PASS","FAIL")')
    widths(ws, {"A": 5, "B": 52, "C": 18, "D": 18, "E": 10})

    wb.save(OUT)
    print(f"built {OUT} — {n} orders, {len(products)} products, "
          f"{len(customers)} customers")


if __name__ == "__main__":
    build()
