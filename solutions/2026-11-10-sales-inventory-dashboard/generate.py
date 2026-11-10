#!/usr/bin/env python3
"""SalesInv — sales & inventory dashboard factory (Excel + SQLite).

Built around a real Freelancer brief ("Sales Insights & Power BI
Dashboard", Siuri, India, INR 750-1,250/hour): a sales and inventory
Power BI-style dashboard (profit, customer, product and monthly trends)
with reusable SQL reporting tables and Excel validation.

What this builds, from fictional Siuri Traders data (May-Oct 2026,
520 orders, 14 products, 24 customers, deterministic data/*.csv):
1. sales_inventory.db — SQLite with the raw tables plus reusable
   reporting views: v_monthly_sales, v_profit_by_product,
   v_profit_by_category, v_customer_ranking, v_inventory_status,
   v_reorder_alerts. The views are the "reusable SQL reporting tables"
   from the brief — any BI tool can query them directly.
2. sales-inventory-dashboard.xlsx — a formula-driven workbook:
   Purchases (stock-in), Sales (Amount/Cost/Margin/Month formulas +
   data-validation dropdowns for Product and Region), Products master,
   Inventory (closing stock, stock value, OUT/LOW/OK status, days of
   cover — all live formulas), Monthly rollup with prior-month
   comparison, Category and Customer drill-downs, Dashboard (14 KPIs +
   3 native charts), Checks (10 integrity checks).

Conventions (LibreOffice-safe, as proven on the 2026-11-09 build):
plain A1 ranges with the column letter repeated on both ends, real date
values (never text dates), percentages stored as 0-1 fractions, numeric
columns guarded to 0 instead of "".
"""
import csv
import datetime
import sqlite3
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent
OUT_XLSX = BASE / "sales-inventory-dashboard.xlsx"
OUT_DB = BASE / "sales_inventory.db"

NAVY = "1F3864"
TEAL = "0F7B7A"
LIGHT = "D9E2F3"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color=NAVY, size=14)
SECTION_FONT = Font(bold=True, color=NAVY, size=11)
INR = "\u20b9#,##0"
PCT1 = "0.0%"
ONE_DEC = "0.0"
DATEFMT = "yyyy-mm-dd"

MONTHS = [("May 2026", datetime.date(2026, 5, 1)),
          ("Jun 2026", datetime.date(2026, 6, 1)),
          ("Jul 2026", datetime.date(2026, 7, 1)),
          ("Aug 2026", datetime.date(2026, 8, 1)),
          ("Sep 2026", datetime.date(2026, 9, 1)),
          ("Oct 2026", datetime.date(2026, 10, 1))]
PERIOD_DAYS = 184  # 1 May - 31 Oct 2026


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


def build_db(products, purchases, sales, customers):
    if OUT_DB.exists():
        OUT_DB.unlink()
    con = sqlite3.connect(OUT_DB)
    cur = con.cursor()
    cur.executescript("""
        CREATE TABLE products (product TEXT PRIMARY KEY, category TEXT,
            unit_price INTEGER, cost_price INTEGER, opening_stock INTEGER,
            reorder_level INTEGER);
        CREATE TABLE customers (customer TEXT PRIMARY KEY, region TEXT);
        CREATE TABLE purchases (purchase_id TEXT PRIMARY KEY, date TEXT,
            product TEXT, qty INTEGER, unit_cost INTEGER,
            FOREIGN KEY (product) REFERENCES products(product));
        CREATE TABLE sales (order_id TEXT PRIMARY KEY, date TEXT,
            customer TEXT, region TEXT, category TEXT, product TEXT,
            qty INTEGER, unit_price INTEGER,
            FOREIGN KEY (product) REFERENCES products(product));
        CREATE VIEW v_monthly_sales AS
            SELECT substr(sales.date, 1, 7) AS month,
                SUM(sales.qty * sales.unit_price) AS revenue,
                COUNT(*) AS orders, SUM(sales.qty) AS units,
                SUM(sales.qty * (sales.unit_price - products.cost_price)) AS margin
            FROM sales JOIN products USING (product)
            GROUP BY month ORDER BY month;
        CREATE VIEW v_profit_by_product AS
            SELECT product, products.category, SUM(sales.qty) AS units,
                SUM(sales.qty * sales.unit_price) AS revenue,
                SUM(sales.qty * (sales.unit_price - products.cost_price)) AS margin
            FROM sales JOIN products USING (product)
            GROUP BY product ORDER BY revenue DESC;
        CREATE VIEW v_profit_by_category AS
            SELECT products.category, SUM(sales.qty) AS units,
                SUM(sales.qty * sales.unit_price) AS revenue,
                SUM(sales.qty * (sales.unit_price - products.cost_price)) AS margin
            FROM sales JOIN products USING (product)
            GROUP BY products.category ORDER BY revenue DESC;
        CREATE VIEW v_customer_ranking AS
            SELECT customer, region, COUNT(*) AS orders,
                SUM(qty * unit_price) AS revenue
            FROM sales GROUP BY customer ORDER BY revenue DESC;
        CREATE VIEW v_inventory_status AS
            SELECT p.product, p.category, p.opening_stock,
                COALESCE(pur.purchased, 0) AS purchased,
                COALESCE(s.sold, 0) AS sold,
                p.opening_stock + COALESCE(pur.purchased, 0)
                    - COALESCE(s.sold, 0) AS closing,
                p.reorder_level,
                (p.opening_stock + COALESCE(pur.purchased, 0)
                    - COALESCE(s.sold, 0)) * p.cost_price AS stock_value,
                CASE WHEN p.opening_stock + COALESCE(pur.purchased, 0)
                        - COALESCE(s.sold, 0) <= 0 THEN 'OUT'
                     WHEN p.opening_stock + COALESCE(pur.purchased, 0)
                        - COALESCE(s.sold, 0) <= p.reorder_level THEN 'LOW'
                     ELSE 'OK' END AS status
            FROM products p
            LEFT JOIN (SELECT product, SUM(qty) AS purchased
                       FROM purchases GROUP BY product) pur
                ON pur.product = p.product
            LEFT JOIN (SELECT product, SUM(qty) AS sold
                       FROM sales GROUP BY product) s
                ON s.product = p.product
            ORDER BY closing;
        CREATE VIEW v_reorder_alerts AS
            SELECT product, category, closing, reorder_level, status
            FROM v_inventory_status WHERE status IN ('LOW', 'OUT')
            ORDER BY closing;
    """)
    cur.executemany("INSERT INTO products VALUES (?,?,?,?,?,?)",
                    [(p["product"], p["category"], int(p["unit_price"]),
                      int(p["cost_price"]), int(p["opening_stock"]),
                      int(p["reorder_level"])) for p in products])
    cur.executemany("INSERT INTO customers VALUES (?,?)",
                    [(c["customer"], c["region"]) for c in customers])
    cur.executemany("INSERT INTO purchases VALUES (?,?,?,?,?)",
                    [(p["purchase_id"], p["date"], p["product"],
                      int(p["qty"]), int(p["unit_cost"])) for p in purchases])
    cur.executemany("INSERT INTO sales VALUES (?,?,?,?,?,?,?,?)",
                    [(s["order_id"], s["date"], s["customer"], s["region"],
                      s["category"], s["product"], int(s["qty"]),
                      int(s["unit_price"])) for s in sales])
    con.commit()
    con.close()


def build():
    products = read_csv("products.csv")
    purchases = read_csv("purchases.csv")
    sales = read_csv("sales.csv")
    customers = read_csv("customers.csv")
    build_db(products, purchases, sales, customers)

    n_sales = len(sales)
    sales_last = n_sales + 1
    n_pur = len(purchases)
    pur_last = n_pur + 1
    n_prod = len(products)
    prod_last = n_prod + 1
    categories = sorted({p["category"] for p in products})

    rev_by_cust = {}
    for r in sales:
        rev_by_cust[r["customer"]] = rev_by_cust.get(r["customer"], 0) \
            + int(r["qty"]) * int(r["unit_price"])
    cust_order = sorted(customers, key=lambda c: -rev_by_cust.get(c["customer"], 0))

    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    widths(ws, {"A": 28, "B": 110})
    ws["A1"] = "SalesInv — Sales & Inventory Dashboard"
    ws["A1"].font = TITLE_FONT
    guide = [
        ("What this is", "A formula-driven sales & inventory dashboard "
         "plus a SQLite reporting database. Every figure on Monthly / "
         "Category / Customers / Inventory / Dashboard is a live Excel "
         "formula over the Sales and Purchases sheets — change an order "
         "and the whole pack updates. The .db file carries the same "
         "numbers as reusable SQL views for Power BI or any BI tool."),
        ("The business", "Siuri Traders, a fictional wholesaler in Siuri, "
         f"India. {n_sales} sales orders and {n_pur} restock lines, "
         f"May-Oct 2026, {n_prod} products in {len(categories)} "
         f"categories, {len(customers)} customers. All data is fictional "
         "and deterministic (data/*.csv, made by make_data.py)."),
        ("Sheets", "Purchases — stock-in lines (Total Cost is a formula). "
         "Sales — order lines (Amount/Cost/Margin/Month are formulas; "
         "Product and Region are validated dropdowns). Products — price, "
         "cost, opening stock, reorder level. Inventory — closing stock, "
         "stock value, OUT/LOW/OK status, days of cover, all formulas. "
         "Monthly — revenue, orders, units, margin + prior-month "
         "comparison. Category / Customers — drill-downs. Dashboard — "
         "14 KPIs + 3 charts. Checks — 10 integrity checks, all PASS."),
        ("How to refresh", "1) Paste new purchase rows at the bottom of "
         "Purchases (columns A-E) and copy the Total Cost formula down. "
         "2) Paste new order rows at the bottom of Sales (columns A-H) "
         "and copy the formulas in I, J, K, L down. 3) Extend the range "
         f"ends in Inventory, Monthly, Category, Customers and Checks "
         f"from rows {sales_last}/{pur_last} to the new last rows. "
         "4) Re-check the Checks sheet. 5) Rebuild the .db by running "
         "generate.py (it reloads data/*.csv — append the new rows there "
         "too)."),
        ("Reading the KPIs", "Best Month / Top Category / Top Customer are "
         "INDEX+MATCH over the live rollups. Reorder Alert counts "
         "products at or below their reorder level (LOW) plus stock-outs "
         "(OUT). Days of Cover = closing stock / average daily sales "
         f"over the {PERIOD_DAYS}-day period."),
        ("Built from", "Freelancer project 'Sales Insights & Power BI "
         "Dashboard' (Siuri, INR 750-1,250/hour, 3 proposals, avg bid "
         "INR 784). This workbook + database is a demo reference "
         "implementation, not client work."),
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
    ws.append(["Product", "Category", "Unit Price (INR)",
               "Cost Price (INR)", "Opening Stock", "Reorder Level"])
    style_header(ws, 1, 6)
    for p in products:
        ws.append([p["product"], p["category"], int(p["unit_price"]),
                   int(p["cost_price"]), int(p["opening_stock"]),
                   int(p["reorder_level"])])
    for row in ws.iter_rows(min_row=2, max_row=prod_last,
                            min_col=3, max_col=4):
        for cell in row:
            cell.number_format = INR
    widths(ws, {"A": 20, "B": 16, "C": 16, "D": 16, "E": 15, "F": 14})

    # ---------------- Purchases ----------------
    ws = wb.create_sheet("Purchases")
    ws.append(["Purchase ID", "Date", "Product", "Qty", "Unit Cost (INR)",
               "Total Cost (INR)"])
    style_header(ws, 1, 6)
    for i, p in enumerate(purchases):
        row = i + 2
        d = datetime.date.fromisoformat(p["date"])
        ws.cell(row=row, column=1, value=p["purchase_id"])
        c = ws.cell(row=row, column=2, value=d)
        c.number_format = DATEFMT
        ws.cell(row=row, column=3, value=p["product"])
        ws.cell(row=row, column=4, value=int(p["qty"]))
        u = ws.cell(row=row, column=5, value=int(p["unit_cost"]))
        u.number_format = INR
        t = ws.cell(row=row, column=6, value=f"=D{row}*E{row}")
        t.number_format = INR
    ws.auto_filter.ref = f"A1:F{pur_last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 12, "B": 12, "C": 20, "D": 8, "E": 16, "F": 17})

    # ---------------- Sales ----------------
    ws = wb.create_sheet("Sales")
    headers = ["Order ID", "Date", "Customer", "Region", "Category",
               "Product", "Qty", "Unit Price (INR)", "Amount (INR)",
               "Cost (INR)", "Margin (INR)", "Month"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, o in enumerate(sales):
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
        a = ws.cell(row=row, column=9, value=f"=G{row}*H{row}")
        a.number_format = INR
        co = ws.cell(row=row, column=10,
                    value=f"=G{row}*VLOOKUP(F{row},Products!$A$2:$D${prod_last},4,FALSE)")
        co.number_format = INR
        m = ws.cell(row=row, column=11, value=f"=I{row}-J{row}")
        m.number_format = INR
        mo = ws.cell(row=row, column=12, value=f'=TEXT(B{row},"mmm yyyy")')
    ws.auto_filter.ref = f"A1:L{sales_last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 12, "B": 12, "C": 22, "D": 10, "E": 16, "F": 20,
                "G": 7, "H": 16, "I": 15, "J": 14, "K": 14, "L": 11})
    # Excel validation: Product must come from the Products master,
    # Region from the four-region list. (The brief asked for Excel
    # validation; these dropdowns are it.)
    dv_prod = DataValidation(type="list",
                             formula1=f"Products!$A$2:$A${prod_last}",
                             allow_blank=False)
    dv_prod.error = "Pick a product from the Products sheet."
    dv_prod.errorTitle = "Unknown product"
    dv_prod.prompt = "Choose a product from the master list."
    dv_prod.promptTitle = "Product"
    dv_prod.showDropDown = False
    dv_prod.showErrorMessage = True
    ws.add_data_validation(dv_prod)
    dv_prod.add(f"F2:F{sales_last}")
    dv_region = DataValidation(type="list",
                               formula1='"North,South,East,West"',
                               allow_blank=False)
    dv_region.error = "Region must be North, South, East or West."
    dv_region.errorTitle = "Unknown region"
    dv_region.showErrorMessage = True
    ws.add_data_validation(dv_region)
    dv_region.add(f"D2:D{sales_last}")

    SI = f"Sales!$I$2:$I${sales_last}"   # amount range
    SG = f"Sales!$G$2:$G${sales_last}"   # qty range
    SK = f"Sales!$K$2:$K${sales_last}"   # margin range
    SL = f"Sales!$L$2:$L${sales_last}"   # month helper range
    SE = f"Sales!$E$2:$E${sales_last}"   # category range
    SC = f"Sales!$C$2:$C${sales_last}"   # customer range
    PI = f"Purchases!$D$2:$D${pur_last}"  # purchase qty range
    PC = f"Purchases!$C$2:$C${pur_last}"  # purchase product range
    PF = f"Purchases!$F$2:$F${pur_last}"  # purchase total cost range

    # ---------------- Inventory ----------------
    ws = wb.create_sheet("Inventory")
    ws.append(["Product", "Category", "Opening Stock", "Purchased",
               "Sold", "Closing Stock", "Cost Price (INR)",
               "Stock Value (INR)", "Reorder Level", "Status",
               "Days of Cover"])
    style_header(ws, 1, 11)
    for i, p in enumerate(products):
        row = i + 2
        ws.cell(row=row, column=1, value=p["product"])
        ws.cell(row=row, column=2,
               value=f"=VLOOKUP($A{row},Products!$A$2:$F${prod_last},2,FALSE)")
        ws.cell(row=row, column=3,
               value=f"=VLOOKUP($A{row},Products!$A$2:$F${prod_last},5,FALSE)")
        ws.cell(row=row, column=4,
               value=f"=SUMIFS({PI},{PC},$A{row})")
        ws.cell(row=row, column=5,
               value=f"=SUMIFS({SG},Sales!$F$2:$F${sales_last},$A{row})")
        ws.cell(row=row, column=6, value=f"=C{row}+D{row}-E{row}")
        g = ws.cell(row=row, column=7,
                   value=f"=VLOOKUP($A{row},Products!$A$2:$F${prod_last},4,FALSE)")
        g.number_format = INR
        h = ws.cell(row=row, column=8, value=f"=F{row}*G{row}")
        h.number_format = INR
        ws.cell(row=row, column=9,
               value=f"=VLOOKUP($A{row},Products!$A$2:$F${prod_last},6,FALSE)")
        ws.cell(row=row, column=10,
               value=f'=IF(F{row}<=0,"OUT",IF(F{row}<=I{row},"LOW","OK"))')
        k = ws.cell(row=row, column=11,
                   value=f"=IF(E{row}=0,0,F{row}/(E{row}/{PERIOD_DAYS}))")
        k.number_format = ONE_DEC
    ws.auto_filter.ref = f"A1:K{prod_last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 20, "B": 16, "C": 14, "D": 11, "E": 9, "F": 14,
                "G": 16, "H": 18, "I": 14, "J": 9, "K": 14})
    inv_last = prod_last

    # ---------------- Monthly ----------------
    ws = wb.create_sheet("Monthly")
    ws.append(["Month", "Revenue (INR)", "Orders", "Units",
               "Avg Order (INR)", "Margin (INR)", "Margin %",
               "Prior Month Revenue (INR)", "MoM Change"])
    style_header(ws, 1, 9)
    for i, (label, _d) in enumerate(MONTHS):
        row = i + 2
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2,
               value=f"=SUMIFS({SI},{SL},$A{row})").number_format = INR
        ws.cell(row=row, column=3, value=f"=COUNTIFS({SL},$A{row})")
        ws.cell(row=row, column=4, value=f"=SUMIFS({SG},{SL},$A{row})")
        e = ws.cell(row=row, column=5, value=f"=IF(C{row}=0,0,B{row}/C{row})")
        e.number_format = INR
        ws.cell(row=row, column=6,
               value=f"=SUMIFS({SK},{SL},$A{row})").number_format = INR
        g = ws.cell(row=row, column=7, value=f"=IF(B{row}=0,0,F{row}/B{row})")
        g.number_format = PCT1
        if i == 0:
            ws.cell(row=row, column=8, value=0).number_format = INR
            ws.cell(row=row, column=9, value=0).number_format = PCT1
        else:
            ws.cell(row=row, column=8, value=f"=B{row - 1}").number_format = INR
            ch = ws.cell(row=row, column=9,
                        value=f"=IF(H{row}=0,0,(B{row}-H{row})/H{row})")
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
    m_last = t - 1

    # ---------------- Category ----------------
    ws = wb.create_sheet("Category")
    ws.append(["Category", "Revenue (INR)", "Orders", "Units",
               "Margin (INR)", "Share of Revenue", "Avg Order (INR)"])
    style_header(ws, 1, 7)
    for i, cat in enumerate(categories):
        row = i + 2
        ws.cell(row=row, column=1, value=cat)
        ws.cell(row=row, column=2,
               value=f"=SUMIFS({SI},{SE},$A{row})").number_format = INR
        ws.cell(row=row, column=3, value=f"=COUNTIFS({SE},$A{row})")
        ws.cell(row=row, column=4, value=f"=SUMIFS({SG},{SE},$A{row})")
        ws.cell(row=row, column=5,
               value=f"=SUMIFS({SK},{SE},$A{row})").number_format = INR
        f = ws.cell(row=row, column=6,
                   value=f"=IF(Dashboard!$B$4=0,0,B{row}/Dashboard!$B$4)")
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
        ws.cell(row=row, column=3, value=f"=COUNTIFS({SC},$A{row})")
        ws.cell(row=row, column=4,
               value=f"=SUMIFS({SI},{SC},$A{row})").number_format = INR
        e = ws.cell(row=row, column=5,
                   value=f"=IF(Dashboard!$B$4=0,0,D{row}/Dashboard!$B$4)")
        e.number_format = PCT1
        f = ws.cell(row=row, column=6, value=f"=IF(C{row}=0,0,D{row}/C{row})")
        f.number_format = INR
    widths(ws, {"A": 24, "B": 10, "C": 9, "D": 17, "E": 17, "F": 16})
    cust_last = len(cust_order) + 1

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "SalesInv — Sales & Inventory Dashboard"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Siuri Traders  •  May-Oct 2026  •  all figures live "
                "formulas over Sales / Purchases / Inventory")
    kpis = [
        ("Total Revenue", f"=SUM(Monthly!B2:B{m_last})", INR),
        ("Total Orders", f"=SUM(Monthly!C2:C{m_last})", None),
        ("Avg Order Value", "=IF(B5=0,0,B4/B5)", INR),
        ("Total Margin", f"=SUM(Monthly!F2:F{m_last})", INR),
        ("Margin %", "=IF(B4=0,0,B7/B4)", PCT1),
        ("Best Month", f"=INDEX(Monthly!A2:A{m_last},MATCH(MAX(Monthly!B2:B{m_last}),Monthly!B2:B{m_last},0))", None),
        ("Best Month Revenue", f"=MAX(Monthly!B2:B{m_last})", INR),
        ("Top Category", f"=INDEX(Category!A2:A{cat_last},MATCH(MAX(Category!B2:B{cat_last}),Category!B2:B{cat_last},0))", None),
        ("Top Customer", f"=INDEX(Customers!A2:A{cust_last},MATCH(MAX(Customers!D2:D{cust_last}),Customers!D2:D{cust_last},0))", None),
        ("Repeat-Customer %", f"=COUNTIF(Customers!C2:C{cust_last},\">1\")/COUNTA(Customers!A2:A{cust_last})", PCT1),
        ("Closing Stock Value", f"=SUM(Inventory!H2:H{inv_last})", INR),
        ("Products on Reorder Alert", f"=COUNTIF(Inventory!J2:J{inv_last},\"LOW\")+COUNTIF(Inventory!J2:J{inv_last},\"OUT\")", None),
        ("Stock-out Products", f"=COUNTIF(Inventory!J2:J{inv_last},\"OUT\")", None),
        ("Avg Days of Cover", f"=AVERAGE(Inventory!K2:K{inv_last})", ONE_DEC),
    ]
    ws["A3"] = "KPI"
    ws["B3"] = "Value"
    style_header(ws, 3, 2)
    for i, (label, formula, fmt) in enumerate(kpis):
        row = 4 + i
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=row, column=2, value=formula)
        if fmt:
            c.number_format = fmt
    widths(ws, {"A": 26, "B": 34, "C": 4, "D": 30, "E": 30})

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
        ("Sales Amount total equals Monthly revenue total",
         f"=SUM({SI})", f"=Monthly!B{t}"),
        ("Sales count equals Monthly orders total",
         f"=COUNTA(Sales!A2:A{sales_last})", f"=Monthly!C{t}"),
        ("Sales Margin total equals Monthly margin total",
         f"=SUM({SK})", f"=Monthly!F{t}"),
        ("Category revenue total equals Total Revenue",
         f"=SUM(Category!B2:B{cat_last})", "=Dashboard!B4"),
        ("Customer revenue total equals Total Revenue",
         f"=SUM(Customers!D2:D{cust_last})", "=Dashboard!B4"),
        ("Stock conservation: opening + purchased - sold equals closing",
         f"=SUM(Products!E2:E{prod_last})+SUM({PI})-SUM({SG})",
         f"=SUM(Inventory!F2:F{inv_last})"),
        ("Products with negative closing stock (must be 0)", 0,
         f"=SUMPRODUCT(--(Inventory!F2:F{inv_last}<0))"),
        ("Sales rows with Amount <= 0 (must be 0)", 0,
         f"=SUMPRODUCT(--({SI}<=0))"),
        ("Sales dated outside May-Oct 2026 (must be 0)", 0,
         f'=COUNTIFS(Sales!$B$2:$B${sales_last},"<"&DATE(2026,5,1))'
         f'+COUNTIFS(Sales!$B$2:$B${sales_last},">"&DATE(2026,10,31))'),
        ("Blank Product or Customer cells in Sales (must be 0)", 0,
         f"=COUNTBLANK(Sales!F2:F{sales_last})"
         f"+COUNTBLANK(Sales!C2:C{sales_last})"),
    ]
    for i, (label, expected, actual) in enumerate(checks):
        row = i + 2
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=2, value=label)
        e = ws.cell(row=row, column=3, value=expected)
        a = ws.cell(row=row, column=4, value=actual)
        if isinstance(expected, (int, float)):
            e.number_format = "0"
        else:
            e.number_format = INR
        a.number_format = INR
        ws.cell(row=row, column=5, value=f'=IF(C{row}=D{row},"PASS","FAIL")')
    widths(ws, {"A": 5, "B": 58, "C": 18, "D": 18, "E": 10})

    wb.save(OUT_XLSX)
    print(f"built {OUT_XLSX.name} + {OUT_DB.name} — {n_sales} sales, "
          f"{n_pur} purchases, {n_prod} products, {len(customers)} customers")


if __name__ == "__main__":
    build()
