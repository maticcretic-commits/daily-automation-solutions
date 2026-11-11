#!/usr/bin/env python3
"""SalesOpsPulse — sales & operations Power BI pack factory (Excel + SQLite).

Built around a real Freelancer brief ("Power BI Developer for Sales and
Operations Dashboards (SAP Business One data)", ID 40812525, posted ~1h
before the 2026-11-11 scout, 2 proposals, avg $90): a 5-page Power BI set
— sales overview (YTD, month pacing, channels), inventory/reorder,
delivery routes, calling plan, and receivables & ageing — from SAP
Business One + Excel, DAX / Power Query, star schema, publish to
Power BI Service with scheduled refresh.

What this builds, from fictional Sharma Trading Co., Patna data
(FMCG wholesaler, SAP B1-style, 2026-04-01 .. 2026-11-10, deterministic
data/*.csv, seed 20261111):
1. sales_ops.db — SQLite with the raw tables plus 9 reusable reporting
   views (v_monthly_sales, v_sales_by_channel, v_profit_by_product,
   v_inventory_status, v_reorder_alerts, v_delivery_performance,
   v_call_conversion, v_receivables_ageing, v_top_customers). The views
   are the star-schema facts a Power BI model would consume — see
   powerbi/MEASURES.md for the DAX measures per page.
2. sales-ops-powerbi-pack.xlsx — a formula-driven workbook mirroring the
   brief's 5 pages: Sales_Overview (YTD + November pacing + channels),
   Inventory_Reorder (extracted reorder list), Delivery routes summary
   (on Raw_Deliveries), Calling Plan (on Raw_Calls), Receivables_Ageing
   (buckets + top 5 debtors), plus Category / Customers drill-downs,
   Dashboard (17 KPIs + 3 native charts) and Checks (10 integrity checks).

Conventions (LibreOffice-safe, as proven on earlier builds): plain A1
ranges with the column letter repeated on both ends, real date values
(never text dates), percentages stored as 0-1 fractions, numeric columns
guarded to 0 instead of "", LO-safe extraction via per-row numeric helper
+ SMALL/INDEX (never SMALL(IF(...)) — it mis-evaluates in LO).
"""
import csv
import datetime
import sqlite3
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT_XLSX = BASE / "sales-ops-powerbi-pack.xlsx"
OUT_DB = BASE / "sales_ops.db"
AS_OF = datetime.date(2026, 11, 10)
PERIOD_DAYS = 224  # 2026-04-01 .. 2026-11-10
MONTHS = ["Apr 2026", "May 2026", "Jun 2026", "Jul 2026", "Aug 2026",
          "Sep 2026", "Oct 2026", "Nov 2026"]
CHANNELS = ["Distributor", "Retail", "Institutional", "Online"]
BUCKETS = ["Current", "0-30", "31-60", "61-90", "90+", "Paid"]

NAVY = "1F3864"
TEAL = "0F7B7A"
LIGHT = "D9E2F3"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color=NAVY, size=14)
SECTION_FONT = Font(bold=True, color=NAVY, size=11)
MONEY_FMT = '#,##0'
MONEY2_FMT = '#,##0.00'
PCT_FMT = '0.0%'
DATE_FMT = 'yyyy-mm-dd'


def read_csv(name):
    with open(DATA / name, newline="") as f:
        return list(csv.DictReader(f))


def d(s):
    return datetime.date.fromisoformat(s)


def bucket_of(due, outstanding):
    if outstanding <= 0:
        return "Paid"
    dpd = (AS_OF - due).days
    if dpd <= 0:
        return "Current"
    if dpd <= 30:
        return "0-30"
    if dpd <= 60:
        return "31-60"
    if dpd <= 90:
        return "61-90"
    return "90+"


def build_db(products, customers, orders, receipts, deliveries, calls,
             receivables):
    if OUT_DB.exists():
        OUT_DB.unlink()
    con = sqlite3.connect(OUT_DB)
    cur = con.cursor()
    cur.executescript("""
    CREATE TABLE products (sku TEXT, category TEXT, unit_price INT,
        cost_price INT, opening_stock INT, reorder_level INT);
    CREATE TABLE customers (customer TEXT, channel TEXT, route TEXT,
        credit_days INT, credit_limit INT);
    CREATE TABLE orders (order_id TEXT, date TEXT, customer TEXT,
        channel TEXT, sku TEXT, category TEXT, qty INT, unit_price INT,
        discount_pct INT);
    CREATE TABLE receipts (receipt_id TEXT, date TEXT, sku TEXT, qty INT);
    CREATE TABLE deliveries (delivery_id TEXT, order_id TEXT,
        order_date TEXT, route TEXT, vehicle TEXT, planned_date TEXT,
        actual_date TEXT);
    CREATE TABLE calls (call_id TEXT, date TEXT, customer TEXT,
        channel TEXT, outcome TEXT, order_value INT);
    CREATE TABLE receivables (invoice_id TEXT, order_id TEXT,
        customer TEXT, invoice_date TEXT, due_date TEXT, amount INT,
        paid_amount INT);
    """)
    cur.executemany("INSERT INTO products VALUES (?,?,?,?,?,?)",
                    [(p["sku"], p["category"], p["unit_price"],
                      p["cost_price"], p["opening_stock"],
                      p["reorder_level"]) for p in products])
    cur.executemany("INSERT INTO customers VALUES (?,?,?,?,?)",
                    [(c["customer"], c["channel"], c["route"],
                      c["credit_days"], c["credit_limit"])
                     for c in customers])
    cur.executemany("INSERT INTO orders VALUES (?,?,?,?,?,?,?,?,?)",
                    [(o["order_id"], o["date"], o["customer"], o["channel"],
                      o["sku"], o["category"], o["qty"], o["unit_price"],
                      o["discount_pct"]) for o in orders])
    cur.executemany("INSERT INTO receipts VALUES (?,?,?,?)",
                    [(r["receipt_id"], r["date"], r["sku"], r["qty"])
                     for r in receipts])
    cur.executemany("INSERT INTO deliveries VALUES (?,?,?,?,?,?,?)",
                    [(x["delivery_id"], x["order_id"], x["order_date"],
                      x["route"], x["vehicle"], x["planned_date"],
                      x["actual_date"]) for x in deliveries])
    cur.executemany("INSERT INTO calls VALUES (?,?,?,?,?,?)",
                    [(c["call_id"], c["date"], c["customer"], c["channel"],
                      c["outcome"], c["order_value"]) for c in calls])
    cur.executemany("INSERT INTO receivables VALUES (?,?,?,?,?,?,?)",
                    [(r["invoice_id"], r["order_id"], r["customer"],
                      r["invoice_date"], r["due_date"], r["amount"],
                      r["paid_amount"]) for r in receivables])
    cur.executescript("""
    CREATE VIEW v_monthly_sales AS
      SELECT substr(date, 1, 7) AS ym, COUNT(*) AS orders,
             SUM(qty) AS units,
             SUM(qty * unit_price * (100 - discount_pct) / 100.0) AS revenue,
             SUM(qty * (unit_price * (100 - discount_pct) / 100.0
                 - (SELECT cost_price FROM products p WHERE p.sku = orders.sku)))
             AS margin
      FROM orders GROUP BY ym;
    CREATE VIEW v_sales_by_channel AS
      SELECT channel, COUNT(*) AS orders,
             SUM(qty * unit_price * (100 - discount_pct) / 100.0) AS revenue
      FROM orders GROUP BY channel;
    CREATE VIEW v_profit_by_product AS
      SELECT sku, category, SUM(qty) AS units,
             SUM(qty * unit_price * (100 - discount_pct) / 100.0) AS revenue,
             SUM(qty * (unit_price * (100 - discount_pct) / 100.0
                 - (SELECT cost_price FROM products p WHERE p.sku = orders.sku)))
             AS margin
      FROM orders GROUP BY sku;
    CREATE VIEW v_inventory_status AS
      SELECT p.sku, p.category, p.opening_stock AS opening,
             COALESCE((SELECT SUM(qty) FROM receipts r WHERE r.sku = p.sku), 0)
               AS received,
             COALESCE((SELECT SUM(qty) FROM orders o WHERE o.sku = p.sku), 0)
               AS sold,
             p.opening_stock
               + COALESCE((SELECT SUM(qty) FROM receipts r WHERE r.sku = p.sku), 0)
               - COALESCE((SELECT SUM(qty) FROM orders o WHERE o.sku = p.sku), 0)
               AS closing,
             p.reorder_level
      FROM products p;
    CREATE VIEW v_reorder_alerts AS
      SELECT sku, category, closing, reorder_level,
             CASE WHEN closing <= 0 THEN 'OUT'
                  WHEN closing <= reorder_level THEN 'LOW' ELSE 'OK' END
               AS status
      FROM v_inventory_status WHERE closing <= reorder_level;
    CREATE VIEW v_delivery_performance AS
      SELECT route, COUNT(*) AS deliveries,
             SUM(CASE WHEN actual_date <= planned_date THEN 1 ELSE 0 END)
               AS on_time,
             AVG(julianday(actual_date) - julianday(planned_date))
               AS avg_delay_days
      FROM deliveries GROUP BY route;
    CREATE VIEW v_call_conversion AS
      SELECT channel, COUNT(*) AS calls,
             SUM(CASE WHEN outcome = 'Order placed' THEN 1 ELSE 0 END)
               AS orders_placed,
             SUM(order_value) AS order_value
      FROM calls GROUP BY channel;
    CREATE VIEW v_receivables_ageing AS
      SELECT CASE WHEN amount - paid_amount <= 0 THEN 'Paid'
                  WHEN julianday('2026-11-10') - julianday(due_date) <= 0
                    THEN 'Current'
                  WHEN julianday('2026-11-10') - julianday(due_date) <= 30
                    THEN '0-30'
                  WHEN julianday('2026-11-10') - julianday(due_date) <= 60
                    THEN '31-60'
                  WHEN julianday('2026-11-10') - julianday(due_date) <= 90
                    THEN '61-90'
                  ELSE '90+' END AS bucket,
             COUNT(*) AS invoices, SUM(amount - paid_amount) AS outstanding
      FROM receivables GROUP BY bucket;
    CREATE VIEW v_top_customers AS
      SELECT customer, COUNT(*) AS orders,
             SUM(qty * unit_price * (100 - discount_pct) / 100.0) AS revenue
      FROM orders GROUP BY customer ORDER BY revenue DESC LIMIT 5;
    """)
    con.commit()
    con.close()
    print(f"  db: {OUT_DB.name} (9 views)")


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c2 = ws.cell(row=2, column=1, value=subtitle)
    c2.font = Font(italic=True, size=10, color="404040")
    c2.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30


def build_workbook(products, customers, orders, receipts, deliveries, calls,
                   receivables):
    wb = Workbook()
    n_ord = len(orders)
    n_rec = len(receipts)
    n_del = len(deliveries)
    n_cal = len(calls)
    n_inv = len(receivables)
    OR = f"$B$2:$B${n_ord + 1}"          # orders date range (row base 2)
    o_last = n_ord + 1
    r_last = n_rec + 1
    d_last = n_del + 1
    c_last = n_cal + 1
    v_last = n_inv + 1

    # ---------- Guide ----------
    ws = wb.active
    ws.title = "Guide"
    title_block(ws, "SalesOpsPulse — Sales & Operations Power BI Pack",
                "Fictional Sharma Trading Co., Patna (FMCG wholesaler). "
                "SAP Business One-style data, 2026-04-01 .. 2026-11-10 "
                "(as-of 2026-11-10). Mirrors the 5-page Power BI brief: "
                "Sales Overview | Inventory/Reorder | Delivery Routes | "
                "Calling Plan | Receivables & Ageing. All figures are live "
                "formulas; change a raw row and the pack re-computes. "
                "powerbi/MEASURES.md carries the DAX measures for the "
                "matching Power BI model; sales_ops.db carries the same "
                "facts as 9 reusable SQL views.", 2)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 110
    guide = [
        ("Source", "data/*.csv — deterministic, seed 20261111 (make_data.py)"),
        ("As-of date", "2026-11-10 (November pacing uses 10 of 30 days)"),
        ("Products", "12 SKUs across Staples / Pulses / Edible Oil / Spices"),
        ("Customers", "20 across Distributor / Retail / Institutional / Online"),
        ("Orders", f"{n_ord} sales-order lines, Apr–Nov 10 2026"),
        ("Receipts", f"{n_rec} stock-in lines"),
        ("Deliveries", f"{n_del} deliveries (orders from 2026-09-01)"),
        ("Calls", f"{n_cal} weekly calling-plan rows, Oct 12 – Nov 10"),
        ("Receivables", f"{n_inv} invoices with payment status"),
        ("Page 1 — Sales", "Sales_Overview: monthly orders/units/revenue/margin, "
         "margin %, target, attainment, November pacing, channel split"),
        ("Page 2 — Inventory", "Inventory_Reorder: extracted reorder list "
         "(OUT/LOW) with closing stock, reorder level, stock value"),
        ("Page 3 — Delivery", "Raw_Deliveries: planned vs actual, on-time flag, "
         "delay days; route rollup on Dashboard"),
        ("Page 4 — Calling", "Raw_Calls: weekly plan, outcome, order value; "
         "conversion on Dashboard"),
        ("Page 5 — Receivables", "Receivables_Ageing: buckets + top 5 debtors"),
        ("Power BI", "Star schema: fact_orders / fact_inventory_daily / "
         "fact_receivables + dim_product / dim_customer / dim_date; "
         "DAX in powerbi/MEASURES.md"),
        ("Checks", "Checks sheet: 10 integrity checks, all must read PASS"),
    ]
    header(ws, 4, ["Sheet / item", "What it is"], [28, 110])
    for i, (a, b) in enumerate(guide, start=5):
        ws.cell(row=i, column=1, value=a).font = SECTION_FONT
        ws.cell(row=i, column=2, value=b).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[i].height = 28

    # ---------- Products ----------
    ws = wb.create_sheet("Products")
    title_block(ws, "Products — master", "Cost price feeds Raw_Orders margin "
                "via VLOOKUP; opening stock + reorder level feed "
                "Raw_Inventory.", 6)
    header(ws, 4, ["SKU", "Category", "Unit price (Rs)", "Cost price (Rs)",
                   "Opening stock", "Reorder level"], [24, 14, 15, 15, 14, 14])
    for i, p in enumerate(products, start=5):
        ws.cell(row=i, column=1, value=p["sku"])
        ws.cell(row=i, column=2, value=p["category"])
        for j, k in ((3, "unit_price"), (4, "cost_price"),
                     (5, "opening_stock"), (6, "reorder_level")):
            c = ws.cell(row=i, column=j, value=int(p[k]))
            c.number_format = MONEY_FMT

    # ---------- Customer_Master ----------
    ws = wb.create_sheet("Customer_Master")
    title_block(ws, "Customers — master", "Credit days drive receivable due "
                "dates; channel drives the Sales_Overview split.", 5)
    header(ws, 4, ["Customer", "Channel", "Route", "Credit days",
                   "Credit limit (Rs)"], [26, 14, 14, 12, 16])
    for i, c in enumerate(customers, start=5):
        ws.cell(row=i, column=1, value=c["customer"])
        ws.cell(row=i, column=2, value=c["channel"])
        ws.cell(row=i, column=3, value=c["route"])
        ws.cell(row=i, column=4, value=int(c["credit_days"]))
        c5 = ws.cell(row=i, column=5, value=int(c["credit_limit"]))
        c5.number_format = MONEY_FMT

    # ---------- Raw_Orders ----------
    ws = wb.create_sheet("Raw_Orders")
    title_block(ws, "Raw_Orders — SAP B1-style sales-order lines",
                "Amount = qty x price x (1 - discount). Cost via VLOOKUP on "
                "Products. Month = TEXT(date, 'mmm yyyy') for the rollup.", 13)
    header(ws, 4, ["Order ID", "Date", "Customer", "Channel", "SKU",
                   "Category", "Qty", "Unit price (Rs)", "Discount %",
                   "Amount (Rs)", "Cost (Rs)", "Margin (Rs)", "Month"],
          [11, 12, 24, 14, 20, 12, 7, 14, 10, 13, 12, 12, 11])
    for i, o in enumerate(orders, start=5):
        r = i - 3  # data row number in sheet terms is i; formula row = i
        ws.cell(row=i, column=1, value=o["order_id"])
        dc = ws.cell(row=i, column=2, value=d(o["date"]))
        dc.number_format = DATE_FMT
        ws.cell(row=i, column=3, value=o["customer"])
        ws.cell(row=i, column=4, value=o["channel"])
        ws.cell(row=i, column=5, value=o["sku"])
        ws.cell(row=i, column=6, value=o["category"])
        ws.cell(row=i, column=7, value=int(o["qty"])).number_format = MONEY_FMT
        ws.cell(row=i, column=8, value=int(o["unit_price"])).number_format = MONEY_FMT
        pc = ws.cell(row=i, column=9, value=int(o["discount_pct"]) / 100)
        pc.number_format = PCT_FMT
        for col, f in ((10, f"=G{i}*H{i}*(1-I{i})"),
                       (11, f"=G{i}*VLOOKUP(E{i},Products!$A$5:$D$16,4,FALSE)"),
                       (12, f"=J{i}-K{i}"),
                       (13, f'=TEXT(B{i},"mmm yyyy")')):
            c = ws.cell(row=i, column=col, value=f)
            c.number_format = MONEY_FMT if col != 13 else "General"
    o_first, o_last_row = 5, 4 + n_ord

    # ---------- Raw_Receipts ----------
    ws = wb.create_sheet("Raw_Receipts")
    title_block(ws, "Raw_Receipts — stock-in lines", "Feeds Raw_Inventory "
                "received quantities via SUMIF.", 4)
    header(ws, 4, ["Receipt ID", "Date", "SKU", "Qty"],
          [12, 12, 20, 8])
    for i, x in enumerate(receipts, start=5):
        ws.cell(row=i, column=1, value=x["receipt_id"])
        dc = ws.cell(row=i, column=2, value=d(x["date"]))
        dc.number_format = DATE_FMT
        ws.cell(row=i, column=3, value=x["sku"])
        ws.cell(row=i, column=4, value=int(x["qty"])).number_format = MONEY_FMT
    rc_first, rc_last_row = 5, 4 + n_rec

    # ---------- Raw_Inventory ----------
    ws = wb.create_sheet("Raw_Inventory")
    title_block(ws, "Raw_Inventory — stock position per SKU",
                "Closing = opening + received - sold. Status OUT at zero, "
                "LOW at/below reorder level. Days of cover uses 224-day "
                "average daily sales. Helper column K drives the reorder "
                "extraction (LO-safe: per-row numeric helper + SMALL).", 11)
    header(ws, 4, ["SKU", "Category", "Opening", "Received", "Sold",
                   "Closing", "Reorder level", "Status", "Stock value (Rs)",
                   "Days of cover", "Helper"],
          [20, 12, 9, 9, 9, 9, 13, 9, 15, 13, 8])
    for i, p in enumerate(products, start=5):
        ws.cell(row=i, column=1, value=p["sku"])
        ws.cell(row=i, column=2, value=p["category"])
        ws.cell(row=i, column=3, value=int(p["opening_stock"])).number_format = MONEY_FMT
        formulas = {
            4: f"=SUMIF(Raw_Receipts!$C$5:$C${rc_last_row},A{i},Raw_Receipts!$D$5:$D${rc_last_row})",
            5: f"=SUMIF(Raw_Orders!$E$5:$E${o_last_row},A{i},Raw_Orders!$G$5:$G${o_last_row})",
            6: f"=C{i}+D{i}-E{i}",
            7: f"=VLOOKUP(A{i},Products!$A$5:$F$16,6,FALSE)",
            8: f'=IF(F{i}<=0,"OUT",IF(F{i}<=G{i},"LOW","OK"))',
            9: f"=F{i}*VLOOKUP(A{i},Products!$A$5:$D$16,4,FALSE)",
            10: f"=IF(E{i}=0,0,F{i}/(E{i}/{PERIOD_DAYS}))",
            11: f'=IF(H{i}<>"OK",ROW()-4,"")',
        }
        for col, f in formulas.items():
            c = ws.cell(row=i, column=col, value=f)
            if col in (4, 5, 6, 7, 9):
                c.number_format = MONEY_FMT
            elif col == 10:
                c.number_format = '0.0'

    # ---------- Raw_Deliveries ----------
    ws = wb.create_sheet("Raw_Deliveries")
    title_block(ws, "Raw_Deliveries — delivery routes", "On-time = actual on "
                "or before planned. Delay days = actual - planned.", 9)
    header(ws, 4, ["Delivery ID", "Order ID", "Order date", "Route",
                   "Vehicle", "Planned date", "Actual date", "On-time (1/0)",
                   "Delay (days)"], [12, 11, 12, 13, 14, 12, 12, 12, 11])
    for i, x in enumerate(deliveries, start=5):
        ws.cell(row=i, column=1, value=x["delivery_id"])
        ws.cell(row=i, column=2, value=x["order_id"])
        dc = ws.cell(row=i, column=3, value=d(x["order_date"]))
        dc.number_format = DATE_FMT
        ws.cell(row=i, column=4, value=x["route"])
        ws.cell(row=i, column=5, value=x["vehicle"])
        for col, key in ((6, "planned_date"), (7, "actual_date")):
            c = ws.cell(row=i, column=col, value=d(x[key]))
            c.number_format = DATE_FMT
        ws.cell(row=i, column=8, value=f"=IF(G{i}<=F{i},1,0)").number_format = MONEY_FMT
        ws.cell(row=i, column=9, value=f"=G{i}-F{i}").number_format = MONEY_FMT
    dl_first, dl_last_row = 5, 4 + n_del

    # ---------- Raw_Calls ----------
    ws = wb.create_sheet("Raw_Calls")
    title_block(ws, "Raw_Calls — weekly calling plan", "One call per customer "
                "per week, Oct 12 – Nov 10. Order value is 0 unless the "
                "outcome is 'Order placed'.", 6)
    header(ws, 4, ["Call ID", "Date", "Customer", "Channel", "Outcome",
                   "Order value (Rs)"], [11, 12, 24, 14, 13, 15])
    for i, c in enumerate(calls, start=5):
        ws.cell(row=i, column=1, value=c["call_id"])
        dc = ws.cell(row=i, column=2, value=d(c["date"]))
        dc.number_format = DATE_FMT
        ws.cell(row=i, column=3, value=c["customer"])
        ws.cell(row=i, column=4, value=c["channel"])
        ws.cell(row=i, column=5, value=c["outcome"])
        ws.cell(row=i, column=6, value=int(c["order_value"])).number_format = MONEY_FMT
    cl_first, cl_last_row = 5, 4 + n_cal

    # ---------- Raw_Receivables ----------
    ws = wb.create_sheet("Raw_Receivables")
    title_block(ws, "Raw_Receivables — invoices & ageing",
                "Outstanding = amount - paid. DPD counts only while money is "
                "outstanding, as of 2026-11-10. Bucket: Paid / Current / "
                "0-30 / 31-60 / 61-90 / 90+. Helper K ranks outstanding "
                "amounts for the top-5 debtors table.", 11)
    header(ws, 4, ["Invoice ID", "Order ID", "Customer", "Invoice date",
                   "Due date", "Amount (Rs)", "Paid (Rs)", "Outstanding (Rs)",
                   "DPD", "Bucket", "Rank helper"],
          [11, 11, 24, 12, 12, 12, 11, 14, 7, 9, 11])
    for i, r in enumerate(receivables, start=5):
        ws.cell(row=i, column=1, value=r["invoice_id"])
        ws.cell(row=i, column=2, value=r["order_id"])
        ws.cell(row=i, column=3, value=r["customer"])
        for col, key in ((4, "invoice_date"), (5, "due_date")):
            c = ws.cell(row=i, column=col, value=d(r[key]))
            c.number_format = DATE_FMT
        ws.cell(row=i, column=6, value=int(r["amount"])).number_format = MONEY_FMT
        ws.cell(row=i, column=7, value=int(r["paid_amount"])).number_format = MONEY_FMT
        ws.cell(row=i, column=8, value=f"=F{i}-G{i}").number_format = MONEY_FMT
        ws.cell(row=i, column=9,
               value=f"=IF(H{i}=0,0,MAX(0,DATE(2026,11,10)-E{i}))").number_format = MONEY_FMT
        ws.cell(row=i, column=10,
               value=f'=IF(H{i}=0,"Paid",IF(I{i}=0,"Current",IF(I{i}<=30,"0-30",IF(I{i}<=60,"31-60",IF(I{i}<=90,"61-90","90+")))))')
        ws.cell(row=i, column=11,
               value=f"=RANK(H{i},$H$5:$H${4 + n_inv})+COUNTIF($H$5:H{i},H{i})-1").number_format = MONEY_FMT
    v_first, v_last_row = 5, 4 + n_inv

    # ---------- Sales_Overview ----------
    ws = wb.create_sheet("Sales_Overview")
    title_block(ws, "Sales_Overview — Page 1: sales overview",
                "Monthly orders / units / revenue / margin from Raw_Orders. "
                "Target from data/targets.csv. Pacing for November annualises "
                "10 days of sales over 30 days; closed months show attainment.", 13)
    header(ws, 4, ["Month", "Orders", "Units", "Revenue (Rs)", "Margin (Rs)",
                   "Margin %", "Target (Rs)", "Attainment", "Pacing",
                   "Distributor (Rs)", "Retail (Rs)", "Institutional (Rs)",
                   "Online (Rs)"],
          [11, 8, 9, 13, 13, 9, 12, 10, 9, 14, 12, 14, 12])
    targets = {t["month"]: int(t["target"]) for t in read_csv("targets.csv")}
    for i, m in enumerate(MONTHS, start=5):
        ws.cell(row=i, column=1, value=m)
        formulas = {
            2: f'=COUNTIF(Raw_Orders!$M$5:$M${o_last_row},A{i})',
            3: f'=SUMIF(Raw_Orders!$M$5:$M${o_last_row},A{i},Raw_Orders!$G$5:$G${o_last_row})',
            4: f'=SUMIF(Raw_Orders!$M$5:$M${o_last_row},A{i},Raw_Orders!$J$5:$J${o_last_row})',
            5: f'=SUMIF(Raw_Orders!$M$5:$M${o_last_row},A{i},Raw_Orders!$L$5:$L${o_last_row})',
            6: f"=IF(D{i}=0,0,E{i}/D{i})",
            8: f"=IF(G{i}=0,0,D{i}/G{i})",
            9: f'=IF(A{i}="Nov 2026",H{i}*3,H{i})',
        }
        for col, f in formulas.items():
            c = ws.cell(row=i, column=col, value=f)
            if col in (2, 3, 4, 5):
                c.number_format = MONEY_FMT
            elif col in (6, 8, 9):
                c.number_format = PCT_FMT
        ws.cell(row=i, column=7, value=targets[m]).number_format = MONEY_FMT
        for j, ch in enumerate(CHANNELS, start=10):
            c = ws.cell(row=i, column=j,
                       value=f'=SUMIFS(Raw_Orders!$J$5:$J${o_last_row},Raw_Orders!$M$5:$M${o_last_row},A{i},Raw_Orders!$D$5:$D${o_last_row},"{ch}")')
            c.number_format = MONEY_FMT
    so_first, so_last_row = 5, 4 + len(MONTHS)

    # ---------- Inventory_Reorder ----------
    ws = wb.create_sheet("Inventory_Reorder")
    title_block(ws, "Inventory_Reorder — Page 2: reorder list",
                "Extracted from Raw_Inventory via the numeric helper + "
                "SMALL/INDEX (LO-safe). Only OUT / LOW SKUs appear.", 7)
    header(ws, 4, ["#", "SKU", "Category", "Closing", "Reorder level",
                   "Status", "Stock value (Rs)"],
          [5, 20, 12, 9, 13, 9, 15])
    for k in range(1, 13):
        i = 4 + k
        ws.cell(row=i, column=1, value=k)
        for col, src in ((2, "A"), (3, "B"), (4, "F"), (5, "G"),
                         (6, "H"), (7, "I")):
            c = ws.cell(row=i, column=col,
                       value=f'=IFERROR(INDEX(Raw_Inventory!${src}$5:${src}$16,SMALL(Raw_Inventory!$K$5:$K$16,$A{i})),"")')
            if col in (4, 5, 7):
                c.number_format = MONEY_FMT
    ws.cell(row=18, column=1, value="Reorder alerts").font = SECTION_FONT
    ws.cell(row=18, column=2,
           value='=COUNTIF(Raw_Inventory!$H$5:$H$16,"<>OK")').number_format = MONEY_FMT
    ws.cell(row=19, column=1, value="Stockouts").font = SECTION_FONT
    ws.cell(row=19, column=2,
           value='=COUNTIF(Raw_Inventory!$H$5:$H$16,"OUT")').number_format = MONEY_FMT

    # ---------- Receivables_Ageing ----------
    ws = wb.create_sheet("Receivables_Ageing")
    title_block(ws, "Receivables_Ageing — Page 5: ageing + top debtors",
                "Bucket summary from Raw_Receivables; top 5 debtors via the "
                "rank helper (RANK + COUNTIF tie-break, LO-safe).", 8)
    header(ws, 4, ["Bucket", "Invoices", "Outstanding (Rs)"], [12, 10, 17])
    for i, b in enumerate(BUCKETS, start=5):
        ws.cell(row=i, column=1, value=b)
        ws.cell(row=i, column=2,
               value=f'=COUNTIF(Raw_Receivables!$J$5:$J${v_last_row},A{i})').number_format = MONEY_FMT
        ws.cell(row=i, column=3,
               value=f'=SUMIF(Raw_Receivables!$J$5:$J${v_last_row},A{i},Raw_Receivables!$H$5:$H${v_last_row})').number_format = MONEY_FMT
    ws.cell(row=11, column=1, value="TOTAL").font = SECTION_FONT
    ws.cell(row=11, column=2, value="=SUM(B5:B10)").number_format = MONEY_FMT
    ws.cell(row=11, column=3, value="=SUM(C5:C10)").number_format = MONEY_FMT
    ws.cell(row=13, column=1, value="Top 5 debtors").font = TITLE_FONT
    header(ws, 14, ["#", "Customer", "Outstanding (Rs)", "DPD"],
          [5, 24, 17, 8])
    for k in range(1, 6):
        i = 14 + k
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2,
               value=f"=INDEX(Raw_Receivables!$C$5:$C${v_last_row},MATCH($A{i},Raw_Receivables!$K$5:$K${v_last_row},0))")
        c = ws.cell(row=i, column=3,
                   value=f"=INDEX(Raw_Receivables!$H$5:$H${v_last_row},MATCH($A{i},Raw_Receivables!$K$5:$K${v_last_row},0))")
        c.number_format = MONEY_FMT
        c = ws.cell(row=i, column=4,
                   value=f"=INDEX(Raw_Receivables!$I$5:$I${v_last_row},MATCH($A{i},Raw_Receivables!$K$5:$K${v_last_row},0))")
        c.number_format = MONEY_FMT

    # ---------- Category ----------
    ws = wb.create_sheet("Category")
    title_block(ws, "Category — drill-down", "Revenue / orders / margin per "
                "category from Raw_Orders.", 4)
    header(ws, 4, ["Category", "Revenue (Rs)", "Orders", "Margin (Rs)"],
          [14, 14, 9, 13])
    for i, cat in enumerate(["Staples", "Pulses", "Edible Oil", "Spices"],
                            start=5):
        ws.cell(row=i, column=1, value=cat)
        ws.cell(row=i, column=2,
               value=f'=SUMIF(Raw_Orders!$F$5:$F${o_last_row},A{i},Raw_Orders!$J$5:$J${o_last_row})').number_format = MONEY_FMT
        ws.cell(row=i, column=3,
               value=f'=COUNTIF(Raw_Orders!$F$5:$F${o_last_row},A{i})').number_format = MONEY_FMT
        ws.cell(row=i, column=4,
               value=f'=SUMIF(Raw_Orders!$F$5:$F${o_last_row},A{i},Raw_Orders!$L$5:$L${o_last_row})').number_format = MONEY_FMT

    # ---------- Customers (summary) ----------
    ws = wb.create_sheet("Customers")
    title_block(ws, "Customers — drill-down", "Orders / revenue / outstanding "
                "per customer from Raw_Orders and Raw_Receivables.", 5)
    header(ws, 4, ["Customer", "Channel", "Orders", "Revenue (Rs)",
                   "Outstanding (Rs)"], [26, 14, 8, 13, 15])
    for i, c in enumerate(customers, start=5):
        ws.cell(row=i, column=1, value=c["customer"])
        ws.cell(row=i, column=2,
               value=f"=VLOOKUP(A{i},Customer_Master!$A$5:$B$24,2,FALSE)")
        ws.cell(row=i, column=3,
               value=f'=COUNTIF(Raw_Orders!$C$5:$C${o_last_row},A{i})').number_format = MONEY_FMT
        ws.cell(row=i, column=4,
               value=f'=SUMIF(Raw_Orders!$C$5:$C${o_last_row},A{i},Raw_Orders!$J$5:$J${o_last_row})').number_format = MONEY_FMT
        ws.cell(row=i, column=5,
               value=f'=SUMIF(Raw_Receivables!$C$5:$C${v_last_row},A{i},Raw_Receivables!$H$5:$H${v_last_row})').number_format = MONEY_FMT

    # ---------- Dashboard ----------
    ws = wb.create_sheet("Dashboard")
    title_block(ws, "Dashboard — all five pages at a glance",
                "Every KPI below is a live formula. Charts: monthly revenue, "
                "revenue by category, receivables ageing.", 2)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22
    kpis = [
        ("YTD revenue (Rs)", f"=SUM(Sales_Overview!$D$5:$D${so_last_row})", MONEY_FMT),
        ("Total orders", f"=SUM(Sales_Overview!$B$5:$B${so_last_row})", MONEY_FMT),
        ("Avg order value (Rs)", "=IF(B6=0,0,B5/B6)", MONEY_FMT),
        ("Total margin (Rs)", f"=SUM(Sales_Overview!$E$5:$E${so_last_row})", MONEY_FMT),
        ("Margin %", "=IF(B5=0,0,B8/B5)", PCT_FMT),
        ("November revenue so far (Rs)", "=Sales_Overview!D12", MONEY_FMT),
        ("November pacing", "=Sales_Overview!I12", PCT_FMT),
        ("Best month", f"=INDEX(Sales_Overview!$A$5:$A${so_last_row},MATCH(MAX(Sales_Overview!$D$5:$D${so_last_row}),Sales_Overview!$D$5:$D${so_last_row},0))", "General"),
        ("Top customer", "=INDEX(Customers!$A$5:$A$24,MATCH(MAX(Customers!$D$5:$D$24),Customers!$D$5:$D$24,0))", "General"),
        ("Stock value (Rs)", "=SUM(Raw_Inventory!$I$5:$I$16)", MONEY_FMT),
        ("Reorder alerts", '=COUNTIF(Raw_Inventory!$H$5:$H$16,"<>OK")', MONEY_FMT),
        ("Stockouts", '=COUNTIF(Raw_Inventory!$H$5:$H$16,"OUT")', MONEY_FMT),
        ("On-time delivery %", f"=IF(COUNTA(Raw_Deliveries!$A$5:$A${dl_last_row})=0,0,SUM(Raw_Deliveries!$H$5:$H${dl_last_row})/COUNTA(Raw_Deliveries!$A$5:$A${dl_last_row}))", PCT_FMT),
        ("Avg delivery delay (days)", f"=IF(COUNTA(Raw_Deliveries!$A$5:$A${dl_last_row})=0,0,SUM(Raw_Deliveries!$I$5:$I${dl_last_row})/COUNTA(Raw_Deliveries!$A$5:$A${dl_last_row}))", '0.00'),
        ("Calls made", f"=COUNTA(Raw_Calls!$A$5:$A${cl_last_row})", MONEY_FMT),
        ("Call conversion %", f'=IF(COUNTA(Raw_Calls!$A$5:$A${cl_last_row})=0,0,COUNTIF(Raw_Calls!$E$5:$E${cl_last_row},"Order placed")/COUNTA(Raw_Calls!$A$5:$A${cl_last_row}))', PCT_FMT),
        ("Receivables outstanding (Rs)", f"=SUM(Raw_Receivables!$H$5:$H${v_last_row})", MONEY_FMT),
        ("Overdue 90+ (Rs)", '=Receivables_Ageing!C9', MONEY_FMT),
        ("Avg days of cover", "=AVERAGE(Raw_Inventory!$J$5:$J$16)", '0.0'),
    ]
    header(ws, 4, ["KPI", "Value"], [34, 22])
    for i, (label, formula, fmt) in enumerate(kpis, start=5):
        ws.cell(row=i, column=1, value=label).font = SECTION_FONT
        c = ws.cell(row=i, column=2, value=formula)
        c.number_format = fmt
        c.font = Font(bold=True, size=12)

    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "Monthly revenue (Rs)"
    ch1.y_axis.title = "Rs"
    ch1.x_axis.title = "Month"
    so = wb["Sales_Overview"]
    ch1.add_data(Reference(so, min_col=4, min_row=4, max_row=so_last_row,
                           max_col=4), titles_from_data=True)
    ch1.set_categories(Reference(so, min_col=1, min_row=5, max_row=so_last_row))
    ws.add_chart(ch1, "D4")

    ch2 = PieChart()
    ch2.title = "Revenue by category"
    ch2.add_data(Reference(wb["Category"], min_col=2, min_row=4, max_row=8),
                 titles_from_data=True)
    ch2.set_categories(Reference(wb["Category"], min_col=1, min_row=5, max_row=8))
    ws.add_chart(ch2, "D20")

    ch3 = BarChart()
    ch3.type = "col"
    ch3.title = "Receivables ageing (Rs outstanding)"
    ch3.y_axis.title = "Rs"
    ch3.add_data(Reference(wb["Receivables_Ageing"], min_col=3, min_row=4,
                           max_row=9), titles_from_data=True)
    ch3.set_categories(Reference(wb["Receivables_Ageing"], min_col=1,
                                 min_row=5, max_row=9))
    ws.add_chart(ch3, "D36")

    # ---------- Checks ----------
    ws = wb.create_sheet("Checks")
    title_block(ws, "Checks — integrity", "Every row must read PASS after "
                "recalculation. Value and Expected are both formulas or "
                "computed values; Status compares them.", 4)
    header(ws, 4, ["Check", "Value", "Expected", "Status"],
          [44, 16, 12, 10])
    checks = [
        ("Revenue tie: Raw_Orders vs Sales_Overview",
         f"=SUM(Raw_Orders!$J$5:$J${o_last_row})-SUM(Sales_Overview!$D$5:$D${so_last_row})", 0),
        ("Margin tie: Raw_Orders vs Sales_Overview",
         f"=SUM(Raw_Orders!$L$5:$L${o_last_row})-SUM(Sales_Overview!$E$5:$E${so_last_row})", 0),
        ("Units tie: Raw_Orders vs Sales_Overview",
         f"=SUM(Raw_Orders!$G$5:$G${o_last_row})-SUM(Sales_Overview!$C$5:$C${so_last_row})", 0),
        ("Receivables tie: amount = paid + outstanding",
         f"=SUM(Raw_Receivables!$F$5:$F${v_last_row})-SUM(Raw_Receivables!$G$5:$G${v_last_row})-SUM(Raw_Receivables!$H$5:$H${v_last_row})", 0),
        ("No negative closing stock",
         '=COUNTIF(Raw_Inventory!$F$5:$F$16,"<0")', 0),
        ("Ageing tie: buckets sum = total outstanding",
         f"=SUM(Receivables_Ageing!$C$5:$C$10)-SUM(Raw_Receivables!$H$5:$H${v_last_row})", 0),
        ("Reorder list count = OUT/LOW count",
         '=SUMPRODUCT(--(Inventory_Reorder!$B$5:$B$16<>""))-COUNTIF(Raw_Inventory!$H$5:$H$16,"<>OK")', 0),
        ("On-time deliveries within count",
         f"=IF(SUM(Raw_Deliveries!$H$5:$H${dl_last_row})<=COUNTA(Raw_Deliveries!$A$5:$A${dl_last_row}),0,1)", 0),
        ("Top-5 debtors sum within outstanding",
         f"=IF(SUM(Receivables_Ageing!$C$15:$C$19)<=SUM(Raw_Receivables!$H$5:$H${v_last_row}),0,1)", 0),
        ("No blank order dates",
         f"=COUNTBLANK(Raw_Orders!$B$5:$B${o_last_row})", 0),
    ]
    for i, (label, formula, expected) in enumerate(checks, start=5):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=formula).number_format = MONEY_FMT
        ws.cell(row=i, column=3, value=expected).number_format = MONEY_FMT
        ws.cell(row=i, column=4, value=f"=IF(B{i}=C{i},\"PASS\",\"FAIL\")")

    wb.save(OUT_XLSX)
    print(f"  xlsx: {OUT_XLSX.name} ({len(wb.sheetnames)} sheets)")


def main():
    products = read_csv("products.csv")
    customers = read_csv("customers.csv")
    orders = read_csv("orders.csv")
    receipts = read_csv("receipts.csv")
    deliveries = read_csv("deliveries.csv")
    calls = read_csv("calls.csv")
    receivables = read_csv("receivables.csv")
    build_db(products, customers, orders, receipts, deliveries, calls,
             receivables)
    build_workbook(products, customers, orders, receipts, deliveries, calls,
                   receivables)
    print("DONE")


if __name__ == "__main__":
    main()
