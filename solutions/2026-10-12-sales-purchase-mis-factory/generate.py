#!/usr/bin/env python3
"""
SalesPulse MIS Factory — builds a formula-driven Excel MIS dashboard workbook
from multiple fictional daily sales/purchase source files.

Simulates the client scenario from Freelancer project "Excel & MIS Reporting
Specialist" (ID 40762180): several Excel workbooks tracking daily sales and
purchase activity -> one consolidated, refreshable MIS workbook with trend
analysis, seasonality, anomaly flags and a management dashboard.

Every derived figure in the output workbook is a LIVE Excel formula pointing
at the raw data ranges — nothing is hard-coded. Paste new rows into the Sales /
Purchases tables and every KPI, chart and check refreshes. Formulas use plain
A1-style references (no structured table refs) so they calculate identically
in Excel and LibreOffice.

Run:  python3 generate.py            (writes data/*.csv + sales-purchase-mis.xlsx)
Test: python3 -m pytest tests/ -q
"""
import csv
import datetime as dt
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT_XLSX = BASE / "sales-purchase-mis.xlsx"
SEED = 20261012

COMPANY = "Verma Trading Co., Patna (fictional demo data)"
PERIOD_START = dt.date(2026, 4, 1)
PERIOD_END = dt.date(2026, 9, 30)

# Formula ranges carry headroom for pasted rows (documented in Guide sheet).
# NOTE: ranges must repeat the column letter on BOTH ends
# ("Sales!$H$2:$H$5000", never "Sales!$H$2:$5000" — the latter is not a valid
# reference and makes LibreOffice/Excel raise #NAME?).
S_AMT = "Sales!$H$2:$H$5000"
S_QTY = "Sales!$F$2:$F$5000"
S_RATE = "Sales!$G$2:$G$5000"
S_DATE = "Sales!$A$2:$A$5000"
S_PROD = "Sales!$D$2:$D$5000"
S_CAT = "Sales!$E$2:$E$5000"
S_COGS = "Sales!$I$2:$I$5000"
S_INV = "Sales!$B$2:$B$5000"
P_AMT = "Purchases!$H$2:$H$2000"
P_PROD = "Purchases!$D$2:$D$2000"
P_COST = "Purchases!$G$2:$G$2000"
P_DATE = "Purchases!$A$2:$A$2000"
P_QTY = "Purchases!$F$2:$F$2000"
PR_PROD = "Products!$B$2:$B$13"
PR_COST = "Products!$D$2:$D$13"
PR_SALES = "Products!$E$2:$E$13"
DAILY_FLAG = "Daily!$G$2:$G$184"

# ---------------------------------------------------------------- data model
PRODUCTS = [  # (code, name, category, base_sale_rate, base_cost_rate)
    ("P001", "Mango Juice 1L", "Beverages", 120, 96),
    ("P002", "Cola 750ml", "Beverages", 45, 36),
    ("P003", "Packaged Water 1L", "Beverages", 20, 15),
    ("P004", "Potato Chips 150g", "Snacks", 50, 38),
    ("P005", "Biscuits 400g", "Snacks", 40, 30),
    ("P006", "Instant Noodles 280g", "Snacks", 56, 42),
    ("P007", "Toned Milk 1L", "Dairy", 66, 54),
    ("P008", "Paneer 200g", "Dairy", 95, 78),
    ("P009", "Curd Cup 400g", "Dairy", 45, 35),
    ("P010", "Detergent 1kg", "Household", 130, 104),
    ("P011", "Dish Soap 500ml", "Household", 99, 79),
    ("P012", "Floor Cleaner 1L", "Household", 110, 88),
]
CATEGORIES = sorted({p[2] for p in PRODUCTS})

CUSTOMERS = [
    "Gupta General Store", "Sharma Provision", "Khan Mart", "Patel Traders",
    "Singh Super Bazaar", "Yadav Daily Needs", "Mishra Stores", "Verma Retail",
    "Das & Sons", "Kumari Kirana", "Raza Brothers", "Nath Provisions",
    "Choudhary Mart", "Tiwari General", "Azad Stores", "Prasad Traders",
    "Sinha Super Mart", "Dubey Kirana", "Ali General Store", "Jha Provisions",
    "Rai Daily Mart", "Pandey Stores", "Sahu Traders", "Mehta Retail",
    "Bhatt General Store",
]
SUPPLIERS = [
    "Bihar Foods Distributors", "Ganga FMCG Supply", "Patliputra Agencies",
    "Magadh Wholesale Co.", "Sone Valley Traders", "Koshi Enterprises",
    "Nalanda Distribution", "Vaishali Supply House",
]

MONTH_FACTOR = {4: 1.00, 5: 1.15, 6: 1.25, 7: 1.05, 8: 0.95, 9: 1.10}
ANOMALY_BULK_DAY = dt.date(2026, 5, 18)    # one huge bulk order -> spike
ANOMALY_DEAD_DAY = dt.date(2026, 8, 2)     # near-zero sales day -> dip


def daterange(a, b):
    d = a
    while d <= b:
        yield d
        d += dt.timedelta(days=1)


def gen_data():
    """Deterministic fictional daily sales + purchase activity."""
    rng = random.Random(SEED)
    sales, purchases = [], []
    inv_n, bil_n = 0, 0
    for day in daterange(PERIOD_START, PERIOD_END):
        weekend = day.weekday() >= 5
        for _ in range(rng.randint(5, 11)):
            code, name, cat, rate, cost = rng.choice(PRODUCTS)
            qty = rng.randint(4, 60)
            uplift = MONTH_FACTOR[day.month] * (1.3 if weekend else 1.0)
            if cat == "Beverages" and day.month in (5, 6):
                uplift *= 1.4
            qty = max(1, int(qty * uplift * rng.uniform(0.7, 1.3)))
            rate_j = round(rate * rng.uniform(0.97, 1.05), 2)
            inv_n += 1
            sales.append({
                "date": day, "invoice": f"INV-26{day.month:02d}-{inv_n:04d}",
                "customer": rng.choice(CUSTOMERS), "product": name,
                "category": cat, "qty": qty, "rate": rate_j,
            })
        if rng.random() < 0.75:  # purchase days
            for _ in range(rng.randint(1, 4)):
                code, name, cat, rate, cost = rng.choice(PRODUCTS)
                qty = rng.randint(50, 400)
                cost_j = round(cost * rng.uniform(0.98, 1.03), 2)
                bil_n += 1
                purchases.append({
                    "date": day, "bill": f"BIL-26{day.month:02d}-{bil_n:04d}",
                    "supplier": rng.choice(SUPPLIERS), "product": name,
                    "category": cat, "qty": qty, "cost_rate": cost_j,
                })
    # injected anomalies (kept obvious so the anomaly sheet flags them)
    sales.append({"date": ANOMALY_BULK_DAY, "invoice": f"INV-2605-{inv_n+1:04d}",
                  "customer": "Singh Super Bazaar", "product": "Mango Juice 1L",
                  "category": "Beverages", "qty": 4000, "rate": 118.0})
    sales = [s for s in sales if s["date"] != ANOMALY_DEAD_DAY]
    return sales, purchases


def write_csvs(sales, purchases):
    DATA.mkdir(parents=True, exist_ok=True)
    # split sales into two "source workbooks" to mirror the client's setup
    mid = dt.date(2026, 7, 1)
    parts = {"source_sales_apr_jun.csv": [s for s in sales if s["date"] < mid],
             "source_sales_jul_sep.csv": [s for s in sales if s["date"] >= mid]}
    for name, rows in parts.items():
        with open(DATA / name, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Date", "Invoice", "Customer", "Product", "Category",
                        "Qty", "Rate"])
            for r in rows:
                w.writerow([r["date"].isoformat(), r["invoice"], r["customer"],
                            r["product"], r["category"], r["qty"], r["rate"]])
    with open(DATA / "source_purchases.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Bill", "Supplier", "Product", "Category",
                    "Qty", "CostRate"])
        for r in purchases:
            w.writerow([r["date"].isoformat(), r["bill"], r["supplier"],
                        r["product"], r["category"], r["qty"], r["cost_rate"]])
    return parts


# ---------------------------------------------------------------- styling
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=16, color="1F4E78")
KPI_FONT = Font(bold=True, size=14, color="1F4E78")
KPI_LBL = Font(size=10, color="595959")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INR = '#,##0.00'
INT = '#,##0'
PCT = '0.0%'


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 22


def add_table(ws, name, headers, data_rows, start_row=1, date_cols=()):
    """Write headers + value rows as a styled data range.

    NOTE: deliberately NOT wrapped in an Excel Table object — LibreOffice
    fails to recalculate cross-sheet references that extend beyond a Table's
    extent (#NAME?), while plain ranges with headroom calculate identically
    in Excel and LibreOffice. The Guide sheet documents the headroom.
    """
    for c, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=c, value=h)
    for r, row in enumerate(data_rows, start_row + 1):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            if c - 1 in date_cols:
                cell.number_format = "yyyy-mm-dd"
    style_header(ws, start_row, len(headers))
    return start_row + len(data_rows)


def autofit(ws, min_width=12, max_width=28):
    for col in ws.columns:
        letter = col[0].column_letter
        width = max((len(str(c.value)) if c.value is not None else 0
                     for c in col), default=0)
        ws.column_dimensions[letter].width = min(max_width,
                                                max(min_width, width + 2))


# ---------------------------------------------------------------- workbook
def build_workbook(sales, purchases):
    wb = Workbook()

    ws = wb.active
    ws.title = "Sales"
    sales_rows = [[s["date"], s["invoice"], s["customer"], s["product"],
                   s["category"], s["qty"], s["rate"], None, None]
                  for s in sales]
    n_sales = add_table(
        ws, "tblSales",
        ["Date", "Invoice", "Customer", "Product", "Category", "Qty",
         "Rate", "Amount", "COGS"],
        sales_rows, date_cols=(0,))
    for r in range(2, n_sales + 1):
        ws.cell(row=r, column=6).number_format = INT
        ws.cell(row=r, column=7, value=None)  # placeholder, set below
        ws.cell(row=r, column=7).number_format = INR
        ws.cell(row=r, column=8, value=f"=F{r}*G{r}")
        ws.cell(row=r, column=8).number_format = INR
        ws.cell(row=r, column=8).border = BORDER
        ws.cell(row=r, column=9,
                value=f"=F{r}*VLOOKUP(D{r},Products!$B$2:$E$13,3,FALSE)")
        ws.cell(row=r, column=9).number_format = INR
        ws.cell(row=r, column=9).border = BORDER
    # restore Rate values (were set as None above only for col G formatting)
    for r, s in enumerate(sales, 2):
        ws.cell(row=r, column=7, value=s["rate"])
        ws.cell(row=r, column=7).number_format = INR
        ws.cell(row=r, column=7).border = BORDER
    # pre-fill formula columns into the headroom so pasted rows calculate
    # immediately (blank-product rows stay blank via the IF guard)
    for r in range(n_sales + 1, 2501):
        ws.cell(row=r, column=8, value=f"=F{r}*G{r}")
        ws.cell(row=r, column=8).number_format = INR
        ws.cell(row=r, column=8).border = BORDER
        ws.cell(row=r, column=9,
                value=f'=IF(D{r}="","",F{r}*VLOOKUP(D{r},'
                      f'Products!$B$2:$E$13,3,FALSE))')
        ws.cell(row=r, column=9).number_format = INR
        ws.cell(row=r, column=9).border = BORDER
    autofit(ws)

    ws = wb.create_sheet("Purchases")
    pur_rows = [[p["date"], p["bill"], p["supplier"], p["product"],
                 p["category"], p["qty"], p["cost_rate"], None]
                for p in purchases]
    n_pur = add_table(
        ws, "tblPurchases",
        ["Date", "Bill", "Supplier", "Product", "Category", "Qty",
         "CostRate", "Amount"],
        pur_rows, date_cols=(0,))
    for r, p in enumerate(purchases, 2):
        ws.cell(row=r, column=7, value=p["cost_rate"])
        ws.cell(row=r, column=7).number_format = INR
        ws.cell(row=r, column=7).border = BORDER
        ws.cell(row=r, column=8, value=f"=F{r}*G{r}")
        ws.cell(row=r, column=8).number_format = INR
        ws.cell(row=r, column=8).border = BORDER
        ws.cell(row=r, column=6).number_format = INT
    for r in range(n_pur + 1, 1501):
        ws.cell(row=r, column=8, value=f"=F{r}*G{r}")
        ws.cell(row=r, column=8).number_format = INR
        ws.cell(row=r, column=8).border = BORDER
    autofit(ws)

    ws = wb.create_sheet("Products")
    prod_rows = [[code, name, cat, None, None, None]
                 for code, name, cat, _, _ in PRODUCTS]
    n_prod = add_table(ws, "tblProducts",
                       ["Code", "Product", "Category", "StdCost",
                        "TotalSales", "TotalQty"],
                       prod_rows)
    for r, (_, name, _, _, _) in enumerate(PRODUCTS, 2):
        ws.cell(row=r, column=4,
                value=f"=AVERAGEIF({P_PROD},B{r},{P_COST})")
        ws.cell(row=r, column=4).number_format = INR
        ws.cell(row=r, column=4).border = BORDER
        ws.cell(row=r, column=5,
                value=f"=SUMIF({S_PROD},B{r},{S_AMT})")
        ws.cell(row=r, column=5).number_format = INR
        ws.cell(row=r, column=5).border = BORDER
        ws.cell(row=r, column=6,
                value=f"=SUMIF({S_PROD},B{r},{S_QTY})")
        ws.cell(row=r, column=6).number_format = INT
        ws.cell(row=r, column=6).border = BORDER
    autofit(ws)
    return wb


def monthly_sheet(wb):
    ws = wb.create_sheet("Monthly")
    for c, h in enumerate(["Month", "Sales", "Purchases", "COGS",
                           "GrossProfit", "Margin%", "SeasonalityIdx"], 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, 1, 7)
    months = [dt.date(2026, m, 1) for m in range(4, 10)]
    for i, m in enumerate(months, 2):
        ws.cell(row=i, column=1, value=m).number_format = "mmm-yy"
        ws.cell(row=i, column=2,
                value=f'=SUMIFS({S_AMT},{S_DATE},">="&A{i},{S_DATE},'
                      f'"<"&EDATE(A{i},1))')
        ws.cell(row=i, column=3,
                value=f'=SUMIFS({P_AMT},{P_DATE},">="&A{i},'
                      f'{P_DATE},"<"&EDATE(A{i},1))')
        ws.cell(row=i, column=4,
                value=f'=SUMIFS({S_COGS},{S_DATE},">="&A{i},{S_DATE},'
                      f'"<"&EDATE(A{i},1))')
        ws.cell(row=i, column=5, value=f"=B{i}-D{i}")
        ws.cell(row=i, column=6, value=f"=IF(B{i}=0,0,E{i}/B{i})")
        ws.cell(row=i, column=7, value="=B{i}/AVERAGE($B$2:$B$7)".format(i=i))
        for c, fmt in ((2, INR), (3, INR), (4, INR), (5, INR), (6, PCT),
                       (7, "0.00")):
            ws.cell(row=i, column=c).number_format = fmt
            ws.cell(row=i, column=c).border = BORDER
        ws.cell(row=i, column=1).border = BORDER
    ws.column_dimensions["A"].width = 12
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 16
    return ws


def daily_sheet(wb):
    ws = wb.create_sheet("Daily")
    for c, h in enumerate(["Date", "Sales", "Avg", "Stdev", "Upper(2s)",
                           "Lower(2s)", "Flag"], 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, 1, 7)
    days = list(daterange(PERIOD_START, PERIOD_END))
    n = len(days) + 1
    for i, d in enumerate(days, 2):
        ws.cell(row=i, column=1, value=d).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2,
                value=f"=SUMIF({S_DATE},A{i},{S_AMT})")
        ws.cell(row=i, column=3, value=f"=AVERAGE($B$2:$B${n})")
        ws.cell(row=i, column=4, value=f"=STDEVP($B$2:$B${n})")
        ws.cell(row=i, column=5, value="=C2+2*D2")
        ws.cell(row=i, column=6, value="=MAX(0,C2-2*D2)")
        ws.cell(row=i, column=7,
                value=f'=IF(OR(B{i}>E{i},B{i}<F{i},AND(B{i}=0,C{i}>0)),'
                      f'"ANOMALY","")')
        for c, fmt in ((2, INR), (3, INR), (4, INR), (5, INR), (6, INR)):
            ws.cell(row=i, column=c).number_format = fmt
            ws.cell(row=i, column=c).border = BORDER
        ws.cell(row=i, column=1).border = BORDER
        ws.cell(row=i, column=7).border = BORDER
    for col, w in zip("ABCDEFG", (12, 14, 14, 12, 14, 14, 12)):
        ws.column_dimensions[col].width = w
    return ws


def dashboard_sheet(wb):
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = COMPANY
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Consolidated MIS dashboard — every figure below is a live "
                "formula over the Sales / Purchases ranges. Paste new rows "
                "into those tables and this page refreshes.")
    ws["A2"].font = Font(italic=True, size=10, color="595959")
    ws.merge_cells("A2:H2")

    kpis = [  # (label, formula)
        ("Total Sales (Rs.)", f"=SUM({S_AMT})"),
        ("Total Purchases (Rs.)", f"=SUM({P_AMT})"),
        ("Gross Profit (Rs.)", f"=SUM({S_AMT})-SUM({S_COGS})"),
        ("Margin %", f"=IF(SUM({S_AMT})=0,0,(SUM({S_AMT})-SUM({S_COGS}))"
                     f"/SUM({S_AMT}))"),
        ("Transactions", f"=COUNTA({S_INV})"),
        ("Avg Ticket (Rs.)", f"=AVERAGE({S_AMT})"),
        ("Anomaly Days", f'=COUNTIF({DAILY_FLAG},"ANOMALY")'),
        ("Products Sold", f"=SUMPRODUCT(--({PR_SALES}>0))"),
    ]
    ws["A4"] = "KEY PERFORMANCE INDICATORS"
    ws["A4"].font = Font(bold=True, size=12, color="1F4E78")
    for i, (lbl, formula) in enumerate(kpis):
        col = (i % 4) * 2 + 1
        row = 5 + (i // 4) * 3
        ws.cell(row=row, column=col, value=lbl).font = KPI_LBL
        v = ws.cell(row=row + 1, column=col, value=formula)
        v.font = KPI_FONT
        v.border = BORDER
        v.number_format = (PCT if "Margin" in lbl else
                          INT if lbl in ("Transactions", "Anomaly Days",
                                         "Products Sold") else INR)
        ws.merge_cells(start_row=row + 1, start_column=col,
                       end_row=row + 1, end_column=col + 1)

    # category summary (formulas) + bar chart
    r0 = 12
    ws.cell(row=r0, column=1, value="Sales by Category").font = Font(bold=True)
    ws.cell(row=r0 + 1, column=1, value="Category")
    ws.cell(row=r0 + 1, column=2, value="Sales")
    style_header(ws, r0 + 1, 2)
    for i, cat in enumerate(CATEGORIES):
        r = r0 + 2 + i
        ws.cell(row=r, column=1, value=cat).border = BORDER
        ws.cell(row=r, column=2,
                value=f"=SUMIF({S_CAT},A{r},{S_AMT})").border = BORDER
        ws.cell(row=r, column=2).number_format = INR

    # top-5 products (formulas) + bar chart
    c0 = 5
    ws.cell(row=r0, column=c0, value="Top 5 Products by Sales").font = \
        Font(bold=True)
    for c, h in ((c0, "Product"), (c0 + 1, "Sales")):
        cell = ws.cell(row=r0 + 1, column=c, value=h)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[r0 + 1].height = 22
    for i in range(5):
        r = r0 + 2 + i
        ws.cell(row=r, column=c0,
                value=f"=INDEX({PR_PROD},MATCH(LARGE({PR_SALES},{i+1}),"
                      f"{PR_SALES},0))").border = BORDER
        ws.cell(row=r, column=c0 + 1,
                value=f"=LARGE({PR_SALES},{i+1})").border = BORDER
        ws.cell(row=r, column=c0 + 1).number_format = INR

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 18

    line = LineChart()
    line.title = "Monthly Sales vs Purchases vs Gross Profit"
    line.style = 10
    line.height, line.width = 7.5, 14
    data = Reference(wb["Monthly"], min_col=2, min_row=1, max_row=7, max_col=5)
    cats = Reference(wb["Monthly"], min_col=1, min_row=2, max_row=7)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    ws.add_chart(line, "A22")

    bar = BarChart()
    bar.title = "Sales by Category"
    bar.style = 10
    bar.height, bar.width = 7.5, 12
    bar.add_data(Reference(ws, min_col=2, min_row=r0 + 1,
                           max_row=r0 + 1 + len(CATEGORIES)),
                 titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=r0 + 2,
                                 max_row=r0 + 1 + len(CATEGORIES)))
    ws.add_chart(bar, "A37")

    bar2 = BarChart()
    bar2.title = "Top 5 Products"
    bar2.style = 10
    bar2.height, bar2.width = 7.5, 12
    bar2.add_data(Reference(ws, min_col=c0 + 1, min_row=r0 + 1,
                            max_row=r0 + 6), titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=c0, min_row=r0 + 2,
                                  max_row=r0 + 6))
    ws.add_chart(bar2, "I22")
    return ws


def checks_sheet(wb):
    ws = wb.create_sheet("Checks")
    ws["A1"] = "INTEGRITY CHECKS — all must read OK"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E78")
    checks = [
        ("Sales amounts = Qty x Rate on every row",
         f"=SUM({S_AMT})-SUMPRODUCT({S_QTY},{S_RATE})",
         '=IF(ABS(B3)<0.01,"OK","FAIL")'),
        ("Purchase amounts = Qty x CostRate on every row",
         f"=SUM({P_AMT})-SUMPRODUCT({P_QTY},{P_COST})",
         '=IF(ABS(B4)<0.01,"OK","FAIL")'),
        ("No negative quantities",
         f"=SUMPRODUCT(--({S_QTY}<0))+SUMPRODUCT(--({P_QTY}<0))",
         '=IF(B5=0,"OK","FAIL")'),
        ("All sale products exist in master",
         f"=SUMPRODUCT(--({S_PROD}<>\"\"),--ISNA(MATCH({S_PROD},{PR_PROD},0)))",
         '=IF(B6=0,"OK","FAIL")'),
        ("Dates inside Apr-Sep 2026",
         f"=SUMPRODUCT(--({S_DATE}<>\"\"),--(({S_DATE}<DATE(2026,4,1))+"
         f"({S_DATE}>DATE(2026,9,30))))",
         '=IF(B7=0,"OK","FAIL")'),
        ("Every product has a StdCost",
         f"=SUMPRODUCT(--({PR_COST}=0))",
         '=IF(B8=0,"OK","FAIL")'),
    ]
    ws["A2"] = "Check"; ws["B2"] = "Difference"; ws["C2"] = "Status"
    style_header(ws, 2, 3)
    for i, (lbl, diff, status) in enumerate(checks, 3):
        ws.cell(row=i, column=1, value=lbl).border = BORDER
        ws.cell(row=i, column=2, value=diff).border = BORDER
        ws.cell(row=i, column=3, value=status).border = BORDER
        ws.cell(row=i, column=3).font = Font(bold=True)
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 10
    return ws


def guide_sheet(wb):
    ws = wb.create_sheet("Guide")
    ws["A1"] = "HANDOVER GUIDE — keeping this MIS fresh"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E78")
    lines = [
        "1. Paste new daily rows at the bottom of the Sales / Purchases data "
        "ranges (rows 2-5000 / 2-2000 carry formula headroom; Amount and COGS "
        "columns are pre-filled with formulas down to the headroom limit — "
        "copy the last formula row down if you outgrow it).",
        "2. Never type over the blue formula columns (Amount, COGS, StdCost, "
        "TotalSales, TotalQty) — they recalculate on paste.",
        "3. Formula ranges carry headroom (Sales rows 2-5000, Purchases 2-2000): "
        "if you outgrow them, extend the range ends in the formulas.",
        "4. Monthly, Daily, Dashboard and Checks sheets need no manual refresh: "
        "open the file and Excel recalculates.",
        "5. Charts read the Monthly / Dashboard tables, so new months appear "
        "once you extend the Monthly sheet's month column.",
        "6. Checks sheet: every row must read OK before the weekly review — "
        "a FAIL pinpoints the exact broken rule.",
        "7. Anomaly flags (Daily sheet) use avg ± 2 standard deviations on "
        "daily sales: investigate any ANOMALY row before the management meeting.",
        "Source files simulated here as data/*.csv (two sales parts + one "
        "purchases file, mirroring the client's several workbooks); in "
        "production, import each client workbook into the Sales / Purchases "
        "tables (Data > From Text/CSV) or paste directly.",
    ]
    for i, ln in enumerate(lines, 3):
        ws.cell(row=i, column=1, value=ln)
        ws.merge_cells(f"A{i}:H{i}")
    ws.column_dimensions["A"].width = 140
    return ws


def main():
    sales, purchases = gen_data()
    write_csvs(sales, purchases)
    wb = build_workbook(sales, purchases)
    monthly_sheet(wb)
    daily_sheet(wb)
    dashboard_sheet(wb)
    checks_sheet(wb)
    guide_sheet(wb)
    wb.active = wb["Dashboard"]
    wb.save(OUT_XLSX)
    print(f"wrote {OUT_XLSX}  ({len(sales)} sales rows, "
          f"{len(purchases)} purchase rows)")


if __name__ == "__main__":
    main()
