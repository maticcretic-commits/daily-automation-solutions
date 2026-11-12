#!/usr/bin/env python3
"""SalesClean Dashboard Factory — dirty sales data -> cleaned, refreshable dashboard.

Built around a real Freelancer brief ("Excel Sales Dashboard & Cleanup",
Freelancer, posted ~17h before the 2026-11-12 scout, 39 proposals,
avg INR 16,683): clean two years of sales data (duplicates, blanks,
inconsistent names) and build a refreshable Excel dashboard — sales by
product / customer / region against target.

What this builds, from fictional Gupta Traders, Patna data (FMCG wholesaler,
2024-11-01 .. 2026-10-31, deterministic data/*.csv, seed 20261112):
1. data/cleaned_sales.csv — the cleaned order lines (machine-readable).
2. sales-cleanup-dashboard.xlsx — a formula-driven workbook:
   Guide, Raw_Dirty (as received + disposition per row), Cleaning_Log
   (per-reason drop counts), Cleaned_Sales (live Amount/Month/Quarter
   formulas), Monthly (COUNTIFS/SUMIFS vs target + attainment),
   By_Product / By_Customer / By_Region (SUMIFS + share), Dashboard
   (9 KPIs + 3 native charts), Checks (10 integrity checks).

Cleaning rules, applied in this fixed order (order matters for the
per-reason counts — a row is dropped at the FIRST rule it fails):
1. BAD_DATE — date empty or unparseable as ISO (YYYY-MM-DD) or
   DD/MM/YYYY. Dash-separated DD-MM-YYYY is rejected as ambiguous.
2. UNKNOWN_CUSTOMER — customer name not in the master after normalisation.
3. UNKNOWN_PRODUCT — product name not in the master after normalisation.
4. BAD_QTY — qty blank, non-numeric, or <= 0.
5. BAD_PRICE — unit_price blank, non-numeric, or <= 0.
6. DUPLICATE — (date, customer, product, qty, price) already seen after
   normalisation (exact duplicates AND case/spacing near-duplicates).

Normalisation: names are matched on a canonical key (lowercase,
alphanumeric only) so "SHARMA DISTRIBUTORS ", "Sharma Distributors"
and "sharma  distributors" all resolve to the master name; dates are
stored as real date values; qty/price as numbers.

Conventions (LibreOffice-safe, as proven on earlier builds): plain A1
ranges with the column letter repeated on both ends, real date values
(never text dates), percentages stored as 0-1 fractions, numeric columns
guarded to 0 instead of "", LO-safe lookups via INDEX/MATCH/MAX
(never MAXIFS, never SMALL(IF(...))).
"""
import csv
import datetime
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT_XLSX = BASE / "sales-cleanup-dashboard.xlsx"
OUT_CLEANED = DATA / "cleaned_sales.csv"

PRODUCTS = [
    ("Rice 25kg", "Staples"),
    ("Wheat Flour 10kg", "Staples"),
    ("Sugar 5kg", "Staples"),
    ("Chana Dal 5kg", "Pulses"),
    ("Moong Dal 2kg", "Pulses"),
    ("Mustard Oil 5L", "Edible Oil"),
    ("Sunflower Oil 5L", "Edible Oil"),
    ("Chilli Powder 1kg", "Spices"),
    ("Turmeric Powder 1kg", "Spices"),
    ("Coriander Powder 1kg", "Spices"),
]
CUSTOMERS = [
    ("Sharma Distributors", "Patna", "Distributor"),
    ("Kumar & Sons", "Patna", "Distributor"),
    ("Verma Trading Co.", "Patna", "Wholesale"),
    ("City Hospital Canteen", "Patna", "Institutional"),
    ("Anand General Store", "Muzaffarpur", "Retail"),
    ("Gupta Kirana", "Muzaffarpur", "Retail"),
    ("Singh Traders", "Gaya", "Distributor"),
    ("Rani Sati Stores", "Gaya", "Retail"),
    ("Maa Durga Mart", "Bhagalpur", "Retail"),
    ("Bhagalpur Wholesale Hub", "Bhagalpur", "Wholesale"),
    ("New Bharat Foods", "Patna", "Institutional"),
    ("Lakshmi Traders", "Muzaffarpur", "Distributor"),
    ("Jai Mata Di Store", "Gaya", "Retail"),
    ("Sahu Brothers", "Bhagalpur", "Distributor"),
    ("College Canteen (Magadh Univ.)", "Gaya", "Institutional"),
    ("Railway Canteen (PNBE)", "Patna", "Institutional"),
]
PRODUCT_MAP = {re.sub(r"[^a-z0-9]", "", p.lower()): (p, c) for p, c in PRODUCTS}
CUSTOMER_MAP = {re.sub(r"[^a-z0-9]", "", c.lower()): (c, r, ch)
                for c, r, ch in CUSTOMERS}
REASONS = ["BAD_DATE", "UNKNOWN_CUSTOMER", "UNKNOWN_PRODUCT",
           "BAD_QTY", "BAD_PRICE", "DUPLICATE"]

NAVY = "1F3864"
TEAL = "0F7B7A"
LIGHT = "D9E2F3"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color=NAVY, size=14)
MONEY_FMT = '#,##0.00'
MONEY0_FMT = '#,##0'
PCT_FMT = '0.0%'
DATE_FMT = 'yyyy-mm-dd'


def canon(name):
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def parse_date(text):
    """ISO YYYY-MM-DD, then DD/MM/YYYY. Returns date or None."""
    text = (text or "").strip()
    if not text:
        return None
    try:
        return datetime.date.fromisoformat(text)
    except ValueError:
        pass
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if m:
        d, mo, y = (int(g) for g in m.groups())
        try:
            return datetime.date(y, mo, d)
        except ValueError:
            return None
    return None


def to_number(text):
    """Blank/non-numeric -> None, else float."""
    text = (text or "").strip() if isinstance(text, str) else text
    if text == "" or text is None:
        return None
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def clean_rows(raw):
    """Apply the six rules in order. Returns (kept, dispositions, counts).

    kept: list of dicts with normalised fields + amount.
    dispositions: {order_id: 'KEEP' or 'DROP:<REASON>'} in raw order.
    counts: {reason: n}.
    """
    kept, dispositions, counts = [], {}, {r: 0 for r in REASONS}
    seen = set()
    for row in raw:
        oid = row["order_id"]
        date = parse_date(row["date"])
        if date is None:
            dispositions[oid] = "DROP:BAD_DATE"
            counts["BAD_DATE"] += 1
            continue
        cust_key = canon(row["customer"])
        if cust_key not in CUSTOMER_MAP:
            dispositions[oid] = "DROP:UNKNOWN_CUSTOMER"
            counts["UNKNOWN_CUSTOMER"] += 1
            continue
        prod_key = canon(row["product"])
        if prod_key not in PRODUCT_MAP:
            dispositions[oid] = "DROP:UNKNOWN_PRODUCT"
            counts["UNKNOWN_PRODUCT"] += 1
            continue
        qty = to_number(row["qty"])
        if qty is None or qty <= 0:
            dispositions[oid] = "DROP:BAD_QTY"
            counts["BAD_QTY"] += 1
            continue
        price = to_number(row["unit_price"])
        if price is None or price <= 0:
            dispositions[oid] = "DROP:BAD_PRICE"
            counts["BAD_PRICE"] += 1
            continue
        cust, region, channel = CUSTOMER_MAP[cust_key]
        prod, cat = PRODUCT_MAP[prod_key]
        key = (date, cust, prod, qty, price)
        if key in seen:
            dispositions[oid] = "DROP:DUPLICATE"
            counts["DUPLICATE"] += 1
            continue
        seen.add(key)
        kept.append({"order_id": oid, "date": date, "customer": cust,
                     "region": region, "channel": channel, "product": prod,
                     "category": cat, "qty": qty, "unit_price": price,
                     "amount": qty * price})
        dispositions[oid] = "KEEP"
    return kept, dispositions, counts


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

def build_workbook(raw, kept, dispositions, counts, targets):
    n_raw = len(raw)
    n_kept = len(kept)
    n_drop = n_raw - n_kept
    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 95
    r = title_block(ws, "SalesClean Dashboard Factory — Guide",
                    "Gupta Traders, Patna (FMCG wholesaler)  •  Nov 2024 – Oct 2026  •  "
                    "Built from the Freelancer brief 'Excel Sales Dashboard & Cleanup'", 2)
    guide = [
        ("What this is", "A working demo of the brief: two years of messy sales data cleaned "
         "and turned into a refreshable Excel dashboard — sales by product, customer "
         "and region against monthly target."),
        ("How to refresh", "1) Paste the new month's raw rows at the bottom of Raw_Dirty. "
         "2) Re-run: python3 generate.py — it re-cleans everything and rebuilds this "
         "workbook. 3) Every figure on Monthly / By_* / Dashboard is a live formula, "
         "so the dashboard updates from the rebuilt Cleaned_Sales sheet."),
        ("Cleaning rules", "Applied in order — BAD_DATE, UNKNOWN_CUSTOMER, UNKNOWN_PRODUCT, "
         "BAD_QTY, BAD_PRICE, DUPLICATE. A row is dropped at the first rule it fails. "
         "Names match on a canonical key (lowercase, alphanumeric only), so case and "
         "spacing variants resolve to the master name. Dates accept ISO or DD/MM/YYYY."),
        ("Sheets", "Raw_Dirty — as received, with a Disposition per row. Cleaning_Log — "
         "drop counts per reason. Cleaned_Sales — kept rows (live Amount / Month / "
         "Quarter formulas). Monthly / By_Product / By_Customer / By_Region — "
         "COUNTIFS / SUMIFS summaries. Dashboard — KPIs + 3 charts. Checks — "
         "10 integrity checks, all must read PASS."),
        ("Fictional data", "All customers, products and figures are invented for the demo. "
         "No real business data is used."),
    ]
    for label, text in guide:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 42
        r += 1

    # ---------------- Raw_Dirty ----------------
    ws = wb.create_sheet("Raw_Dirty")
    ws.sheet_properties.tabColor = "C00000"
    headers = ["order_id", "date", "customer", "region", "channel", "product",
               "category", "qty", "unit_price", "disposition"]
    ws.column_dimensions["A"].width = 12
    for col, w in zip("BCDEFGHIJ", [13, 32, 13, 14, 22, 12, 10, 11, 18]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Raw sales data — as received",
                     f"{n_raw} rows. Disposition: KEEP, or DROP:<reason> "
                     "(first failed rule). Re-run generate.py after pasting new rows.", 10)
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 10)
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:J{hr + n_raw}"
    for i, row in enumerate(raw):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=row["order_id"])
        d = parse_date(row["date"])
        c = ws.cell(row=r, column=2, value=d if d else row["date"])
        if d:
            c.number_format = DATE_FMT
        ws.cell(row=r, column=3, value=row["customer"])
        ws.cell(row=r, column=4, value=row["region"])
        ws.cell(row=r, column=5, value=row["channel"])
        ws.cell(row=r, column=6, value=row["product"])
        ws.cell(row=r, column=7, value=row["category"])
        q = to_number(row["qty"])
        ws.cell(row=r, column=8, value=q if q is not None else row["qty"])
        p = to_number(row["unit_price"])
        ws.cell(row=r, column=9, value=p if p is not None else row["unit_price"])
        if p is not None:
            ws.cell(row=r, column=9).number_format = MONEY_FMT
        ws.cell(row=r, column=10, value=dispositions[row["order_id"]])
    raw_first, raw_last = hr + 1, hr + n_raw

    # ---------------- Cleaning_Log ----------------
    ws = wb.create_sheet("Cleaning_Log")
    ws.sheet_properties.tabColor = "ED7D31"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 72
    hr = title_block(ws, "Cleaning log — rows dropped per reason",
                     "Rule order: BAD_DATE, UNKNOWN_CUSTOMER, UNKNOWN_PRODUCT, "
                     "BAD_QTY, BAD_PRICE, DUPLICATE.", 3)
    for c, h in enumerate(["reason", "rows", "what it means"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 3)
    notes = {
        "BAD_DATE": "Date empty or unparseable (ISO or DD/MM/YYYY only; dash DD-MM-YYYY rejected).",
        "UNKNOWN_CUSTOMER": "Customer name not in the master list after normalisation.",
        "UNKNOWN_PRODUCT": "Product name not in the master list after normalisation.",
        "BAD_QTY": "Quantity blank, non-numeric, or <= 0.",
        "BAD_PRICE": "Unit price blank, non-numeric, or <= 0.",
        "DUPLICATE": "Same (date, customer, product, qty, price) already kept — exact or case/spacing near-duplicate.",
    }
    for i, reason in enumerate(REASONS):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=reason)
        ws.cell(row=r, column=2, value=counts[reason])
        ws.cell(row=r, column=3, value=notes[reason])
    tr = hr + 1 + len(REASONS)
    ws.cell(row=tr, column=1, value="TOTAL DROPPED").font = Font(bold=True)
    ws.cell(row=tr, column=2,
           value=f"=SUM(B{hr + 1}:B{hr + len(REASONS)})").font = Font(bold=True)
    ws.cell(row=tr + 1, column=1, value="ROWS KEPT").font = Font(bold=True)
    ws.cell(row=tr + 1, column=2, value=n_kept).font = Font(bold=True)
    ws.cell(row=tr + 2, column=1, value="ROWS RECEIVED").font = Font(bold=True)
    ws.cell(row=tr + 2, column=2, value=n_raw).font = Font(bold=True)
    log_first, log_last = hr + 1, hr + len(REASONS)
    log_total_row = tr

    # ---------------- Cleaned_Sales ----------------
    ws = wb.create_sheet("Cleaned_Sales")
    ws.sheet_properties.tabColor = TEAL
    headers = ["order_id", "date", "customer", "region", "channel", "product",
               "category", "qty", "unit_price", "amount", "month", "quarter"]
    widths = [12, 13, 32, 13, 14, 22, 12, 10, 11, 13, 10, 10]
    for col, w in zip("ABCDEFGHIJKL", widths):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Cleaned sales — kept rows",
                     "Names normalised to master lists; dates are real date values. "
                     "Amount, Month and Quarter are live formulas.", 12)
    for c, h in enumerate(headers, 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 12)
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:L{hr + n_kept}"
    for i, row in enumerate(kept):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=row["order_id"])
        c = ws.cell(row=r, column=2, value=row["date"])
        c.number_format = DATE_FMT
        ws.cell(row=r, column=3, value=row["customer"])
        ws.cell(row=r, column=4, value=row["region"])
        ws.cell(row=r, column=5, value=row["channel"])
        ws.cell(row=r, column=6, value=row["product"])
        ws.cell(row=r, column=7, value=row["category"])
        ws.cell(row=r, column=8, value=row["qty"])
        c = ws.cell(row=r, column=9, value=row["unit_price"])
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=10, value=f"=H{r}*I{r}")
        c.number_format = MONEY_FMT
        ws.cell(row=r, column=11, value=f'=TEXT(B{r},"yyyy-mm")')
        ws.cell(row=r, column=12,
               value=f'="Q"&INT((MONTH(B{r})-1)/3)+1&" "&YEAR(B{r})')
    cl_first, cl_last = hr + 1, hr + n_kept
    cl = f"Cleaned_Sales!$A${cl_first}:$L${cl_last}"
    amt = f"Cleaned_Sales!$J${cl_first}:$J${cl_last}"
    mon = f"Cleaned_Sales!$K${cl_first}:$K${cl_last}"
    prod = f"Cleaned_Sales!$F${cl_first}:$F${cl_last}"
    cust = f"Cleaned_Sales!$C${cl_first}:$C${cl_last}"
    regn = f"Cleaned_Sales!$D${cl_first}:$D${cl_last}"

    # ---------------- Monthly ----------------
    ws = wb.create_sheet("Monthly")
    ws.sheet_properties.tabColor = NAVY
    for col, w in zip("ABCDEF", [12, 10, 16, 14, 12, 12]):
        ws.column_dimensions[col].width = w
    hr = title_block(ws, "Monthly sales vs target",
                     "Orders and revenue are COUNTIFS / SUMIFS over Cleaned_Sales. "
                     "Attainment = revenue / target.", 6)
    for c, h in enumerate(["month", "orders", "revenue", "target",
                           "attainment", "avg order"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 6)
    months = sorted(targets)
    for i, m in enumerate(months):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=m)
        ws.cell(row=r, column=2, value=f"=COUNTIFS({mon},$A{r})")
        c = ws.cell(row=r, column=3, value=f"=SUMIFS({amt},{mon},$A{r})")
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=4, value=targets[m])
        c.number_format = MONEY_FMT
        c = ws.cell(row=r, column=5, value=f"=IF($D{r}=0,0,$C{r}/$D{r})")
        c.number_format = PCT_FMT
        c = ws.cell(row=r, column=6, value=f"=IF($B{r}=0,0,$C{r}/$B{r})")
        c.number_format = MONEY_FMT
    tr = hr + 1 + len(months)
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, fmt in ((2, None), (3, MONEY_FMT), (4, MONEY_FMT)):
        c = ws.cell(row=tr, column=col,
                   value=f"=SUM({get_column_letter(col)}{hr + 1}:"
                         f"{get_column_letter(col)}{hr + len(months)})")
        c.font = Font(bold=True)
        if fmt:
            c.number_format = fmt
    c = ws.cell(row=tr, column=5, value=f"=IF($D{tr}=0,0,$C{tr}/$D{tr})")
    c.font = Font(bold=True)
    c.number_format = PCT_FMT
    c = ws.cell(row=tr, column=6, value=f"=IF($B{tr}=0,0,$C{tr}/$B{tr})")
    c.font = Font(bold=True)
    c.number_format = MONEY_FMT
    mo_first, mo_last = hr + 1, hr + len(months)
    mo_total = tr

    # ---------------- By_Product / By_Customer / By_Region ----------------
    def summary_sheet(name, key_col, items, label):
        w = wb.create_sheet(name)
        w.sheet_properties.tabColor = TEAL
        for col, wd in zip("ABCDE", [32, 14, 10, 16, 10]):
            w.column_dimensions[col].width = wd
        h = title_block(w, f"Sales by {label}",
                        "Orders and revenue are COUNTIFS / SUMIFS over Cleaned_Sales; "
                        "share is revenue / column total.", 5)
        for c, hh in enumerate([label, "category / region", "orders",
                                "revenue", "share"], 1):
            w.cell(row=h, column=c, value=hh)
        style_header(w, h, 5)
        for i, (item, extra) in enumerate(items):
            r = h + 1 + i
            w.cell(row=r, column=1, value=item)
            w.cell(row=r, column=2, value=extra)
            w.cell(row=r, column=3,
                  value=f"=COUNTIFS({key_col},$A{r})")
            c = w.cell(row=r, column=4,
                      value=f"=SUMIFS({amt},{key_col},$A{r})")
            c.number_format = MONEY_FMT
            c = w.cell(row=r, column=5,
                      value=f"=IF(SUM($D${h + 1}:$D${h + len(items)})=0,0,"
                            f"$D{r}/SUM($D${h + 1}:$D${h + len(items)}))")
            c.number_format = PCT_FMT
        t = h + 1 + len(items)
        w.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
        for col in (3, 4):
            c = w.cell(row=t, column=col,
                      value=f"=SUM({get_column_letter(col)}{h + 1}:"
                            f"{get_column_letter(col)}{h + len(items)})")
            c.font = Font(bold=True)
            if col == 4:
                c.number_format = MONEY_FMT
        c = w.cell(row=t, column=5, value=1)
        c.font = Font(bold=True)
        c.number_format = PCT_FMT
        return h, t

    bp_h, bp_t = summary_sheet("By_Product", prod,
                               [(p, c) for p, c in PRODUCTS], "product")
    bc_h, bc_t = summary_sheet("By_Customer", cust,
                               [(c, r) for c, r, _ in CUSTOMERS], "customer")
    regions = sorted({r for _, r, _ in CUSTOMERS})
    br_h, br_t = summary_sheet("By_Region", regn,
                               [(r, "") for r in regions], "region")

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 40
    hr = title_block(ws, "Dashboard — Gupta Traders, two years of sales",
                     "Every KPI below is a live formula over Monthly / By_* / "
                     "Cleaned_Sales. Period Nov 2024 – Oct 2026.", 3)
    kpis = [
        ("Total revenue", f"=Monthly!C{mo_total}", MONEY_FMT),
        ("Total orders", f"=Monthly!B{mo_total}", None),
        ("Average order value", f"=IF(B{hr + 2}=0,0,B{hr + 1}/B{hr + 2})", MONEY_FMT),
        ("Overall attainment vs target", f"=Monthly!E{mo_total}", PCT_FMT),
        ("Rows received (raw)", f"=COUNTA(Raw_Dirty!A{raw_first}:A{raw_last})", None),
        ("Rows kept after cleaning", f"=COUNTA(Cleaned_Sales!A{cl_first}:A{cl_last})", None),
        ("Clean rate", f"=IF(B{hr + 5}=0,0,B{hr + 6}/B{hr + 5})", PCT_FMT),
        ("Top product by revenue",
         f"=INDEX(By_Product!A{bp_h + 1}:A{bp_t - 1},MATCH(MAX(By_Product!D{bp_h + 1}:D{bp_t - 1}),By_Product!D{bp_h + 1}:D{bp_t - 1},0))", None),
        ("Top customer by revenue",
         f"=INDEX(By_Customer!A{bc_h + 1}:A{bc_t - 1},MATCH(MAX(By_Customer!D{bc_h + 1}:D{bc_t - 1}),By_Customer!D{bc_h + 1}:D{bc_t - 1},0))", None),
        ("Top region by revenue",
         f"=INDEX(By_Region!A{br_h + 1}:A{br_t - 1},MATCH(MAX(By_Region!D{br_h + 1}:D{br_t - 1}),By_Region!D{br_h + 1}:D{br_t - 1},0))", None),
        ("Best month by revenue",
         f"=INDEX(Monthly!A{mo_first}:A{mo_last},MATCH(MAX(Monthly!C{mo_first}:C{mo_last}),Monthly!C{mo_first}:C{mo_last},0))", None),
    ]
    # Average-order formula references the two KPI rows above it; fix refs:
    kpis[2] = ("Average order value",
               f"=IF(B{hr + 2}=0,0,B{hr + 1}/B{hr + 2})", MONEY_FMT)
    for i, (label, formula, fmt) in enumerate(kpis):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=formula)
        if fmt:
            c.number_format = fmt
    dash_first_kpi = hr + 1
    # Charts
    ch1 = BarChart()
    ch1.type, ch1.style, ch1.title = "col", 10, "Monthly revenue"
    ch1.y_axis.title, ch1.x_axis.title = "Revenue (INR)", "Month"
    ch1.height, ch1.width = 7.5, 14
    ch1.add_data(Reference(wb["Monthly"], min_col=3, min_row=mo_first - 1,
                           max_row=mo_last), titles_from_data=True)
    ch1.set_categories(Reference(wb["Monthly"], min_col=1, min_row=mo_first,
                                 max_row=mo_last))
    ch1.shape = 4
    ws.add_chart(ch1, f"A{hr + len(kpis) + 3}")
    ch2 = BarChart()
    ch2.type, ch2.style, ch2.title = "col", 10, "Revenue by product"
    ch2.y_axis.title = "Revenue (INR)"
    ch2.height, ch2.width = 7.5, 14
    ch2.add_data(Reference(wb["By_Product"], min_col=4, min_row=bp_h,
                           max_row=bp_t - 1), titles_from_data=True)
    ch2.set_categories(Reference(wb["By_Product"], min_col=1, min_row=bp_h + 1,
                                 max_row=bp_t - 1))
    ch2.shape = 4
    ws.add_chart(ch2, f"A{hr + len(kpis) + 20}")
    ch3 = PieChart()
    ch3.title = "Revenue share by region"
    ch3.height, ch3.width = 7.5, 14
    ch3.add_data(Reference(wb["By_Region"], min_col=4, min_row=br_h,
                           max_row=br_t - 1), titles_from_data=True)
    ch3.set_categories(Reference(wb["By_Region"], min_col=1, min_row=br_h + 1,
                                 max_row=br_t - 1))
    ws.add_chart(ch3, f"A{hr + len(kpis) + 37}")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws.sheet_properties.tabColor = "70AD47"
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    hr = title_block(ws, "Integrity checks — all must read PASS",
                     "Each check is a live formula. A FAIL means the rebuild or the "
                     "source data changed — do not trust the dashboard until it is PASS.", 4)
    for c, h in enumerate(["check", "value", "expected", "result"], 1):
        ws.cell(row=hr, column=c, value=h)
    style_header(ws, hr, 4)
    checks = [
        ("Raw rows = kept + dropped",
         f"=COUNTA(Raw_Dirty!A{raw_first}:A{raw_last})-COUNTA(Cleaned_Sales!A{cl_first}:A{cl_last})-Cleaning_Log!B{log_total_row}",
         0),
        ("Revenue tie-out: Cleaned amount = Monthly revenue",
         f"=SUM(Cleaned_Sales!J{cl_first}:J{cl_last})-Monthly!C{mo_total}",
         0),
        ("No blanks in Cleaned key fields",
         f"=COUNTBLANK(Cleaned_Sales!B{cl_first}:I{cl_last})",
         0),
        ("No non-positive amounts",
         f'=COUNTIF(Cleaned_Sales!J{cl_first}:J{cl_last},"<=0")',
         0),
        ("Months covered",
         f"=COUNTA(Monthly!A{mo_first}:A{mo_last})",
         24),
        ("Attainment within 0–200% (MAX)",
         f"=MAX(Monthly!E{mo_first}:E{mo_last})",
         "<=2"),
        ("Product revenue ties to total",
         f"=SUM(By_Product!D{bp_h + 1}:D{bp_t - 1})-Monthly!C{mo_total}",
         0),
        ("Customer revenue ties to total",
         f"=SUM(By_Customer!D{bc_h + 1}:D{bc_t - 1})-Monthly!C{mo_total}",
         0),
        ("Region revenue ties to total",
         f"=SUM(By_Region!D{br_h + 1}:D{br_t - 1})-Monthly!C{mo_total}",
         0),
        ("Cleaning log total = rows dropped",
         f"=Cleaning_Log!B{log_total_row}-(COUNTA(Raw_Dirty!A{raw_first}:A{raw_last})-COUNTA(Cleaned_Sales!A{cl_first}:A{cl_last}))",
         0),
    ]
    for i, (label, formula, expected) in enumerate(checks):
        r = hr + 1 + i
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=formula)
        if isinstance(expected, str):
            ws.cell(row=r, column=3, value=expected)
            ws.cell(row=r, column=4,
                   value=f'=IF(B{r}<=2,"PASS","FAIL")')
        else:
            ws.cell(row=r, column=3, value=expected)
            ws.cell(row=r, column=4,
                   value=f'=IF(ABS(B{r}-C{r})<0.01,"PASS","FAIL")')
    return wb, {"n_raw": n_raw, "n_kept": n_kept, "n_drop": n_drop,
                "dash_first_kpi": dash_first_kpi, "mo_total": mo_total}


def main():
    with open(DATA / "raw_sales_dirty.csv", newline="") as f:
        raw = list(csv.DictReader(f))
    with open(DATA / "targets.csv", newline="") as f:
        targets = {r["month"]: int(r["target"]) for r in csv.DictReader(f)}
    kept, dispositions, counts = clean_rows(raw)
    with open(OUT_CLEANED, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["order_id", "date", "customer",
                                          "region", "channel", "product",
                                          "category", "qty", "unit_price",
                                          "amount"])
        w.writeheader()
        for row in kept:
            w.writerow({**row, "date": row["date"].isoformat()})
    wb, meta = build_workbook(raw, kept, dispositions, counts, targets)
    wb.save(OUT_XLSX)
    print(f"raw={meta['n_raw']} kept={meta['n_kept']} "
          f"dropped={meta['n_drop']} counts={counts}")
    print(f"wrote {OUT_CLEANED} and {OUT_XLSX}")


if __name__ == "__main__":
    main()
