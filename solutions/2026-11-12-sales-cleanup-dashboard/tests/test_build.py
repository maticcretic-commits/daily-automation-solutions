#!/usr/bin/env python3
"""Tests for SalesClean Dashboard Factory (build-2026-11-12).

- Rebuilds the workbook + cleaned CSV from generate.py (fresh).
- Regex-scans every formula for short-form range literals
  ($H$2:$5000 — invalid in LibreOffice; must repeat the column letter).
- Independently re-implements the cleaning rules from the spec
  (separate code from generate.py) and checks kept/dropped counts,
  per-reason counts, and every kept row's normalised fields.
- Recalculates the workbook with LibreOffice headless and compares
  every Monthly row, every By_Product / By_Customer / By_Region row,
  every Dashboard KPI and every Check against the same ground truth.
- Asserts all Checks read PASS and no cell carries a formula error.
"""
import csv
import datetime
import re
import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
SRC_DIR = Path("/tmp/salesclean-src-20261112")
OUTDIR = Path("/tmp/salesclean-recalc-20261112")
XLSX = BASE / "sales-cleanup-dashboard.xlsx"
CLEANED_CSV = BASE / "data" / "cleaned_sales.csv"
SHORT_RANGE = re.compile(r"\$[A-Z]+\$\d+:\$\d+")
ERRORS = {"#VALUE!", "#NAME?", "#REF!", "#DIV/0!", "#N/A", "#NUM!",
          "Err:508", "Err:509", "Err:510", "Err:511", "Err:512", "Err:513",
          "Err:514", "Err:515", "Err:516", "Err:517", "Err:518", "Err:519",
          "Err:520", "Err:521", "Err:522", "Err:523", "Err:524", "Err:525",
          "Err:526", "Err:527", "Err:528", "Err:529", "Err:530", "Err:531",
          "Err:532"}
TOL_MONEY = 1.0
TOL_RATIO = 0.002

PRODUCTS = {
    "rice25kg": ("Rice 25kg", "Staples"),
    "wheatflour10kg": ("Wheat Flour 10kg", "Staples"),
    "sugar5kg": ("Sugar 5kg", "Staples"),
    "chanadal5kg": ("Chana Dal 5kg", "Pulses"),
    "moongdal2kg": ("Moong Dal 2kg", "Pulses"),
    "mustardoil5l": ("Mustard Oil 5L", "Edible Oil"),
    "sunfloweroil5l": ("Sunflower Oil 5L", "Edible Oil"),
    "chillipowder1kg": ("Chilli Powder 1kg", "Spices"),
    "turmericpowder1kg": ("Turmeric Powder 1kg", "Spices"),
    "corianderpowder1kg": ("Coriander Powder 1kg", "Spices"),
}
CUSTOMERS = {
    "sharmadistributors": ("Sharma Distributors", "Patna"),
    "kumarsons": ("Kumar & Sons", "Patna"),
    "vermatradingco": ("Verma Trading Co.", "Patna"),
    "cityhospitalcanteen": ("City Hospital Canteen", "Patna"),
    "anandgeneralstore": ("Anand General Store", "Muzaffarpur"),
    "guptakirana": ("Gupta Kirana", "Muzaffarpur"),
    "singhtraders": ("Singh Traders", "Gaya"),
    "ranisatistores": ("Rani Sati Stores", "Gaya"),
    "maadurgamart": ("Maa Durga Mart", "Bhagalpur"),
    "bhagalpurwholesalehub": ("Bhagalpur Wholesale Hub", "Bhagalpur"),
    "newbharatfoods": ("New Bharat Foods", "Patna"),
    "lakshmitraders": ("Lakshmi Traders", "Muzaffarpur"),
    "jaimatadistore": ("Jai Mata Di Store", "Gaya"),
    "sahubrothers": ("Sahu Brothers", "Bhagalpur"),
    "collegecanteenmagadhuniv": ("College Canteen (Magadh Univ.)", "Gaya"),
    "railwaycanteenpnbe": ("Railway Canteen (PNBE)", "Patna"),
}


def canon(name):
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def parse_date(text):
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


def num(text):
    text = (text or "").strip() if isinstance(text, str) else text
    if text == "" or text is None:
        return None
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def read_csv(name):
    with open(BASE / "data" / name, newline="") as f:
        return list(csv.DictReader(f))


def approx(a, b, tol, what):
    assert a is not None and b is not None, f"{what}: None value {a} vs {b}"
    assert abs(a - b) <= tol, f"{what}: {a} != {b} (tol {tol})"


def ground_truth():
    """Independent re-implementation of the cleaning spec."""
    raw = read_csv("raw_sales_dirty.csv")
    kept, counts = [], {r: 0 for r in
                        ["BAD_DATE", "UNKNOWN_CUSTOMER", "UNKNOWN_PRODUCT",
                         "BAD_QTY", "BAD_PRICE", "DUPLICATE"]}
    seen = set()
    for row in raw:
        date = parse_date(row["date"])
        if date is None:
            counts["BAD_DATE"] += 1
            continue
        ck = canon(row["customer"])
        if ck not in CUSTOMERS:
            counts["UNKNOWN_CUSTOMER"] += 1
            continue
        pk = canon(row["product"])
        if pk not in PRODUCTS:
            counts["UNKNOWN_PRODUCT"] += 1
            continue
        qty = num(row["qty"])
        if qty is None or qty <= 0:
            counts["BAD_QTY"] += 1
            continue
        price = num(row["unit_price"])
        if price is None or price <= 0:
            counts["BAD_PRICE"] += 1
            continue
        cust, region = CUSTOMERS[ck]
        prod, cat = PRODUCTS[pk]
        key = (date, cust, prod, qty, price)
        if key in seen:
            counts["DUPLICATE"] += 1
            continue
        seen.add(key)
        kept.append({"order_id": row["order_id"], "date": date,
                     "customer": cust, "region": region, "product": prod,
                     "category": cat, "qty": qty, "unit_price": price,
                     "amount": qty * price})
    targets = {r["month"]: int(r["target"])
               for r in read_csv("targets.csv")}
    monthly = {}
    for k in kept:
        m = k["date"].strftime("%Y-%m")
        monthly.setdefault(m, {"orders": 0, "revenue": 0.0})
        monthly[m]["orders"] += 1
        monthly[m]["revenue"] += k["amount"]
    by_product, by_customer, by_region = {}, {}, {}
    for k in kept:
        for d, key in ((by_product, k["product"]),
                       (by_customer, k["customer"]),
                       (by_region, k["region"])):
            d.setdefault(key, {"orders": 0, "revenue": 0.0})
            d[key]["orders"] += 1
            d[key]["revenue"] += k["amount"]
    total_rev = sum(k["amount"] for k in kept)
    return {"raw_n": len(raw), "kept": kept, "counts": counts,
            "targets": targets, "monthly": monthly,
            "by_product": by_product, "by_customer": by_customer,
            "by_region": by_region, "total_rev": total_rev,
            "total_orders": len(kept)}


def sheet_rows(ws, header_row=4):
    """Rows below a title block: header at `header_row`, data after."""
    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if all(v is None for v in row):
            continue
        rows.append(row)
    return rows


def find_header_row(ws, first_col_value):
    for row in ws.iter_rows(values_only=True):
        if row and row[0] == first_col_value:
            return row
    return None


def test_cleaning_counts(gt):
    assert gt["raw_n"] == 1430, gt["raw_n"]
    assert len(gt["kept"]) == 1336, len(gt["kept"])
    assert gt["counts"] == {"BAD_DATE": 8, "UNKNOWN_CUSTOMER": 4,
                            "UNKNOWN_PRODUCT": 4, "BAD_QTY": 18,
                            "BAD_PRICE": 15, "DUPLICATE": 45}, gt["counts"]
    # cleaned CSV matches the independent ground truth row for row
    rows = read_csv("cleaned_sales.csv")
    assert len(rows) == len(gt["kept"])
    for got, exp in zip(rows, gt["kept"]):
        assert got["order_id"] == exp["order_id"]
        assert got["date"] == exp["date"].isoformat()
        assert got["customer"] == exp["customer"]
        assert got["product"] == exp["product"]
        approx(float(got["amount"]), exp["amount"], TOL_MONEY, "amount")


def test_formula_hygiene():
    wb = load_workbook(XLSX)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    bad = SHORT_RANGE.search(cell.value)
                    assert not bad, f"{ws.title}!{cell.coordinate}: {cell.value}"


def test_recalc(gt):
    if OUTDIR.exists():
        shutil.rmtree(OUTDIR)
    OUTDIR.mkdir(parents=True)
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx",
                    "--outdir", str(OUTDIR), str(XLSX)],
                   check=True, capture_output=True, timeout=300)
    recalc = OUTDIR / XLSX.name
    assert recalc.exists(), "LibreOffice produced no output"
    wb = load_workbook(recalc, data_only=True)

    def no_errors(ws):
        for row in ws.iter_rows():
            for cell in row:
                assert cell.value not in ERRORS, \
                    f"{ws.title}!{cell.coordinate}: {cell.value}"

    # Monthly
    ws = wb["Monthly"]
    rows = [r for r in ws.iter_rows(min_row=5, values_only=True)
            if r[0] is not None]
    assert len(rows) == 25, f"Monthly rows: {len(rows)}"  # 24 months + TOTAL
    for r in rows[:24]:
        m = gt["monthly"][r[0]]
        assert r[1] == m["orders"], f"{r[0]} orders {r[1]} != {m['orders']}"
        approx(r[2], m["revenue"], TOL_MONEY, f"{r[0]} revenue")
        approx(r[3], gt["targets"][r[0]], 0.01, f"{r[0]} target")
        approx(r[4], m["revenue"] / gt["targets"][r[0]], TOL_RATIO,
               f"{r[0]} attainment")
    tot = rows[24]
    assert tot[0] == "TOTAL"
    assert tot[1] == gt["total_orders"]
    approx(tot[2], gt["total_rev"], TOL_MONEY, "total revenue")
    approx(tot[3], sum(gt["targets"].values()), 0.01, "total target")

    # By_Product / By_Customer / By_Region
    for sheet, gt_map in (("By_Product", gt["by_product"]),
                          ("By_Customer", gt["by_customer"]),
                          ("By_Region", gt["by_region"])):
        ws = wb[sheet]
        rows = [r for r in ws.iter_rows(min_row=5, values_only=True)
                if r[0] is not None]
        assert len(rows) == len(gt_map) + 1, f"{sheet} rows: {len(rows)}"
        for r in rows[:-1]:
            exp = gt_map[r[0]]
            assert r[2] == exp["orders"], f"{sheet} {r[0]} orders"
            approx(r[3], exp["revenue"], TOL_MONEY, f"{sheet} {r[0]} revenue")
            approx(r[4], exp["revenue"] / gt["total_rev"], TOL_RATIO,
                   f"{sheet} {r[0]} share")
        approx(rows[-1][3], gt["total_rev"], TOL_MONEY, f"{sheet} total")

    # Dashboard KPIs
    ws = wb["Dashboard"]
    vals = {r[0]: r[1] for r in ws.iter_rows(min_row=5, values_only=True)
            if r[0] is not None}
    approx(vals["Total revenue"], gt["total_rev"], TOL_MONEY, "KPI revenue")
    assert vals["Total orders"] == gt["total_orders"]
    approx(vals["Average order value"],
           gt["total_rev"] / gt["total_orders"], TOL_MONEY, "KPI AOV")
    approx(vals["Overall attainment vs target"],
           gt["total_rev"] / sum(gt["targets"].values()), TOL_RATIO,
           "KPI attainment")
    assert vals["Rows received (raw)"] == 1430
    assert vals["Rows kept after cleaning"] == 1336
    approx(vals["Clean rate"], 1336 / 1430, TOL_RATIO, "KPI clean rate")
    top_prod = max(gt["by_product"].items(), key=lambda kv: kv[1]["revenue"])[0]
    top_cust = max(gt["by_customer"].items(),
                   key=lambda kv: kv[1]["revenue"])[0]
    top_reg = max(gt["by_region"].items(),
                  key=lambda kv: kv[1]["revenue"])[0]
    assert vals["Top product by revenue"] == top_prod, vals["Top product by revenue"]
    assert vals["Top customer by revenue"] == top_cust
    assert vals["Top region by revenue"] == top_reg
    best_month = max(gt["monthly"].items(),
                     key=lambda kv: kv[1]["revenue"])[0]
    assert vals["Best month by revenue"] == best_month

    # Checks — all PASS, no formula errors anywhere
    ws = wb["Checks"]
    rows = [r for r in ws.iter_rows(min_row=5, values_only=True)
            if r[0] is not None]
    assert len(rows) == 10, f"Checks rows: {len(rows)}"
    for r in rows:
        assert r[3] == "PASS", f"check FAILED: {r[0]} = {r[1]}"
    for ws in wb.worksheets:
        no_errors(ws)


def main():
    sys.path.insert(0, str(BASE))
    import generate
    generate.main()  # fresh rebuild
    gt = ground_truth()
    test_cleaning_counts(gt)
    test_formula_hygiene()
    test_recalc(gt)
    print(f"ALL TESTS PASS — raw {gt['raw_n']}, kept {len(gt['kept'])}, "
          f"total revenue {gt['total_rev']:,.2f}, "
          f"orders {gt['total_orders']}, months {len(gt['monthly'])}")


if __name__ == "__main__":
    main()
