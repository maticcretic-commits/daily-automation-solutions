#!/usr/bin/env python3
"""Tests for SalesPerf BI Pack (build-2026-11-13).

- Rebuilds the workbook + CSV exports from generate.py (fresh).
- Regex-scans every formula for short-form range literals
  ($H$2:$5000 — invalid in LibreOffice; must repeat the column letter).
- Independently aggregates data/*.csv (separate code from generate.py)
  and checks Monthly / Quarterly / By_Product / By_Customer / By_Region /
  By_Category rows, Dashboard KPIs (incl. the live filter cells), the two
  CSV exports, and every Check — against LibreOffice headless recalc.
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
OUTDIR = Path("/tmp/salesperf-recalc-20261113")
XLSX = BASE / "sales-performance-bi-pack.xlsx"
SHORT_RANGE = re.compile(r"\$[A-Z]+\$\d+:\$\d+")
ERRORS = {"#VALUE!", "#NAME?", "#REF!", "#DIV/0!", "#N/A", "#NUM!",
          "Err:508", "Err:509", "Err:510", "Err:511", "Err:512", "Err:513",
          "Err:514", "Err:515", "Err:516", "Err:517", "Err:518", "Err:519",
          "Err:520", "Err:521", "Err:522", "Err:523", "Err:524", "Err:525",
          "Err:526", "Err:527", "Err:528", "Err:529", "Err:530", "Err:531",
          "Err:532"}
TOL_MONEY = 1.0
TOL_RATIO = 0.002


def read_csv(name):
    with open(BASE / "data" / name, newline="") as f:
        return list(csv.DictReader(f))


def approx(a, b, tol, what):
    assert a is not None and b is not None, f"{what}: None value {a} vs {b}"
    assert abs(a - b) <= tol, f"{what}: {a} != {b} (tol {tol})"


def ground_truth():
    orders = read_csv("orders.csv")
    targets = {r["month"]: int(r["target"]) for r in read_csv("targets.csv")}
    for o in orders:
        o["qty"] = int(o["qty"])
        o["unit_price"] = float(o["unit_price"])
        o["amount"] = float(o["amount"])
        o["date"] = datetime.date.fromisoformat(o["date"])
        o["month"] = o["date"].strftime("%Y-%m")
        o["quarter"] = (f"Q{(o['date'].month - 1) // 3 + 1} "
                        f"{o['date'].year}")
    monthly, quarterly = {}, {}
    by_product, by_customer, by_region, by_category = {}, {}, {}, {}
    for o in orders:
        m = monthly.setdefault(o["month"], {"orders": 0, "revenue": 0.0})
        m["orders"] += 1
        m["revenue"] += o["amount"]
        q = quarterly.setdefault(o["quarter"], {"orders": 0, "revenue": 0.0})
        q["orders"] += 1
        q["revenue"] += o["amount"]
        for d, key in ((by_product, o["product"]),
                       (by_customer, o["customer"]),
                       (by_region, o["region"]),
                       (by_category, o["category"])):
            e = d.setdefault(key, {"orders": 0, "qty": 0, "revenue": 0.0})
            e["orders"] += 1
            e["qty"] += o["qty"]
            e["revenue"] += o["amount"]
    total_rev = sum(o["amount"] for o in orders)
    total_tgt = sum(targets.values())
    return {"orders": orders, "targets": targets, "monthly": monthly,
            "quarterly": quarterly, "by_product": by_product,
            "by_customer": by_customer, "by_region": by_region,
            "by_category": by_category, "total_rev": total_rev,
            "total_tgt": total_tgt, "total_orders": len(orders)}


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

    def data_rows(ws, first_col=0):
        return [r for r in ws.iter_rows(min_row=5, values_only=True)
                if r[first_col] is not None]

    # Monthly: 10 months + TOTAL
    ws = wb["Monthly"]
    rows = data_rows(ws)
    assert len(rows) == 11, f"Monthly rows: {len(rows)}"
    for r in rows[:10]:
        m = gt["monthly"][r[0]]
        assert r[1] == m["orders"], f"{r[0]} orders {r[1]} != {m['orders']}"
        approx(r[2], m["revenue"], TOL_MONEY, f"{r[0]} revenue")
        approx(r[3], gt["targets"][r[0]], 0.01, f"{r[0]} target")
        approx(r[4], m["revenue"] / gt["targets"][r[0]], TOL_RATIO,
               f"{r[0]} pct_to_goal")
        approx(r[5], m["revenue"] / m["orders"], TOL_MONEY, f"{r[0]} AOV")
    tot = rows[10]
    assert tot[0] == "TOTAL"
    assert tot[1] == gt["total_orders"]
    approx(tot[2], gt["total_rev"], TOL_MONEY, "monthly total revenue")
    approx(tot[3], gt["total_tgt"], 0.01, "monthly total target")
    approx(tot[4], gt["total_rev"] / gt["total_tgt"], TOL_RATIO,
           "monthly total pct")

    # Quarterly: 4 quarters + TOTAL
    ws = wb["Quarterly"]
    rows = data_rows(ws)
    assert len(rows) == 5, f"Quarterly rows: {len(rows)}"
    for r in rows[:4]:
        q = gt["quarterly"][r[0]]
        assert r[1] == q["orders"], f"{r[0]} orders"
        approx(r[2], q["revenue"], TOL_MONEY, f"{r[0]} revenue")
        qt = sum(gt["targets"][m] for m in gt["monthly"]
                 if f"Q{(int(m[5:7]) - 1) // 3 + 1} {m[:4]}" == r[0])
        approx(r[3], qt, 0.01, f"{r[0]} target")
        approx(r[4], q["revenue"] / qt, TOL_RATIO, f"{r[0]} pct")
        on_tgt = sum(1 for m, v in gt["monthly"].items()
                     if f"Q{(int(m[5:7]) - 1) // 3 + 1} {m[:4]}" == r[0]
                     and v["revenue"] / gt["targets"][m] >= 1)
        assert r[5] == on_tgt, f"{r[0]} months_on_target {r[5]} != {on_tgt}"
    approx(rows[4][2], gt["total_rev"], TOL_MONEY, "quarterly total revenue")

    # By_Product / By_Customer / By_Region / By_Category
    for sheet, gt_map, ord_col, rev_col, share_col in (
            ("By_Product", gt["by_product"], 2, 4, 6),
            ("By_Customer", gt["by_customer"], 2, 3, 5),
            ("By_Region", gt["by_region"], 1, 2, 3),
            ("By_Category", gt["by_category"], 1, 2, 3)):
        ws = wb[sheet]
        rows = data_rows(ws)
        assert len(rows) == len(gt_map) + 1, f"{sheet} rows: {len(rows)}"
        for r in rows[:-1]:
            exp = gt_map[r[0]]
            assert r[ord_col] == exp["orders"], f"{sheet} {r[0]} orders"
            approx(r[rev_col], exp["revenue"], TOL_MONEY,
                   f"{sheet} {r[0]} revenue")
            approx(r[share_col], exp["revenue"] / gt["total_rev"],
                   TOL_RATIO, f"{sheet} {r[0]} share")
        approx(rows[-1][rev_col], gt["total_rev"], TOL_MONEY,
               f"{sheet} total revenue")

    # Dashboard KPIs + live filters (defaults: month 2026-10, category
    # = first alphabetically = Edible Oil)
    ws = wb["Dashboard"]
    vals = {}
    for r in ws.iter_rows(values_only=True):
        if r[0] and isinstance(r[1], (int, float, str)):
            vals[r[0]] = r[1]
    approx(vals["Total revenue"], gt["total_rev"], TOL_MONEY, "KPI revenue")
    approx(vals["Total target"], gt["total_tgt"], 0.01, "KPI target")
    approx(vals["YTD pct to goal"], gt["total_rev"] / gt["total_tgt"],
           TOL_RATIO, "KPI pct")
    assert vals["Total orders"] == gt["total_orders"]
    approx(vals["Average order value"],
           gt["total_rev"] / gt["total_orders"], TOL_MONEY, "KPI AOV")
    on_tgt = sum(1 for m, v in gt["monthly"].items()
                 if v["revenue"] / gt["targets"][m] >= 1)
    assert vals["Months on target (of 10)"] == on_tgt, vals["Months on target (of 10)"]
    best_month = max(gt["monthly"].items(),
                     key=lambda kv: kv[1]["revenue"])[0]
    assert vals["Best month by revenue"] == best_month
    approx(vals["Best month revenue"], gt["monthly"][best_month]["revenue"],
           TOL_MONEY, "KPI best month revenue")
    top_prod = max(gt["by_product"].items(),
                   key=lambda kv: kv[1]["revenue"])[0]
    top_cust = max(gt["by_customer"].items(),
                   key=lambda kv: kv[1]["revenue"])[0]
    top_reg = max(gt["by_region"].items(),
                  key=lambda kv: kv[1]["revenue"])[0]
    top_cat = max(gt["by_category"].items(),
                  key=lambda kv: kv[1]["revenue"])[0]
    assert vals["Top product by revenue"] == top_prod
    assert vals["Top customer by revenue"] == top_cust
    assert vals["Top region by revenue"] == top_reg
    assert vals["Top category by revenue"] == top_cat
    approx(vals["Revenue vs target gap"],
           gt["total_rev"] - gt["total_tgt"], TOL_MONEY, "KPI gap")
    # Live filter cells: default month 2026-10, category Edible Oil
    filt = sum(o["amount"] for o in gt["orders"]
               if o["month"] == "2026-10" and o["category"] == "Edible Oil")
    approx(vals["Filtered revenue (month × category)"], filt, TOL_MONEY,
           "filter month x category")
    filt_m = sum(o["amount"] for o in gt["orders"] if o["month"] == "2026-10")
    approx(vals["Filtered revenue (month only)"], filt_m, TOL_MONEY,
           "filter month")
    filt_c = sum(o["amount"] for o in gt["orders"]
                 if o["category"] == "Edible Oil")
    approx(vals["Filtered revenue (category only)"], filt_c, TOL_MONEY,
           "filter category")

    # Checks — all PASS, no formula errors anywhere
    ws = wb["Checks"]
    rows = data_rows(ws)
    assert len(rows) == 10, f"Checks rows: {len(rows)}"
    for r in rows:
        assert r[3] == "PASS", f"check FAILED: {r[0]} = {r[1]} vs {r[2]}"
    for ws in wb.worksheets:
        no_errors(ws)


def test_exports(gt):
    rows = read_csv("monthly_export.csv")
    assert len(rows) == 10, f"monthly export rows: {len(rows)}"
    for r in rows:
        m = gt["monthly"][r["month"]]
        assert int(r["orders"]) == m["orders"]
        approx(float(r["revenue"]), m["revenue"], TOL_MONEY, "export revenue")
        approx(float(r["pct_to_goal"]),
               m["revenue"] / gt["targets"][r["month"]], TOL_RATIO,
               "export pct")
    rows = read_csv("category_export.csv")
    assert len(rows) == 4, f"category export rows: {len(rows)}"
    for r in rows:
        c = gt["by_category"][r["category"]]
        assert int(r["orders"]) == c["orders"]
        approx(float(r["revenue"]), c["revenue"], TOL_MONEY,
               "export cat revenue")
        approx(float(r["share"]), c["revenue"] / gt["total_rev"], TOL_RATIO,
               "export cat share")


def main():
    sys.path.insert(0, str(BASE))
    import generate
    generate.main()  # fresh rebuild
    gt = ground_truth()
    test_formula_hygiene()
    test_recalc(gt)
    test_exports(gt)
    best_month = max(gt["monthly"].items(),
                     key=lambda kv: kv[1]["revenue"])[0]
    top_prod = max(gt["by_product"].items(),
                   key=lambda kv: kv[1]["revenue"])[0]
    top_cust = max(gt["by_customer"].items(),
                   key=lambda kv: kv[1]["revenue"])[0]
    print(f"ALL TESTS PASS — orders {gt['total_orders']}, "
          f"total revenue {gt['total_rev']:,.2f}, "
          f"attainment {gt['total_rev'] / gt['total_tgt'] * 100:.1f}%, "
          f"best month {best_month}, top product {top_prod}, "
          f"top customer {top_cust}")


if __name__ == "__main__":
    main()
