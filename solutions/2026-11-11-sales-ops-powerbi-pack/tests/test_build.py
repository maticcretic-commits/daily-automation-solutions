#!/usr/bin/env python3
"""Tests for SalesOpsPulse (build-2026-11-11).

- Rebuilds the workbook + SQLite DB from generate.py (fresh).
- Regex-scans every formula for short-form range literals
  ($H$2:$5000 — invalid in LibreOffice; must repeat the column letter).
- Verifies every SQLite view (9 views) against an independent Python
  ground truth from data/*.csv.
- Recalculates the workbook with LibreOffice headless and compares every
  Sales_Overview row, every Raw_Inventory row, every Receivables_Ageing
  bucket, every Category row, the top-5 debtors, every Customers row and
  every Dashboard KPI against the same ground truth.
- Asserts all Checks read PASS and no cell carries a formula error.
"""
import csv
import datetime
import re
import shutil
import sqlite3
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
SRC_DIR = Path("/tmp/salesops-src-20261111")
OUTDIR = Path("/tmp/salesops-recalc-20261111")
XLSX = BASE / "sales-ops-powerbi-pack.xlsx"
DB = BASE / "sales_ops.db"
AS_OF = datetime.date(2026, 11, 10)
PERIOD_DAYS = 224
MONTHS = ["Apr 2026", "May 2026", "Jun 2026", "Jul 2026", "Aug 2026",
          "Sep 2026", "Oct 2026", "Nov 2026"]
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


def approx(a, b, tol):
    assert a is not None and b is not None, f"None value: {a} vs {b}"
    assert abs(a - b) <= tol, f"{a} != {b} (tol {tol})"


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


def ground_truth():
    products = {p["sku"]: p for p in read_csv("products.csv")}
    customers = read_csv("customers.csv")
    orders = read_csv("orders.csv")
    receipts = read_csv("receipts.csv")
    deliveries = read_csv("deliveries.csv")
    calls = read_csv("calls.csv")
    receivables = read_csv("receivables.csv")
    targets = {t["month"]: int(t["target"]) for t in read_csv("targets.csv")}

    monthly = {m: {"orders": 0, "units": 0, "rev": 0.0, "margin": 0.0,
                   "chan": {c: 0.0 for c in
                            ["Distributor", "Retail", "Institutional",
                             "Online"]}} for m in MONTHS}
    cats = {}
    custs = {c["customer"]: {"orders": 0, "rev": 0.0, "out": 0}
             for c in customers}
    inv = {}
    for p in products.values():
        inv[p["sku"]] = {"opening": int(p["opening_stock"]), "recv": 0,
                         "sold": 0, "reorder": int(p["reorder_level"]),
                         "cost": int(p["cost_price"]),
                         "category": p["category"]}
    for r in receipts:
        inv[r["sku"]]["recv"] += int(r["qty"])
    for o in orders:
        d = datetime.date.fromisoformat(o["date"])
        m = d.strftime("%b %Y")
        qty = int(o["qty"])
        price = int(o["unit_price"])
        disc = int(o["discount_pct"])
        amt = qty * price * (100 - disc) / 100
        marg = qty * (price * (100 - disc) / 100
                      - int(products[o["sku"]]["cost_price"]))
        monthly[m]["orders"] += 1
        monthly[m]["units"] += qty
        monthly[m]["rev"] += amt
        monthly[m]["margin"] += marg
        monthly[m]["chan"][o["channel"]] += amt
        c = cats.setdefault(o["category"],
                            {"rev": 0.0, "orders": 0, "margin": 0.0})
        c["rev"] += amt
        c["orders"] += 1
        c["margin"] += marg
        custs[o["customer"]]["orders"] += 1
        custs[o["customer"]]["rev"] += amt
        inv[o["sku"]]["sold"] += qty
    for v in inv.values():
        v["closing"] = v["opening"] + v["recv"] - v["sold"]
        v["value"] = v["closing"] * v["cost"]
        v["status"] = ("OUT" if v["closing"] <= 0
                       else "LOW" if v["closing"] <= v["reorder"] else "OK")
        v["cover"] = 0 if v["sold"] == 0 else v["closing"] / (v["sold"] / PERIOD_DAYS)
    buckets = {b: {"n": 0, "out": 0} for b in
               ["Current", "0-30", "31-60", "61-90", "90+", "Paid"]}
    total_out = 0
    debtors = []
    for idx, r in enumerate(receivables):
        out = int(r["amount"]) - int(r["paid_amount"])
        b = bucket_of(datetime.date.fromisoformat(r["due_date"]), out)
        buckets[b]["n"] += 1
        buckets[b]["out"] += out
        total_out += out
        custs[r["customer"]]["out"] += out
        dpd = 0 if out == 0 else max(0, (AS_OF - datetime.date.fromisoformat(r["due_date"])).days)
        debtors.append((r["customer"], out, dpd, idx))
    debtors.sort(key=lambda x: (-x[1], x[3]))
    top5 = debtors[:5]
    on_time = sum(1 for x in deliveries
                  if x["actual_date"] <= x["planned_date"])
    delay = sum((datetime.date.fromisoformat(x["actual_date"])
                 - datetime.date.fromisoformat(x["planned_date"])).days
                for x in deliveries)
    placed = sum(1 for c in calls if c["outcome"] == "Order placed")
    total_rev = sum(m["rev"] for m in monthly.values())
    total_margin = sum(m["margin"] for m in monthly.values())
    total_orders = sum(m["orders"] for m in monthly.values())
    nov = monthly["Nov 2026"]
    best_month = max(MONTHS, key=lambda m: monthly[m]["rev"])
    top_cust = max(custs, key=lambda c: custs[c]["rev"])
    stock_value = sum(v["value"] for v in inv.values())
    alerts = sum(1 for v in inv.values() if v["status"] != "OK")
    stockout = sum(1 for v in inv.values() if v["status"] == "OUT")
    avg_cover = sum(v["cover"] for v in inv.values()) / len(inv)
    kpis = {
        "rev": total_rev, "orders": total_orders,
        "aov": total_rev / total_orders, "margin": total_margin,
        "margin_pct": total_margin / total_rev,
        "nov_rev": nov["rev"], "nov_pacing": (nov["rev"] / targets["Nov 2026"]) * 3,
        "best_month": best_month, "top_cust": top_cust,
        "stock_value": stock_value, "alerts": alerts, "stockout": stockout,
        "on_time_pct": on_time / len(deliveries),
        "avg_delay": delay / len(deliveries),
        "calls": len(calls), "conv": placed / len(calls),
        "outstanding": total_out, "over90": buckets["90+"]["out"],
        "avg_cover": avg_cover,
    }
    return {"monthly": monthly, "cats": cats, "custs": custs, "inv": inv,
            "buckets": buckets, "top5": top5, "kpis": kpis,
            "targets": targets, "deliveries": deliveries, "calls": calls}


def test_rebuild_and_formulas():
    r = subprocess.run([sys.executable, "generate.py"], cwd=BASE,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr[-2000:]
    assert XLSX.exists() and DB.exists()
    wb = load_workbook(XLSX)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    assert not SHORT_RANGE.search(c.value), \
                        f"short range in {ws.title}!{c.coordinate}: {c.value}"
    print("  rebuild + short-range scan OK")


def test_db_views(gt):
    con = sqlite3.connect(DB)
    cur = con.cursor()
    rows = cur.execute("SELECT ym, orders, revenue FROM v_monthly_sales"
                       ).fetchall()
    assert len(rows) == 8, rows
    for ym, orders, rev in rows:
        m = datetime.date(int(ym[:4]), int(ym[5:7]), 1).strftime("%b %Y")
        approx(orders, gt["monthly"][m]["orders"], 0)
        approx(rev, gt["monthly"][m]["rev"], TOL_MONEY)
    rows = cur.execute("SELECT channel, revenue FROM v_sales_by_channel"
                       ).fetchall()
    assert len(rows) == 4
    for ch, rev in rows:
        approx(rev, sum(gt["monthly"][m]["chan"][ch] for m in MONTHS),
               TOL_MONEY)
    rows = cur.execute("SELECT sku, closing FROM v_inventory_status"
                       ).fetchall()
    assert len(rows) == 12
    for sku, closing in rows:
        assert closing == gt["inv"][sku]["closing"], sku
    rows = cur.execute("SELECT sku, status FROM v_reorder_alerts").fetchall()
    assert len(rows) == gt["kpis"]["alerts"], rows
    assert all(s in ("LOW", "OUT") for _, s in rows)
    rows = cur.execute("SELECT bucket, invoices, outstanding "
                       "FROM v_receivables_ageing").fetchall()
    seen = {b: (n, o) for b, n, o in rows}
    for b, v in gt["buckets"].items():
        if v["n"]:
            assert seen[b][0] == v["n"], b
            approx(seen[b][1], v["out"], TOL_MONEY)
    rows = cur.execute("SELECT customer FROM v_top_customers").fetchall()
    assert [r[0] for r in rows] == sorted(
        gt["custs"], key=lambda c: -gt["custs"][c]["rev"])[:5]
    rows = cur.execute("SELECT channel, calls FROM v_call_conversion"
                       ).fetchall()
    assert sum(n for _, n in rows) == gt["kpis"]["calls"]
    rows = cur.execute("SELECT route, deliveries FROM v_delivery_performance"
                       ).fetchall()
    assert sum(n for _, n in rows) == len(gt["deliveries"])
    con.close()
    print("  DB views OK (9 views)")


def recalc():
    if SRC_DIR.exists():
        shutil.rmtree(SRC_DIR)
    if OUTDIR.exists():
        shutil.rmtree(OUTDIR)
    SRC_DIR.mkdir(parents=True)
    OUTDIR.mkdir(parents=True)
    shutil.copy(XLSX, SRC_DIR / XLSX.name)
    r = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx",
                        "--outdir", str(OUTDIR), str(SRC_DIR / XLSX.name)],
                       capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, r.stderr[-2000:]
    out = OUTDIR / XLSX.name
    assert out.exists(), "LO produced no output"
    wb = load_workbook(out, data_only=True)
    errors = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    errors.append(f"{ws.title}!{c.coordinate}={c.value}")
    assert not errors, errors[:10]
    return wb


def test_recalc_values(gt):
    wb = recalc()
    so = wb["Sales_Overview"]
    for i, m in enumerate(MONTHS, start=5):
        assert so.cell(row=i, column=1).value == m
        g = gt["monthly"][m]
        approx(so.cell(row=i, column=2).value, g["orders"], 0)
        approx(so.cell(row=i, column=3).value, g["units"], 0)
        approx(so.cell(row=i, column=4).value, g["rev"], TOL_MONEY)
        approx(so.cell(row=i, column=5).value, g["margin"], TOL_MONEY)
        approx(so.cell(row=i, column=6).value,
               g["margin"] / g["rev"] if g["rev"] else 0, TOL_RATIO)
        approx(so.cell(row=i, column=7).value, gt["targets"][m], 0)
        approx(so.cell(row=i, column=8).value,
               g["rev"] / gt["targets"][m], TOL_RATIO)
        for j, ch in enumerate(["Distributor", "Retail", "Institutional",
                                "Online"], start=10):
            approx(so.cell(row=i, column=j).value, g["chan"][ch], TOL_MONEY)
    inv_ws = wb["Raw_Inventory"]
    for i, p in enumerate(read_csv("products.csv"), start=5):
        sku = p["sku"]
        g = gt["inv"][sku]
        assert inv_ws.cell(row=i, column=1).value == sku
        approx(inv_ws.cell(row=i, column=3).value, g["opening"], 0)
        approx(inv_ws.cell(row=i, column=4).value, g["recv"], 0)
        approx(inv_ws.cell(row=i, column=5).value, g["sold"], 0)
        approx(inv_ws.cell(row=i, column=6).value, g["closing"], 0)
        assert inv_ws.cell(row=i, column=8).value == g["status"], sku
        approx(inv_ws.cell(row=i, column=9).value, g["value"], TOL_MONEY)
        approx(inv_ws.cell(row=i, column=10).value, g["cover"], 0.05)
    ra = wb["Receivables_Ageing"]
    for i, b in enumerate(["Current", "0-30", "31-60", "61-90", "90+",
                           "Paid"], start=5):
        assert ra.cell(row=i, column=1).value == b
        approx(ra.cell(row=i, column=2).value, gt["buckets"][b]["n"], 0)
        approx(ra.cell(row=i, column=3).value, gt["buckets"][b]["out"],
               TOL_MONEY)
    for k, (cust, out, dpd, _idx) in enumerate(gt["top5"], start=15):
        assert ra.cell(row=k, column=2).value == cust, (k, cust)
        approx(ra.cell(row=k, column=3).value, out, TOL_MONEY)
        approx(ra.cell(row=k, column=4).value, dpd, 0)
    cat = wb["Category"]
    for i, cname in enumerate(["Staples", "Pulses", "Edible Oil", "Spices"],
                              start=5):
        g = gt["cats"][cname]
        assert cat.cell(row=i, column=1).value == cname
        approx(cat.cell(row=i, column=2).value, g["rev"], TOL_MONEY)
        approx(cat.cell(row=i, column=3).value, g["orders"], 0)
        approx(cat.cell(row=i, column=4).value, g["margin"], TOL_MONEY)
    cu = wb["Customers"]
    for i, c in enumerate(read_csv("customers.csv"), start=5):
        cname = c["customer"]
        g = gt["custs"][cname]
        assert cu.cell(row=i, column=1).value == cname
        approx(cu.cell(row=i, column=3).value, g["orders"], 0)
        approx(cu.cell(row=i, column=4).value, g["rev"], TOL_MONEY)
        approx(cu.cell(row=i, column=5).value, g["out"], TOL_MONEY)
    db = wb["Dashboard"]
    k = gt["kpis"]
    checks = [
        (5, k["rev"], TOL_MONEY), (6, k["orders"], 0),
        (7, k["aov"], TOL_MONEY), (8, k["margin"], TOL_MONEY),
        (9, k["margin_pct"], TOL_RATIO), (10, k["nov_rev"], TOL_MONEY),
        (11, k["nov_pacing"], TOL_RATIO),
        (14, k["stock_value"], TOL_MONEY), (15, k["alerts"], 0),
        (16, k["stockout"], 0), (17, k["on_time_pct"], TOL_RATIO),
        (18, k["avg_delay"], 0.05), (19, k["calls"], 0),
        (20, k["conv"], TOL_RATIO), (21, k["outstanding"], TOL_MONEY),
        (22, k["over90"], TOL_MONEY), (23, k["avg_cover"], 0.05),
    ]
    for row, exp, tol in checks:
        approx(db.cell(row=row, column=2).value, exp, tol)
    assert db.cell(row=12, column=2).value == k["best_month"]
    assert db.cell(row=13, column=2).value == k["top_cust"]
    ch = wb["Checks"]
    for i in range(5, 15):
        assert ch.cell(row=i, column=4).value == "PASS", \
            f"check row {i}: {ch.cell(row=i, column=1).value} = " \
            f"{ch.cell(row=i, column=2).value}"
    print("  LO recalc values OK (overview, inventory, ageing, category, "
          "customers, 19 KPIs, 10 checks PASS)")


def main():
    gt = ground_truth()
    test_rebuild_and_formulas()
    test_db_views(gt)
    test_recalc_values(gt)
    print("ALL TESTS PASSED")


if __name__ == "__main__":
    main()
