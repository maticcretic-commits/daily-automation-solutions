#!/usr/bin/env python3
"""Tests for SalesDrill (build-2026-11-09).

- Rebuilds the workbook from generate.py (fresh).
- Regex-scans every formula for short-form range literals
  ($H$2:$5000 — invalid in LibreOffice; must repeat the column letter).
- Recalculates with LibreOffice headless and compares every Monthly row
  (revenue, orders, units, avg order, margin, margin %, MoM), every
  Category row, every Customer row and every Dashboard KPI against an
  independent Python ground truth computed from data/*.csv.
- Asserts all Checks read PASS and no cell carries a formula error.
"""
import csv
import datetime
import re
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
OUTDIR = Path("/tmp/salesdrill-recalc-20261109")
XLSX = BASE / "sales-drilldown-dashboard.xlsx"
SHORT_RANGE = re.compile(r"\$[A-Z]+\$\d+:\$\d+")
ERRORS = {"#VALUE!", "#NAME?", "#REF!", "#DIV/0!", "#N/A", "#NUM!",
          "Err:508", "Err:509", "Err:510", "Err:511", "Err:512", "Err:513",
          "Err:514", "Err:515", "Err:516", "Err:517", "Err:518", "Err:519",
          "Err:520", "Err:521", "Err:522", "Err:523", "Err:524", "Err:525",
          "Err:526", "Err:527", "Err:528", "Err:529", "Err:530", "Err:531",
          "Err:532"}


def read_csv(name):
    with open(BASE / "data" / name, newline="") as f:
        return list(csv.DictReader(f))


def ground_truth():
    orders = read_csv("orders.csv")
    products = {p["product"]: p for p in read_csv("products.csv")}
    customers = read_csv("customers.csv")
    months = ["Apr 2026", "May 2026", "Jun 2026", "Jul 2026",
              "Aug 2026", "Sep 2026"]
    monthly = {m: {"rev": 0, "orders": 0, "units": 0, "margin": 0}
               for m in months}
    cats = {}
    custs = {}
    total_rev = 0
    total_margin = 0
    for o in orders:
        d = datetime.date.fromisoformat(o["date"])
        m = d.strftime("%b %Y")
        qty = int(o["qty"])
        price = int(o["unit_price"])
        cost = int(products[o["product"]]["cost_price"])
        amt = qty * price
        marg = qty * (price - cost)
        monthly[m]["rev"] += amt
        monthly[m]["orders"] += 1
        monthly[m]["units"] += qty
        monthly[m]["margin"] += marg
        c = cats.setdefault(o["category"],
                            {"rev": 0, "orders": 0, "units": 0, "margin": 0})
        c["rev"] += amt
        c["orders"] += 1
        c["units"] += qty
        c["margin"] += marg
        u = custs.setdefault(o["customer"], {"orders": 0, "rev": 0})
        u["orders"] += 1
        u["rev"] += amt
        total_rev += amt
        total_margin += marg
    total_orders = len(orders)
    best_month = max(months, key=lambda m: monthly[m]["rev"])
    top_cat = max(cats, key=lambda c: cats[c]["rev"])
    top_cust = max(custs, key=lambda c: custs[c]["rev"])
    repeat = sum(1 for u in custs.values() if u["orders"] > 1) / len(custs)
    kpis = {
        "rev": total_rev,
        "orders": total_orders,
        "aov": total_rev / total_orders,
        "margin": total_margin,
        "margin_pct": total_margin / total_rev,
        "best_month": best_month,
        "best_month_rev": monthly[best_month]["rev"],
        "top_cat": top_cat,
        "top_cust": top_cust,
        "repeat": repeat,
    }
    return monthly, cats, custs, kpis


def close(a, b, tol):
    return a is not None and b is not None and abs(a - b) <= tol


def main():
    failures = []

    # Rebuild fresh.
    r = subprocess.run([sys.executable, str(BASE / "generate.py")],
                       capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode != 0:
        print(r.stderr)
        failures.append("generate.py failed")

    # Short-form range scan over every formula cell.
    wb = load_workbook(XLSX)
    for sh in wb.sheetnames:
        ws = wb[sh]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    if SHORT_RANGE.search(cell.value):
                        failures.append(
                            f"short-form range in {sh}!{cell.coordinate}: "
                            f"{cell.value}")

    # LibreOffice headless recalc.
    OUTDIR.mkdir(parents=True, exist_ok=True)
    for old in OUTDIR.glob("*.xlsx"):
        old.unlink()
    r = subprocess.run(
        ["soffice", "--headless", "--calc", "--convert-to", "xlsx",
         "--outdir", str(OUTDIR), str(XLSX)],
        capture_output=True, text=True, timeout=180)
    recalc = OUTDIR / XLSX.name
    if not recalc.exists():
        print(r.stdout[-2000:])
        print(r.stderr[-2000:])
        failures.append("LibreOffice recalc produced no output")
        print(f"FAILURES: {len(failures)}")
        for f in failures:
            print(" -", f)
        sys.exit(1)
    rb = load_workbook(recalc, data_only=True)

    # No formula errors anywhere.
    for sh in rb.sheetnames:
        ws = rb[sh]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value in ERRORS:
                    failures.append(f"formula error {sh}!{cell.coordinate}: "
                                    f"{cell.value}")

    monthly, cats, custs, kpis = ground_truth()
    months = ["Apr 2026", "May 2026", "Jun 2026", "Jul 2026",
              "Aug 2026", "Sep 2026"]

    ws = rb["Monthly"]
    prev_rev = None
    for i, m in enumerate(months):
        row = i + 2
        got = [ws.cell(row=row, column=c).value for c in range(1, 10)]
        gt = monthly[m]
        exp_rev, exp_ord, exp_units = gt["rev"], gt["orders"], gt["units"]
        exp_aov = exp_rev / exp_ord
        exp_marg = gt["margin"]
        exp_mpct = exp_marg / exp_rev
        exp_mom = 0 if prev_rev in (None, 0) else (exp_rev - prev_rev) / prev_rev
        checks = [
            ("label", got[0], m, got[0] == m),
            ("revenue", got[1], exp_rev, close(got[1], exp_rev, 0.01)),
            ("orders", got[2], exp_ord, got[2] == exp_ord),
            ("units", got[3], exp_units, got[3] == exp_units),
            ("aov", got[4], exp_aov, close(got[4], exp_aov, 0.01)),
            ("margin", got[5], exp_marg, close(got[5], exp_marg, 0.01)),
            ("margin%", got[6], exp_mpct, close(got[6], exp_mpct, 1e-6)),
            ("mom", got[8], exp_mom, close(got[8], exp_mom, 1e-6)),
        ]
        for name, g, e, ok in checks:
            if not ok:
                failures.append(f"Monthly {m} {name}: got {g}, want {e}")
        prev_rev = exp_rev
    # Monthly TOTAL row.
    t = len(months) + 2
    tot = [ws.cell(row=t, column=c).value for c in range(1, 8)]
    if not close(tot[1], kpis["rev"], 0.01):
        failures.append(f"Monthly TOTAL revenue: got {tot[1]}, want {kpis['rev']}")
    if tot[2] != kpis["orders"]:
        failures.append(f"Monthly TOTAL orders: got {tot[2]}, want {kpis['orders']}")
    if not close(tot[5], kpis["margin"], 0.01):
        failures.append(f"Monthly TOTAL margin: got {tot[5]}, want {kpis['margin']}")

    ws = rb["Category"]
    for i, cat in enumerate(sorted(cats)):
        row = i + 2
        got = [ws.cell(row=row, column=c).value for c in range(1, 8)]
        gt = cats[cat]
        if got[0] != cat:
            failures.append(f"Category row {row} label: got {got[0]}, want {cat}")
        if not close(got[1], gt["rev"], 0.01):
            failures.append(f"Category {cat} revenue: got {got[1]}, want {gt['rev']}")
        if got[2] != gt["orders"]:
            failures.append(f"Category {cat} orders: got {got[2]}, want {gt['orders']}")
        if got[3] != gt["units"]:
            failures.append(f"Category {cat} units: got {got[3]}, want {gt['units']}")
        if not close(got[4], gt["margin"], 0.01):
            failures.append(f"Category {cat} margin: got {got[4]}, want {gt['margin']}")
        if not close(got[5], gt["rev"] / kpis["rev"], 1e-6):
            failures.append(f"Category {cat} share: got {got[5]}, "
                            f"want {gt['rev'] / kpis['rev']}")

    ws = rb["Customers"]
    cust_sorted = sorted(custs, key=lambda c: -custs[c]["rev"])
    for i, name in enumerate(cust_sorted):
        row = i + 2
        got = [ws.cell(row=row, column=c).value for c in range(1, 7)]
        gt = custs[name]
        if got[0] != name:
            failures.append(f"Customers row {row} label: got {got[0]}, want {name}")
        if got[2] != gt["orders"]:
            failures.append(f"Customers {name} orders: got {got[2]}, want {gt['orders']}")
        if not close(got[3], gt["rev"], 0.01):
            failures.append(f"Customers {name} revenue: got {got[3]}, want {gt['rev']}")
        if not close(got[4], gt["rev"] / kpis["rev"], 1e-6):
            failures.append(f"Customers {name} share: got {got[4]}")

    ws = rb["Dashboard"]
    dk = [(4, kpis["rev"], 0.01, "Total Revenue"),
          (5, kpis["orders"], 0, "Total Orders"),
          (6, kpis["aov"], 0.01, "Avg Order Value"),
          (7, kpis["margin"], 0.01, "Total Margin"),
          (8, kpis["margin_pct"], 1e-6, "Margin %"),
          (10, kpis["best_month_rev"], 0.01, "Best Month Revenue"),
          (13, kpis["repeat"], 1e-6, "Repeat-Customer %")]
    for row, exp, tol, name in dk:
        got = ws.cell(row=row, column=2).value
        ok = (got == exp) if tol == 0 else close(got, exp, tol)
        if not ok:
            failures.append(f"Dashboard {name}: got {got}, want {exp}")
    for row, exp, name in [(9, kpis["best_month"], "Best Month"),
                           (11, kpis["top_cat"], "Top Category"),
                           (12, kpis["top_cust"], "Top Customer")]:
        got = ws.cell(row=row, column=2).value
        if got != exp:
            failures.append(f"Dashboard {name}: got {got}, want {exp}")

    ws = rb["Checks"]
    for row in range(2, 10):
        res = ws.cell(row=row, column=5).value
        if res != "PASS":
            failures.append(
                f"Check {row - 1} ({ws.cell(row=row, column=2).value}): "
                f"{res} — expected "
                f"{ws.cell(row=row, column=3).value}, actual "
                f"{ws.cell(row=row, column=4).value}")

    n_checks = 6 + 4 + 30 + 10  # monthly cells + category + customers + KPIs
    if failures:
        print(f"FAILURES ({len(failures)} of ~{n_checks} checks):")
        for f in failures:
            print(" -", f)
        sys.exit(1)
    print(f"ALL PASS — Monthly 6/6, Category 4/4, Customers 30/30, "
          f"Dashboard KPIs 10/10, Checks 8/8, no formula errors, "
          f"no short-form ranges. Total revenue "
          f"\u20b9{kpis['rev']:,} · margin \u20b9{kpis['margin']:,} "
          f"({kpis['margin_pct']:.1%}) · best month {kpis['best_month']} "
          f"· top category {kpis['top_cat']} · top customer {kpis['top_cust']} "
          f"· repeat {kpis['repeat']:.1%}")


if __name__ == "__main__":
    main()
