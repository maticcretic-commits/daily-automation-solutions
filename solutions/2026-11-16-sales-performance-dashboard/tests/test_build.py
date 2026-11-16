#!/usr/bin/env python3
"""Tests for SalesPerformanceDashboard.

Runs make_data.py + generate.py, then verifies the workbook two ways:
1. Formula presence — key cells in the generated xlsx are live formulas.
2. LibreOffice headless recalc vs an independent Python re-implementation
   from data/orders.csv + data/targets.csv: per-category x month revenue
   and target, dashboard KPIs at default filters (All / All), monthly and
   category tables, top product, and all 10 Checks reading PASS.
"""
import csv
import subprocess
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data"
XLSX = BASE / "sales-performance-dashboard.xlsx"

TOL = 0.01
CATEGORIES = ["Staples", "Beverages", "Personal Care", "Home Care",
              "Snacks"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 11)]
PRODUCTS = ["Atta 10kg", "Basmati Rice 5kg", "Sugar 1kg", "Tea 500g",
            "Instant Coffee 200g", "Bath Soap (pack of 4)",
            "Shampoo 650ml", "Detergent Powder 3kg",
            "Floor Cleaner 1L", "Biscuits 400g", "Namkeen 500g"]


def ground_truth():
    """Independent recomputation straight from the CSVs."""
    with open(DATA / "orders.csv", newline="") as f:
        orders = list(csv.DictReader(f))
    with open(DATA / "targets.csv", newline="") as f:
        targets = list(csv.DictReader(f))
    gt = {"n": len(orders), "monthly": defaultdict(
        lambda: [0.0, 0, 0.0]), "cats": defaultdict(
        lambda: [0.0, 0.0]), "prods": defaultdict(lambda: [0.0, 0]),
          "months": defaultdict(lambda: [0.0, 0.0])}
    gt["revenue"] = sum(float(o["Revenue"]) for o in orders)
    gt["qty"] = sum(int(o["Qty"]) for o in orders)
    gt["target"] = sum(int(t["Target"]) for t in targets)
    for o in orders:
        rev = float(o["Revenue"])
        gt["monthly"][(o["Category"], o["Month"])][0] += rev
        gt["monthly"][(o["Category"], o["Month"])][1] += 1
        gt["cats"][o["Category"]][0] += rev
        gt["prods"][o["Product"]][0] += rev
        gt["prods"][o["Product"]][1] += 1
        gt["months"][o["Month"]][0] += rev
    for t in targets:
        gt["monthly"][(t["Category"], t["Month"])][2] += int(t["Target"])
        gt["cats"][t["Category"]][1] += int(t["Target"])
        gt["months"][t["Month"]][1] += int(t["Target"])
    gt["top_product"] = max(gt["prods"], key=lambda p: gt["prods"][p][0])
    gt["top_revenue"] = gt["prods"][gt["top_product"]][0]
    return gt


def recalc(path):
    """LibreOffice headless recalc; returns a data_only workbook."""
    with tempfile.TemporaryDirectory() as tmp:
        out = subprocess.run(
            ["soffice", "--headless", "--convert-to", "xlsx",
             "--outdir", tmp, str(path)],
            capture_output=True, text=True, timeout=300)
        assert out.returncode == 0, f"soffice failed: {out.stderr[:500]}"
        recalc_path = Path(tmp) / path.name
        assert recalc_path.exists(), "recalc produced no file"
        return openpyxl.load_workbook(recalc_path, data_only=True)


def test_build():
    r = subprocess.run([sys.executable, "make_data.py"], cwd=BASE,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    r = subprocess.run([sys.executable, "generate.py"], cwd=BASE,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert XLSX.exists(), "xlsx not built"
    gt = ground_truth()
    n = gt["n"]
    ld = 4 + n

    # --- formula presence ---
    wb = openpyxl.load_workbook(XLSX, data_only=False)
    assert wb.sheetnames == ["Guide", "Data", "Targets", "Monthly",
                             "Dashboard", "Checks"], wb.sheetnames
    d = wb["Data"]
    assert str(d["I5"].value) == "=G5*H5", d["I5"].value
    assert str(d["J5"].value) == "=LEFT(A5,7)", d["J5"].value
    assert str(d[f"I{ld + 1}"].value) == f"=SUM(I5:I{ld})"
    assert wb["Dashboard"]["B3"].value == "All"
    assert wb["Dashboard"]["B4"].value == "All"
    assert str(wb["Dashboard"]["B7"].value).startswith("=SUMPRODUCT(")
    assert str(wb["Dashboard"]["B8"].value).startswith("=SUMPRODUCT(")
    assert str(wb["Dashboard"]["B9"].value).startswith("=IF(B8=0")
    assert str(wb["Dashboard"]["B12"].value) == "=MAX(B35:B46)"
    assert str(wb["Dashboard"]["B13"].value).startswith("=INDEX(")
    assert str(wb["Dashboard"]["B16"].value).startswith("=SUMPRODUCT(")
    assert str(wb["Monthly"]["C5"].value).startswith("=SUMIFS(")
    assert str(wb["Monthly"]["E5"].value).startswith('=IF(D5=0')
    assert str(wb["Checks"]["D5"].value).startswith("=COUNTA(")
    assert len(wb["Dashboard"]._charts) == 2, "expected 2 charts"

    # --- recalc vs ground truth ---
    rw = recalc(XLSX)

    def val(sheet, cell):
        v = rw[sheet][cell].value
        assert v is not None, f"{sheet}!{cell} is empty after recalc"
        return v

    ERRORS = ("#VALUE!", "#DIV/0!", "#NAME?", "#REF!", "#N/A", "#NULL!",
              "Err:")
    for ws in rw.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and \
                        c.value.startswith(ERRORS):
                    raise AssertionError(
                        f"formula error in {ws.title}!{c.coordinate}: "
                        f"{c.value}")

    # Data totals
    assert abs(val("Data", f"I{ld + 1}") - gt["revenue"]) < TOL
    assert val("Data", f"G{ld + 1}") == gt["qty"]

    # Targets total
    assert abs(val("Targets", "C55") - gt["target"]) < TOL

    # Monthly: January Staples and October Snacks
    mo = rw["Monthly"]
    # rows are category-major: Staples Jan = row 5, Snacks Oct = row 54
    assert abs(mo.cell(row=5, column=3).value
               - gt["monthly"][("Staples", "2026-01")][0]) < TOL
    assert abs(mo.cell(row=5, column=4).value
               - gt["monthly"][("Staples", "2026-01")][2]) < TOL
    assert mo.cell(row=5, column=6).value == \
        gt["monthly"][("Staples", "2026-01")][1]
    assert abs(mo.cell(row=54, column=3).value
               - gt["monthly"][("Snacks", "2026-10")][0]) < TOL
    assert abs(mo.cell(row=54, column=4).value
               - gt["monthly"][("Snacks", "2026-10")][2]) < TOL
    assert abs(val("Monthly", "C55") - gt["revenue"]) < TOL
    assert abs(val("Monthly", "D55") - gt["target"]) < TOL
    assert val("Monthly", "F55") == n

    # Dashboard KPIs at default filters (All / All)
    db = rw["Dashboard"]
    assert abs(val("Dashboard", "B7") - gt["revenue"]) < TOL
    assert abs(val("Dashboard", "B8") - gt["target"]) < TOL
    assert abs(val("Dashboard", "B9") - gt["revenue"] / gt["target"]) < 0.001
    assert val("Dashboard", "B10") == n
    assert abs(val("Dashboard", "B11") - gt["revenue"] / n) < TOL
    assert abs(val("Dashboard", "B12") - gt["top_revenue"]) < TOL
    assert val("Dashboard", "B13") == gt["top_product"]

    # Monthly table: January and October rows
    assert abs(val("Dashboard", "B16")
               - gt["months"]["2026-01"][0]) < TOL
    assert abs(val("Dashboard", "C16")
               - gt["months"]["2026-01"][1]) < TOL
    assert abs(val("Dashboard", "B25")
               - gt["months"]["2026-10"][0]) < TOL
    assert abs(val("Dashboard", "C25")
               - gt["months"]["2026-10"][1]) < TOL

    # Category table sums to dashboard revenue; spot-check Staples
    cat_rev = sum(val("Dashboard", f"B{r}") for r in range(28, 33))
    assert abs(cat_rev - gt["revenue"]) < TOL
    assert abs(val("Dashboard", "B28") - gt["cats"]["Staples"][0]) < TOL
    assert abs(val("Dashboard", "C28") - gt["cats"]["Staples"][1]) < TOL

    # Product table: every product ties to ground truth
    for i, prod in enumerate(PRODUCTS):
        r = 35 + i
        assert db.cell(row=r, column=1).value == prod
        assert abs(db.cell(row=r, column=2).value
                   - gt["prods"][prod][0]) < TOL, prod
        assert db.cell(row=r, column=3).value == gt["prods"][prod][1], prod

    # Checks: all 10 PASS
    ck = rw["Checks"]
    for row in range(5, 15):
        assert ck.cell(row=row, column=5).value == "PASS", \
            f"check row {row} ({ck.cell(row=row, column=2).value}) = " \
            f"{ck.cell(row=row, column=5).value}"

    # orders.csv ties to the same totals
    with open(DATA / "orders.csv", newline="") as f:
        rows = list(csv.DictReader(f))
    assert len(rows) == n
    assert abs(sum(float(x["Revenue"]) for x in rows)
               - gt["revenue"]) < TOL

    print(f"OK: {n} orders revenue Rs {gt['revenue']:,.2f} "
          f"target Rs {gt['target']:,.2f} "
          f"({gt['revenue'] / gt['target']:.1%} to goal) "
          f"top product {gt['top_product']} "
          f"Rs {gt['top_revenue']:,.2f}")


if __name__ == "__main__":
    test_build()
    print("ALL TESTS PASS")
