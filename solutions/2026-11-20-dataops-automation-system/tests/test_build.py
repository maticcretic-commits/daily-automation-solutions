#!/usr/bin/env python3
"""Tests for DataOps Automation System.

Runs generate.py, then verifies the workbook two ways:
1. Formula presence — derived columns on raw sheets, SUMIFS Monthly,
   SQL cross-check columns, Dashboard KPIs + dropdown validation,
   3 charts exist.
2. LibreOffice headless recalc vs an independent Python re-implementation
   from data/*.csv: monthly revenue/COGS/gross profit/purchases/expenses/
   net profit/orders, SQL cross-check diffs, Dashboard KPIs for the
   default focus month (2026-10), category totals, and all 10 Checks
   reading PASS.
"""
import csv
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data"
XLSX = BASE / "dataops-automation-system.xlsx"

TOL = 0.01


def load(name):
    with open(DATA / name, newline="") as f:
        return list(csv.DictReader(f))


def ground_truth():
    sales = load("sales.csv")
    purch = load("purchases.csv")
    exp = load("expenses.csv")
    sql = load("sql_summary.csv")
    months = sorted({r["date"][:7] for r in sales})
    m = {}
    for mo in months:
        s = [r for r in sales if r["date"].startswith(mo)]
        rev = sum(int(r["qty"]) * float(r["unit_price"]) for r in s)
        cogs = sum(int(r["qty"]) * float(r["unit_cost"]) for r in s)
        p = [r for r in purch if r["date"].startswith(mo)]
        pur = sum(int(r["qty"]) * float(r["unit_cost"]) for r in p)
        e = [r for r in exp if r["date"].startswith(mo)]
        ex = sum(float(r["amount"]) for r in e)
        m[mo] = {"rev": rev, "cogs": cogs, "gp": rev - cogs,
                 "purch": pur, "exp": ex, "net": rev - cogs - ex,
                 "orders": len(s)}
    cats = {}
    for r in sales:
        cats[r["category"]] = cats.get(r["category"], 0) + \
            int(r["qty"]) * float(r["unit_price"])
    chans = {}
    for r in sales:
        chans[r["channel"]] = chans.get(r["channel"], 0) + \
            int(r["qty"]) * float(r["unit_price"])
    sql_gt = {r["month"]: r for r in sql}
    return {"months": months, "m": m, "cats": cats, "chans": chans,
            "sql": sql_gt, "n_sales": len(sales), "n_purch": len(purch),
            "n_exp": len(exp)}


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


def check_no_errors(rw):
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


def assert_close(actual, expected, label, tol=TOL):
    assert actual is not None, f"{label} is empty after recalc"
    assert abs(actual - expected) < tol, \
        f"{label}: got {actual}, expected {expected}"


def test_build():
    r = subprocess.run([sys.executable, "generate.py"], cwd=BASE,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert XLSX.exists(), "xlsx not built"
    gt = ground_truth()
    months = gt["months"]
    assert len(months) == 7, months
    assert gt["n_sales"] == 252
    assert gt["n_purch"] == 126
    assert gt["n_exp"] == 42

    # SQL summary matches the independent Python computation.
    for mo in months:
        s = gt["sql"][mo]
        assert_close(float(s["sql_revenue"]), gt["m"][mo]["rev"],
                     f"SQL revenue {mo}")
        assert_close(float(s["sql_cogs"]), gt["m"][mo]["cogs"],
                     f"SQL cogs {mo}")
        assert_close(float(s["sql_gross_profit"]), gt["m"][mo]["gp"],
                     f"SQL gp {mo}")
        assert int(s["sql_orders"]) == gt["m"][mo]["orders"], \
            f"SQL orders {mo}"
        assert_close(float(s["sql_purchases"]), gt["m"][mo]["purch"],
                     f"SQL purchases {mo}")
        assert_close(float(s["sql_expenses"]), gt["m"][mo]["exp"],
                     f"SQL expenses {mo}")

    # --- formula presence ---
    wb = openpyxl.load_workbook(XLSX, data_only=False)
    assert wb.sheetnames == ["Guide", "Sales", "Purchases", "Expenses",
                             "Monthly", "Dashboard", "Checks"], wb.sheetnames
    s = wb["Sales"]
    assert str(s["J2"].value) == "=E2*F2"
    assert str(s["K2"].value) == "=E2*G2"
    assert str(s["L2"].value) == "=J2-K2"
    assert str(s["M2"].value) == "=LEFT(B2,7)"
    p = wb["Purchases"]
    assert str(p["H2"].value) == "=F2*G2"
    assert str(p["I2"].value) == "=LEFT(B2,7)"
    e = wb["Expenses"]
    assert str(e["E2"].value) == "=LEFT(A2,7)"
    mo = wb["Monthly"]
    assert str(mo["B2"].value).startswith("=SUMIFS(Sales!$J$2:$J$253,")
    assert str(mo["D2"].value) == "=B2-C2"
    assert str(mo["G2"].value) == "=D2-F2"
    assert str(mo["K2"].value) == "=B2-I2"
    assert str(mo["L2"].value) == "=D2-J2"
    db = wb["Dashboard"]
    assert db["B2"].value == "2026-10", "default focus month"
    assert len(db.data_validations.dataValidation) == 1
    assert str(db["B4"].value).startswith("=SUMIFS(Sales!$J$2:$J$253,")
    assert str(db["B6"].value).startswith("=SUMIFS(Sales!$L$2:$L$253,")
    assert str(db["B7"].value) == "=IF(B4=0,0,B6/B4)"
    assert str(db["B9"].value) == "=B6-B8"
    assert len(db._charts) == 3, "expected 3 charts"
    ck = wb["Checks"]
    assert str(ck["E2"].value).startswith('=IF(ABS(')

    # --- recalc vs ground truth ---
    rw = recalc(XLSX)
    check_no_errors(rw)

    mo_r = rw["Monthly"]
    tot_row = 9
    for j, mo_name in enumerate(months, start=2):
        g = gt["m"][mo_name]
        assert_close(mo_r.cell(row=j, column=2).value, g["rev"],
                     f"Monthly revenue {mo_name}")
        assert_close(mo_r.cell(row=j, column=3).value, g["cogs"],
                     f"Monthly cogs {mo_name}")
        assert_close(mo_r.cell(row=j, column=4).value, g["gp"],
                     f"Monthly gp {mo_name}")
        assert_close(mo_r.cell(row=j, column=5).value, g["purch"],
                     f"Monthly purchases {mo_name}")
        assert_close(mo_r.cell(row=j, column=6).value, g["exp"],
                     f"Monthly expenses {mo_name}")
        assert_close(mo_r.cell(row=j, column=7).value, g["net"],
                     f"Monthly net {mo_name}")
        assert mo_r.cell(row=j, column=8).value == g["orders"], \
            f"Monthly orders {mo_name}"
        assert_close(mo_r.cell(row=j, column=11).value, 0,
                     f"Revenue SQL diff {mo_name}")
        assert_close(mo_r.cell(row=j, column=12).value, 0,
                     f"GP SQL diff {mo_name}")
    total_rev = sum(gt["m"][x]["rev"] for x in months)
    total_gp = sum(gt["m"][x]["gp"] for x in months)
    total_net = sum(gt["m"][x]["net"] for x in months)
    assert_close(mo_r.cell(row=tot_row, column=2).value, total_rev,
                 "Monthly total revenue")
    assert_close(mo_r.cell(row=tot_row, column=7).value, total_net,
                 "Monthly total net")

    # Dashboard KPIs for the default focus month 2026-10.
    db_r = rw["Dashboard"]
    oct_g = gt["m"]["2026-10"]
    assert_close(db_r["B4"].value, oct_g["rev"], "Dashboard Oct revenue")
    assert_close(db_r["B5"].value, oct_g["cogs"], "Dashboard Oct COGS")
    assert_close(db_r["B6"].value, oct_g["gp"], "Dashboard Oct GP")
    assert_close(db_r["B7"].value, oct_g["gp"] / oct_g["rev"],
                 "Dashboard Oct margin", 0.0001)
    assert_close(db_r["B8"].value, oct_g["exp"], "Dashboard Oct expenses")
    assert_close(db_r["B9"].value, oct_g["net"], "Dashboard Oct net")
    assert db_r["B10"].value == oct_g["orders"], "Dashboard Oct orders"

    # Category table on Dashboard (rows after the trend table).
    crow = 14 + len(months) + 2  # header row of the category table
    cats_sorted = sorted(gt["cats"])
    for i, cat in enumerate(cats_sorted):
        row = crow + 1 + i
        assert db_r.cell(row=row, column=1).value == cat, \
            f"category row {row}"
        assert_close(db_r.cell(row=row, column=2).value, gt["cats"][cat],
                     f"category revenue {cat}")

    # --- Checks: all 10 PASS ---
    ck_r = rw["Checks"]
    for row in range(2, 12):
        assert ck_r.cell(row=row, column=5).value == "PASS", \
            f"check row {row} " \
            f"({ck_r.cell(row=row, column=2).value}) = " \
            f"{ck_r.cell(row=row, column=5).value}"

    print(f"OK: revenue {total_rev:,.0f} gross profit {total_gp:,.0f} "
          f"({total_gp / total_rev:.1%}) net {total_net:,.0f} "
          f"Oct revenue {oct_g['rev']:,.0f} Oct net {oct_g['net']:,.0f}")


if __name__ == "__main__":
    test_build()
    print("ALL TESTS PASS")
