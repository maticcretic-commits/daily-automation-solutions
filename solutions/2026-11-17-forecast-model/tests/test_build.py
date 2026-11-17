#!/usr/bin/env python3
"""Tests for ForecastForge.

Runs generate.py, then verifies the workbook two ways:
1. Formula presence — key cells in the generated xlsx are live formulas,
   the Dashboard scenario selector defaults to Base, 2 charts exist.
2. LibreOffice headless recalc vs an independent Python re-implementation
   of the model from data/base_inputs.csv + data/scenarios.csv: the full
   P&L, cash flow, balance sheet and dashboard for the Base scenario,
   then the same for Pessimistic (scenario switched via Dashboard!B3),
   zero formula errors anywhere, and all 10 Checks reading PASS.
"""
import csv
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data"
XLSX = BASE / "forecast-model.xlsx"

TOL = 0.01
PCT_TOL = 0.0001


def load_csvs():
    with open(DATA / "base_inputs.csv", newline="") as f:
        vals = {r["key"]: float(r["value"]) for r in csv.DictReader(f)}
    with open(DATA / "scenarios.csv", newline="") as f:
        scenarios = {r["scenario"]: {k: float(v) for k, v in r.items()
                                     if k != "scenario"}
                     for r in csv.DictReader(f)}
    return vals, scenarios


def model(vals, s):
    """Independent recomputation of the whole forecast."""
    v = vals
    rev = [v["base_revenue_fy26"]]
    for g in (s["growth_y1"], s["growth_y2"], s["growth_y3"]):
        rev.append(rev[-1] * (1 + g))
    cogs = [v["cogs_fy26"]] + [rev[i + 1] * (1 - s["gross_margin"])
                               for i in range(3)]
    gp = [r - c for r, c in zip(rev, cogs)]
    opex_lines = []
    for key in ("salaries_fy26", "rent_fy26", "marketing_fy26",
                "admin_fy26"):
        line = [v[key]]
        for _ in range(3):
            line.append(line[-1] * (1 + s["opex_growth"]))
        opex_lines.append(line)
    opex = [sum(line[i] for line in opex_lines) for i in range(4)]
    ebitda = [g - o for g, o in zip(gp, opex)]
    nfa = [v["opening_net_fixed_assets"]]
    dep = [v["depreciation_fy26"]]
    for _ in range(3):
        dep.append(v["depreciation_rate"] * (nfa[-1] + v["annual_capex"]))
        nfa.append(nfa[-1] + v["annual_capex"] - dep[-1])
    debt = [v["opening_debt"]]
    for _ in range(3):
        debt.append(debt[-1] - v["annual_debt_repayment"])
    interest = [v["interest_fy26"]] + [v["interest_rate"] * debt[i]
                                        for i in range(3)]
    pbt = [e - d - i for e, d, i in zip(ebitda, dep, interest)]
    tax = [max(0.0, p) * v["tax_rate"] for p in pbt]
    ni = [p - t for p, t in zip(pbt, tax)]
    rec = [v["opening_receivables"]] + \
        [v["receivables_pct_revenue"] * rev[i + 1] for i in range(3)]
    inv = [v["opening_inventory"]] + \
        [v["inventory_pct_cogs"] * cogs[i + 1] for i in range(3)]
    pay = [v["opening_payables"]] + \
        [v["payables_pct_cogs"] * cogs[i + 1] for i in range(3)]
    cf = [ni[i + 1] + dep[i + 1] - (rec[i + 1] - rec[i])
          - (inv[i + 1] - inv[i]) + (pay[i + 1] - pay[i])
          - v["annual_capex"] - v["annual_debt_repayment"]
          for i in range(3)]
    cash = [v["opening_cash"]]
    for f in cf:
        cash.append(cash[-1] + f)
    re = [v["opening_retained_earnings"]]
    for i in range(3):
        re.append(re[-1] + ni[i + 1])
    assets = [cash[i] + rec[i] + inv[i] + nfa[i] for i in range(4)]
    le = [pay[i] + debt[i] + v["share_capital"] + re[i] for i in range(4)]
    check = [a - li for a, li in zip(assets, le)]
    return {
        "rev": rev, "cogs": cogs, "gp": gp, "opex_lines": opex_lines,
        "opex": opex, "ebitda": ebitda, "dep": dep, "nfa": nfa,
        "debt": debt, "interest": interest, "pbt": pbt, "tax": tax,
        "ni": ni, "rec": rec, "inv": inv, "pay": pay, "cf": cf,
        "cash": cash, "re": re, "assets": assets, "le": le,
        "check": check,
    }


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


def verify_model(rw, gt, label):
    """Compare a recalculated workbook against ground truth `gt`."""
    def val(sheet, cell):
        return rw[sheet][cell].value

    # --- P&L: rows 5..18, cols B..E (FY26..FY29) ---
    pl_rows = {
        5: gt["rev"], 6: gt["cogs"], 7: gt["gp"],
        8: gt["opex_lines"][0], 9: gt["opex_lines"][1],
        10: gt["opex_lines"][2], 11: gt["opex_lines"][3],
        12: gt["opex"], 13: gt["ebitda"], 14: gt["dep"],
        15: gt["interest"], 16: gt["pbt"], 17: gt["tax"], 18: gt["ni"],
    }
    pl = rw["P&L"]
    for r, series in pl_rows.items():
        for j, expected in enumerate(series):
            assert_close(pl.cell(row=r, column=2 + j).value, expected,
                         f"{label} P&L row {r} col {j}")

    # --- CashFlow: rows 5..14, cols B..D (FY27..FY29) ---
    cf = rw["CashFlow"]
    for j in range(3):
        assert_close(cf.cell(row=5, column=2 + j).value, gt["ni"][j + 1],
                     f"{label} CF net income FY{27 + j}")
        assert_close(cf.cell(row=6, column=2 + j).value, gt["dep"][j + 1],
                     f"{label} CF depreciation FY{27 + j}")
        assert_close(cf.cell(row=12, column=2 + j).value, gt["cf"][j],
                     f"{label} CF net cash flow FY{27 + j}")
        assert_close(cf.cell(row=13, column=2 + j).value, gt["cash"][j],
                     f"{label} CF opening cash FY{27 + j}")
        assert_close(cf.cell(row=14, column=2 + j).value, gt["cash"][j + 1],
                     f"{label} CF closing cash FY{27 + j}")

    # --- BalanceSheet: rows 5..15, cols B..E ---
    bs = rw["BalanceSheet"]
    bs_rows = {
        5: gt["cash"], 6: gt["rec"], 7: gt["inv"], 8: gt["nfa"],
        9: gt["assets"], 10: gt["pay"], 11: gt["debt"],
        13: gt["re"], 14: gt["le"], 15: gt["check"],
    }
    for r, series in bs_rows.items():
        for j, expected in enumerate(series):
            assert_close(bs.cell(row=r, column=2 + j).value, expected,
                         f"{label} BS row {r} col {j}")
    # Share capital is flat
    for j in range(4):
        assert_close(bs.cell(row=12, column=2 + j).value, 1000000,
                     f"{label} BS share capital col {j}")

    # --- Dashboard KPIs (Base/Pessimistic FY29) ---
    db = rw["Dashboard"]
    assert_close(val("Dashboard", "B7"), gt["rev"][3],
                 f"{label} Dashboard revenue FY29")
    assert_close(val("Dashboard", "B8"), gt["ebitda"][3],
                 f"{label} Dashboard EBITDA FY29")
    assert_close(val("Dashboard", "B9"), gt["ni"][3],
                 f"{label} Dashboard net income FY29")
    assert_close(val("Dashboard", "B10"), gt["cash"][3],
                 f"{label} Dashboard closing cash FY29")
    assert_close(val("Dashboard", "B11"), gt["ebitda"][3] / gt["rev"][3],
                 f"{label} Dashboard EBITDA margin", PCT_TOL)
    assert_close(val("Dashboard", "B12"),
                 (gt["rev"][3] / gt["rev"][0]) ** (1 / 3) - 1,
                 f"{label} Dashboard CAGR", PCT_TOL)
    assert_close(val("Dashboard", "B13"), sum(gt["ni"][1:]),
                 f"{label} Dashboard 3yr net income")

    # --- Dashboard summary table rows 16..19 ---
    for i in range(4):
        assert_close(db.cell(row=16 + i, column=2).value, gt["rev"][i],
                     f"{label} summary revenue row {i}")
        assert_close(db.cell(row=16 + i, column=3).value, gt["ebitda"][i],
                     f"{label} summary EBITDA row {i}")
        assert_close(db.cell(row=16 + i, column=4).value, gt["ni"][i],
                     f"{label} summary net income row {i}")
        assert_close(db.cell(row=16 + i, column=5).value, gt["cash"][i],
                     f"{label} summary closing cash row {i}")

    # --- Checks: all 10 PASS ---
    ck = rw["Checks"]
    for row in range(5, 15):
        assert ck.cell(row=row, column=5).value == "PASS", \
            f"{label} check row {row} " \
            f"({ck.cell(row=row, column=2).value}) = " \
            f"{ck.cell(row=row, column=5).value}"


def test_build():
    r = subprocess.run([sys.executable, "generate.py"], cwd=BASE,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert XLSX.exists(), "xlsx not built"
    vals, scenarios = load_csvs()
    gt_base = model(vals, scenarios["Base"])
    gt_pess = model(vals, scenarios["Pessimistic"])

    # --- formula presence ---
    wb = openpyxl.load_workbook(XLSX, data_only=False)
    assert wb.sheetnames == ["Guide", "Assumptions", "P&L", "CashFlow",
                             "BalanceSheet", "Dashboard",
                             "Checks"], wb.sheetnames
    assert wb["Dashboard"]["B3"].value == "Base"
    assert str(wb["Assumptions"]["B35"].value) == "=Dashboard!B3"
    assert str(wb["Assumptions"]["B36"].value).startswith("=INDEX(")
    assert str(wb["P&L"]["C5"].value) == "=B5*(1+Assumptions!$B$36)"
    assert str(wb["P&L"]["C6"].value) == "=C5*(1-Assumptions!$B$39)"
    assert str(wb["P&L"]["C14"].value) == \
        "=Assumptions!B22*(BalanceSheet!B8+Assumptions!B19)"
    assert str(wb["CashFlow"]["B7"].value).startswith(
        "=-(BalanceSheet!C6")
    assert str(wb["BalanceSheet"]["C8"].value).startswith(
        "=B8+Assumptions!")
    assert str(wb["BalanceSheet"]["C15"].value) == "=C9-C14"
    assert str(wb["Dashboard"]["B11"].value).startswith("=IF(")
    assert str(wb["Checks"]["D5"].value) == "=BalanceSheet!B15"
    assert wb["Checks"]["C5"].value == 0, "Expected must be numeric"
    assert len(wb["Dashboard"]._charts) == 2, "expected 2 charts"
    # Pessimistic really does go cash-negative (the documented gap)
    assert gt_pess["cash"][2] < 0, gt_pess["cash"]
    assert gt_base["cash"][3] > 0, gt_base["cash"]

    # --- Base scenario recalc ---
    rw = recalc(XLSX)
    check_no_errors(rw)
    verify_model(rw, gt_base, "Base")

    # --- Pessimistic scenario recalc (switch via Dashboard!B3) ---
    with tempfile.TemporaryDirectory() as tmp:
        pess_path = Path(tmp) / "forecast-model-pess.xlsx"
        shutil.copy(XLSX, pess_path)
        wb2 = openpyxl.load_workbook(pess_path)
        wb2["Dashboard"]["B3"] = "Pessimistic"
        wb2.save(pess_path)
        rw2 = recalc(pess_path)
    check_no_errors(rw2)
    assert rw2["Dashboard"]["B4"].value == "Pessimistic"
    verify_model(rw2, gt_pess, "Pessimistic")

    print(f"OK: Base FY29 revenue Rs {gt_base['rev'][3]:,.0f} "
          f"EBITDA Rs {gt_base['ebitda'][3]:,.0f} "
          f"net income Rs {gt_base['ni'][3]:,.0f} "
          f"closing cash Rs {gt_base['cash'][3]:,.0f}; "
          f"Pessimistic FY28 closing cash Rs {gt_pess['cash'][2]:,.0f} "
          f"(funding gap, as documented)")


if __name__ == "__main__":
    test_build()
    print("ALL TESTS PASS")
