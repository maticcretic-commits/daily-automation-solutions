#!/usr/bin/env python3
"""Tests for LenderMatch (build-2026-11-08).

- Rebuilds the workbook from generate.py (fresh).
- Regex-scans every formula for short-form range literals
  ($H$2:$5000 — invalid in LibreOffice; must repeat the column letter).
- Recalculates with LibreOffice headless and compares every Comparison
  row (eligibility, mid rate, payment, total interest, fees, DSCR,
  score, rank) and every Dashboard KPI against an independent Python
  ground truth computed from data/lenders.csv.
- Asserts all Checks read PASS and no cell carries a formula error.
"""
import csv
import re
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
OUTDIR = Path("/tmp/lendermatch-recalc-20261108")
XLSX = BASE / "lendermatch.xlsx"
SHORT_RANGE = re.compile(r"\$[A-Z]+\$\d+:\$\d+")
ERRORS = {"#VALUE!", "#NAME?", "#REF!", "#DIV/0!", "#N/A", "#NUM!", "Err:508",
          "Err:509", "Err:510", "Err:511", "Err:512", "Err:513", "Err:514",
          "Err:515", "Err:516", "Err:517", "Err:518", "Err:519", "Err:520",
          "Err:521", "Err:522", "Err:523", "Err:524", "Err:525", "Err:526",
          "Err:527", "Err:528", "Err:529", "Err:530", "Err:531", "Err:532"}

PRICE = 2_400_000
EQUITY_PCT = 0.20
LOAN = PRICE * (1 - EQUITY_PCT)
REVENUE = 6_500_000
EBITDA_M = 145_000
CREDIT = 720
OTHER_DEBT_M = 12_000

WEIGHTS = {"rate": 35, "speed": 20, "ltv": 20, "fee": 15, "tenor": 10}


def read_lenders():
    with open(BASE / "data" / "lenders.csv", newline="") as f:
        return list(csv.DictReader(f))


def pmt(rate_monthly, nper, pv):
    return rate_monthly * pv / (1 - (1 + rate_monthly) ** -nper)


def ground_truth():
    lenders = read_lenders()
    mids = [(float(r["min_rate"]) + float(r["max_rate"])) / 2 for r in lenders]
    closes = [int(r["close_days"]) for r in lenders]
    ltvs = [float(r["ltv"]) for r in lenders]
    fees = [float(r["orig_fee"]) for r in lenders]
    tenors = [int(r["tenor_yrs"]) for r in lenders]
    rows = {}
    for i, r in enumerate(lenders):
        eligible = (LOAN <= int(r["max_loan"])
                    and LOAN <= PRICE * float(r["ltv"])
                    and REVENUE >= int(r["min_revenue"])
                    and CREDIT >= int(r["min_credit"]))
        mid = mids[i]
        pay = pmt(mid / 12, int(r["tenor_yrs"]) * 12, LOAN) if eligible else None
        total_int = pay * int(r["tenor_yrs"]) * 12 - LOAN if eligible else None
        up = LOAN * float(r["orig_fee"]) if eligible else None
        dscr = EBITDA_M / (pay + OTHER_DEBT_M) if eligible else None
        if eligible:
            score = (
                WEIGHTS["rate"] * (max(mids) - mid) / (max(mids) - min(mids))
                + WEIGHTS["speed"] * (max(closes) - closes[i]) / (max(closes) - min(closes))
                + WEIGHTS["ltv"] * (ltvs[i] - min(ltvs)) / (max(ltvs) - min(ltvs))
                + WEIGHTS["fee"] * (max(fees) - fees[i]) / (max(fees) - min(fees))
                + WEIGHTS["tenor"] * (tenors[i] - min(tenors)) / (max(tenors) - min(tenors))
            )
        else:
            score = 0.0
        rows[r["lender"]] = {
            "type": r["type"], "eligible": eligible, "mid": mid, "pay": pay,
            "total_int": total_int, "upfront": up, "dscr": dscr,
            "close": closes[i], "score": score,
        }
    scores = sorted((v["score"] for v in rows.values() if v["eligible"]), reverse=True)
    rank_of = {s: k + 1 for k, s in enumerate(scores)}
    for v in rows.values():
        v["rank"] = rank_of[v["score"]] if v["eligible"] else None
    elig = [v for v in rows.values() if v["eligible"]]
    best = max(elig, key=lambda v: v["score"])
    best_name = [k for k, v in rows.items() if v is best][0]
    kpis = {
        "eligible_count": len(elig),
        "best_score": best["score"],
        "best_pick": best_name,
        "lowest_rate": min(v["mid"] for v in elig),
        "lowest_pay": min(v["pay"] for v in elig),
        "fastest_close": min(v["close"] for v in elig),
        "highest_ltv": max(float(r["ltv"]) for r in lenders
                           if rows[r["lender"]]["eligible"]),
        "best_total_int": best["total_int"],
        "best_dscr": best["dscr"],
        "best_upfront": best["upfront"],
        "loan": LOAN,
    }
    return rows, kpis


def close(a, b, tol):
    return abs(a - b) <= tol


def main():
    failures = []

    # Rebuild fresh.
    r = subprocess.run([sys.executable, str(BASE / "generate.py")],
                       capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode != 0:
        print(r.stderr)
        sys.exit(f"generate.py failed ({r.returncode})")

    wb = load_workbook(XLSX)
    # 1. Short-form range scan over every formula cell.
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    if SHORT_RANGE.search(c.value):
                        bad.append(f"{ws.title}!{c.coordinate}: {c.value}")
    if bad:
        failures.append(f"short-form ranges ({len(bad)}): " + "; ".join(bad[:5]))
    else:
        print("OK  no short-form range literals")

    # 2. LibreOffice headless recalc.
    OUTDIR.mkdir(parents=True, exist_ok=True)
    for old in OUTDIR.glob("*"):
        old.unlink()
    r = subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx",
                        "--outdir", str(OUTDIR), str(XLSX)],
                       capture_output=True, text=True, timeout=180)
    recalc = OUTDIR / "lendermatch.xlsx"
    if not recalc.exists():
        print(r.stdout[-2000:])
        print(r.stderr[-2000:])
        sys.exit("LibreOffice recalc produced no file")
    print("OK  LibreOffice recalc done")
    rb = load_workbook(recalc, data_only=True)

    # 3. No formula errors anywhere in the recalculated values.
    errs = []
    for ws in rb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    errs.append(f"{ws.title}!{c.coordinate}={c.value}")
    if errs:
        failures.append(f"formula errors: " + "; ".join(errs[:8]))
    else:
        print("OK  no formula errors in recalculated workbook")

    rows, kpis = ground_truth()
    comp = rb["Comparison"]
    # 4. Per-lender comparison rows.
    for i, name in enumerate(rows, start=2):
        gt = rows[name]
        got_name = comp.cell(row=i, column=1).value
        if got_name != name:
            failures.append(f"Comparison row {i} lender {got_name!r} != {name!r}")
            continue
        got_elig = comp.cell(row=i, column=3).value
        if got_elig != ("Yes" if gt["eligible"] else "No"):
            failures.append(f"{name}: eligible {got_elig!r} != {gt['eligible']}")
        if not close(comp.cell(row=i, column=4).value, gt["mid"], 1e-9):
            failures.append(f"{name}: mid rate {comp.cell(row=i, column=4).value} != {gt['mid']}")
        pay = comp.cell(row=i, column=5).value
        if gt["eligible"]:
            if not (isinstance(pay, (int, float)) and close(pay, gt["pay"], 0.01)):
                failures.append(f"{name}: payment {pay} != {gt['pay']:.2f}")
            if not close(comp.cell(row=i, column=6).value, gt["total_int"], 0.5):
                failures.append(f"{name}: total interest {comp.cell(row=i, column=6).value} != {gt['total_int']:.2f}")
            if not close(comp.cell(row=i, column=7).value, gt["upfront"], 0.01):
                failures.append(f"{name}: upfront {comp.cell(row=i, column=7).value} != {gt['upfront']:.2f}")
            if not close(comp.cell(row=i, column=8).value, gt["dscr"], 0.005):
                failures.append(f"{name}: DSCR {comp.cell(row=i, column=8).value} != {gt['dscr']:.4f}")
            if not close(comp.cell(row=i, column=10).value, gt["score"], 0.05):
                failures.append(f"{name}: score {comp.cell(row=i, column=10).value} != {gt['score']:.2f}")
            if comp.cell(row=i, column=11).value != gt["rank"]:
                failures.append(f"{name}: rank {comp.cell(row=i, column=11).value} != {gt['rank']}")
        else:
            if pay not in (None, "", 0):
                failures.append(f"{name}: ineligible but payment={pay!r}")
            if comp.cell(row=i, column=10).value not in (0, 0.0):
                failures.append(f"{name}: ineligible but score={comp.cell(row=i, column=10).value!r}")
    print("OK  Comparison rows checked vs ground truth")

    # 5. Dashboard KPIs.
    dash = rb["Dashboard"]
    got = {
        "eligible_count": dash["B3"].value,
        "best_score": dash["B4"].value,
        "best_pick": dash["B5"].value,
        "lowest_rate": dash["B6"].value,
        "lowest_pay": dash["B7"].value,
        "fastest_close": dash["B8"].value,
        "highest_ltv": dash["B9"].value,
        "best_total_int": dash["B10"].value,
        "best_dscr": dash["B11"].value,
        "best_upfront": dash["B12"].value,
        "loan": dash["B13"].value,
    }
    tol = {"best_score": 0.05, "lowest_rate": 1e-9, "lowest_pay": 0.01,
           "best_total_int": 0.5, "best_dscr": 0.005, "best_upfront": 0.01,
           "highest_ltv": 1e-9}
    for k, exp in kpis.items():
        g = got[k]
        if k in ("best_pick",):
            ok = g == exp
        elif k in ("eligible_count", "fastest_close", "loan"):
            ok = g == exp
        else:
            ok = isinstance(g, (int, float)) and close(g, exp, tol[k])
        if not ok:
            failures.append(f"Dashboard {k}: {g!r} != {exp!r}")
    print("OK  Dashboard KPIs checked vs ground truth")
    print(f"    best pick: {got['best_pick']}  score {got['best_score']:.1f}  "
          f"lowest payment ${got['lowest_pay']:,.2f}/mo  eligible {got['eligible_count']}/14")

    # 6. Checks sheet all PASS.
    checks = rb["Checks"]
    bad_checks = [checks.cell(row=i, column=1).value
                  for i in range(3, 13)
                  if checks.cell(row=i, column=2).value != "PASS"]
    if bad_checks:
        failures.append("Checks not PASS: " + "; ".join(bad_checks))
    else:
        print("OK  Checks 10/10 PASS")

    if failures:
        print(f"\n{len(failures)} FAILURE(S):")
        for f in failures:
            print(" -", f)
        sys.exit(1)
    print("\nALL TESTS PASS")


if __name__ == "__main__":
    main()
