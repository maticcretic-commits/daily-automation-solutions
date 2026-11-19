#!/usr/bin/env python3
"""Tests for DueDiligence Review Pack.

Runs generate.py, then verifies the workbook two ways:
1. Formula presence — key cells are live formulas, the Dashboard
   defaults, 3 charts exist.
2. LibreOffice headless recalc vs an independent Python re-implementation
   from data/*.csv: derived P&L/balance-sheet lines, all 12 ratios plus
   the growth block, the 12-month bank reconciliation (book balances
   from receipts/payments, bank balances from the item ledger), loan
   totals and LT/current split, contingent liabilities, all 9 red-flag
   triggers, the risk score and verdict, and all 10 Checks reading PASS.
"""
import csv
import datetime as dt
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data"
XLSX = BASE / "due-diligence-review.xlsx"

TOL = 0.01
PCT_TOL = 0.0001


def load(name):
    with open(DATA / name, newline="") as f:
        return list(csv.DictReader(f))


def month_end(y, m):
    if m == 12:
        return dt.date(y, 12, 31)
    return dt.date(y, m + 1, 1) - dt.timedelta(days=1)


def ground_truth():
    fin_rows = load("financials.csv")
    fin = {r["item"]: [float(r["fy24"]), float(r["fy25"]),
                       float(r["fy26"])] for r in fin_rows}
    bank_rows = load("bankrec.csv")
    items_rows = load("bank_items.csv")
    loan_rows = load("loans.csv")
    cont_rows = load("contingent.csv")
    flag_rows = load("redflags.csv")
    meta = {r["key"]: r["value"] for r in load("meta.csv")}
    as_of = dt.date.fromisoformat(meta["as_of"])

    # Derived P&L / balance sheet (per year index 0..2).
    gp = [fin["Revenue"][i] - fin["COGS"][i] for i in range(3)]
    ebitda = [gp[i] - fin["Operating expenses"][i] for i in range(3)]
    ebit = [ebitda[i] - fin["Depreciation"][i] for i in range(3)]
    pbt = [ebit[i] - fin["Interest"][i] for i in range(3)]
    ni = [pbt[i] - fin["Tax"][i] for i in range(3)]
    nw = [fin["Share capital"][i] + fin["Reserves"][i] for i in range(3)]
    tle = [nw[i] + fin["Long-term borrowings"][i]
           + fin["Short-term borrowings"][i] + fin["Trade payables"][i]
           + fin["Other current liabilities"][i] for i in range(3)]
    assets = [fin["Fixed assets (net)"][i] + fin["Inventory"][i]
              + fin["Trade receivables"][i] + fin["Cash & bank"][i]
              + fin["Other current assets"][i] for i in range(3)]

    def ratio(fn):
        return [fn(i) for i in range(3)]

    ratios = {
        "current": ratio(lambda i: (
            fin["Inventory"][i] + fin["Trade receivables"][i]
            + fin["Cash & bank"][i] + fin["Other current assets"][i])
            / (fin["Short-term borrowings"][i] + fin["Trade payables"][i]
               + fin["Other current liabilities"][i])),
        "quick": ratio(lambda i: (
            fin["Trade receivables"][i] + fin["Cash & bank"][i]
            + fin["Other current assets"][i])
            / (fin["Short-term borrowings"][i] + fin["Trade payables"][i]
               + fin["Other current liabilities"][i])),
        "de": ratio(lambda i: (
            fin["Long-term borrowings"][i]
            + fin["Short-term borrowings"][i]) / nw[i]),
        "ic": ratio(lambda i: ebit[i] / fin["Interest"][i]),
        "gm": ratio(lambda i: gp[i] / fin["Revenue"][i]),
        "ebitda_m": ratio(lambda i: ebitda[i] / fin["Revenue"][i]),
        "nm": ratio(lambda i: ni[i] / fin["Revenue"][i]),
        "roce": ratio(lambda i: ebit[i] / (nw[i]
                      + fin["Long-term borrowings"][i])),
        "roe": ratio(lambda i: ni[i] / nw[i]),
        "recv_days": ratio(lambda i: fin["Trade receivables"][i]
                           / fin["Revenue"][i] * 365),
        "pay_days": ratio(lambda i: fin["Trade payables"][i]
                          / fin["COGS"][i] * 365),
        "inv_days": ratio(lambda i: fin["Inventory"][i]
                          / fin["COGS"][i] * 365),
    }
    cagr = (fin["Revenue"][2] / fin["Revenue"][0]) ** 0.5 - 1
    cash_growth = ((fin["Cash & bank"][2] - fin["Cash & bank"][0])
                   / fin["Cash & bank"][0])
    recv_growth = ((fin["Trade receivables"][2]
                    - fin["Trade receivables"][0])
                   / fin["Trade receivables"][0])
    rev_growth = ((fin["Revenue"][2] - fin["Revenue"][0])
                  / fin["Revenue"][0])

    # Bank reconciliation from the item ledger.
    ledger = []
    for r in items_rows:
        ledger.append({
            "item": r["item"],
            "date": dt.date.fromisoformat(r["date"]),
            "amount": float(r["amount"]),
            "effect": float(r["effect"]),
            "cleared": (dt.date.fromisoformat(r["cleared"])
                        if r["cleared"] else None),
        })
    book = fin["Cash & bank"][1]  # opening = FY25 cash
    bank_gt = []
    for br in bank_rows:
        book = round(book + float(br["receipts"])
                     - float(br["payments"]), 2)
        y, m = br["month"].split("-")
        end = month_end(int(y), int(m))
        diff = round(sum(
            it["effect"] for it in ledger
            if it["date"] <= end
            and (it["cleared"] is None or it["cleared"] > end)), 2)
        bank_gt.append({"month": br["month"], "book": book,
                        "bank": round(book - diff, 2), "diff": diff})
    open_items = [it for it in ledger if it["cleared"] is None]
    open_days = [(as_of - it["date"]).days for it in open_items]
    open_effect = round(sum(it["effect"] for it in open_items), 2)

    # Loans + contingent.
    loan_total = sum(float(r["outstanding"]) for r in loan_rows)
    lt_total = sum(float(r["lt_portion"]) for r in loan_rows)
    cur_total = sum(float(r["current_portion"]) for r in loan_rows)
    cont_total = sum(float(r["amount"]) for r in cont_rows)
    cont_pct = cont_total / nw[2]

    # Red-flag triggers, reimplemented independently.
    trig = {
        "F1": ratios["recv_days"][2] > ratios["recv_days"][1]
              > ratios["recv_days"][0],
        "F2": cont_pct > 0.2,
        "F3": any(r["facility"].startswith("Director loan")
                  and float(r["outstanding"]) > 0
                  and float(r["rate"]) == 0 for r in loan_rows),
        "F4": max(ratios["ic"]) < 2,
        "F5": max(ratios["quick"]) < 1,
        "F6": max(open_days) > 30,
        "F7": float(cont_rows[0]["amount"]) > 0,
        "F8": recv_growth > rev_growth,
        "F9": True,  # auditor emphasis of matter: Yes (data cell B2)
    }
    weights = {r["id"]: int(r["weight"]) for r in flag_rows}
    score = sum(weights[k] for k, v in trig.items() if v)
    n_trig = sum(1 for v in trig.values() if v)
    sev_count = {}
    for r in flag_rows:
        if trig[r["id"]]:
            sev_count[r["severity"]] = sev_count.get(r["severity"], 0) + 1
    verdict = ("HIGH — proceed only with price protection and conditions"
               if score >= 30 else
               "MODERATE — conditions required" if score >= 15 else
               "LOW — no material flags")

    return {
        "fin": fin, "gp": gp, "ebitda": ebitda, "ebit": ebit, "pbt": pbt,
        "ni": ni, "nw": nw, "tle": tle, "assets": assets,
        "ratios": ratios, "cagr": cagr, "cash_growth": cash_growth,
        "recv_growth": recv_growth, "rev_growth": rev_growth,
        "bank_gt": bank_gt, "open_items": open_items,
        "open_days": open_days, "open_effect": open_effect,
        "loan_total": loan_total, "lt_total": lt_total,
        "cur_total": cur_total, "cont_total": cont_total,
        "cont_pct": cont_pct, "trig": trig, "score": score,
        "n_trig": n_trig, "sev_count": sev_count, "verdict": verdict,
        "as_of": as_of,
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


def test_build():
    r = subprocess.run([sys.executable, "generate.py"], cwd=BASE,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert XLSX.exists(), "xlsx not built"
    gt = ground_truth()
    assert gt["score"] == 31, f"risk score {gt['score']}, expected 31"
    assert gt["n_trig"] == 9, f"{gt['n_trig']} flags trigger, expected 9"
    assert gt["open_effect"] == 1.55
    assert gt["bank_gt"][-1]["book"] == 175
    assert gt["bank_gt"][-1]["bank"] == 173.45

    # --- formula presence ---
    wb = openpyxl.load_workbook(XLSX, data_only=False)
    assert wb.sheetnames == ["Guide", "Financials", "Ratios", "BankRec",
                             "Loans", "RedFlags", "Dashboard",
                             "Checks"], wb.sheetnames
    fin = wb["Financials"]
    assert str(fin["B6"].value) == "=B4-B5"
    assert str(fin["B8"].value) == "=B6-B7"
    assert str(fin["B19"].value) == "=B17+B18"
    assert str(fin["B24"].value) == "=SUM(B19:B23)"
    assert str(fin["B31"].value) == "=B30-B24"
    rat = wb["Ratios"]
    assert "Financials!B26" in str(rat["B4"].value)
    assert str(rat["B17"].value) == "=(Financials!D4/Financials!B4)^(1/2)-1"
    assert str(rat["B19"].value) == "=Loans!B17"
    br = wb["BankRec"]
    assert str(br["D4"].value) == "=Financials!C28+B4-C4"
    assert str(br["F4"].value) == "=D4-E4"
    assert str(br["E20"].value) == "=$B$2-B20"
    assert str(br["B25"].value) == "=D24-F15"
    ln = wb["Loans"]
    assert str(ln["D9"].value) == "=SUM(D4:D8)"
    assert str(ln["B17"].value) == "=B16/Financials!D19"
    rf = wb["RedFlags"]
    assert str(rf["F4"].value).startswith("=IF(AND(Ratios!D13")
    assert str(rf["F9"].value).startswith("=IF(MAX(BankRec!E20:E23)")
    assert str(rf["G13"].value) == "=SUM(G4:G12)"
    db = wb["Dashboard"]
    assert str(db["B4"].value) == "=BankRec!B2"
    assert str(db["B17"].value) == '=COUNTIF(RedFlags!F4:F12,"TRIGGERED")'
    assert str(db["B18"].value).startswith("=COUNTIFS(")
    assert str(db["B22"].value).startswith("=IF(B21>=30,")
    assert str(wb["Checks"]["E4"].value).startswith("=IF(ABS(")
    assert len(db._charts) == 3, "expected 3 charts"

    # --- recalc vs ground truth ---
    rw = recalc(XLSX)
    check_no_errors(rw)

    # Financials derived lines
    fin_r = rw["Financials"]
    derived = {6: gt["gp"], 8: gt["ebitda"], 10: gt["ebit"],
               12: gt["pbt"], 14: gt["ni"], 19: gt["nw"],
               24: gt["tle"], 30: gt["assets"]}
    for row, vals in derived.items():
        for j in range(3):
            assert_close(fin_r.cell(row=row, column=2 + j).value,
                         vals[j], f"Financials row {row} FY{24 + j}")
    for j in range(3):
        assert_close(fin_r.cell(row=31, column=2 + j).value, 0,
                     f"Financials balance check FY{24 + j}")

    # Ratios
    rat_r = rw["Ratios"]
    ratio_rows = {"current": 4, "quick": 5, "de": 6, "ic": 7, "gm": 8,
                  "ebitda_m": 9, "nm": 10, "roce": 11, "roe": 12,
                  "recv_days": 13, "pay_days": 14, "inv_days": 15}
    for key, row in ratio_rows.items():
        tol = PCT_TOL if key in ("gm", "ebitda_m", "nm", "roce", "roe") \
            else TOL
        for j in range(3):
            assert_close(rat_r.cell(row=row, column=2 + j).value,
                         gt["ratios"][key][j],
                         f"Ratios {key} FY{24 + j}", tol)
    assert_close(rat_r["B17"].value, gt["cagr"], "Revenue CAGR", PCT_TOL)
    assert_close(rat_r["B18"].value, gt["cash_growth"], "Cash growth",
                 PCT_TOL)
    assert_close(rat_r["B19"].value, gt["cont_pct"], "Contingent % NW",
                 PCT_TOL)
    assert_close(rat_r["B20"].value, gt["recv_growth"],
                 "Receivables growth", PCT_TOL)
    assert_close(rat_r["B21"].value, gt["rev_growth"], "Revenue growth",
                 PCT_TOL)

    # BankRec: book, bank, difference per month
    br_r = rw["BankRec"]
    assert br_r["B2"].value == gt["as_of"] or \
        str(br_r["B2"].value)[:10] == "2026-03-31", \
        f"as-of date {br_r['B2'].value}"
    for i, m in enumerate(gt["bank_gt"]):
        row = 4 + i
        assert_close(br_r.cell(row=row, column=4).value, m["book"],
                     f"BankRec book {m['month']}")
        assert_close(br_r.cell(row=row, column=5).value, m["bank"],
                     f"BankRec bank {m['month']}")
        assert_close(br_r.cell(row=row, column=6).value, m["diff"],
                     f"BankRec diff {m['month']}")
    for i, days in enumerate(gt["open_days"]):
        row = 20 + i
        assert br_r.cell(row=row, column=5).value == days, \
            f"open-item days row {row}: " \
            f"{br_r.cell(row=row, column=5).value} != {days}"
    assert_close(br_r["D24"].value, gt["open_effect"],
                 "BankRec open-item total")
    assert_close(br_r["B25"].value, 0, "BankRec reconciliation check")

    # Loans + contingent
    ln_r = rw["Loans"]
    assert_close(ln_r["D9"].value, gt["loan_total"], "Loan total")
    assert_close(ln_r["I9"].value, gt["lt_total"], "Loan LT total")
    assert_close(ln_r["J9"].value, gt["cur_total"], "Loan current total")
    assert_close(ln_r["I9"].value + ln_r["J9"].value, gt["loan_total"],
                 "Loan LT+current = total")
    assert_close(ln_r["B16"].value, gt["cont_total"], "Contingent total")
    assert_close(ln_r["B17"].value, gt["cont_pct"], "Contingent % NW",
                 PCT_TOL)

    # RedFlags: statuses and points
    rf_r = rw["RedFlags"]
    for i, fid in enumerate(["F1", "F2", "F3", "F4", "F5", "F6", "F7",
                             "F8", "F9"]):
        row = 4 + i
        want = "TRIGGERED" if gt["trig"][fid] else "CLEAR"
        assert rf_r.cell(row=row, column=6).value == want, \
            f"{fid} status {rf_r.cell(row=row, column=6).value} != {want}"
    assert_close(rf_r["G13"].value, gt["score"], "Risk score")

    # Dashboard KPIs
    db_r = rw["Dashboard"]
    assert_close(db_r["B8"].value, 2480, "Dashboard revenue FY26")
    assert_close(db_r["B9"].value, gt["cagr"], "Dashboard CAGR", PCT_TOL)
    assert_close(db_r["B10"].value, gt["ratios"]["ebitda_m"][2],
                 "Dashboard EBITDA margin", PCT_TOL)
    assert_close(db_r["B11"].value, gt["ratios"]["current"][2],
                 "Dashboard current ratio")
    assert_close(db_r["B12"].value, gt["ratios"]["quick"][2],
                 "Dashboard quick ratio")
    assert_close(db_r["B13"].value, gt["ratios"]["de"][2],
                 "Dashboard D/E")
    assert_close(db_r["B14"].value, gt["ratios"]["ic"][2],
                 "Dashboard interest coverage")
    assert_close(db_r["B15"].value, gt["cont_total"],
                 "Dashboard contingent total")
    assert_close(db_r["B16"].value, gt["cont_pct"], "Dashboard cont %",
                 PCT_TOL)
    assert_close(db_r["B17"].value, gt["n_trig"], "Dashboard flags")
    assert_close(db_r["B18"].value, gt["sev_count"].get("High", 0),
                 "Dashboard high flags")
    assert_close(db_r["B19"].value, gt["sev_count"].get("Medium", 0),
                 "Dashboard medium flags")
    assert_close(db_r["B20"].value, gt["sev_count"].get("Low", 0),
                 "Dashboard low flags")
    assert_close(db_r["B21"].value, gt["score"], "Dashboard risk score")
    assert db_r["B22"].value == gt["verdict"], \
        f"verdict {db_r['B22'].value!r} != {gt['verdict']!r}"
    assert_close(db_r["B28"].value, gt["n_trig"], "Severity total")

    # --- Checks: all 10 PASS ---
    ck = rw["Checks"]
    for row in range(4, 14):
        assert ck.cell(row=row, column=5).value == "PASS", \
            f"check row {row} " \
            f"({ck.cell(row=row, column=2).value}) = " \
            f"{ck.cell(row=row, column=5).value}"

    print(f"OK: revenue FY26 2,480 CAGR {gt['cagr']:.1%} EBITDA margin "
          f"{gt['ratios']['ebitda_m'][2]:.1%} current "
          f"{gt['ratios']['current'][2]:.2f} D/E "
          f"{gt['ratios']['de'][2]:.2f} contingent "
          f"{gt['cont_total']} ({gt['cont_pct']:.1%} NW) risk score "
          f"{gt['score']} verdict HIGH")


if __name__ == "__main__":
    test_build()
    print("ALL TESTS PASS")
