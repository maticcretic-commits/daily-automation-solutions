#!/usr/bin/env python3
"""Tests for OptionsSignal.

Runs generate.py, then verifies the workbook two ways:
1. Formula presence — key cells in the generated xlsx are live formulas,
   the Dashboard defaults, 3 charts exist.
2. LibreOffice headless recalc vs an independent Python re-implementation
   from data/option_chain.csv: per-strike derived columns (total OI, PCR,
   % changes, CE/PE signals, max-pain payouts), the Signals combined
   reads, all Dashboard KPIs and signal counts, zero formula errors
   anywhere, and all 10 Checks reading PASS.
"""
import csv
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data"
XLSX = BASE / "options-signal-dashboard.xlsx"

TOL = 0.01
PCT_TOL = 0.0001


def load_chain():
    with open(DATA / "option_chain.csv", newline="") as f:
        rows = list(csv.DictReader(f))
    with open(DATA / "meta.csv", newline="") as f:
        meta = {r["key"]: r["value"] for r in csv.DictReader(f)}
    return rows, meta


def signal(oi_chg, ltp, prev):
    up = ltp > prev
    if oi_chg > 0 and up:
        return "Long Build-up"
    if oi_chg > 0 and not up:
        return "Short Build-up"
    if oi_chg < 0 and up:
        return "Short Covering"
    return "Long Unwinding"


def combined(ce_sig, pe_sig):
    if ce_sig == "Long Build-up" and pe_sig == "Short Covering":
        return "Strong Bullish"
    if ce_sig == "Short Build-up" and pe_sig == "Long Build-up":
        return "Strong Bearish"
    if ce_sig == "Long Build-up" or pe_sig == "Short Covering":
        return "Bullish"
    if ce_sig == "Short Build-up" or pe_sig == "Long Build-up":
        return "Bearish"
    return "Neutral"


def ground_truth(rows):
    strikes = [int(r["strike"]) for r in rows]
    ce_oi = [int(r["ce_oi"]) for r in rows]
    pe_oi = [int(r["pe_oi"]) for r in rows]
    ce_chg = [int(r["ce_oi_chg"]) for r in rows]
    pe_chg = [int(r["pe_oi_chg"]) for r in rows]
    ce_ltp = [float(r["ce_ltp"]) for r in rows]
    ce_prev = [float(r["ce_prev_close"]) for r in rows]
    pe_ltp = [float(r["pe_ltp"]) for r in rows]
    pe_prev = [float(r["pe_prev_close"]) for r in rows]
    payouts = []
    for s in strikes:
        payouts.append(
            sum(o * max(s - k, 0) for o, k in zip(ce_oi, strikes))
            + sum(o * max(k - s, 0) for o, k in zip(pe_oi, strikes)))
    ce_sig = [signal(c, l, p) for c, l, p in zip(ce_chg, ce_ltp, ce_prev)]
    pe_sig = [signal(c, l, p) for c, l, p in zip(pe_chg, pe_ltp, pe_prev)]
    comb = [combined(c, p) for c, p in zip(ce_sig, pe_sig)]
    return {
        "strikes": strikes, "ce_oi": ce_oi, "pe_oi": pe_oi,
        "ce_chg": ce_chg, "pe_chg": pe_chg,
        "ce_ltp": ce_ltp, "ce_prev": ce_prev,
        "pe_ltp": pe_ltp, "pe_prev": pe_prev,
        "payouts": payouts,
        "ce_sig": ce_sig, "pe_sig": pe_sig, "combined": comb,
        "total_ce": sum(ce_oi), "total_pe": sum(pe_oi),
        "pcr": sum(pe_oi) / sum(ce_oi),
        "max_pain": strikes[payouts.index(min(payouts))],
        "min_payout": min(payouts),
        "support": strikes[pe_oi.index(max(pe_oi))],
        "resistance": strikes[ce_oi.index(max(ce_oi))],
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
    rows, meta = load_chain()
    gt = ground_truth(rows)
    assert int(meta["spot"]) == 25480
    assert meta["expiry"] == "2026-11-26"

    # --- formula presence ---
    wb = openpyxl.load_workbook(XLSX, data_only=False)
    assert wb.sheetnames == ["Guide", "OptionChain", "Signals",
                             "Dashboard", "Checks"], wb.sheetnames
    oc = wb["OptionChain"]
    assert str(oc["L4"].value) == "=B4+G4"
    assert str(oc["N4"].value) == "=IF(B4=0,0,G4/B4)"
    assert str(oc["Q4"].value).startswith('=IF(AND(C4>0,O4>0)')
    assert str(oc["R4"].value).startswith('=IF(AND(H4>0,P4>0)')
    assert "SUMPRODUCT" in str(oc["S4"].value)
    assert str(wb["Signals"]["D4"].value).startswith(
        '=IF(AND(B4="Long Build-up"')
    assert str(wb["Dashboard"]["B6"].value) == "=ROUND(B4/50,0)*50"
    assert str(wb["Dashboard"]["B12"].value) == "=B10/B9"
    assert "INDEX(OptionChain!A4:A32" in str(wb["Dashboard"]["B15"].value)
    assert str(wb["Checks"]["D5"].value) == "=Dashboard!B9"
    assert wb["Checks"]["C5"].value == gt["total_ce"], \
        "Expected must be numeric ground truth"
    assert len(wb["Dashboard"]._charts) == 3, "expected 3 charts"
    # Ground truth is interesting: all four signal types occur, and
    # support/resistance bracket the ATM strike.
    assert set(gt["ce_sig"]) == {"Long Build-up", "Short Build-up",
                                 "Short Covering", "Long Unwinding"}
    assert set(gt["pe_sig"]) == {"Long Build-up", "Short Build-up",
                                 "Short Covering", "Long Unwinding"}
    assert gt["support"] < 25500 < gt["resistance"], (
        gt["support"], gt["resistance"])

    # --- recalc vs ground truth ---
    rw = recalc(XLSX)
    check_no_errors(rw)

    # OptionChain derived columns, rows 4..32
    oc_r = rw["OptionChain"]
    for i in range(29):
        row = 4 + i
        assert_close(oc_r.cell(row=row, column=12).value,
                     gt["ce_oi"][i] + gt["pe_oi"][i],
                     f"OptionChain total OI row {row}")
        assert_close(oc_r.cell(row=row, column=13).value,
                     gt["ce_chg"][i] + gt["pe_chg"][i],
                     f"OptionChain total OI chg row {row}")
        assert_close(oc_r.cell(row=row, column=14).value,
                     gt["pe_oi"][i] / gt["ce_oi"][i],
                     f"OptionChain PCR row {row}", PCT_TOL)
        assert_close(oc_r.cell(row=row, column=15).value,
                     (gt["ce_ltp"][i] - gt["ce_prev"][i])
                     / gt["ce_prev"][i],
                     f"OptionChain CE %chg row {row}", PCT_TOL)
        assert_close(oc_r.cell(row=row, column=16).value,
                     (gt["pe_ltp"][i] - gt["pe_prev"][i])
                     / gt["pe_prev"][i],
                     f"OptionChain PE %chg row {row}", PCT_TOL)
        assert oc_r.cell(row=row, column=17).value == gt["ce_sig"][i], \
            f"CE signal row {row}"
        assert oc_r.cell(row=row, column=18).value == gt["pe_sig"][i], \
            f"PE signal row {row}"
        assert_close(oc_r.cell(row=row, column=19).value,
                     gt["payouts"][i],
                     f"OptionChain max-pain payout row {row}")

    # Signals sheet
    sg = rw["Signals"]
    for i in range(29):
        row = 4 + i
        assert sg.cell(row=row, column=4).value == gt["combined"][i], \
            f"combined read row {row}"
        flag = sg.cell(row=row, column=6).value
        if gt["strikes"][i] == gt["max_pain"]:
            assert flag == "MAX PAIN", f"max pain flag row {row}"
        else:
            assert flag in ("", None), f"max pain flag row {row}"

    # Dashboard KPIs
    db = rw["Dashboard"]
    assert_close(db["B4"].value, 25480, "Dashboard spot")
    assert_close(db["B6"].value, 25500, "Dashboard ATM strike")
    assert_close(db["B9"].value, gt["total_ce"], "Dashboard total CE OI")
    assert_close(db["B10"].value, gt["total_pe"], "Dashboard total PE OI")
    assert_close(db["B11"].value, gt["total_ce"] + gt["total_pe"],
                 "Dashboard total OI")
    assert_close(db["B12"].value, gt["pcr"], "Dashboard PCR", PCT_TOL)
    assert_close(db["B13"].value, sum(gt["ce_chg"]),
                 "Dashboard net CE OI change")
    assert_close(db["B14"].value, sum(gt["pe_chg"]),
                 "Dashboard net PE OI change")
    assert_close(db["B15"].value, gt["max_pain"],
                 "Dashboard max pain strike")
    assert_close(db["B16"].value, gt["min_payout"],
                 "Dashboard max pain payout")
    assert_close(db["B17"].value, gt["support"], "Dashboard support")
    assert_close(db["B18"].value, gt["resistance"],
                 "Dashboard resistance")
    bullish = sum(1 for c in gt["combined"]
                  if c in ("Bullish", "Strong Bullish"))
    bearish = sum(1 for c in gt["combined"]
                  if c in ("Bearish", "Strong Bearish"))
    neutral = sum(1 for c in gt["combined"] if c == "Neutral")
    assert_close(db["B19"].value, bullish, "Dashboard bullish strikes")
    assert_close(db["B20"].value, bearish, "Dashboard bearish strikes")
    assert_close(db["B21"].value, neutral, "Dashboard neutral strikes")

    # Signal-count table rows 25..28
    for i, sig in enumerate(["Long Build-up", "Short Build-up",
                             "Short Covering", "Long Unwinding"]):
        assert_close(db.cell(row=25 + i, column=2).value,
                     gt["ce_sig"].count(sig), f"CE count {sig}")
        assert_close(db.cell(row=25 + i, column=3).value,
                     gt["pe_sig"].count(sig), f"PE count {sig}")
    assert_close(db["B29"].value, 29, "CE signal total")
    assert_close(db["C29"].value, 29, "PE signal total")

    # --- Checks: all 10 PASS ---
    ck = rw["Checks"]
    for row in range(5, 15):
        assert ck.cell(row=row, column=5).value == "PASS", \
            f"check row {row} " \
            f"({ck.cell(row=row, column=2).value}) = " \
            f"{ck.cell(row=row, column=5).value}"

    print(f"OK: total CE OI {gt['total_ce']:,} total PE OI "
          f"{gt['total_pe']:,} PCR {gt['pcr']:.3f} max pain "
          f"{gt['max_pain']} support {gt['support']} resistance "
          f"{gt['resistance']} bullish {bullish} bearish {bearish} "
          f"neutral {neutral}")


if __name__ == "__main__":
    test_build()
    print("ALL TESTS PASS")
