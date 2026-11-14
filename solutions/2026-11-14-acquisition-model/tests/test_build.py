#!/usr/bin/env python3
"""Tests for AcquisitionModel.

Runs make_data.py + generate.py, then verifies the workbook two ways:
1. Formula presence — key cells in the generated xlsx are live formulas.
2. LibreOffice headless recalc vs an independent Python re-implementation
   of the model (Base scenario): historicals, projected revenue/EBITDA/
   FCF per year, DCF (PV of FCF, terminal value, enterprise value,
   equity value, implied price/share, premium vs asking), the
   sensitivity-grid base cell, and all 10 Checks reading PASS.
"""
import csv
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data"
XLSX = BASE / "acquisition-model.xlsx"

TOL_MONEY = 0.5
TOL_RATIO = 1e-6


def ground_truth():
    """Independent re-implementation of the Base-scenario model."""
    rev = 4200000.0
    opex = 1450000.0
    wc_prev = rev * 0.12
    fcf = []
    revs, ebitdas = [], []
    for _ in range(5):
        rev = rev * 1.08
        cogs = rev * (1 - 0.38)
        gp = rev - cogs
        opex = opex * 1.05
        ebitda = gp - opex
        ebit = ebitda - 220000
        ebt = ebit - 90000
        nopat = ebit * 0.75
        wc = rev * 0.12
        dwc = wc - wc_prev
        wc_prev = wc
        fcf.append(nopat + 220000 - 180000 - dwc)
        revs.append(rev)
        ebitdas.append(ebitda)
    pv_fcf = [f / (1.11 ** (i + 1)) for i, f in enumerate(fcf)]
    tv = fcf[-1] * 1.025 / (0.11 - 0.025)
    pv_tv = tv / (1.11 ** 5)
    ev = sum(pv_fcf) + pv_tv
    equity = ev - 650000
    return {
        "revs": revs, "ebitdas": ebitdas, "fcf": fcf,
        "pv_fcf": pv_fcf, "tv": tv, "pv_tv": pv_tv, "ev": ev,
        "equity": equity, "pps": equity / 100000,
        "premium": (equity - 2400000) / 2400000,
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


def test_build():
    r = subprocess.run([sys.executable, "make_data.py"], cwd=BASE,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    r = subprocess.run([sys.executable, "generate.py"], cwd=BASE,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert XLSX.exists(), "xlsx not built"

    # --- formula presence ---
    wb = openpyxl.load_workbook(XLSX, data_only=False)
    assert wb.sheetnames == ["Guide", "Assumptions", "Historical",
                             "Scenarios", "Projections", "Valuation",
                             "Dashboard", "Checks"], wb.sheetnames
    p = wb["Projections"]
    assert str(p["C6"].value).startswith("="), p["C6"].value
    assert str(p["C22"].value).startswith("="), p["C22"].value
    assert p["B3"].value == "Base"
    v = wb["Valuation"]
    assert str(v["B11"].value).startswith("="), v["B11"].value
    assert str(v["D25"].value).startswith("="), v["D25"].value

    # --- recalc vs ground truth ---
    gt = ground_truth()
    rw = recalc(XLSX)

    def val(sheet, cell):
        v = rw[sheet][cell].value
        assert v is not None, f"{sheet}!{cell} is empty after recalc"
        return v

    # no formula errors anywhere
    for ws in rw.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#"):
                    raise AssertionError(
                        f"formula error in {ws.title}!{c.coordinate}: "
                        f"{c.value}")

    h = rw["Historical"]
    assert abs(h["D5"].value - 4200000) < TOL_MONEY
    assert abs(h["D7"].value - 1596000) < TOL_MONEY  # gross profit 2025
    assert abs(h["D12"].value - 154000) < TOL_MONEY  # EBITDA 2025
    assert abs(h["D18"].value - (-156000)) < TOL_MONEY  # net income 2025

    for j in range(5):
        col = "CDEFG"[j]
        assert abs(val("Projections", f"{col}6") - gt["revs"][j]) < TOL_MONEY, col
        assert abs(val("Projections", f"{col}11") - gt["ebitdas"][j]) < TOL_MONEY, col
        assert abs(val("Projections", f"{col}22") - gt["fcf"][j]) < TOL_MONEY, col
        assert abs(val("Valuation", f"{'BCDEF'[j]}7") - gt["pv_fcf"][j]) < TOL_MONEY, col
    assert abs(val("Valuation", "B9") - gt["tv"]) < TOL_MONEY
    assert abs(val("Valuation", "B10") - gt["pv_tv"]) < TOL_MONEY
    assert abs(val("Valuation", "B11") - gt["ev"]) < TOL_MONEY
    assert abs(val("Valuation", "B13") - gt["equity"]) < TOL_MONEY
    assert abs(val("Valuation", "B15") - gt["pps"]) < 0.01
    assert abs(val("Valuation", "B18") - gt["premium"]) < 1e-4
    assert val("Valuation", "B19").startswith("PASS"), val("Valuation", "B19")
    assert abs(val("Valuation", "D25") - gt["ev"]) < TOL_MONEY  # sensitivity base
    # sensitivity corners move the right way
    assert val("Valuation", "B23") > val("Valuation", "F27")  # low WACC/high g > high WACC/low g
    # dashboard KPIs
    assert abs(val("Dashboard", "B7") - gt["ev"]) < TOL_MONEY
    assert abs(val("Dashboard", "B8") - gt["equity"]) < TOL_MONEY
    assert abs(val("Dashboard", "B13") - gt["revs"][-1]) < TOL_MONEY
    # scenario table: Bull 2030 revenue > Base > Bear
    assert (val("Dashboard", "B23") > val("Dashboard", "B22") >
            val("Dashboard", "B24"))
    # checks all PASS
    for row in range(5, 15):
        assert rw["Checks"][f"D{row}"].value == "PASS", \
            f"check row {row} = {rw['Checks'][f'D{row}'].value}"

    print(f"OK: EV ${gt['ev']:,.0f} equity ${gt['equity']:,.0f} "
          f"pps ${gt['pps']:.2f} premium {gt['premium']:.1%} "
          f"FCF 2026 ${gt['fcf'][0]:,.0f} FCF 2030 ${gt['fcf'][-1]:,.0f}")


if __name__ == "__main__":
    test_build()
    print("ALL TESTS PASS")
