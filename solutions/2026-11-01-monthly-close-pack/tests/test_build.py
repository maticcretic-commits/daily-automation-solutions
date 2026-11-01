#!/usr/bin/env python3
"""Build the workbook, recalc it headlessly in LibreOffice, and compare
every checked figure against the independent Python ground truth.

LO quirk honoured: --convert-to runs on a COPY in a separate input dir
with --outdir elsewhere (converting in place overwrites the input file).
Formulas are A1-style only; a formula error in the recalculated copy is a
failure, not a warning. A second regression scan rejects short-form range
literals ($H$2:$5000) in every formula cell of the builder.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generate  # noqa: E402
import build_workbook  # noqa: E402
import main as ground_truth  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(HERE, "monthly-close-pack.xlsx")
ERRORS = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "Err:522")
BAD_RANGE = re.compile(r"\$[A-Z]+\$\d+:\$\d+")


def close(a, b, tol):
    return abs(a - b) <= tol


def main():
    generate.main()
    build_workbook.main()
    truth = ground_truth.compute()

    with open(os.path.join(HERE, "build_workbook.py")) as f:
        src = f.read()
    bad = [m.group(0) for m in BAD_RANGE.finditer(src)]
    assert not bad, f"short-form range literals in builder: {bad[:5]}"

    tmp = tempfile.mkdtemp(prefix="closepack-")
    indir, outdir = os.path.join(tmp, "in"), os.path.join(tmp, "out")
    os.makedirs(indir)
    os.makedirs(outdir)
    shutil.copy(XLSX, os.path.join(indir, "monthly-close-pack.xlsx"))
    r = subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", outdir,
         os.path.join(indir, "monthly-close-pack.xlsx")],
        capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, f"soffice failed: {r.stderr[:500]}"
    recalc = os.path.join(outdir, "monthly-close-pack.xlsx")
    assert os.path.exists(recalc), "LO produced no output"
    wb = load_workbook(recalc, data_only=True)

    failures = []

    def check(name, got, want, tol):
        if got is None or isinstance(got, str) or not close(got, want, tol):
            failures.append(f"{name}: got {got!r}, want {want}")

    for sheet in ("Dashboard", "PL_Monthly", "Balance_Sheet", "CashFlow",
                  "Budget_vs_Actual", "Checks", "Assumptions"):
        for row in wb[sheet].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    failures.append(f"formula error {sheet}!{c.coordinate}: {c.value}")

    ws = wb["Dashboard"]
    check("Dashboard revenue", ws["B4"].value, truth["revenue"], 0.01)
    check("Dashboard net income", ws["B5"].value, truth["net_income"], 0.01)
    check("Dashboard margin", ws["B6"].value, truth["margin"], 1e-6)
    check("Dashboard cash", ws["B7"].value, truth["cash_close"], 0.01)
    check("Dashboard AR", ws["B8"].value, truth["ar_close"], 0.01)
    ws = wb["PL_Monthly"]
    check("PL total revenue", ws["B12"].value, truth["revenue"], 0.01)
    check("PL total net", ws["G12"].value, truth["net_income"], 0.01)
    for m in range(1, 11):
        check(f"PL month {m} revenue", ws.cell(row=m + 1, column=2).value,
              truth["monthly"][m]["revenue"], 0.01)
        check(f"PL month {m} net", ws.cell(row=m + 1, column=7).value,
              truth["monthly"][m]["net"], 0.01)
    ws = wb["Balance_Sheet"]
    check("BS assets", ws["D5"].value, truth["assets"], 0.01)
    check("BS equity", ws["D11"].value, truth["equity_close"], 0.01)
    check("BS balance gap", ws["D13"].value, truth["balance_gap"], 0.01)
    check("BS AR", ws["D3"].value, truth["ar_close"], 0.01)
    check("BS AP", ws["D6"].value, truth["ap_close"], 0.01)
    check("BS equipment", ws["D4"].value, truth["equip_close"], 0.01)
    ws = wb["CashFlow"]
    check("CF tie-out", ws["B11"].value, 0, 0.01)
    check("CF closing cash", ws["B10"].value, truth["cash_close"], 0.01)
    ws = wb["Budget_vs_Actual"]
    for i, (cat, _typ) in enumerate(build_workbook.CATEGORIES, start=2):
        check(f"BvA actual {cat}", ws.cell(row=i, column=3).value,
              truth["bva"][cat]["actual"], 0.01)
        check(f"BvA budget {cat}", ws.cell(row=i, column=4).value,
              truth["bva"][cat]["budget"], 0.01)
    ws = wb["Checks"]
    for r in range(2, 12):
        if ws.cell(row=r, column=4).value != "PASS":
            failures.append(f"Checks row {r} ({ws.cell(row=r, column=1).value}): "
                            f"result {ws.cell(row=r, column=2).value!r} not PASS")
    ws = wb["Assumptions"]
    check("Assumptions opening cash", ws["B2"].value, 850000, 0)
    check("Assumptions AR share", ws["B9"].value, 0.20, 1e-9)

    assert not failures, "FAILURES:\n" + "\n".join(failures)
    print(f"ALL CHECKS PASS ({truth['txn_count']} transactions, "
          f"revenue Rs {truth['revenue']:,}, net Rs {truth['net_income']:,}, "
          f"balance gap {truth['balance_gap']})")


if __name__ == "__main__":
    main()
