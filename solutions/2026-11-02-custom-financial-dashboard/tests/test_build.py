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
XLSX = os.path.join(HERE, "finance-dashboard.xlsx")
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

    tmp = tempfile.mkdtemp(prefix="financepulse-")
    indir, outdir = os.path.join(tmp, "in"), os.path.join(tmp, "out")
    os.makedirs(indir)
    os.makedirs(outdir)
    shutil.copy(XLSX, os.path.join(indir, "finance-dashboard.xlsx"))
    r = subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", outdir,
         os.path.join(indir, "finance-dashboard.xlsx")],
        capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, f"soffice failed: {r.stderr[:500]}"
    recalc = os.path.join(outdir, "finance-dashboard.xlsx")
    assert os.path.exists(recalc), "LO produced no output"
    wb = load_workbook(recalc, data_only=True)

    failures = []

    def check(name, got, want, tol):
        if got is None or isinstance(got, str) or not close(got, want, tol):
            failures.append(f"{name}: got {got!r}, want {want}")

    for sheet in ("Dashboard", "Monthly", "Income", "Expenses", "Sales",
                  "Investments", "Checks", "Assumptions"):
        for row in wb[sheet].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    failures.append(f"formula error {sheet}!{c.coordinate}: {c.value}")

    ws = wb["Dashboard"]
    check("Dashboard total income", ws["B4"].value, truth["tot_income"], 0.01)
    check("Dashboard total expenses", ws["B5"].value, truth["tot_expenses"], 0.01)
    check("Dashboard net", ws["B6"].value, truth["net"], 0.01)
    check("Dashboard cash", ws["B7"].value, truth["cash_close"], 0.01)
    check("Dashboard sales revenue", ws["B8"].value, truth["sales_rev"], 0.01)
    check("Dashboard sales cost", ws["B9"].value, truth["sales_cost"], 0.01)
    check("Dashboard gross profit", ws["B10"].value, truth["gross"], 0.01)
    check("Dashboard savings rate", ws["B11"].value, truth["savings"], 1e-6)
    check("Dashboard portfolio value", ws["B12"].value, truth["inv_value"], 0.01)
    check("Dashboard invested", ws["B13"].value, truth["inv_invested"], 0.01)
    check("Dashboard gain", ws["B14"].value, truth["inv_gain"], 0.01)
    check("Dashboard gain %", ws["B15"].value, truth["inv_gain_pct"], 1e-6)
    for i, cat in enumerate(build_workbook.EXP_CATEGORIES, start=19):
        check(f"Expense mix {cat}", ws.cell(row=i, column=2).value,
              truth["exp_by_cat"][cat], 0.01)
    ws = wb["Monthly"]
    check("Monthly total income", ws["C12"].value, truth["tot_income"], 0.01)
    check("Monthly total net", ws["E12"].value, truth["net"], 0.01)
    check("Monthly closing cash", ws["I12"].value, truth["cash_close"], 0.01)
    for i, (key, _label) in enumerate(build_workbook.MONTHS, start=2):
        check(f"Monthly {key} income", ws.cell(row=i, column=3).value,
              truth["monthly"][key]["income"], 0.01)
        check(f"Monthly {key} net", ws.cell(row=i, column=5).value,
              truth["monthly"][key]["net"], 0.01)
        check(f"Monthly {key} gross", ws.cell(row=i, column=8).value,
              truth["monthly"][key]["gross"], 0.01)
    ws = wb["Investments"]
    check("Investments total row value", ws["G10"].value, truth["inv_value"], 0.01)
    check("Investments total row invested", ws["F10"].value, truth["inv_invested"], 0.01)
    ws = wb["Checks"]
    for r in range(2, 12):
        if ws.cell(row=r, column=4).value != "PASS":
            failures.append(f"Checks row {r} ({ws.cell(row=r, column=1).value}): "
                            f"result {ws.cell(row=r, column=2).value!r} not PASS")
    ws = wb["Assumptions"]
    check("Assumptions opening cash", ws["B2"].value, 125000, 0)

    assert not failures, "FAILURES:\n" + "\n".join(failures)
    print(f"ALL CHECKS PASS ({truth['income_count']} income, "
          f"{truth['expense_count']} expenses, {truth['sales_count']} sales rows; "
          f"income Rs {truth['tot_income']:,}, expenses Rs {truth['tot_expenses']:,}, "
          f"net Rs {truth['net']:,}, cash Rs {truth['cash_close']:,}, "
          f"portfolio Rs {truth['inv_value']:,.0f}, gain {truth['inv_gain_pct']:.1%})")


if __name__ == "__main__":
    main()
