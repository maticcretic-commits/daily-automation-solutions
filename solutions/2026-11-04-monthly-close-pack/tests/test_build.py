#!/usr/bin/env python3
"""Build the workbook, recalc it headlessly in LibreOffice, and compare
every checked figure against the independent Python ground truth.

LO quirks honoured: --convert-to runs on a COPY in a separate input dir
with --outdir elsewhere (converting in place overwrites the input file).
Formulas are A1-style only; a formula error in the recalculated copy is a
failure, not a warning. A second regression scan rejects short-form range
literals ($H$2:$5000) in every formula cell of the builder.
"""
import csv
import datetime
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

ACCOUNTS = ["Cash", "Accounts Receivable", "Inventory", "Prepaid Expenses",
            "Equipment", "Accumulated Depreciation", "Accounts Payable",
            "GST Payable", "Short-term Loan", "Owner's Capital",
            "Retained Earnings", "Drawings", "Sales Revenue",
            "Other Income", "Cost of Goods Sold", "Rent Expense",
            "Salaries Expense", "Utilities Expense", "Depreciation Expense",
            "Marketing Expense", "Bank Charges"]


def close(a, b, tol):
    return abs(a - b) <= tol


def main():
    generate.main()
    build_workbook.main()
    truth = ground_truth.compute()
    n = truth["lines"]
    last = 1 + n + build_workbook.HEADROOM
    tb_last = 1 + len(ACCOUNTS)

    with open(os.path.join(HERE, "build_workbook.py")) as f:
        src = f.read()
    bad = [m.group(0) for m in BAD_RANGE.finditer(src)]
    assert not bad, f"short-form range literals in builder: {bad[:5]}"

    tmp = tempfile.mkdtemp(prefix="close-pack-")
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

    for sheet in ("Dashboard", "Journal", "TrialBalance", "IncomeStatement",
                  "BalanceSheet", "CashFlow", "Monthly", "Budget",
                  "Assumptions", "Checks"):
        for row in wb[sheet].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    failures.append(
                        f"formula error {sheet}!{c.coordinate}: {c.value}")

    # --- Trial balance: closing balances per account. ---
    ws = wb["TrialBalance"]
    for i, acct in enumerate(ACCOUNTS, start=2):
        signed = truth["closing"][acct]
        want_dr = max(signed, 0)
        want_cr = max(-signed, 0)
        check(f"TB {acct} closing Dr", ws.cell(row=i, column=8).value,
              want_dr, 0.01)
        check(f"TB {acct} closing Cr", ws.cell(row=i, column=9).value,
              want_cr, 0.01)
        check(f"TB {acct} Dr movements",
              ws.cell(row=i, column=6).value,
              truth["dr_move"].get(acct, 0), 0.01)
        check(f"TB {acct} Cr movements",
              ws.cell(row=i, column=7).value,
              truth["cr_move"].get(acct, 0), 0.01)
    tr = tb_last + 1
    check("TB total Dr", ws.cell(row=tr, column=8).value, truth["tb_dr"], 0.01)
    check("TB total Cr", ws.cell(row=tr, column=9).value, truth["tb_cr"], 0.01)

    # --- Income statement. ---
    ws = wb["IncomeStatement"]
    check("IS revenue", ws["B4"].value, truth["revenue"], 0.01)
    check("IS COGS", ws["B5"].value, truth["cogs"], 0.01)
    check("IS gross", ws["B6"].value, truth["gross"], 0.01)
    check("IS opex", ws["B13"].value, truth["expenses"], 0.01)
    check("IS net", ws["B14"].value, truth["net_income"], 0.01)

    # --- Balance sheet. ---
    ws = wb["BalanceSheet"]
    check("BS assets", ws["B8"].value, truth["assets"], 0.01)
    check("BS liabilities", ws["B13"].value, truth["liabilities"], 0.01)
    check("BS equity", ws["B20"].value, truth["equity"], 0.01)
    check("BS retained closing", ws["B18"].value, truth["retained"], 0.01)
    check("BS check", ws["B22"].value, 0, 0.005)

    # --- Cash flow. ---
    ws = wb["CashFlow"]
    check("CF operating", ws["B9"].value, truth["cf_operating"], 0.01)
    check("CF investing", ws["B11"].value, truth["cf_investing"], 0.01)
    check("CF financing", ws["B15"].value, truth["cf_financing"], 0.01)
    check("CF total", ws["B16"].value, truth["cf_total"], 0.01)
    check("CF cash change per TB", ws["B19"].value, truth["cash_change"], 0.01)
    check("CF check diff", ws["B20"].value, 0, 0.005)

    # --- Monthly + Budget. ---
    ws = wb["Monthly"]
    for i, (key, _label) in enumerate(build_workbook.MONTHS, start=2):
        m = truth["monthly"][key]
        check(f"Monthly {key} revenue", ws.cell(row=i, column=3).value,
              m["revenue"], 0.01)
        check(f"Monthly {key} COGS", ws.cell(row=i, column=4).value,
              m["cogs"], 0.01)
        check(f"Monthly {key} net", ws.cell(row=i, column=7).value,
              m["net"], 0.01)
    mtot = 2 + len(build_workbook.MONTHS)
    check("Monthly total revenue", ws.cell(row=mtot, column=3).value,
          truth["revenue"], 0.01)
    check("Monthly total net", ws.cell(row=mtot, column=7).value,
          truth["net_income"], 0.01)
    # Journal lines per month, counted independently from the CSV.
    lines = list(csv.DictReader(
        open(os.path.join(HERE, "data", "journal.csv"))))
    from collections import Counter
    per_month = Counter(
        datetime.date.fromisoformat(r["date"]).year * 100
        + datetime.date.fromisoformat(r["date"]).month for r in lines)
    for i, (key, _label) in enumerate(build_workbook.MONTHS, start=2):
        check(f"Monthly {key} lines", ws.cell(row=i, column=8).value,
              per_month[key], 0)

    ws = wb["Budget"]
    for i, (key, _label) in enumerate(build_workbook.MONTHS, start=2):
        b = truth["budget_rows"][key]
        check(f"Budget {key} sales actual", ws.cell(row=i, column=4).value,
              b["sales_actual"], 0.01)
        check(f"Budget {key} net actual", ws.cell(row=i, column=8).value,
              b["net_actual"], 0.01)
        check(f"Budget {key} sales variance", ws.cell(row=i, column=5).value,
              b["sales_actual"] - b["sales_budget"], 0.01)

    # --- Dashboard KPIs + as-of date. ---
    ws = wb["Dashboard"]
    check("Dashboard revenue", ws["B4"].value, truth["revenue"], 0.01)
    check("Dashboard net", ws["B8"].value, truth["net_income"], 0.01)
    check("Dashboard assets", ws["B10"].value, truth["assets"], 0.01)
    check("Dashboard journal lines", ws["B15"].value, truth["lines"], 0)
    if ws["B2"].value != datetime.datetime(2026, 11, 4):
        failures.append(f"Dashboard as-of: got {ws['B2'].value!r}")

    # --- Checks sheet: all 12 must PASS. ---
    ws = wb["Checks"]
    for r in range(2, 14):
        if ws.cell(row=r, column=4).value != "PASS":
            failures.append(f"Checks row {r} ({ws.cell(row=r, column=1).value}): "
                            f"difference {ws.cell(row=r, column=2).value!r} "
                            "not PASS")

    # Headroom: the last Journal row must still carry a live Month Key
    # formula, and the workbook must carry 3 charts.
    wb2 = load_workbook(XLSX)
    fcell = wb2["Journal"].cell(row=last, column=6).value
    if not (isinstance(fcell, str) and fcell.startswith("=")):
        failures.append(f"headroom missing: Journal F{last} = {fcell!r}")
    if len(wb2["Dashboard"]._charts) != 3:
        failures.append(f"expected 3 Dashboard charts, got "
                        f"{len(wb2['Dashboard']._charts)}")

    assert not failures, "FAILURES:\n" + "\n".join(failures)
    print(f"ALL CHECKS PASS ({truth['lines']} journal lines, "
          f"{truth['entries']} entries, {len(ACCOUNTS)} accounts; revenue "
          f"Rs {truth['revenue']:,}, net income Rs {truth['net_income']:,}, "
          f"assets Rs {truth['assets']:,}, BS diff {truth['bs_diff']}, "
          f"cash change Rs {truth['cash_change']:,} = CF total)")


if __name__ == "__main__":
    main()
