#!/usr/bin/env python3
"""Build the workbook, recalc it headlessly in LibreOffice, and compare
every checked figure against the independent Python ground truth.

LO quirk honoured: --convert-to runs on a COPY in a separate input dir
with --outdir elsewhere (converting in place overwrites the input file).
Formulas are A1-style only; a formula error in the recalculated copy is a
failure, not a warning. A second regression scan rejects short-form range
literals ($H$2:$5000) in every formula cell of the builder.
"""
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
XLSX = os.path.join(HERE, "payment-tracker.xlsx")
ERRORS = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "Err:522")
BAD_RANGE = re.compile(r"\$[A-Z]+\$\d+:\$\d+")


def close(a, b, tol):
    return abs(a - b) <= tol


def main():
    generate.main()
    build_workbook.main()
    truth = ground_truth.compute()
    n = truth["count"]
    last = 1 + n + build_workbook.HEADROOM
    cust_last = 1 + truth["customers"]

    with open(os.path.join(HERE, "build_workbook.py")) as f:
        src = f.read()
    bad = [m.group(0) for m in BAD_RANGE.finditer(src)]
    assert not bad, f"short-form range literals in builder: {bad[:5]}"

    tmp = tempfile.mkdtemp(prefix="paytracker-")
    indir, outdir = os.path.join(tmp, "in"), os.path.join(tmp, "out")
    os.makedirs(indir)
    os.makedirs(outdir)
    shutil.copy(XLSX, os.path.join(indir, "payment-tracker.xlsx"))
    r = subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", outdir,
         os.path.join(indir, "payment-tracker.xlsx")],
        capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, f"soffice failed: {r.stderr[:500]}"
    recalc = os.path.join(outdir, "payment-tracker.xlsx")
    assert os.path.exists(recalc), "LO produced no output"
    wb = load_workbook(recalc, data_only=True)

    failures = []

    def check(name, got, want, tol):
        if got is None or isinstance(got, str) or not close(got, want, tol):
            failures.append(f"{name}: got {got!r}, want {want}")

    for sheet in ("Dashboard", "Payments", "Customers", "Monthly",
                  "Overdue", "Checks", "Assumptions"):
        for row in wb[sheet].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    failures.append(
                        f"formula error {sheet}!{c.coordinate}: {c.value}")

    ws = wb["Dashboard"]
    check("Dashboard total invoiced", ws["B4"].value, truth["tot_invoiced"], 0.01)
    check("Dashboard total collected", ws["B5"].value, truth["tot_collected"], 0.01)
    check("Dashboard total pending", ws["B6"].value, truth["tot_pending"], 0.01)
    check("Dashboard collection rate", ws["B7"].value, truth["collection_rate"], 1e-6)
    check("Dashboard invoices", ws["B8"].value, truth["count"], 0)
    check("Dashboard paid", ws["B9"].value, truth["paid_count"], 0)
    check("Dashboard overdue count", ws["B10"].value, truth["overdue_count"], 0)
    check("Dashboard overdue amount", ws["B11"].value, truth["overdue_amt"], 0.01)
    check("Dashboard customers", ws["B12"].value, truth["customers"], 0)
    check("Dashboard avg overdue", ws["B13"].value, truth["avg_overdue"], 0.01)
    check("Dashboard partial", ws["B14"].value, truth["partial_count"], 0)
    if ws["B2"].value != datetime.datetime(2026, 11, 3):
        failures.append(f"Dashboard as-of: got {ws['B2'].value!r}")
    for i, status in enumerate(build_workbook.STATUSES, start=17):
        check(f"Status mix {status}", ws.cell(row=i, column=2).value,
              truth["status_counts"].get(status, 0), 0)
    check("Status total", ws["B22"].value, truth["count"], 0)

    ws = wb["Monthly"]
    check("Monthly total invoiced", ws["C9"].value, truth["tot_invoiced"], 0.01)
    check("Monthly total collected", ws["D9"].value, truth["tot_collected"], 0.01)
    check("Monthly total pending", ws["E9"].value, truth["tot_pending"], 0.01)
    check("Monthly cumulative", ws["G9"].value, truth["tot_collected"], 0.01)
    for i, (key, _label) in enumerate(build_workbook.MONTHS, start=2):
        m = truth["monthly"][key]
        check(f"Monthly {key} invoiced", ws.cell(row=i, column=3).value,
              m["invoiced"], 0.01)
        check(f"Monthly {key} collected", ws.cell(row=i, column=4).value,
              m["collected"], 0.01)
        check(f"Monthly {key} pending", ws.cell(row=i, column=5).value,
              m["pending"], 0.01)
        check(f"Monthly {key} count", ws.cell(row=i, column=8).value,
              m["count"], 0)

    ws = wb["Customers"]
    check("Customers total invoiced", ws.cell(row=cust_last + 1, column=3).value,
          truth["tot_invoiced"], 0.01)
    check("Customers total pending", ws.cell(row=cust_last + 1, column=5).value,
          truth["tot_pending"], 0.01)
    for i, (cust, c) in enumerate(sorted(truth["customers_detail"].items()),
                                  start=2):
        check(f"Customer {cust} invoiced", ws.cell(row=i, column=3).value,
              c["invoiced"], 0.01)
        check(f"Customer {cust} pending", ws.cell(row=i, column=5).value,
              c["pending"], 0.01)
        check(f"Customer {cust} overdue amt", ws.cell(row=i, column=6).value,
              c["overdue_amt"], 0.01)

    ws = wb["Overdue"]
    listed = [ws.cell(row=r, column=2).value
              for r in range(2, 52) if ws.cell(row=r, column=2).value]
    want_list = [r["invoice_no"] for r in truth["overdue_list"][:50]]
    if listed != want_list:
        failures.append(f"Overdue list mismatch: got {listed[:5]}..., "
                        f"want {want_list[:5]}... ({len(listed)} listed, "
                        f"{len(want_list)} expected)")

    ws = wb["Checks"]
    for r in range(2, 12):
        if ws.cell(row=r, column=4).value != "PASS":
            failures.append(f"Checks row {r} ({ws.cell(row=r, column=1).value}): "
                            f"result {ws.cell(row=r, column=2).value!r} not PASS")

    # Headroom: the last Payments row must still carry live formulas, and
    # the workbook must carry 3 charts and conditional-formatting rules.
    wb2 = load_workbook(XLSX)
    gcell = wb2["Payments"].cell(row=last, column=7).value
    if not (isinstance(gcell, str) and gcell.startswith("=")):
        failures.append(f"headroom missing: Payments G{last} = {gcell!r}")
    if len(wb2["Dashboard"]._charts) != 3:
        failures.append(f"expected 3 Dashboard charts, got "
                        f"{len(wb2['Dashboard']._charts)}")
    cf_rules = sum(len(v) for v in
                   wb2["Payments"].conditional_formatting._cf_rules.values())
    if cf_rules < 5:
        failures.append(f"Payments conditional-formatting rules missing: {cf_rules}")

    assert not failures, "FAILURES:\n" + "\n".join(failures)
    print(f"ALL CHECKS PASS ({truth['count']} invoices, "
          f"{truth['customers']} customers; invoiced "
          f"Rs {truth['tot_invoiced']:,}, collected "
          f"Rs {truth['tot_collected']:,}, pending "
          f"Rs {truth['tot_pending']:,}, overdue {truth['overdue_count']} / "
          f"Rs {truth['overdue_amt']:,})")


if __name__ == "__main__":
    main()
