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
XLSX = os.path.join(HERE, "lease-disparity.xlsx")
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

    tmp = tempfile.mkdtemp(prefix="leasedisparity-")
    indir, outdir = os.path.join(tmp, "in"), os.path.join(tmp, "out")
    os.makedirs(indir)
    os.makedirs(outdir)
    shutil.copy(XLSX, os.path.join(indir, "lease-disparity.xlsx"))
    r = subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", outdir,
         os.path.join(indir, "lease-disparity.xlsx")],
        capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, f"soffice failed: {r.stderr[:500]}"
    recalc = os.path.join(outdir, "lease-disparity.xlsx")
    assert os.path.exists(recalc), "LO produced no output"
    wb = load_workbook(recalc, data_only=True)

    failures = []

    def check(name, got, want, tol):
        if got is None or isinstance(got, str) or not close(got, want, tol):
            failures.append(f"{name}: got {got!r}, want {want}")

    for sheet in ("Dashboard", "Leases", "Schedule", "Checks", "Assumptions"):
        for row in wb[sheet].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    failures.append(f"formula error {sheet}!{c.coordinate}: {c.value}")

    ws = wb["Dashboard"]
    check("Dashboard leases", ws["B3"].value, truth["count"], 0)
    check("Dashboard contract", ws["B4"].value, truth["total_contract"], 0.01)
    check("Dashboard market", ws["B5"].value, truth["total_market"], 0.01)
    check("Dashboard monthly disparity", ws["B6"].value, truth["total_disparity"], 0.01)
    check("Dashboard annual disparity", ws["B7"].value, truth["total_annual"], 0.01)
    check("Dashboard undiscounted", ws["B8"].value, truth["total_undiscounted"], 0.01)
    check("Dashboard PV", ws["B9"].value, truth["total_pv"], 0.05)
    check("Dashboard under-rented", ws["B10"].value, truth["under"], 0)
    check("Dashboard over-rented", ws["B11"].value, truth["over"], 0)
    check("Dashboard at-market", ws["B12"].value, truth["at_market"], 0)
    check("Dashboard avg months", ws["B13"].value, truth["avg_months"], 1e-6)
    check("Dashboard yield", ws["B14"].value, truth["annual_yield"], 1e-9)

    ws = wb["Leases"]
    for i, lease in enumerate(truth["leases"], start=2):
        check(f"Leases {lease['id']} disparity", ws.cell(row=i, column=8).value,
              lease["disparity"], 0.01)
        check(f"Leases {lease['id']} PV", ws.cell(row=i, column=11).value,
              lease["pv"], 0.05)
        got_status = ws.cell(row=i, column=12).value
        if got_status != lease["status"]:
            failures.append(f"Leases {lease['id']} status: got {got_status!r}, "
                            f"want {lease['status']!r}")
    check("Leases TOTAL PV", ws.cell(row=14, column=11).value, truth["total_pv"], 0.05)

    ws = wb["Schedule"]
    check("Schedule month 1 monthly", ws.cell(row=2, column=2).value,
          truth["schedule"][0]["monthly"], 0.01)
    check("Schedule month 13 monthly", ws.cell(row=14, column=2).value,
          truth["schedule"][12]["monthly"], 0.01)
    check("Schedule month 60 monthly", ws.cell(row=61, column=2).value,
          truth["schedule"][59]["monthly"], 0.01)
    check("Schedule cumulative PV", ws.cell(row=61, column=5).value,
          truth["schedule_pv"], 1.0)

    ws = wb["Checks"]
    for r in range(2, 11):
        if ws.cell(row=r, column=5).value != "PASS":
            failures.append(f"Checks row {r} not PASS: "
                            f"expected={ws.cell(row=r, column=2).value!r} "
                            f"got={ws.cell(row=r, column=3).value!r}")

    print(f"truth: leases={truth['count']} contract=${truth['total_contract']:,.0f}/mo "
          f"market=${truth['total_market']:,.0f}/mo disparity=${truth['total_disparity']:,.0f}/mo "
          f"PV=${truth['total_pv']:,.2f} schedulePV=${truth['schedule_pv']:,.2f} "
          f"under={truth['under']} over={truth['over']} at_market={truth['at_market']}")
    if failures:
        print(f"{len(failures)} FAILURES:")
        for f in failures:
            print(" -", f)
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
