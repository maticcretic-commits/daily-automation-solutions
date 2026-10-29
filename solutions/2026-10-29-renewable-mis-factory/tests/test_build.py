#!/usr/bin/env python3
"""Build the workbook, recalc it headlessly in LibreOffice, and compare
every checked figure against the independent Python ground truth.

LO quirk honoured: --convert-to runs on a COPY in a separate input dir
with --outdir elsewhere (converting in place overwrites the input file).
Formulas are A1-style only; a formula error in the recalculated copy is a
failure, not a warning. A second regression scan rejects short-form range
literals ($H$2:$5000) in every formula cell.
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
XLSX = os.path.join(HERE, "renewable-mis.xlsx")
ERRORS = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "Err:522")
BAD_RANGE = re.compile(r"\$[A-Z]+\$\d+:\$\d+")


def close(a, b, tol):
    return abs(a - b) <= tol


def main():
    generate.main()
    build_workbook.main()
    truth = ground_truth.compute()

    # Regression: no short-form range literal in any formula of the builder.
    with open(os.path.join(HERE, "build_workbook.py")) as f:
        src = f.read()
    bad = [m.group(0) for m in BAD_RANGE.finditer(src)]
    assert not bad, f"short-form range literals in builder: {bad[:5]}"

    tmp = tempfile.mkdtemp(prefix="renewablemis-")
    indir, outdir = os.path.join(tmp, "in"), os.path.join(tmp, "out")
    os.makedirs(indir)
    os.makedirs(outdir)
    shutil.copy(XLSX, os.path.join(indir, "renewable-mis.xlsx"))
    r = subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", outdir,
         os.path.join(indir, "renewable-mis.xlsx")],
        capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, f"soffice failed: {r.stderr[:500]}"
    recalc = os.path.join(outdir, "renewable-mis.xlsx")
    assert os.path.exists(recalc), "LO produced no output"
    wb = load_workbook(recalc, data_only=True)

    failures = []

    def check(name, got, want, tol):
        if got is None or isinstance(got, str) or not close(got, want, tol):
            failures.append(f"{name}: got {got!r}, want {want}")

    for sheet in ("Dashboard", "Monthly", "State", "Technology", "Checks"):
        for row in wb[sheet].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    failures.append(f"formula error {sheet}!{c.coordinate}: {c.value}")

    # Dashboard KPIs: B3 sites, B4 planned, B5 actual, B6 achievement,
    # B7 planned rev, B8 actual rev, B9 rev achievement, B10 curtailment,
    # B11 curtailment %, B12 outage, B13 availability.
    ws = wb["Dashboard"]
    check("Dashboard sites", ws["B3"].value, truth["sites"], 0)
    check("Dashboard planned MWh", ws["B4"].value, truth["planned_mwh"], 0)
    check("Dashboard actual MWh", ws["B5"].value, truth["actual_mwh"], 0)
    check("Dashboard achievement", ws["B6"].value, truth["achievement"], 1e-6)
    check("Dashboard planned rev", ws["B7"].value, truth["planned_rev"], 0.02)
    check("Dashboard actual rev", ws["B8"].value, truth["actual_rev"], 0.02)
    check("Dashboard rev achievement", ws["B9"].value, truth["rev_achievement"], 1e-6)
    check("Dashboard curtailment", ws["B10"].value, truth["curtailment"], 0)
    check("Dashboard curtailment %", ws["B11"].value, truth["curtailment_pct"], 1e-6)
    check("Dashboard outage", ws["B12"].value, truth["outage"], 0)
    check("Dashboard availability", ws["B13"].value, truth["availability"], 1e-6)

    # Monthly rows 2-8 in MONTHS order.
    ws = wb["Monthly"]
    for i, m in enumerate(["2026-04", "2026-05", "2026-06", "2026-07",
                           "2026-08", "2026-09", "2026-10"], start=2):
        t = truth["monthly"][m]
        check(f"Monthly {m} planned", ws.cell(row=i, column=3).value, t["planned"], 0)
        check(f"Monthly {m} actual", ws.cell(row=i, column=4).value, t["actual"], 0)
        check(f"Monthly {m} achievement", ws.cell(row=i, column=6).value, t["achievement"], 1e-6)
        check(f"Monthly {m} planned rev", ws.cell(row=i, column=7).value, t["planned_rev"], 0.02)
        check(f"Monthly {m} actual rev", ws.cell(row=i, column=8).value, t["actual_rev"], 0.02)
        check(f"Monthly {m} availability", ws.cell(row=i, column=12).value, t["availability"], 1e-6)
    check("Monthly TOTAL planned", ws["C9"].value, truth["planned_mwh"], 0)
    check("Monthly TOTAL actual", ws["D9"].value, truth["actual_mwh"], 0)

    # State rows 2-12 in STATES order; Technology rows 2-5 in TECHS order.
    ws = wb["State"]
    states = ["Rajasthan", "Gujarat", "Maharashtra", "Karnataka", "Tamil Nadu",
              "Telangana", "Andhra Pradesh", "Madhya Pradesh", "Uttar Pradesh",
              "UAE", "Singapore"]
    for i, s in enumerate(states, start=2):
        t = truth["state"][s]
        check(f"State {s} planned", ws.cell(row=i, column=3).value, t["planned"], 0)
        check(f"State {s} actual", ws.cell(row=i, column=4).value, t["actual"], 0)
        check(f"State {s} achievement", ws.cell(row=i, column=5).value, t["achievement"], 1e-6)
    ws = wb["Technology"]
    for i, tech in enumerate(["Solar", "Wind", "Hybrid", "Rooftop"], start=2):
        t = truth["technology"][tech]
        check(f"Tech {tech} planned", ws.cell(row=i, column=3).value, t["planned"], 0)
        check(f"Tech {tech} actual", ws.cell(row=i, column=4).value, t["actual"], 0)
        check(f"Tech {tech} achievement", ws.cell(row=i, column=5).value, t["achievement"], 1e-6)

    # Checks sheet: every Status must read PASS.
    ws = wb["Checks"]
    for r in range(2, 10):
        if ws.cell(row=r, column=5).value != "PASS":
            failures.append(f"Checks row {r} not PASS: "
                            f"expected={ws.cell(row=r, column=2).value!r} "
                            f"got={ws.cell(row=r, column=3).value!r}")

    print(f"truth: sites={truth['sites']} planned={truth['planned_mwh']} MWh "
          f"actual={truth['actual_mwh']} MWh achievement={truth['achievement']:.1%} "
          f"planned_rev=Rs {truth['planned_rev']:,.2f}L actual_rev=Rs {truth['actual_rev']:,.2f}L "
          f"curtailment={truth['curtailment']} MWh availability={truth['availability']:.2%}")
    if failures:
        print(f"{len(failures)} FAILURES:")
        for f in failures:
            print(" -", f)
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
