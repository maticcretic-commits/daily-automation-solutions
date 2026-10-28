#!/usr/bin/env python3
"""Build the workbook, recalc it headlessly in LibreOffice, and compare
every checked figure against the independent Python ground truth.

LO quirk honoured: --convert-to runs on a COPY in a separate input dir
with --outdir elsewhere (converting in place overwrites the input file).
Formulas are A1-style only; a formula error in the recalculated copy is a
failure, not a warning.
"""
import os
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
XLSX = os.path.join(HERE, "statementpulse.xlsx")
ERRORS = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "Err:522")


def close(a, b, tol):
    return abs(a - b) <= tol


def main():
    generate.main()
    build_workbook.main()
    truth = ground_truth.compute()

    tmp = tempfile.mkdtemp(prefix="statementpulse-")
    indir, outdir = os.path.join(tmp, "in"), os.path.join(tmp, "out")
    os.makedirs(indir)
    os.makedirs(outdir)
    shutil.copy(XLSX, os.path.join(indir, "statementpulse.xlsx"))
    r = subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", outdir,
         os.path.join(indir, "statementpulse.xlsx")],
        capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, f"soffice failed: {r.stderr[:500]}"
    recalc = os.path.join(outdir, "statementpulse.xlsx")
    assert os.path.exists(recalc), "LO produced no output"
    wb = load_workbook(recalc, data_only=True)

    failures = []

    def check(name, got, want, tol):
        if got is None or isinstance(got, str) or not close(got, want, tol):
            failures.append(f"{name}: got {got!r}, want {want}")

    for sheet in ("Dashboard", "Statements", "Ratios", "Trend", "Checks"):
        for row in wb[sheet].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    failures.append(f"formula error {sheet}!{c.coordinate}: {c.value}")

    # Statements sheet rows: P&L 3-13, BS 15-33, CF 35-48 (see builder).
    ws = wb["Statements"]
    st_rows = {"Revenue": 3, "Cost of goods sold": 4, "Gross profit": 5,
               "Operating expenses": 6, "EBITDA": 7, "Depreciation": 8,
               "EBIT": 9, "Interest expense": 10, "Profit before tax": 11,
               "Tax expense": 12, "Profit after tax": 13,
               "Cash and bank": 15, "Trade receivables": 16,
               "Inventories": 17, "Other current assets": 18,
               "Total current assets": 19,
               "Property plant and equipment (net)": 20,
               "Other non-current assets": 21,
               "Total non-current assets": 22, "TOTAL ASSETS": 23,
               "Trade payables": 24, "Short-term borrowings": 25,
               "Other current liabilities": 26,
               "Total current liabilities": 27, "Long-term borrowings": 28,
               "Share capital": 29, "Retained earnings": 30,
               "Total equity": 31, "TOTAL LIABILITIES + EQUITY": 32,
               "Balance check (assets - L+E; must be 0)": 33,
               "Cash flow from operations": 41, "Capital expenditure": 42,
               "Change in borrowings": 43, "Dividends paid": 44,
               "Cash flow from financing": 45, "Net change in cash": 46,
               "Opening cash": 47, "Closing cash": 48,
               "Closing vs balance-sheet cash (must be 0)": 49}
    for year, col in (("FY2024", 2), ("FY2025", 3), ("FY2026", 4)):
        t = truth["years"][year]
        check(f"{year} revenue", ws.cell(row=3, column=col).value, t["revenue"], 1.0)
        check(f"{year} gross", ws.cell(row=5, column=col).value, t["gross"], 1.0)
        check(f"{year} ebitda", ws.cell(row=7, column=col).value, t["ebitda"], 1.0)
        check(f"{year} ebit", ws.cell(row=9, column=col).value, t["ebit"], 1.0)
        check(f"{year} pat", ws.cell(row=13, column=col).value, t["pat"], 1.0)
        check(f"{year} total assets", ws.cell(row=23, column=col).value, t["total_assets"], 1.0)
        check(f"{year} total L+E", ws.cell(row=32, column=col).value, t["total_le"], 1.0)
        check(f"{year} balance gap", ws.cell(row=33, column=col).value, 0, 1.0)
        check(f"{year} CFO", ws.cell(row=41, column=col).value, t["cfo"], 1.0)
        check(f"{year} closing cash", ws.cell(row=48, column=col).value, t["closing_cash"], 1.0)
        check(f"{year} CF tie", ws.cell(row=49, column=col).value, 0, 1.0)

    # Ratios sheet rows 2-14 in builder order.
    ws = wb["Ratios"]
    ratio_keys = ["current_ratio", "quick_ratio", "debt_equity",
                  "interest_coverage", "gross_margin", "net_margin", "roa",
                  "roe", "recv_days", "inv_days", "pay_days",
                  "cash_conversion", "working_capital"]
    for year, col in (("FY2024", 2), ("FY2025", 3), ("FY2026", 4)):
        t = truth["years"][year]
        for i, key in enumerate(ratio_keys):
            tol = 1e-3 if "margin" in key or key in ("roa", "roe") else 0.01
            if key == "working_capital":
                tol = 1.0
            check(f"{year} {key}", ws.cell(row=2 + i, column=col).value, t[key], tol)

    # Trend sheet: rows 2-13 months, row 14 total.
    ws = wb["Trend"]
    for i, m in enumerate(truth["monthly"]):
        r = i + 2
        assert ws.cell(row=r, column=1).value == m["Month"], f"month row {r}"
        check(f"{m['Month']} revenue", ws.cell(row=r, column=2).value, m["Revenue"], 1.0)
        check(f"{m['Month']} gross", ws.cell(row=r, column=4).value,
              m["Revenue"] - m["COGS"], 1.0)
        check(f"{m['Month']} surplus", ws.cell(row=r, column=6).value,
              m["Revenue"] - m["COGS"] - m["Opex"], 1.0)
        check(f"{m['Month']} WC", ws.cell(row=r, column=11).value,
              m["Receivables"] + m["Inventory"] - m["Payables"], 1.0)
    check("trend total revenue", ws.cell(row=14, column=2).value,
          truth["trend_total_revenue"], 1.0)
    check("trend total COGS", ws.cell(row=14, column=3).value,
          truth["trend_total_cogs"], 1.0)

    # Dashboard cards B4:B9.
    ws = wb["Dashboard"]
    d = truth["dashboard"]
    check("dash revenue", ws["B4"].value, d["revenue"], 1.0)
    check("dash PAT", ws["B5"].value, d["pat"], 1.0)
    check("dash net margin", ws["B6"].value, d["net_margin"], 1e-3)
    check("dash current ratio", ws["B7"].value, d["current_ratio"], 1e-3)
    check("dash recv days", ws["B8"].value, d["recv_days"], 0.01)
    check("dash conversion", ws["B9"].value, d["cash_conversion"], 1e-3)

    # Checks: all 8 must read TRUE.
    ws = wb["Checks"]
    for r in range(2, 10):
        if ws.cell(row=r, column=2).value is not True:
            failures.append(f"check row {r} ({ws.cell(row=r, column=1).value}): "
                            f"{ws.cell(row=r, column=2).value!r}")

    shutil.rmtree(tmp, ignore_errors=True)
    if failures:
        print(f"FAIL ({len(failures)}):")
        for f in failures:
            print("  -", f)
        return 1
    print("ALL PASS — statements, ratios, trend, dashboard, checks (8), "
          "zero formula errors")
    t = truth["years"]["FY2026"]
    print(f"FY2026 revenue {t['revenue']:,} · PAT {t['pat']:,} · net margin "
          f"{t['net_margin']:.1%} · current ratio {t['current_ratio']:.2f} · "
          f"receivables {t['recv_days']:.0f} days · CFO/PAT {t['cash_conversion']:.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
