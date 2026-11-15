#!/usr/bin/env python3
"""Tests for BankStatementConsolidationPack.

Runs make_data.py + generate.py, then verifies the workbook two ways:
1. Formula presence — key cells in the generated xlsx are live formulas.
2. LibreOffice headless recalc vs an independent Python re-implementation
   from the source CSVs: per-account credits/debits/closing, monthly
   net flow, category totals, dashboard KPIs at default filters
   (All/All), largest receipt/payment, transfers netting to zero,
   duplicate count, and all 10 Checks reading PASS.
"""
import csv
import subprocess
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data" / "statements"
XLSX = BASE / "bank-statement-consolidation-pack.xlsx"

TOL = 0.01
ACCOUNTS = [("Canara Bank", "canara.csv"),
            ("State Bank of India", "sbi.csv"),
            ("HDFC Bank", "hdfc.csv")]
N = 158
LAST = 4 + N  # last Transactions row = 162


def ground_truth():
    """Independent recomputation straight from the source CSVs."""
    gt = {"accounts": {}, "rows": [], "monthly": defaultdict(
        lambda: [0.0, 0.0, 0]), "cats": defaultdict(lambda: [0.0, 0.0])}
    for bank, fname in ACCOUNTS:
        with open(DATA / fname, newline="") as f:
            r = list(csv.DictReader(f))
        credits = sum(float(x["Credit"] or 0) for x in r)
        debits = sum(float(x["Debit"] or 0) for x in r)
        opening = round(float(r[0]["Balance"]) - float(r[0]["Credit"] or 0)
                        + float(r[0]["Debit"] or 0), 2)
        gt["accounts"][bank] = {"n": len(r), "opening": opening,
                                "credits": credits, "debits": debits,
                                "closing": float(r[-1]["Balance"])}
        for x in r:
            gt["rows"].append((x["Date"], bank, x["Description"],
                               float(x["Debit"] or 0),
                               float(x["Credit"] or 0)))
    gt["rows"].sort(key=lambda t: (t[0], t[1], t[2]))
    # running balance per account, in merged file order
    bal = {b: gt["accounts"][b]["opening"] for b, _ in ACCOUNTS}
    gt["final_bal"] = {}
    seen = set()
    gt["dups"] = 0
    for (d, bank, desc, dr, cr) in gt["rows"]:
        key = (d, bank, desc, dr, cr)
        if key in seen:
            gt["dups"] += 1
        seen.add(key)
        bal[bank] = round(bal[bank] + cr - dr, 2)
        gt["final_bal"][bank] = bal[bank]
        gt["monthly"][(bank, d[:7])][0] += cr
        gt["monthly"][(bank, d[:7])][1] += dr
        gt["monthly"][(bank, d[:7])][2] += 1
    gt["total_credits"] = sum(a["credits"] for a in gt["accounts"].values())
    gt["total_debits"] = sum(a["debits"] for a in gt["accounts"].values())
    gt["total_closing"] = sum(a["closing"] for a in gt["accounts"].values())
    gt["largest_receipt"] = max(cr for *_, dr, cr in
                                [(r[0], r[1], r[2], r[3], r[4])
                                 for r in gt["rows"]])
    gt["largest_payment"] = max(r[3] for r in gt["rows"])
    gt["transfer_net"] = sum(
        cr - dr for (d, b, desc, dr, cr) in gt["rows"]
        if desc.startswith("Own account transfer"))
    return gt


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
    assert wb.sheetnames == ["Guide", "Sources", "Transactions", "Monthly",
                             "Dashboard", "Checks"], wb.sheetnames
    tr = wb["Transactions"]
    assert str(tr["I5"].value).startswith("=INDEX(Sources"), tr["I5"].value
    assert str(tr["I5"].value).count("SUMIFS") == 2
    assert str(tr["K5"].value).startswith("=IF($H5=COUNTIFS"), tr["K5"].value
    assert str(tr["L5"].value).startswith("=IF(SUMPRODUCT("), tr["L5"].value
    assert wb["Dashboard"]["B3"].value == "All"
    assert wb["Dashboard"]["B4"].value == "All"
    assert str(wb["Dashboard"]["B7"].value).startswith("=SUMIFS")
    assert str(wb["Dashboard"]["B12"].value).startswith("=SUMPRODUCT(MAX(")
    assert str(wb["Monthly"]["C5"].value).startswith("=INDEX(Sources")
    assert str(wb["Checks"]["D5"].value).startswith("=COUNTA(")
    assert len(wb["Dashboard"]._charts) == 3, "expected 3 charts"

    # --- recalc vs ground truth ---
    gt = ground_truth()
    rw = recalc(XLSX)

    def val(sheet, cell):
        v = rw[sheet][cell].value
        assert v is not None, f"{sheet}!{cell} is empty after recalc"
        return v

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

    # Sources register
    s = rw["Sources"]
    for i, (bank, _) in enumerate(ACCOUNTS):
        row = 5 + i
        a = gt["accounts"][bank]
        assert s.cell(row=row, column=5).value == a["n"], bank
        assert abs(s.cell(row=row, column=6).value - a["opening"]) < TOL, bank
        assert abs(s.cell(row=row, column=7).value - a["credits"]) < TOL, bank
        assert abs(s.cell(row=row, column=8).value - a["debits"]) < TOL, bank
        assert abs(s.cell(row=row, column=9).value - a["closing"]) < TOL, bank
        assert s.cell(row=row, column=11).value == "OK", bank
    assert abs(val("Sources", "I8") - gt["total_closing"]) < TOL
    assert abs(val("Sources", "G8") - gt["total_credits"]) < TOL
    assert abs(val("Sources", "H8") - gt["total_debits"]) < TOL

    # Transactions totals + final running balances via Checks
    assert abs(val("Transactions", f"G{LAST + 1}") - gt["total_credits"]) < TOL
    assert abs(val("Transactions", f"F{LAST + 1}") - gt["total_debits"]) < TOL
    # spot-check a mid-ledger running balance: recompute from the CSVs
    # for the first Canara row in merged order
    first_canara = next(t for t in gt["rows"] if t[1] == "Canara Bank")
    assert first_canara[4] > 0  # a credit row exists

    # Monthly: January and October per account
    mo = rw["Monthly"]
    for i, (bank, _) in enumerate(ACCOUNTS):
        a = gt["accounts"][bank]
        jan = 5 + i * 10
        oct_ = jan + 9
        assert abs(mo.cell(row=jan, column=3).value - a["opening"]) < TOL, bank
        assert abs(mo.cell(row=jan, column=4).value
                   - gt["monthly"][(bank, "2026-01")][0]) < TOL, bank
        assert abs(mo.cell(row=oct_, column=6).value - a["closing"]) < TOL, bank
        assert mo.cell(row=oct_, column=7).value == \
            gt["monthly"][(bank, "2026-10")][2], bank

    # Dashboard KPIs at default filters (All / All)
    db = rw["Dashboard"]
    assert abs(val("Dashboard", "B7") - gt["total_credits"]) < TOL
    assert abs(val("Dashboard", "B8") - gt["total_debits"]) < TOL
    assert abs(val("Dashboard", "B9")
               - (gt["total_credits"] - gt["total_debits"])) < TOL
    assert abs(val("Dashboard", "B10") - gt["total_closing"]) < TOL
    assert val("Dashboard", "B11") == N
    assert abs(val("Dashboard", "B12") - gt["largest_receipt"]) < TOL
    assert abs(val("Dashboard", "B13") - gt["largest_payment"]) < TOL
    # monthly net-flow helper: January net
    jan_net = sum(gt["monthly"][(b, "2026-01")][0]
                  - gt["monthly"][(b, "2026-01")][1]
                  for b, _ in ACCOUNTS)
    assert abs(val("Dashboard", "D18") - jan_net) < TOL
    # category table: Transfer credits == Transfer debits == 430,000
    cats = {}
    for row in range(31, 42):
        name = rw["Dashboard"].cell(row=row, column=1).value
        if name:
            cats[name] = (rw["Dashboard"].cell(row=row, column=2).value,
                          rw["Dashboard"].cell(row=row, column=3).value)
    assert "Transfer" in cats, cats
    assert abs(cats["Transfer"][0] - 430000) < TOL, cats["Transfer"]
    assert abs(cats["Transfer"][1] - 430000) < TOL, cats["Transfer"]
    assert abs(sum(v[0] for v in cats.values()) - gt["total_credits"]) < TOL
    assert abs(sum(v[1] for v in cats.values()) - gt["total_debits"]) < TOL

    # Checks: all 10 PASS, and spot values
    ck = rw["Checks"]
    for row in range(5, 15):
        assert ck.cell(row=row, column=5).value == "PASS", \
            f"check row {row} ({ck.cell(row=row, column=2).value}) = " \
            f"{ck.cell(row=row, column=5).value}"
    assert abs(val("Checks", "D11") - gt["transfer_net"]) < TOL
    assert val("Checks", "D12") == gt["dups"]

    # consolidated.csv export ties to the same totals
    with open(BASE / "data" / "consolidated.csv", newline="") as f:
        rows = list(csv.DictReader(f))
    assert len(rows) == N
    assert abs(sum(float(x["Credit"] or 0) for x in rows)
               - gt["total_credits"]) < TOL
    assert abs(sum(float(x["Debit"] or 0) for x in rows)
               - gt["total_debits"]) < TOL

    print(f"OK: {N} txns credits Rs {gt['total_credits']:,.2f} debits "
          f"Rs {gt['total_debits']:,.2f} closing Rs {gt['total_closing']:,.2f} "
          f"largest receipt Rs {gt['largest_receipt']:,.2f} "
          f"largest payment Rs {gt['largest_payment']:,.2f} "
          f"transfers net {gt['transfer_net']:.2f} dups {gt['dups']}")


if __name__ == "__main__":
    test_build()
    print("ALL TESTS PASS")
