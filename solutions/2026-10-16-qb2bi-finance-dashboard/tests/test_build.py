"""pytest suite for the QB2BI finance dashboard.

Includes a LibreOffice headless recalc: the workbook is re-saved by LO and every
key figure is compared against an independent Python ground truth computed
directly from data/journals.csv (not via the builder's own code paths).
"""
import csv
import shutil
import sqlite3
import subprocess
from pathlib import Path

import pytest
from openpyxl import load_workbook

HERE = Path(__file__).parent.parent
XLSX = HERE / "qb2bi-finance-dashboard.xlsx"
DB = HERE / "qb2bi.sqlite"


def journal_lines():
    with open(HERE / "data" / "journals.csv", newline="") as f:
        return list(csv.DictReader(f))


def gt():
    """Independent ground truth straight from the CSV."""
    lines = journal_lines()
    def net(acct, upto=None, side="dr"):
        t = 0.0
        for r in lines:
            if r["acct_no"] != acct:
                continue
            if upto and r["date"] > upto:
                continue
            dr, cr = float(r["debit"]), float(r["credit"])
            t += dr - cr if side == "dr" else cr - dr
        return round(t, 2)
    rev = net("4000", side="cr")
    cogs = net("5000", side="dr")
    exps = sum(net(a, side="dr") for a in
               ["5100", "5200", "5300", "5400", "5500", "5600", "5700", "5800"])
    np_ = round(rev - cogs - exps, 2)
    assets = round(sum(net(a, "2026-09-30", "dr") for a in
                       ["1000", "1100", "1200", "1300", "1500", "1510"]), 2)
    liab = round(sum(net(a, "2026-09-30", "cr") for a in ["2000", "2100", "2200"]), 2)
    equity = round(net("3000", "2026-09-30", "cr") + net("3100", "2026-09-30", "cr") + np_, 2)
    cash = net("1000", "2026-09-30", "dr")
    return dict(rev=rev, cogs=cogs, exps=round(exps, 2), np=np_,
                assets=assets, liab=liab, equity=equity, cash=cash)


@pytest.fixture(scope="module")
def recalc():
    """Recalculate the workbook with LibreOffice headless; return data_only wb."""
    outdir = HERE / "tests" / "_recalc"
    srcdir = HERE / "tests" / "_recalc_src"
    for d in (outdir, srcdir):
        if d.exists():
            shutil.rmtree(d)
        d.mkdir(parents=True)
    src = srcdir / "recalc_source.xlsx"
    shutil.copy(XLSX, src)  # LO keeps the input basename: outdir must differ from srcdir
    r = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx",
                        "--outdir", str(outdir), str(src)],
                       capture_output=True, text=True, timeout=300)
    assert (outdir / "recalc_source.xlsx").exists(), f"LO convert failed: {r.stderr[:500]}"
    return load_workbook(outdir / "recalc_source.xlsx", data_only=True)


def test_files_exist():
    for p in ["generate.py", "build_db.py", "build_workbook.py", "README.md",
              "qb2bi.sqlite", "qb2bi-finance-dashboard.xlsx",
              "data/chart_of_accounts.csv", "data/journals.csv",
              "data/customers.csv", "data/vendors.csv"]:
        assert (HERE / p).exists(), f"missing {p}"


def test_ledger_balances():
    con = sqlite3.connect(DB)
    dr, cr = con.execute("SELECT ROUND(SUM(debit),2), ROUND(SUM(credit),2) "
                         "FROM journal_lines").fetchone()
    con.close()
    assert abs(dr - cr) < 0.01, f"ledger out of balance {dr} vs {cr}"
    assert dr > 1_000_000, "dataset suspiciously small"


def test_sheet_inventory():
    wb = load_workbook(XLSX)
    expect = {"Guide", "Raw_Journals", "TrialBalance", "PL", "BalanceSheet",
              "Dashboard", "AI_Report", "Checks"}
    assert expect <= set(wb.sheetnames), wb.sheetnames


def test_all_checks_pass(recalc):
    ws = recalc["Checks"]
    results = [ws.cell(row=r, column=3).value for r in range(4, 12)]
    assert results == ["PASS"] * 8, f"checks failed: {results}"


def test_pl_figures_match(recalc):
    g = gt()
    ws = recalc["PL"]
    # YTD col T = 20; rows: rev 4, cogs 5, gp 6, texp 15, np 16
    assert abs(ws.cell(row=4, column=20).value - g["rev"]) < 1.0, "YTD revenue"
    assert abs(ws.cell(row=5, column=20).value - g["cogs"]) < 1.0, "YTD COGS"
    assert abs(ws.cell(row=16, column=20).value - g["np"]) < 1.0, "YTD net profit"
    assert abs(ws.cell(row=15, column=20).value - g["exps"]) < 1.0, "YTD expenses"


def test_balance_sheet_matches(recalc):
    g = gt()
    ws = recalc["BalanceSheet"]
    assert abs(ws.cell(row=9, column=3).value - g["assets"]) < 1.0, "total assets"
    assert abs(ws.cell(row=14, column=3).value - g["liab"]) < 1.0, "total liabilities"
    assert abs(ws.cell(row=19, column=3).value - g["equity"]) < 1.0, "total equity"
    assert ws.cell(row=21, column=3).value == "✓ BALANCED"
    assert abs(g["assets"] - (g["liab"] + g["equity"])) < 0.01, "A = L + E ground truth"


def test_dashboard_kpis_match(recalc):
    g = gt()
    ws = recalc["Dashboard"]
    # KPI values live in A4, A6, A8, A10, A12, A14, A16, A18, A20
    vals = [ws.cell(row=r, column=1).value for r in (4, 6, 12, 14)]
    assert abs(vals[0] - g["rev"]) < 1.0, "KPI revenue"
    assert abs(vals[1] - g["np"]) < 1.0, "KPI net profit"
    assert abs(vals[2] - g["cash"]) < 1.0, "KPI cash"
    gm = ws.cell(row=8, column=1).value
    assert abs(gm - (g["rev"] - g["cogs"]) / g["rev"]) < 0.001, "KPI gross margin"


def test_tb_ties_to_ledger(recalc):
    g = gt()
    ws = recalc["TrialBalance"]
    # revenue row: acct 4000 is row 16 (21 accounts, alpha order -> 4000 is 13th: row 4+12=16)
    row4000 = next(r for r in range(4, 30)
                   if ws.cell(row=r, column=1).value == "4000")
    net_rev = ws.cell(row=row4000, column=7).value - ws.cell(row=row4000, column=6).value
    assert abs(net_rev - g["rev"]) < 1.0, "TB revenue ties to ground truth"


def test_charts_present():
    wb = load_workbook(XLSX)
    assert len(wb["Dashboard"]._charts) == 3, "dashboard needs 3 charts"


def test_ai_report_sections():
    wb = load_workbook(XLSX, data_only=True)
    text = "\n".join(str(c.value or "") for row in wb["AI_Report"].iter_rows()
                     for c in row)
    for s in ["Executive summary", "Revenue trend", "Expense drivers",
              "Cash & working capital", "Recommendations"]:
        assert s in text, f"AI_Report missing section: {s}"
