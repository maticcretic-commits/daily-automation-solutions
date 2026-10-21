"""DPRPulse tests: LO headless recalc vs independent Python ground truth."""
import os
import shutil
import subprocess
import sys

import pytest
from openpyxl import load_workbook

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from main import ground_truth, SCENARIOS  # noqa: E402

ROOT = os.path.join(os.path.dirname(__file__), "..")
XLSX = os.path.join(ROOT, "dprpulse.xlsx")
RECALC_DIR = os.path.join(ROOT, "recalc-out")
ERRORS = ("#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#REF!", "#NUM!", "Err:509",
          "Err:508", "Err:502")


@pytest.fixture(scope="module")
def recalc():
    if os.path.isdir(RECALC_DIR):
        shutil.rmtree(RECALC_DIR)
    os.makedirs(RECALC_DIR, exist_ok=True)
    r = subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "--outdir",
         RECALC_DIR, XLSX],
        capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, r.stderr[-2000:]
    out = os.path.join(RECALC_DIR, "dprpulse.xlsx")
    assert os.path.exists(out)
    return load_workbook(out, data_only=True)


@pytest.fixture(scope="module")
def gt():
    return ground_truth()


def col(t, y0=False):
    # year t -> 1-based column: sheets with a Y0 column use t+2 (Y0->B),
    # sheets starting at Y1 use t+1 (Y1->B)
    return (t + 2) if y0 else (t + 1)


def test_files_exist():
    for p in ["dprpulse.xlsx", "README.md", "generate.py",
              "build_workbook.py", "main.py",
              "data/assumptions.csv", "data/project_cost.csv",
              "data/utilisation.csv", "tests/test_build.py"]:
        assert os.path.exists(os.path.join(ROOT, p)), p


def test_no_assistant_mentions():
    import re
    for p in ["README.md", "generate.py", "build_workbook.py", "main.py"]:
        text = open(os.path.join(ROOT, p)).read()
        assert "Muse" not in text, p
        assert not re.search(r"(?i)\bmuse\b", text), p


def test_no_formula_errors(recalc):
    bad = []
    for ws in recalc.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith(("#", "Err:")):
                    bad.append(f"{ws.title}!{cell.coordinate}={v}")
    assert not bad, bad[:10]


def test_project_cost(recalc, gt):
    ws = recalc["Project_Cost"]
    assert ws["B9"].value == pytest.approx(gt["project_cost"], abs=0.01)
    assert ws["B11"].value is True


def test_means_of_finance(recalc, gt):
    ws = recalc["Means_of_Finance"]
    assert ws["B4"].value == pytest.approx(gt["project_cost"], abs=0.01)
    assert ws["B6"].value == pytest.approx(gt["promoter_margin"], abs=1e-6)
    assert ws["B7"].value == pytest.approx(3.0, abs=1e-6)
    assert ws["B9"].value is True


def test_loan_schedule(recalc, gt):
    ws = recalc["Loan_Schedule"]
    assert ws["I6"].value == pytest.approx(gt["loan_closing_y7"], abs=0.01)
    assert ws["C3"].value == pytest.approx(gt["loan_interest_y1"], abs=0.01)
    assert ws["B8"].value is True
    # annuity equal across Y2..Y7
    pays = [ws.cell(row=4, column=c).value for c in range(5, 10)]
    assert all(p == pytest.approx(gt["annuity"], abs=0.01) for p in pays)


def test_pl(recalc, gt):
    ws = recalc["PL"]
    for t in range(1, 8):
        c = col(t)
        assert ws.cell(row=2, column=c).value == pytest.approx(
            gt["revenue"][t - 1], abs=0.05), f"revenue Y{t}"
        assert ws.cell(row=9, column=c).value == pytest.approx(
            gt["pat"][t - 1], abs=0.05), f"PAT Y{t}"
        assert ws.cell(row=10, column=c).value == pytest.approx(
            gt["cash_accruals"][t - 1], abs=0.05), f"cash accruals Y{t}"


def test_dscr(recalc, gt):
    ws = recalc["CashFlow_DSCR"]
    for t in range(1, 8):
        c = col(t, y0=True)
        assert ws.cell(row=6, column=c).value == pytest.approx(
            gt["dscr"][t - 1], abs=0.01), f"DSCR Y{t}"
    assert ws["B9"].value == pytest.approx(gt["avg_dscr"], abs=0.01)
    assert ws["B10"].value == pytest.approx(gt["min_dscr_repay"], abs=0.01)
    assert ws["B10"].value >= 1.25, "base min DSCR must clear the bank norm"
    assert ws["B13"].value == gt["payback"]


def test_npv_irr(recalc, gt):
    ws = recalc["CashFlow_DSCR"]
    assert ws["B11"].value == pytest.approx(gt["npv"], abs=1.0), "NPV"
    assert ws["B12"].value == pytest.approx(gt["irr"], abs=0.005), "IRR"
    assert ws["B11"].value > 0
    assert ws["B12"].value > 0.12


def test_balance_sheet_identity(recalc, gt):
    ws = recalc["BalanceSheet"]
    for t in range(0, 8):
        c = col(t, y0=True)
        gap = ws.cell(row=11, column=c).value
        assert gap == pytest.approx(gt["identity_gap"][t], abs=0.01), f"BS Y{t}"
        assert abs(gap) < 0.01
    assert min(ws.cell(row=4, column=c).value
               for c in range(2, 10)) >= 0, "cash never negative"


def test_ratios(recalc, gt):
    ws = recalc["Ratios"]
    assert ws["B8"].value == pytest.approx(gt["npv"], abs=1.0)
    assert ws["B9"].value == pytest.approx(gt["irr"], abs=0.005)
    assert ws["B11"].value == pytest.approx(gt["break_even_pct"], abs=0.002)
    assert ws["B12"].value is True


def test_sensitivity(recalc, gt):
    ws = recalc["Sensitivity"]
    for s, sc in enumerate(gt["scenarios"]):
        R = 2 + s * 25
        assert ws.cell(row=R + 21, column=2).value == pytest.approx(
            sc["min_dscr_repay"], abs=0.02), f"{sc['name']} minDSCR"
        assert ws.cell(row=R + 23, column=2).value == pytest.approx(
            sc["npv"], abs=2.0), f"{sc['name']} NPV"
        assert ws.cell(row=R + 24, column=2).value == pytest.approx(
            sc["irr"], abs=0.01), f"{sc['name']} IRR"
    # S1 mild stress still services debt; S2/S3 flag the raw-material risk
    assert ws["B48"].value >= 1.25
    assert ws["B73"].value < 1.25
    assert ws["B98"].value < ws["B23"].value


def test_all_checks_true(recalc):
    ws = recalc["Checks"]
    for r in range(2, 12):
        assert ws.cell(row=r, column=2).value is True, \
            f"check failed: {ws.cell(row=r, column=1).value}"


def test_dashboard_kpis(recalc, gt):
    ws = recalc["Dashboard"]
    assert ws["B4"].value == pytest.approx(240.0, abs=0.01)
    assert ws["B7"].value == pytest.approx(gt["avg_dscr"], abs=0.01)
    assert ws["B8"].value == pytest.approx(gt["min_dscr_repay"], abs=0.01)
    # risk flags are formula text, non-empty
    assert isinstance(ws["A14"].value, str) and ws["A14"].value.startswith(
        ("RISK", "OK"))
    assert isinstance(ws["A15"].value, str) and ws["A15"].value.startswith(
        ("RISK", "OK"))
