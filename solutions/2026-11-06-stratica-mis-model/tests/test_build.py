#!/usr/bin/env python3
"""Tests for StraticaMIS: generator + LibreOffice recalc vs Python ground truth."""
import csv
import re
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
OUT = BASE / "stratica-mis-model.xlsx"
TMP = Path("/tmp/stratitamis-test")

OPEN_CASH = 1_200_000
OPEN_RECV = 1_800_000
OPEN_INV = 1_500_000
OPEN_FA = 4_000_000
OPEN_PAY = 1_000_000
OPEN_LOAN = 2_500_000
OPEN_CAPITAL = 3_000_000
OPEN_RE = 2_000_000
DEPN = 60_000
LOAN_REPAY = 100_000


def ground_truth(rows):
    rev = [r[1] for r in rows]
    cogs = [r[2] for r in rows]
    opex = [[r[3], r[4], r[5], r[6], r[7]] for r in rows]
    capex = [r[8] for r in rows]
    gp = [a - b for a, b in zip(rev, cogs)]
    topex = [sum(x) for x in opex]
    ebitda = [a - b for a, b in zip(gp, topex)]
    loan = OPEN_LOAN
    interest, pbt, tax, np_ = [], [], [], []
    for e in ebitda:
        intr = round(loan * 0.01)
        loan -= LOAN_REPAY
        interest.append(intr)
        p = e - DEPN - intr
        pbt.append(p)
        t = round(0.25 * max(p, 0))
        tax.append(t)
        np_.append(p - t)
    recv = [round(r * 0.50) for r in rev]
    inv = [round(c * 0.80) for c in cogs]
    pay = [round(c * 0.45) for c in cogs]
    fa = [OPEN_FA]
    for c in capex:
        fa.append(fa[-1] + c - DEPN)
    fa = fa[1:]
    re_ = [OPEN_RE]
    for n in np_:
        re_.append(re_[-1] + n)
    re_ = re_[1:]
    # Indirect cash flow.
    cash = [OPEN_CASH]
    ocf = []
    prev_recv, prev_inv, prev_pay = OPEN_RECV, OPEN_INV, OPEN_PAY
    for i in range(len(rows)):
        op = (np_[i] + DEPN - (recv[i] - prev_recv) - (inv[i] - prev_inv)
              + (pay[i] - prev_pay))
        ocf.append(op)
        cash.append(cash[-1] + op - capex[i] - LOAN_REPAY)
        prev_recv, prev_inv, prev_pay = recv[i], inv[i], pay[i]
    cash = cash[1:]
    # Budget.
    brev = [round(1_900_000 * (1.04 ** i)) for i in range(len(rows))]
    bgp = [round(b * 0.45) for b in brev]
    bebitda = [g - 520_000 for g in bgp]
    bnp = []
    for e in bebitda:
        p = e - DEPN - 25_000
        bnp.append(p - round(0.25 * max(p, 0)))
    return {
        "fy_rev": sum(rev),
        "fy_np": sum(np_),
        "net_margin": sum(np_) / sum(rev),
        "ebitda_margin": sum(ebitda) / sum(rev),
        "close_cash": cash[-1],
        "current_ratio": (cash[-1] + recv[-1] + inv[-1]) / pay[-1],
        "recv_days": recv[-1] / rev[-1] * 30,
        "inv_days": inv[-1] / cogs[-1] * 30,
        "budget_attain": sum(rev) / sum(brev),
        "fy_ocf": sum(ocf),
        "fy_gp": sum(gp),
        "fy_ebitda": sum(ebitda),
        "fy_tax": sum(tax),
        "fy_pbt": sum(pbt),
        "loan_mar": OPEN_LOAN - 12 * LOAN_REPAY,
        "fa_mar": OPEN_FA + sum(capex) - 12 * DEPN,
        "re_mar": OPEN_RE + sum(np_),
        "bs_diff_max": 0,
    }


def test_all():
    subprocess.run([sys.executable, str(BASE / "generate.py")], check=True, cwd=BASE)
    assert OUT.exists()
    with open(BASE / "data" / "inputs.csv") as f:
        rows = [(r["month"], int(r["revenue"]), int(r["cogs"]), int(r["salaries"]),
                 int(r["rent"]), int(r["marketing"]), int(r["admin"]),
                 int(r["other_opex"]), int(r["capex"]))
                for r in csv.DictReader(f)]
    assert len(rows) == 12
    gt = ground_truth(rows)

    # No short-form ranges, no structured refs in any formula.
    wb0 = openpyxl.load_workbook(OUT)
    assert wb0.sheetnames == ["Guide", "Inputs", "PL", "BS", "CF", "Budget",
                               "Dashboard", "Checks"], wb0.sheetnames
    for ws in wb0.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    assert not re.search(r"\$[A-Z]+\$\d+:\$\d+", cell.value), cell.value
                    assert "[" not in cell.value, cell.value
    assert len(wb0["Dashboard"]._charts) == 3

    # LibreOffice headless recalc.
    if TMP.exists():
        shutil.rmtree(TMP)
    TMP.mkdir(parents=True)
    shutil.copy(OUT, TMP / OUT.name)
    r = subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx",
                        "--outdir", str(TMP / "recalc"), str(TMP / OUT.name)],
                       capture_output=True, text=True, timeout=180)
    recalc = TMP / "recalc" / OUT.name
    assert recalc.exists(), r.stdout[-500:] + r.stderr[-500:]
    wb = openpyxl.load_workbook(recalc, data_only=True)
    # Zero formula errors anywhere.
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                assert not (isinstance(cell.value, str) and cell.value.startswith("#")), \
                    f"{ws.title} {cell.coordinate} = {cell.value}"
    checks = [wb["Checks"].cell(row=r, column=2).value for r in range(2, 10)]
    assert checks == ["PASS"] * 8, checks

    def close(got, exp, key):
        assert got is not None, key
        assert abs(got - exp) <= max(1.0, abs(exp) * 1e-6), f"{key}: got {got}, expected {exp}"

    pl, bs, cf, bg, db = wb["PL"], wb["BS"], wb["CF"], wb["Budget"], wb["Dashboard"]
    # FY totals on PL.
    close(pl.cell(row=14, column=2).value, gt["fy_rev"], "fy_rev")
    close(pl.cell(row=14, column=4).value, gt["fy_gp"], "fy_gp")
    close(pl.cell(row=14, column=11).value, gt["fy_ebitda"], "fy_ebitda")
    close(pl.cell(row=14, column=14).value, gt["fy_pbt"], "fy_pbt")
    close(pl.cell(row=14, column=15).value, gt["fy_tax"], "fy_tax")
    close(pl.cell(row=14, column=16).value, gt["fy_np"], "fy_np")
    # March balance sheet.
    close(bs.cell(row=13, column=2).value, gt["close_cash"], "close_cash")
    close(bs.cell(row=13, column=8).value, gt["loan_mar"], "loan_mar")
    close(bs.cell(row=13, column=5).value, gt["fa_mar"], "fa_mar")
    close(bs.cell(row=13, column=10).value, gt["re_mar"], "re_mar")
    close(bs.cell(row=13, column=6).value, bs.cell(row=13, column=11).value, "bs_balances_mar")
    # FY cash flow + budget totals.
    close(cf.cell(row=14, column=7).value, gt["fy_ocf"], "fy_ocf")
    close(bg.cell(row=14, column=5).value, gt["budget_attain"] - 1, "budget_var_pct")
    # Dashboard KPIs mirror the model.
    kpi_keys = ["fy_rev", "fy_np", "net_margin", "ebitda_margin", "close_cash",
                "current_ratio", "recv_days", "inv_days", "budget_attain", "fy_ocf"]
    for i, key in enumerate(kpi_keys):
        close(db.cell(row=3 + i, column=2).value, gt[key], f"kpi:{key}")
    print("ALL TESTS PASS")
    print(f"fy_rev={gt['fy_rev']:,} fy_np={gt['fy_np']:,} net_margin={gt['net_margin']:.1%} "
          f"ebitda_margin={gt['ebitda_margin']:.1%} close_cash={gt['close_cash']:,} "
          f"current_ratio={gt['current_ratio']:.2f} recv_days={gt['recv_days']:.1f} "
          f"inv_days={gt['inv_days']:.1f} budget_attain={gt['budget_attain']:.1%} "
          f"fy_ocf={gt['fy_ocf']:,}")


if __name__ == "__main__":
    test_all()
