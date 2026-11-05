#!/usr/bin/env python3
"""Tests for PortfolioPulse: generator + LibreOffice recalc vs Python ground truth."""
import csv
import math
import re
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
OUT = BASE / "portfolio-performance.xlsx"
TMP = Path("/tmp/portfoliopulse-test")


def mean(xs):
    return sum(xs) / len(xs)


def stdev(xs):
    m = mean(xs)
    return math.sqrt(sum((x - m) ** 2 for x in xs) / (len(xs) - 1))


def slope(ys, xs):
    mx, my = mean(xs), mean(ys)
    return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / sum((x - mx) ** 2 for x in xs)


def ground_truth(rows):
    p = [r[1] for r in rows]
    b = [r[2] for r in rows]
    rf = [r[3] for r in rows]
    pe = [x - y for x, y in zip(p, rf)]
    be = [x - y for x, y in zip(b, rf)]
    act = [x - y for x, y in zip(p, b)]
    wp, wb = [1.0], [1.0]
    for x in p:
        wp.append(wp[-1] * (1 + x))
    for x in b:
        wb.append(wb[-1] * (1 + x))
    wp, wb = wp[1:], wb[1:]
    peak_p, dd_p = -1e18, []
    for v in wp:
        peak_p = max(peak_p, v)
        dd_p.append(v / peak_p - 1)
    peak_b, dd_b = -1e18, []
    for v in wb:
        peak_b = max(peak_b, v)
        dd_b.append(v / peak_b - 1)
    beta = slope(p, b)
    return {
        "months": len(p),
        "cum_p": wp[-1] - 1,
        "cum_b": wb[-1] - 1,
        "ann_p": wp[-1] ** (12 / len(p)) - 1,
        "ann_b": wb[-1] ** (12 / len(b)) - 1,
        "vol_m": stdev(p),
        "vol_a": stdev(p) * math.sqrt(12),
        "vol_b_a": stdev(b) * math.sqrt(12),
        "sharpe_p": mean(pe) / stdev(p) * math.sqrt(12),
        "sharpe_b": mean(be) / stdev(b) * math.sqrt(12),
        "beta": beta,
        "alpha_m": mean(pe) - beta * mean(be),
        "alpha_a": (mean(pe) - beta * mean(be)) * 12,
        "corr": sum((x - mean(p)) * (y - mean(b)) for x, y in zip(p, b))
        / math.sqrt(sum((x - mean(p)) ** 2 for x in p) * sum((y - mean(b)) ** 2 for y in b)),
        "maxdd_p": min(dd_p),
        "maxdd_b": min(dd_b),
        "best": max(p),
        "worst": min(p),
        "win": sum(1 for x in p if x > 0) / len(p),
        "te": stdev(act) * math.sqrt(12),
        "ir": mean(act) / stdev(act) * math.sqrt(12),
    }


def test_all():
    subprocess.run([sys.executable, str(BASE / "generate.py")], check=True, cwd=BASE)
    assert OUT.exists()
    with open(BASE / "data" / "returns.csv") as f:
        rows = [(r["month"], float(r["portfolio"]), float(r["benchmark"]), float(r["risk_free"]))
                for r in csv.DictReader(f)]
    assert len(rows) == 22
    gt = ground_truth(rows)

    # No short-form ranges, no structured refs in any formula.
    wb0 = openpyxl.load_workbook(OUT)
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
    m = wb["Metrics"]
    vals = [m.cell(row=r, column=2).value for r in range(2, 23)]
    keys = ["months", "cum_p", "cum_b", "ann_p", "ann_b", "vol_m", "vol_a", "vol_b_a",
            "sharpe_p", "sharpe_b", "beta", "alpha_m", "alpha_a", "corr",
            "maxdd_p", "maxdd_b", "best", "worst", "win", "te", "ir"]
    for key, got in zip(keys, vals):
        exp = gt[key]
        assert got is not None, key
        assert abs(got - exp) <= max(1e-6, abs(exp) * 1e-4), f"{key}: got {got}, expected {exp}"
    # Dashboard KPIs mirror Metrics.
    db = wb["Dashboard"]
    assert abs(db.cell(row=3, column=2).value - gt["cum_p"]) < 1e-6
    assert abs(db.cell(row=7, column=2).value - gt["sharpe_p"]) < 1e-4
    print("ALL TESTS PASS")
    print(f"cum_p={gt['cum_p']:.4%} cum_b={gt['cum_b']:.4%} ann_p={gt['ann_p']:.4%} "
          f"sharpe_p={gt['sharpe_p']:.3f} beta={gt['beta']:.3f} alpha_a={gt['alpha_a']:.4%} "
          f"maxdd_p={gt['maxdd_p']:.4%} win={gt['win']:.1%}")


if __name__ == "__main__":
    test_all()
