#!/usr/bin/env python3
"""Tests for ExecPortfolio (build-2026-11-07).

- Rebuilds the workbook from generate.py (fresh).
- Regex-scans every formula for short-form range literals
  ($H$2:$5000 — invalid in LibreOffice; must repeat the column letter).
- Recalculates with LibreOffice headless and compares every KPI,
  per-project rollup, monthly total and programme total against an
  independent Python ground truth computed from the CSVs.
- Asserts all Checks read PASS and no cell carries a formula error.
"""
import csv
import re
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
OUTDIR = Path("/tmp/exec-recalc-20261107")
XLSX = BASE / "exec-portfolio.xlsx"
AS_OF = "2026-11-07"
SHORT_RANGE = re.compile(r"\$[A-Z]+\$\d+:\$\d+")
ERRORS = {"#VALUE!", "#NAME?", "#REF!", "#DIV/0!", "#N/A", "#NUM!", "Err:508",
          "Err:509", "Err:510", "Err:511", "Err:512", "Err:513", "Err:514",
          "Err:515", "Err:516", "Err:517", "Err:518", "Err:519", "Err:520",
          "Err:521", "Err:522", "Err:523", "Err:524", "Err:525", "Err:526",
          "Err:527", "Err:528", "Err:529", "Err:530", "Err:531", "Err:532"}


def read_csv(name):
    with open(BASE / "data" / name, newline="") as f:
        return list(csv.DictReader(f))


def ground_truth():
    projects = read_csv("projects.csv")
    tasks = read_csv("tasks.csv")
    risks = read_csv("risks.csv")
    spend = read_csv("spend.csv")
    budget = {p["project_id"]: int(p["budget"]) for p in projects}
    prog = {p["project_id"]: p["programme"] for p in projects}
    spent = defaultdict(int)
    for s in spend:
        spent[s["project_id"]] += int(s["amount"])
    per_project = {}
    for pid in budget:
        p_tasks = [t for t in tasks if t["project_id"] == pid]
        done = sum(1 for t in p_tasks if t["status"] == "Completed")
        overdue = sum(1 for t in p_tasks
                      if t["due"] < AS_OF and t["status"] != "Completed")
        pcts = [float(t["pct"]) for t in p_tasks]
        avg_pct = sum(pcts) / len(pcts)
        var = (spent[pid] - budget[pid]) / budget[pid]
        rag = ("Green" if var <= 0.1 and overdue == 0 else
               "Red" if var > 0.25 or overdue >= 5 else "Amber")
        top_risk = max(int(r["likelihood"]) * int(r["impact"])
                       for r in risks if r["project_id"] == pid)
        per_project[pid] = dict(spent=spent[pid], var=var, avg_pct=avg_pct,
                                n=len(p_tasks), done=done, overdue=overdue,
                                rag=rag, top_risk=top_risk)
    months = sorted({s["month"] for s in spend})
    monthly = {m: sum(int(s["amount"]) for s in spend if s["month"] == m)
               for m in months}
    programmes = sorted({p["programme"] for p in projects})
    prog_totals = {g: (sum(budget[pid] for pid in budget if prog[pid] == g),
                       sum(spent[pid] for pid in budget if prog[pid] == g))
                   for g in programmes}
    return dict(
        n_proj=len(projects), n_task=len(tasks), n_risk=len(risks),
        n_spend=len(spend), budget=budget, spent=spent,
        total_budget=sum(budget.values()), total_spent=sum(spent.values()),
        per_project=per_project, monthly=monthly, prog_totals=prog_totals,
        green=sum(1 for v in per_project.values() if v["rag"] == "Green"),
        red=sum(1 for v in per_project.values() if v["rag"] == "Red"),
        overdue_total=sum(v["overdue"] for v in per_project.values()),
        done_total=sum(v["done"] for v in per_project.values()),
        max_risk=max(int(r["likelihood"]) * int(r["impact"]) for r in risks),
        avg_completion=sum(v["avg_pct"] for v in per_project.values()) / len(per_project),
    )


def approx(a, b, tol=0.005):
    assert abs(a - b) <= tol, f"{a} != {b} (tol {tol})"


def main():
    # 1. Fresh build.
    r = subprocess.run([sys.executable, str(BASE / "generate.py")],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    print(r.stdout.strip())

    gt = ground_truth()
    assert gt["n_proj"] == 12 and gt["n_task"] == 60
    assert gt["n_risk"] == 24 and gt["n_spend"] == 72

    # 2. Short-form range scan over every formula cell.
    wb = load_workbook(XLSX)
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    if SHORT_RANGE.search(c.value):
                        bad.append(f"{ws.title}!{c.coordinate}: {c.value}")
    assert not bad, "short-form ranges found:\n" + "\n".join(bad)
    print("range-literal scan: clean")

    # 3. LibreOffice recalc.
    OUTDIR.mkdir(exist_ok=True)
    r = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx",
                        "--outdir", str(OUTDIR), str(XLSX)],
                       capture_output=True, text=True, timeout=300)
    assert r.returncode == 0, r.stderr
    recalc = OUTDIR / "exec-portfolio.xlsx"
    assert recalc.exists(), "LO produced no output"
    wb2 = load_workbook(recalc, data_only=True)

    # No formula errors anywhere.
    errs = []
    for ws in wb2.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and (c.value in ERRORS
                                                 or c.value.startswith("Err:")):
                    errs.append(f"{ws.title}!{c.coordinate}: {c.value}")
    assert not errs, "formula errors:\n" + "\n".join(errs)
    print("formula errors: none")

    # 4. Dashboard KPIs vs ground truth.
    db = wb2["Dashboard"]
    approx(db["B4"].value, gt["total_budget"], 1)
    approx(db["B5"].value, gt["total_spent"], 1)
    approx(db["B6"].value, (gt["total_spent"] - gt["total_budget"]) / gt["total_budget"])
    assert db["B7"].value == gt["green"], (db["B7"].value, gt["green"])
    assert db["B8"].value == gt["red"], (db["B8"].value, gt["red"])
    assert db["B9"].value == gt["overdue_total"]
    assert db["B10"].value == f"{gt['done_total']} / {gt['n_task']}"
    approx(db["B11"].value, gt["done_total"] / gt["n_task"])
    approx(db["B12"].value, gt["avg_completion"])
    assert db["B13"].value == gt["max_risk"]
    print(f"KPIs OK: budget ${gt['total_budget']:,} spent ${gt['total_spent']:,} "
          f"green {gt['green']} red {gt['red']} overdue {gt['overdue_total']} "
          f"max risk {gt['max_risk']}")

    # 5. Per-project rollup vs ground truth.
    pf = wb2["Portfolio"]
    for i, pid in enumerate(sorted(gt["per_project"])):
        r_ = 3 + i
        assert pf.cell(row=r_, column=1).value == pid
        v = gt["per_project"][pid]
        approx(pf.cell(row=r_, column=5).value, v["spent"], 1)
        approx(pf.cell(row=r_, column=6).value, v["var"])
        approx(pf.cell(row=r_, column=7).value, v["avg_pct"])
        assert pf.cell(row=r_, column=8).value == v["n"]
        assert pf.cell(row=r_, column=9).value == v["done"]
        assert pf.cell(row=r_, column=10).value == v["overdue"], (pid, pf.cell(row=r_, column=10).value, v["overdue"])
        assert pf.cell(row=r_, column=11).value == v["rag"], (pid, pf.cell(row=r_, column=11).value, v["rag"])
        assert pf.cell(row=r_, column=12).value == v["top_risk"]
    print("portfolio rollup: 12/12 projects match")

    # 6. Monthly + programme totals.
    mo = wb2["Monthly"]
    for i, m in enumerate(sorted(gt["monthly"])):
        approx(mo.cell(row=2 + i, column=2).value, gt["monthly"][m], 1)
    approx(mo.cell(row=8, column=2).value, gt["total_spent"], 1)
    prog_rows = {}
    for row in db.iter_rows(min_row=17, max_row=20, max_col=5):
        prog_rows[row[0].value] = (row[2].value, row[3].value)
    for g, (b, s) in gt["prog_totals"].items():
        approx(prog_rows[g][0], b, 1)
        approx(prog_rows[g][1], s, 1)
    print("monthly + programme totals: match")

    # 7. Checks all PASS.
    ck = wb2["Checks"]
    results = [ck.cell(row=r, column=2).value for r in range(2, 12)]
    assert all(v == "PASS" for v in results), results
    print(f"checks: {len(results)}/{len(results)} PASS")
    print("ALL TESTS PASS")


if __name__ == "__main__":
    main()
