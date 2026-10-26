#!/usr/bin/env python3
"""Independent ground truth for FarmPulse: sums straight from the CSVs.

Used by the tests to verify the LibreOffice-recalculated workbook. No
workbook formulas are read here — this is a separate computation path.
"""
import csv
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
QUARTERS = ["2025-Q3", "2025-Q4", "2026-Q1", "2026-Q2"]
ENTERPRISES = ["Dairy", "Wheat", "Vegetables", "Poultry", "Sheep"]
ACTIVITIES = ["Milk Sales", "Calf Sales", "Grain Sales", "Straw Sales",
              "Vegetable Sales", "Market Stall Sales", "Egg Sales",
              "Broiler Sales", "Wool Sales", "Lamb Sales"]
EXP_CATS = ["Feed", "Labour", "Veterinary", "Fuel & Machinery",
            "Fertiliser & Seed", "Transport", "Utilities", "Other"]


def read(name):
    with open(os.path.join(HERE, "data", name), newline="") as f:
        return list(csv.DictReader(f))


def ground_truth():
    inc = read("income.csv")
    exp = read("expense.csv")
    total_inc = round(sum(float(r["amount"]) for r in inc), 2)
    total_exp = round(sum(float(r["amount"]) for r in exp), 2)
    net = round(total_inc - total_exp, 2)
    margin = round(net / total_inc, 4) if total_inc else 0.0

    def sums(rows, key, quarters=True):
        out = {}
        for r in rows:
            k = r[key]
            out.setdefault(k, {q: 0.0 for q in QUARTERS})
            out[k][r["quarter"]] += float(r["amount"])
        for k in out:
            for q in QUARTERS:
                out[k][q] = round(out[k][q], 2)
        return out

    inc_ent = sums(inc, "enterprise")
    exp_ent = sums(exp, "enterprise")
    ent_profit = {}
    for e in ENTERPRISES:
        ent_profit[e] = {q: round(inc_ent.get(e, {}).get(q, 0.0)
                                   - exp_ent.get(e, {}).get(q, 0.0), 2)
                         for q in QUARTERS}
    best = max(ENTERPRISES,
               key=lambda e: sum(ent_profit[e].values()))
    q_inc = {q: round(sum(float(r["amount"]) for r in inc
                          if r["quarter"] == q), 2) for q in QUARTERS}
    q_exp = {q: round(sum(float(r["amount"]) for r in exp
                          if r["quarter"] == q), 2) for q in QUARTERS}
    return {
        "total_income": total_inc, "total_expense": total_exp,
        "net_profit": net, "margin": margin, "best_enterprise": best,
        "income_rows": len(inc), "expense_rows": len(exp),
        "ent_profit": ent_profit,
        "inc_activity": sums(inc, "activity"),
        "exp_category": sums(exp, "category"),
        "q_income": q_inc, "q_expense": q_exp,
    }


if __name__ == "__main__":
    import json
    print(json.dumps(ground_truth(), indent=2, default=str))
