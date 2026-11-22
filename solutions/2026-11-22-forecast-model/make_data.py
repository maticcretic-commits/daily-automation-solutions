#!/usr/bin/env python3
"""Write data/assumptions.csv (the workbook's only inputs) and
data/ground_truth.json (the independent Python model, for tests).

Run:  python3 make_data.py
"""
import csv
import json
import os

import model

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")


def main():
    os.makedirs(DATA, exist_ok=True)
    d = model.DRIVERS
    rows = [
        ("company", d["company"]),
        ("rev_base", d["rev_base"]),
        ("growth_base", d["growth"]["Base"]),
        ("growth_optimistic", d["growth"]["Optimistic"]),
        ("growth_pessimistic", d["growth"]["Pessimistic"]),
        ("cogs_base", d["cogs_pct"]["Base"]),
        ("cogs_optimistic", d["cogs_pct"]["Optimistic"]),
        ("cogs_pessimistic", d["cogs_pct"]["Pessimistic"]),
        ("salary_base", d["salary_base"]),
        ("salary_growth", d["salary_growth"]),
        ("rent", d["rent"]),
        ("utilities", d["utilities"]),
        ("other_opex", d["other_opex"]),
        ("marketing_pct", d["marketing_pct"]),
        ("loan_opening", d["loan_opening"]),
        ("loan_rate", d["loan_rate"]),
        ("loan_principal", d["loan_principal"]),
        ("tax_rate", d["tax_rate"]),
        ("plant_opening", d["plant_opening"]),
        ("plant_life_months", d["plant_life_months"]),
        ("capex_m7", d["capex"][7]),
        ("capex_m13", d["capex"][13]),
        ("wc_recv_days", d["wc_recv_days"]),
        ("wc_inv_days", d["wc_inv_days"]),
        ("wc_pay_days", d["wc_pay_days"]),
        ("cash_opening", d["cash_opening"]),
        ("recv_opening", d["recv_opening"]),
        ("inv_opening", d["inv_opening"]),
        ("pay_opening", d["pay_opening"]),
        ("equity", d["equity"]),
        ("retained_opening", d["retained_opening"]),
    ]
    with open(os.path.join(DATA, "assumptions.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["key", "value"])
        w.writerows(rows)
    gt = model.build_ground_truth()
    with open(os.path.join(DATA, "ground_truth.json"), "w") as f:
        json.dump(gt, f, indent=1)
    print(f"Wrote {len(rows)} assumptions, "
          f"ground truth: {len(gt['months'])} months, "
          f"{len(gt['years'])} years")


if __name__ == "__main__":
    main()
