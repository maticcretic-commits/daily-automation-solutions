#!/usr/bin/env python3
"""Generate deterministic synthetic farm transactions for FarmPulse.

Fictional farm, clearly labelled SYNTHETIC — no client data. Seeded so every
run produces identical CSVs. Dates Jul 2025 – Jun 2026, four quarters.
"""
import csv
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "data")
os.makedirs(OUT, exist_ok=True)

ENTERPRISES = ["Dairy", "Wheat", "Vegetables", "Poultry", "Sheep"]
ACTIVITIES = {
    "Dairy": ["Milk Sales", "Calf Sales"],
    "Wheat": ["Grain Sales", "Straw Sales"],
    "Vegetables": ["Vegetable Sales", "Market Stall Sales"],
    "Poultry": ["Egg Sales", "Broiler Sales"],
    "Sheep": ["Wool Sales", "Lamb Sales"],
}
EXP_CATS = ["Feed", "Labour", "Veterinary", "Fuel & Machinery",
            "Fertiliser & Seed", "Transport", "Utilities", "Other"]
QUARTERS = [("2025-Q3", "2025-07-01", "2025-09-30"),
            ("2025-Q4", "2025-10-01", "2025-12-31"),
            ("2026-Q1", "2026-01-01", "2026-03-31"),
            ("2026-Q2", "2026-04-01", "2026-06-30")]
PADDOCKS = ["Paddock A", "Paddock B", "Paddock C"]
HERDS = ["Herd 1", "Herd 2", "Flock A"]
COST_CENTRES = ["CC-100 Operations", "CC-200 Feed & Vet",
                "CC-300 Machinery", "CC-400 Admin"]


def rand_date(rng, start, end):
    from datetime import date, timedelta
    s = date.fromisoformat(start)
    e = date.fromisoformat(end)
    return s + timedelta(days=rng.randint(0, (e - s).days))


def main():
    rng = random.Random(20261026)
    income = []
    for i in range(172):
        q, qs, qe = QUARTERS[i % 4]
        ent = ENTERPRISES[(i // 4) % len(ENTERPRISES)]
        act = ACTIVITIES[ent][i % len(ACTIVITIES[ent])]
        income.append({
            "date": rand_date(rng, qs, qe).isoformat(),
            "quarter": q, "enterprise": ent, "activity": act,
            "paddock_herd": rng.choice(PADDOCKS + HERDS),
            "cost_centre": rng.choice(COST_CENTRES),
            "amount": round(rng.uniform(250, 9500), 2),
        })
    expense = []
    for i in range(164):
        q, qs, qe = QUARTERS[i % 4]
        ent = ENTERPRISES[(i // 4) % len(ENTERPRISES)]
        expense.append({
            "date": rand_date(rng, qs, qe).isoformat(),
            "quarter": q, "enterprise": ent,
            "category": EXP_CATS[i % len(EXP_CATS)],
            "paddock_herd": rng.choice(PADDOCKS + HERDS),
            "cost_centre": rng.choice(COST_CENTRES),
            "amount": round(rng.uniform(120, 6200), 2),
        })
    with open(os.path.join(OUT, "income.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["date", "quarter", "enterprise",
                                         "activity", "paddock_herd",
                                         "cost_centre", "amount"])
        w.writeheader()
        w.writerows(income)
    with open(os.path.join(OUT, "expense.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["date", "quarter", "enterprise",
                                         "category", "paddock_herd",
                                         "cost_centre", "amount"])
        w.writeheader()
        w.writerows(expense)
    print(f"income rows: {len(income)}, expense rows: {len(expense)}")


if __name__ == "__main__":
    main()
