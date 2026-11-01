#!/usr/bin/env python3
"""Generate deterministic synthetic data for the monthly close pack demo.

Fictional business: Verma Consulting Services Pvt. Ltd., Vadodara —
a service business. Period: Jan-Oct 2026. All figures invented for
this demo; no real client data. Seed fixed so re-runs are identical.
"""
import csv
import datetime
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
SEED = 20261101

# (category, type, monthly base, monthly trend, noise)
PLAN = [
    ("Consulting fees", "Revenue", 340000, 6000, 30000),
    ("Monthly retainers", "Revenue", 155000, 1500, 6000),
    ("Training workshops", "Revenue", 65000, 2500, 22000),
    ("Contractor costs", "Direct", 58000, 800, 9000),
    ("Salaries", "Operating", 225000, 1200, 7000),
    ("Rent", "Operating", 45000, 0, 0),
    ("Software & tools", "Operating", 21000, 400, 2500),
    ("Travel", "Operating", 20000, 900, 7000),
    ("Marketing", "Operating", 24000, 800, 7000),
    ("Professional fees", "Operating", 16000, 500, 6000),
    ("Utilities", "Operating", 9500, 100, 1200),
    ("Other expenses", "Operating", 9000, 200, 3500),
]

# Budget stance per category: budget total as a share of the actual
# total is NOT fixed here (budgets are set per month below); instead
# each category gets a monthly budget bias so variances are realistic.
BUDGET_BIAS = {
    "Consulting fees": 0.97, "Monthly retainers": 1.00,
    "Training workshops": 0.94, "Contractor costs": 1.02,
    "Salaries": 1.02, "Rent": 1.00, "Software & tools": 1.00,
    "Travel": 0.88, "Marketing": 0.92, "Professional fees": 1.03,
    "Utilities": 1.00, "Other expenses": 1.01,
}

CLIENTS = ["Sharma Textiles", "Patel Logistics", "Mehta Foods",
           "Desai Pharma", "Iyer Engineering", "Khan Retail Group"]
VENDORS = {"Contractor costs": ["Contractor A", "Contractor B"],
           "Salaries": ["Payroll"], "Rent": ["Landlord — CG Road"],
           "Software & tools": ["SaaS vendor"], "Travel": ["Travel desk"],
           "Marketing": ["Ad platform"], "Professional fees": ["CA firm"],
           "Utilities": ["Utility provider"], "Other expenses": ["Misc vendor"]}


def split_total(rng, total, n):
    """Split an integer total into n positive parts summing to total."""
    if n == 1:
        return [total]
    cuts = sorted(rng.randint(1, total - 1) for _ in range(n - 1))
    parts, prev = [], 0
    for c in cuts + [total]:
        parts.append(c - prev)
        prev = c
    return parts


def main():
    rng = random.Random(SEED)
    os.makedirs(DATA, exist_ok=True)
    txns, budgets = [], []
    for m in range(1, 11):
        month = datetime.date(2026, m, 1)
        for category, typ, base, trend, noise in PLAN:
            total = int(round(base + trend * (m - 1) + rng.uniform(-noise, noise)))
            total = max(total, 1000)
            # Budget for the month: bias applied, small jitter.
            budget = int(round(total * BUDGET_BIAS[category]
                               * rng.uniform(0.98, 1.02)))
            budgets.append([month.isoformat(), category, typ, budget])
            n = rng.choice([2, 3, 3, 4])
            parts = split_total(rng, total, n)
            for i, amt in enumerate(parts):
                day = rng.randint(1, 28)
                if typ == "Revenue":
                    who = rng.choice(CLIENTS)
                    desc = f"{category} — {who}"
                else:
                    who = rng.choice(VENDORS[category])
                    desc = f"{category} — {who}"
                txns.append([datetime.date(2026, m, day).isoformat(),
                             desc, category, typ, amt])
    txns.sort(key=lambda r: (r[0], r[2]))
    with open(os.path.join(DATA, "transactions.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["date", "description", "category", "type", "amount"])
        w.writerows(txns)
    with open(os.path.join(DATA, "budget.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "category", "type", "budget"])
        w.writerows(budgets)
    print(f"transactions: {len(txns)} rows; budget: {len(budgets)} rows")


if __name__ == "__main__":
    main()
