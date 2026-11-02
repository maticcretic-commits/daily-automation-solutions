#!/usr/bin/env python3
"""Generate deterministic synthetic data for the FinancePulse demo.

Fictional household: A. Sharma, Patna — teaching salary plus freelance
design income and a small side business (candles, planners, spice
packs, tutoring), with a small investment portfolio. Period: Jan-Oct
2026. All figures invented for this demo; no real client data. Seed
fixed so re-runs are identical.
"""
import csv
import datetime
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
SEED = 20261102

EXPENSE_PLAN = [  # (category, monthly base, noise, rows per month)
    ("Housing", 18000, 1500, 2),
    ("Food", 14000, 2500, 4),
    ("Transport", 4500, 1200, 2),
    ("Utilities", 3800, 800, 2),
    ("Healthcare", 2500, 1800, 1),
    ("Education", 3000, 1200, 1),
    ("Entertainment", 2200, 1100, 1),
    ("Shopping", 3500, 1800, 1),
    ("Subscriptions", 1200, 200, 1),
    ("Other", 1800, 900, 1),
]

SALES_PLAN = [  # (product, unit price, cost per unit, qty range)
    ("Handmade candles", 349, 140, (4, 18)),
    ("Printed planners", 499, 190, (3, 12)),
    ("Spice packs", 249, 95, (6, 25)),
    ("Tutoring sessions", 600, 60, (2, 8)),
]

INVESTMENTS = [  # (holding, type, units, buy price, current price)
    ("Nifty 50 ETF", "ETF", 120, 245.50, 278.20),
    ("Gold ETF", "ETF", 85, 62.40, 71.15),
    ("HDFC Bank shares", "Shares", 40, 1680.00, 1745.50),
    ("Tata Motors shares", "Shares", 60, 985.00, 912.30),
    ("Fixed deposit", "Deposit", 1, 150000.00, 158400.00),
    ("PPF", "Deposit", 1, 50000.00, 52400.00),
    ("Liquid fund", "Mutual fund", 950, 108.20, 112.85),
    ("ELSS fund", "Mutual fund", 420, 88.60, 96.40),
]


def main():
    rng = random.Random(SEED)
    os.makedirs(DATA, exist_ok=True)

    income = []
    for m in range(1, 11):
        income.append({"date": datetime.date(2026, m, 1).isoformat(),
                       "source": "Teaching salary", "category": "Salary",
                       "amount": 85000})
        for _ in range(rng.randint(1, 3)):
            income.append({"date": datetime.date(2026, m, rng.randint(2, 28)).isoformat(),
                           "source": "Freelance design", "category": "Freelance",
                           "amount": rng.randrange(4000, 18001, 500)})
        if m % 3 == 1:
            income.append({"date": datetime.date(2026, m, 15).isoformat(),
                           "source": "Bank interest", "category": "Interest",
                           "amount": rng.randrange(1200, 3201, 100)})
    income.sort(key=lambda r: r["date"])
    with open(os.path.join(DATA, "income.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["date", "source", "category", "amount"])
        w.writeheader()
        w.writerows(income)

    expenses = []
    for m in range(1, 11):
        for cat, base, noise, nrows in EXPENSE_PLAN:
            share = base / nrows
            for _ in range(nrows):
                amt = max(100, round(share + rng.uniform(-noise, noise) / nrows, 0))
                expenses.append({"date": datetime.date(2026, m, rng.randint(1, 28)).isoformat(),
                                 "category": cat,
                                 "description": f"{cat} — {datetime.date(2026, m, 1):%b %Y}",
                                 "amount": int(amt)})
    expenses.sort(key=lambda r: r["date"])
    with open(os.path.join(DATA, "expenses.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["date", "category", "description", "amount"])
        w.writeheader()
        w.writerows(expenses)

    sales = []
    for m in range(1, 11):
        for product, price, cost, (lo, hi) in SALES_PLAN:
            for _ in range(rng.randint(1, 2)):
                sales.append({"date": datetime.date(2026, m, rng.randint(1, 28)).isoformat(),
                              "product": product, "qty": rng.randint(lo, hi),
                              "unit_price": price, "cost_per_unit": cost})
    sales.sort(key=lambda r: r["date"])
    with open(os.path.join(DATA, "sales.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["date", "product", "qty", "unit_price", "cost_per_unit"])
        w.writeheader()
        w.writerows(sales)

    with open(os.path.join(DATA, "investments.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["holding", "type", "units", "buy_price", "current_price"])
        w.writeheader()
        for holding, typ, units, buy, cur in INVESTMENTS:
            w.writerow({"holding": holding, "type": typ, "units": units,
                        "buy_price": buy, "current_price": cur})

    print(f"income {len(income)} rows, expenses {len(expenses)} rows, "
          f"sales {len(sales)} rows, investments {len(INVESTMENTS)} holdings")


if __name__ == "__main__":
    main()
