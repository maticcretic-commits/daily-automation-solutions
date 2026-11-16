#!/usr/bin/env python3
"""Generate deterministic fictional sales data for SalesPerformanceDashboard.

Fictional business: Ganga Traders, Patna — a wholesale FMCG distributor,
Jan–Oct 2026. All figures are invented for the demo; no real client data.

Outputs (under data/):
  orders.csv  — one row per order: Date, OrderID, Category, Product,
                Region, Channel, Qty, UnitPrice, Revenue, Month
  targets.csv — monthly revenue target per category (round numbers).

Method: targets are set first as round numbers with a gentle monthly
growth trend; each category-month's orders are then sized so their
revenue lands near target x a deterministic factor in [0.88, 1.08],
so % to goal varies realistically between ~88% and ~108%.
Seed fixed at 20261116 — re-running reproduces the files exactly.
"""
import csv
import random
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
SEED = 20261116

PRODUCTS = {
    "Staples": [("Atta 10kg", 420), ("Basmati Rice 5kg", 310),
                ("Sugar 1kg", 45)],
    "Beverages": [("Tea 500g", 260), ("Instant Coffee 200g", 340)],
    "Personal Care": [("Bath Soap (pack of 4)", 180),
                      ("Shampoo 650ml", 290)],
    "Home Care": [("Detergent Powder 3kg", 330),
                  ("Floor Cleaner 1L", 150)],
    "Snacks": [("Biscuits 400g", 95), ("Namkeen 500g", 120)],
}
# Monthly target base per category (Rs), before the growth trend.
TARGET_BASE = {"Staples": 420000, "Beverages": 180000,
               "Personal Care": 240000, "Home Care": 200000,
               "Snacks": 150000}
REGIONS = ["Patna", "Muzaffarpur", "Gaya", "Bhagalpur", "Darbhanga"]
CHANNELS = ["Wholesale", "Retail", "Online"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 11)]
DAYS_IN_MONTH = {1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31,
                 8: 31, 9: 30, 10: 31}


def main():
    rng = random.Random(SEED)
    DATA.mkdir(parents=True, exist_ok=True)

    targets = []
    for cat, base in TARGET_BASE.items():
        for i, month in enumerate(MONTHS):
            target = round(base * (1 + 0.02 * i) / 5000) * 5000
            targets.append({"Category": cat, "Month": month,
                            "Target": target})
    target_of = {(t["Category"], t["Month"]): t["Target"] for t in targets}

    orders = []
    for cat, prods in PRODUCTS.items():
        for i, month in enumerate(MONTHS):
            mnum = i + 1
            factor = round(rng.uniform(0.88, 1.08), 4)
            month_total = target_of[(cat, month)] * factor
            n = rng.randint(4, 6)
            weights = [rng.uniform(0.5, 1.5) for _ in range(n)]
            wsum = sum(weights)
            for k in range(n):
                product, price = rng.choice(prods)
                revenue = month_total * weights[k] / wsum
                qty = max(1, round(revenue / price))
                revenue = round(qty * price, 2)
                day = rng.randint(1, DAYS_IN_MONTH[mnum])
                orders.append({
                    "Date": f"2026-{mnum:02d}-{day:02d}",
                    "Category": cat, "Product": product,
                    "Region": rng.choice(REGIONS),
                    "Channel": rng.choice(CHANNELS),
                    "Qty": qty, "UnitPrice": price,
                    "Revenue": revenue, "Month": month})
    orders.sort(key=lambda o: (o["Date"], o["Category"], o["Product"]))
    for j, o in enumerate(orders, start=1):
        o2 = {"OrderID": f"GT-2026-{j:04d}"}
        o2.update(o)
        orders[j - 1] = o2

    with open(DATA / "orders.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=[
            "Date", "OrderID", "Category", "Product", "Region",
            "Channel", "Qty", "UnitPrice", "Revenue", "Month"])
        w.writeheader()
        w.writerows(orders)
    with open(DATA / "targets.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Category", "Month", "Target"])
        w.writeheader()
        w.writerows(targets)
    total = sum(o["Revenue"] for o in orders)
    ttotal = sum(t["Target"] for t in targets)
    print(f"orders.csv: {len(orders)} rows, revenue Rs {total:,.2f}")
    print(f"targets.csv: {len(targets)} rows, target Rs {ttotal:,.2f} "
          f"({total / ttotal:.1%} of target)")


if __name__ == "__main__":
    main()
