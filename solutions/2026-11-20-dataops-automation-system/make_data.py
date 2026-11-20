#!/usr/bin/env python3
"""Generate deterministic fictional source data for DataOps Automation System.

Fictional company: Narmada Traders, Indore (wholesale FMCG distributor).
Period: Apr-Oct 2026. All values are invented for the demo; nothing here
is a real client's data.

Outputs (into data/):
  products.csv  — product master (product, category, unit_cost, unit_price)
  sales.csv     — one row per order line
  purchases.csv — one row per purchase-order line
  expenses.csv  — one row per expense
  meta.csv      — company / period / seed
"""
import csv
import random
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
SEED = 20261120

PRODUCTS = [
    ("Rice 25kg", "Grains", 980, 1150),
    ("Wheat Flour 10kg", "Grains", 340, 420),
    ("Sugar 1kg", "Sweeteners", 38, 46),
    ("Edible Oil 1L", "Oils", 128, 155),
    ("Tea 500g", "Beverages", 210, 260),
    ("Toor Dal 1kg", "Pulses", 118, 145),
    ("Spice Mix 200g", "Spices", 72, 95),
    ("Biscuits 400g", "Snacks", 34, 45),
    ("Bathing Soap 100g", "Personal Care", 22, 30),
    ("Detergent 1kg", "Home Care", 88, 110),
]

SUPPLIERS = ["Shree Agro Mills", "Malwa Foods Pvt Ltd", "Narmada Agro",
             "Central India Traders", "Vindhya Food Products"]
CHANNELS = ["Wholesale", "Retail", "Online"]
REGIONS = ["Indore", "Ujjain", "Bhopal", "Dewas", "Ratlam"]

MONTHS = ["2026-04", "2026-05", "2026-06", "2026-07",
          "2026-08", "2026-09", "2026-10"]
# Festive-season lift in Sep-Oct (Diwali trading), quieter monsoon Jun-Jul.
MONTH_WEIGHT = {"2026-04": 1.0, "2026-05": 1.05, "2026-06": 0.9,
                "2026-07": 0.95, "2026-08": 1.0, "2026-09": 1.15,
                "2026-10": 1.35}

FIXED_EXPENSES = [("Rent", 8000), ("Salaries", 22000)]
VAR_EXPENSES = [("Transport", 3000, 9000), ("Utilities", 2000, 6000),
                ("Marketing", 3000, 12000), ("Miscellaneous", 500, 3000)]


def days_in_month(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    if m == 12:
        return 31
    import datetime as _dt
    return (_dt.date(y, m + 1, 1) - _dt.date(y, m, 1)).days


def make_data():
    rng = random.Random(SEED)
    DATA.mkdir(parents=True, exist_ok=True)

    with open(DATA / "products.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["product", "category", "unit_cost", "unit_price"])
        for p in PRODUCTS:
            w.writerow(p)

    sales = []
    order_no = 1000
    for ym in MONTHS:
        n = int(round(34 * MONTH_WEIGHT[ym]))
        dim = days_in_month(ym)
        for _ in range(n):
            order_no += 1
            prod, _cat, cost, price = rng.choice(PRODUCTS)
            qty = rng.randint(5, 80)
            day = rng.randint(1, dim)
            sales.append({
                "order_id": f"SO-{order_no}",
                "date": f"{ym}-{day:02d}",
                "product": prod,
                "category": _cat,
                "qty": qty,
                "unit_price": price,
                "unit_cost": cost,
                "channel": rng.choices(CHANNELS, weights=[55, 30, 15])[0],
                "region": rng.choice(REGIONS),
            })
    with open(DATA / "sales.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(sales[0].keys()))
        w.writeheader()
        w.writerows(sales)

    purchases = []
    po_no = 500
    for ym in MONTHS:
        n = int(round(17 * MONTH_WEIGHT[ym]))
        dim = days_in_month(ym)
        for _ in range(n):
            po_no += 1
            prod, _cat, cost, _price = rng.choice(PRODUCTS)
            qty = rng.randint(20, 200)
            day = rng.randint(1, dim)
            purchases.append({
                "po_id": f"PO-{po_no}",
                "date": f"{ym}-{day:02d}",
                "supplier": rng.choice(SUPPLIERS),
                "product": prod,
                "category": _cat,
                "qty": qty,
                "unit_cost": cost,
            })
    with open(DATA / "purchases.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(purchases[0].keys()))
        w.writeheader()
        w.writerows(purchases)

    expenses = []
    for ym in MONTHS:
        dim = days_in_month(ym)
        for cat, amt in FIXED_EXPENSES:
            expenses.append({"date": f"{ym}-01", "category": cat,
                             "amount": amt,
                             "note": f"{cat} for {ym}"})
        for cat, lo, hi in VAR_EXPENSES:
            day = rng.randint(2, dim)
            expenses.append({"date": f"{ym}-{day:02d}", "category": cat,
                             "amount": rng.randint(lo, hi),
                             "note": f"{cat} — {ym}"})
    with open(DATA / "expenses.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(expenses[0].keys()))
        w.writeheader()
        w.writerows(expenses)

    with open(DATA / "meta.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["key", "value"])
        w.writerow(["company", "Narmada Traders, Indore"])
        w.writerow(["period", "2026-04 to 2026-10"])
        w.writerow(["as_of", "2026-10-31"])
        w.writerow(["seed", SEED])
        w.writerow(["sales_rows", len(sales)])
        w.writerow(["purchase_rows", len(purchases)])
        w.writerow(["expense_rows", len(expenses)])

    print(f"sales={len(sales)} purchases={len(purchases)} "
          f"expenses={len(expenses)}")


if __name__ == "__main__":
    make_data()
