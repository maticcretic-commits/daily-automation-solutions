#!/usr/bin/env python3
"""Generate the fictional source data for SalesInv (build-2026-11-10).

Siuri Traders, a fictional wholesaler in Siuri, India, May-Oct 2026.
Deterministic (seed 20261110). Writes data/*.csv:
  products.csv  — product, category, unit_price, cost_price,
                  opening_stock, reorder_level
  customers.csv — customer, region
  sales.csv     — order_id, date, customer, region, category, product,
                  qty, unit_price
  purchases.csv — purchase_id, date, product, qty, unit_cost

Inventory is engineered so the demo shows all three stock states:
one product ends OUT of stock (closing 0), two end LOW (closing at or
below their reorder level), the rest OK. Purchases are derived from the
sales actually generated (sold + target closing - opening), split across
several restock lines per product.
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
RNG = random.Random(20261110)

PRODUCTS = [
    # product, category, unit_price, cost_price, opening_stock, reorder_level
    ("LED Bulb 9W", "Electronics", 89, 55, 120, 30),
    ("Ceiling Fan", "Electronics", 1450, 980, 45, 15),
    ("Extension Board", "Electronics", 349, 210, 90, 25),
    ("Mixer Grinder", "Electronics", 1899, 1350, 35, 12),
    ("Basmati Rice 5kg", "Grocery", 649, 540, 150, 40),
    ("Mustard Oil 1L", "Grocery", 189, 150, 180, 40),
    ("Wheat Atta 10kg", "Grocery", 459, 380, 140, 35),
    ("Sugar 5kg", "Grocery", 289, 240, 160, 35),
    ("Ball Pen Box", "Stationery", 120, 78, 200, 40),
    ("A4 Notebook", "Stationery", 85, 55, 220, 40),
    ("Geometry Box", "Stationery", 149, 95, 130, 30),
    ("Steel Tiffin", "Home & Kitchen", 299, 190, 80, 20),
    ("Pressure Cooker 5L", "Home & Kitchen", 1199, 820, 40, 12),
    ("Plastic Chair", "Home & Kitchen", 449, 300, 70, 20),
]

CUSTOMERS = [
    ("Sharma General Store", "North"), ("Gupta Traders", "North"),
    ("Patel Kirana", "West"), ("Khan Stores", "North"),
    ("Verma Mart", "East"), ("Singh Provisions", "North"),
    ("Das Traders", "East"), ("Iyer Stores", "South"),
    ("Reddy Mart", "South"), ("Nair Traders", "South"),
    ("Joshi General Store", "West"), ("Mehta Kirana", "West"),
    ("Chopra Stores", "North"), ("Agarwal Mart", "East"),
    ("Bose Traders", "East"), ("Ghosh Stores", "East"),
    ("Menon Mart", "South"), ("Pillai Traders", "South"),
    ("Kulkarni Stores", "West"), ("Rao General Store", "South"),
    ("Saxena Mart", "North"), ("Bhatia Traders", "West"),
    ("Chawla Stores", "North"), ("Kapoor Kirana", "East"),
]

START = date(2026, 5, 1)
END = date(2026, 10, 31)
DAYS = (END - START).days + 1  # 184

OUT_PRODUCT = "Plastic Chair"          # engineered to end OUT (closing 0)
LOW_PRODUCTS = ["Ceiling Fan", "Sugar 5kg"]  # engineered to end LOW


def rand_date():
    # Slightly weighted toward recent months, like real order flow.
    return START + timedelta(days=int(RNG.betavariate(2.2, 1.6) * (DAYS - 1)))


def main():
    DATA.mkdir(parents=True, exist_ok=True)
    prod_by_name = {p[0]: p for p in PRODUCTS}

    with open(DATA / "products.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["product", "category", "unit_price", "cost_price",
                    "opening_stock", "reorder_level"])
        w.writerows(PRODUCTS)

    with open(DATA / "customers.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["customer", "region"])
        w.writerows(CUSTOMERS)

    # ---- sales: 5 guaranteed per product, then random fill to 520 ----
    sales = []
    for name, _cat, price, _cost, _op, _ro in PRODUCTS:
        for _ in range(5):
            cust, region = RNG.choice(CUSTOMERS)
            sales.append([rand_date(), cust, region, _cat, name,
                          RNG.randint(1, 6), price])
    while len(sales) < 520:
        name, cat, price, _cost, _op, _ro = RNG.choice(PRODUCTS)
        cust, region = RNG.choice(CUSTOMERS)
        sales.append([rand_date(), cust, region, cat, name,
                      RNG.randint(1, 12), price])
    sales.sort(key=lambda s: (s[0], s[1]))
    sold = {}
    for s in sales:
        sold[s[4]] = sold.get(s[4], 0) + s[5]

    # ---- engineer the stock states ----
    # OUT product: add drain sales so opening - sold = 0.
    name = OUT_PRODUCT
    drain = prod_by_name[name][4] - sold[name]
    if drain > 0:
        cust, region = RNG.choice(CUSTOMERS)
        cat = prod_by_name[name][1]
        price = prod_by_name[name][2]
        while drain > 0:
            q = min(drain, 12)
            sales.append([rand_date(), cust, region, cat, name, q, price])
            drain -= q
    # LOW products: add sales so closing lands below the reorder level.
    for name in LOW_PRODUCTS:
        _p, _c, price, _cost, opening, reorder = prod_by_name[name]
        cat = prod_by_name[name][1]
        closing = opening - sold.get(name, 0)
        need = closing - (reorder - 3)
        if need > 0:
            cust, region = RNG.choice(CUSTOMERS)
            while need > 0:
                q = min(need, 12)
                sales.append([rand_date(), cust, region, cat, name, q, price])
                need -= q
    sales.sort(key=lambda s: (s[0], s[1]))
    sold = {}
    for s in sales:
        sold[s[4]] = sold.get(s[4], 0) + s[5]

    with open(DATA / "sales.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["order_id", "date", "customer", "region", "category",
                    "product", "qty", "unit_price"])
        for i, (d, cust, region, cat, name, qty, price) in enumerate(sales, 1):
            w.writerow([f"SO-{i:04d}", d.isoformat(), cust, region, cat,
                        name, qty, price])

    # ---- purchases: sold + target closing - opening, split in lines ----
    purchases = []
    targets = {}
    for name, _cat, _price, _cost, opening, reorder in PRODUCTS:
        if name == OUT_PRODUCT:
            targets[name] = 0
        elif name in LOW_PRODUCTS:
            targets[name] = reorder - 3
        else:
            targets[name] = RNG.randint(10, 50)
    for name, _cat, _price, cost, opening, _ro in PRODUCTS:
        total = sold.get(name, 0) + targets[name] - opening
        if total <= 0:
            continue
        lines = RNG.randint(6, 10)
        base, rem = divmod(total, lines)
        for k in range(lines):
            q = base + (1 if k < rem else 0)
            if q <= 0:
                continue
            purchases.append([rand_date(), name, q, cost])
    purchases.sort(key=lambda p: p[0])
    with open(DATA / "purchases.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["purchase_id", "date", "product", "qty", "unit_cost"])
        for i, (d, name, qty, cost) in enumerate(purchases, 1):
            w.writerow([f"PO-{i:04d}", d.isoformat(), name, qty, cost])

    # ---- report ----
    print(f"sales: {len(sales)} orders, purchases: {len(purchases)} lines")
    for name, _c, _p, _co, opening, reorder in PRODUCTS:
        closing = opening + sum(q for _d, n, q, _u in purchases if n == name) \
            - sold.get(name, 0)
        status = "OUT" if closing <= 0 else ("LOW" if closing <= reorder
                                              else "OK")
        print(f"  {name:20s} sold={sold.get(name, 0):4d} "
              f"closing={closing:4d} reorder={reorder:3d} {status}")
    assert sum(1 for name, *_ in PRODUCTS
               if (lambda c: c <= 0)(prod_by_name[name][4]
               + sum(q for _d, n, q, _u in purchases if n == name)
               - sold.get(name, 0))) == 1, "exactly one OUT product expected"


if __name__ == "__main__":
    main()
