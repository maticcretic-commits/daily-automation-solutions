#!/usr/bin/env python3
"""Generate the fictional sales data for SalesPerf BI Pack (build-2026-11-13).

Fictional business: Agarwal Traders, Patna — FMCG wholesaler.
Period: 2026-01-01 .. 2026-10-31 (10 months). Deterministic (seed 20261113).

Writes:
- data/orders.csv — one row per order (order_id, date, customer, region,
  channel, product, category, qty, unit_price, amount).
- data/targets.csv — monthly revenue targets. Demo targets are set as a
  fixed percentage of each month's actual revenue (factors cycle
  0.92–1.08) so the dashboard's attainment bands (red <80%, amber
  80–99%, green >=100%) are all visible; in production the client
  supplies real targets. The method is stated here and in the README —
  the targets are an input file, not computed by the workbook.
- data/products.csv, data/customers.csv — the masters the workbook
  looks up against (category via INDEX/MATCH, region via VLOOKUP).
"""
import csv
import datetime
import random
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)
SEED = 20261113

PRODUCTS = [
    ("Rice 25kg", "Staples", 1250),
    ("Basmati Rice 5kg", "Staples", 420),
    ("Wheat Flour 10kg", "Staples", 480),
    ("Sugar 5kg", "Staples", 240),
    ("Chana Dal 5kg", "Pulses", 310),
    ("Toor Dal 5kg", "Pulses", 340),
    ("Moong Dal 2kg", "Pulses", 260),
    ("Mustard Oil 5L", "Edible Oil", 720),
    ("Sunflower Oil 5L", "Edible Oil", 640),
    ("Groundnut Oil 5L", "Edible Oil", 690),
    ("Chilli Powder 1kg", "Spices", 280),
    ("Turmeric Powder 1kg", "Spices", 190),
    ("Coriander Powder 1kg", "Spices", 160),
    ("Garam Masala 500g", "Spices", 230),
]

CUSTOMERS = [
    ("Agarwal Distributors", "Patna", "Distributor"),
    ("Sharma Distributors", "Patna", "Distributor"),
    ("Kumar & Sons", "Patna", "Distributor"),
    ("Verma Trading Co.", "Patna", "Wholesale"),
    ("City Hospital Canteen", "Patna", "Institutional"),
    ("New Bharat Foods", "Patna", "Institutional"),
    ("Railway Canteen (PNBE)", "Patna", "Institutional"),
    ("Anand General Store", "Muzaffarpur", "Retail"),
    ("Gupta Kirana", "Muzaffarpur", "Retail"),
    ("Lakshmi Traders", "Muzaffarpur", "Distributor"),
    ("Sahu Brothers", "Bhagalpur", "Distributor"),
    ("Maa Durga Mart", "Bhagalpur", "Retail"),
    ("Bhagalpur Wholesale Hub", "Bhagalpur", "Wholesale"),
    ("Singh Traders", "Gaya", "Distributor"),
    ("Rani Sati Stores", "Gaya", "Retail"),
    ("Jai Mata Di Store", "Gaya", "Retail"),
    ("College Canteen (Magadh Univ.)", "Gaya", "Institutional"),
    ("Bhatia General Store", "Patna", "Retail"),
    ("Siuri Traders", "Siuri", "Wholesale"),
    ("Dhanbad Provisions", "Dhanbad", "Wholesale"),
    ("Ranchi Wholesale Mart", "Ranchi", "Distributor"),
    ("Bokaro Kirana House", "Bokaro", "Retail"),
    ("Jamshedpur Foods", "Jamshedpur", "Wholesale"),
    ("Hazaribagh Traders", "Hazaribagh", "Retail"),
    ("Deoghar General Store", "Deoghar", "Retail"),
    ("Motihari Distributors", "Motihari", "Distributor"),
]

# Target factor per month (Jan..Oct): attainment will read ~1/factor.
TARGET_FACTORS = [0.94, 1.05, 0.98, 1.08, 0.92, 1.30, 0.97, 1.06, 0.93, 1.04]


def main():
    rng = random.Random(SEED)
    start = datetime.date(2026, 1, 1)
    end = datetime.date(2026, 10, 31)
    orders = []
    oid = 0
    day = start
    while day <= end:
        month_idx = day.month - 1
        growth = 1.0 + 0.06 * month_idx  # gentle growth through the year
        weekend = 0.55 if day.weekday() >= 5 else 1.0
        n = rng.choices([2, 3, 4, 5], weights=[1, 3, 4, 2])[0]
        n = max(1, round(n * growth * weekend))
        for _ in range(n):
            oid += 1
            prod, cat, base_price = rng.choice(PRODUCTS)
            cust, region, channel = rng.choice(CUSTOMERS)
            qty = rng.choices([5, 10, 15, 20, 25, 40, 50, 80, 100],
                              weights=[3, 5, 5, 4, 3, 2, 2, 1, 1])[0]
            price = round(base_price * rng.uniform(0.92, 1.10), 2)
            orders.append({
                "order_id": f"ORD-{oid:05d}",
                "date": day.isoformat(),
                "customer": cust, "region": region, "channel": channel,
                "product": prod, "category": cat,
                "qty": qty, "unit_price": price,
                "amount": round(qty * price, 2),
            })
        day += datetime.timedelta(days=1)

    with open(DATA / "orders.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(orders[0].keys()))
        w.writeheader()
        w.writerows(orders)

    monthly_rev = {}
    for o in orders:
        m = o["date"][:7]
        monthly_rev[m] = monthly_rev.get(m, 0.0) + o["amount"]
    with open(DATA / "targets.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "target"])
        for i, m in enumerate(sorted(monthly_rev)):
            target = round(monthly_rev[m] * TARGET_FACTORS[i] / 1000) * 1000
            w.writerow([m, target])

    with open(DATA / "products.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["product", "category", "base_price"])
        w.writerows(PRODUCTS)
    with open(DATA / "customers.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["customer", "region", "channel"])
        w.writerows(CUSTOMERS)

    total = sum(o["amount"] for o in orders)
    print(f"orders={len(orders)} total_revenue={total:,.2f} "
          f"months={len(monthly_rev)}")
    for m in sorted(monthly_rev):
        print(f"  {m} revenue={monthly_rev[m]:,.2f}")


if __name__ == "__main__":
    main()
