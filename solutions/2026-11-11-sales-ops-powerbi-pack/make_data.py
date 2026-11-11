#!/usr/bin/env python3
"""make_data.py — deterministic fictional SAP Business One-style data for
SalesOpsPulse (build-2026-11-11).

Company: Sharma Trading Co., Patna — FMCG wholesaler.
Period: 2026-04-01 .. 2026-11-10 (as-of 2026-11-10). Seed 20261111.

Outputs (data/*.csv):
- products.csv — 12 SKUs, 4 categories, price/cost/opening stock/reorder level
- customers.csv — 20 customers, 4 channels, route, credit days/limit
- orders.csv — ~550 sales orders (SAP B1-style sales-order lines)
- receipts.csv — ~90 stock-in lines
- deliveries.csv — deliveries for orders from 2026-09-01 (~150)
- calls.csv — weekly calling-plan rows, 2026-10-12 .. 2026-11-10 (~100)
- receivables.csv — ~130 invoices with payment status (credit customers)
"""
import csv
import datetime
import random
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)
SEED = 20261111
AS_OF = datetime.date(2026, 11, 10)

PRODUCTS = [
    # sku, category, unit_price, cost_price, opening_stock, reorder_level
    ("Rice 25kg", "Staples", 1450, 1280, 180, 60),
    ("Wheat Flour 10kg", "Staples", 420, 365, 240, 80),
    ("Sugar 5kg", "Staples", 230, 205, 300, 100),
    ("Toor Dal 5kg", "Pulses", 640, 560, 150, 50),
    ("Chana Dal 5kg", "Pulses", 310, 270, 180, 60),
    ("Moong Dal 5kg", "Pulses", 520, 455, 120, 40),
    ("Mustard Oil 5L", "Edible Oil", 780, 690, 140, 50),
    ("Sunflower Oil 5L", "Edible Oil", 690, 610, 160, 55),
    ("Refined Oil 1L", "Edible Oil", 165, 145, 400, 120),
    ("Turmeric Powder 1kg", "Spices", 240, 205, 200, 70),
    ("Chilli Powder 1kg", "Spices", 280, 240, 180, 60),
    ("Coriander Powder 1kg", "Spices", 210, 180, 190, 65),
]

CUSTOMERS = [
    # name, channel, route, credit_days, credit_limit
    ("Singh Provisions", "Distributor", "Patna City", 30, 500000),
    ("Gupta Traders", "Distributor", "Danapur", 30, 400000),
    ("Verma Wholesale", "Distributor", "Hajipur", 21, 350000),
    ("Kumar & Sons", "Distributor", "Muzaffarpur", 30, 450000),
    ("Sharma Distributors", "Distributor", "Gaya", 21, 300000),
    ("Ramesh General Store", "Retail", "Patna City", 7, 80000),
    ("Priya Mart", "Retail", "Patna City", 7, 60000),
    ("Anand Kirana", "Retail", "Danapur", 14, 70000),
    ("Sita Stores", "Retail", "Hajipur", 7, 50000),
    ("Mohan Traders", "Retail", "Muzaffarpur", 14, 90000),
    ("Lakshmi General", "Retail", "Gaya", 7, 55000),
    ("City Bazar", "Retail", "Patna City", 14, 100000),
    ("New Mart", "Retail", "Danapur", 7, 65000),
    ("Grand Hotel Patna", "Institutional", "Patna City", 15, 250000),
    ("Shanti Canteen Services", "Institutional", "Danapur", 15, 180000),
    ("Heritage Banquet Hall", "Institutional", "Hajipur", 15, 200000),
    ("City Hospital Canteen", "Institutional", "Patna City", 15, 150000),
    ("QuickBuy Online", "Online", "Patna City", 0, 100000),
    ("BazaarKart Seller", "Online", "Danapur", 0, 80000),
    ("FreshKart Retail", "Online", "Hajipur", 0, 90000),
]

VEHICLES = ["BR01-AA-1010", "BR01-AB-2210", "BR01-AC-3340", "BR01-AD-4450"]
MONTHLY_TARGET = {  # Nov target is the pacing benchmark
    "Apr 2026": 260000, "May 2026": 280000, "Jun 2026": 300000,
    "Jul 2026": 320000, "Aug 2026": 330000, "Sep 2026": 350000,
    "Oct 2026": 380000, "Nov 2026": 400000,
}


def write_csv(name, header, rows):
    with open(DATA / name, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"  {name}: {len(rows)} rows")


def main():
    rng = random.Random(SEED)
    write_csv("products.csv",
              ["sku", "category", "unit_price", "cost_price",
               "opening_stock", "reorder_level"],
              PRODUCTS)
    write_csv("customers.csv",
              ["customer", "channel", "route", "credit_days", "credit_limit"],
              CUSTOMERS)

    prod_by_sku = {p[0]: p for p in PRODUCTS}
    cust_by_name = {c[0]: c for c in CUSTOMERS}
    # Channel weights for picking a customer on an order.
    cust_weights = []
    for c in CUSTOMERS:
        cust_weights.append({"Distributor": 5, "Retail": 3,
                             "Institutional": 2, "Online": 2}[c[1]])

    orders = []
    day = datetime.date(2026, 4, 1)
    oid = 1
    while day <= AS_OF:
        n = 1 if day.weekday() == 6 else rng.randint(1, 4)
        for _ in range(n):
            cust = rng.choices(CUSTOMERS, weights=cust_weights)[0]
            sku = rng.choice(PRODUCTS)
            ch = cust[1]
            qty = (rng.randint(20, 60) if ch == "Distributor"
                   else rng.randint(10, 30) if ch == "Institutional"
                   else rng.randint(5, 20) if ch == "Retail"
                   else rng.randint(5, 15))
            disc = rng.choices([0, 2, 5], weights=[60, 25, 15])[0]
            orders.append([f"SO-{oid:04d}", day.isoformat(), cust[0], ch,
                           sku[0], sku[1], qty, sku[2], disc])
            oid += 1
        day += datetime.timedelta(days=1)
    write_csv("orders.csv",
              ["order_id", "date", "customer", "channel", "sku", "category",
               "qty", "unit_price", "discount_pct"], orders)

    receipts = []
    rid = 1
    day = datetime.date(2026, 4, 3)
    while day <= datetime.date(2026, 10, 31):
        if rng.random() < 0.45:
            sku = rng.choice(PRODUCTS)
            receipts.append([f"RC-{rid:03d}", day.isoformat(), sku[0],
                             rng.randint(50, 200)])
            rid += 1
        day += datetime.timedelta(days=1)
    # Balance pass: no SKU may close negative (the Checks sheet asserts
    # zero negative closings). Top up any short SKU with one balancing
    # receipt; the +30 buffer keeps Sugar/Moong/Coriander LOW (alerts
    # stay live) without going OUT.
    sold_by_sku = {}
    for o in orders:
        sold_by_sku[o[4]] = sold_by_sku.get(o[4], 0) + int(o[6])
    recv_by_sku = {}
    for r in receipts:
        recv_by_sku[r[2]] = recv_by_sku.get(r[2], 0) + int(r[3])
    for p in PRODUCTS:
        closing = (p[4] + recv_by_sku.get(p[0], 0)
                   - sold_by_sku.get(p[0], 0))
        if closing < 0:
            receipts.append([f"RC-{rid:03d}", "2026-08-15", p[0],
                             -closing + 30])
            rid += 1
    write_csv("receipts.csv", ["receipt_id", "date", "sku", "qty"], receipts)

    deliveries = []
    did = 1
    for o in orders:
        odate = datetime.date.fromisoformat(o[1])
        if odate < datetime.date(2026, 9, 1):
            continue
        cust = cust_by_name[o[2]]
        planned = odate + datetime.timedelta(days=rng.randint(1, 2))
        delay = rng.choices([0, 1, 2, 3], weights=[70, 20, 7, 3])[0]
        actual = planned + datetime.timedelta(days=delay)
        if actual > AS_OF:
            actual = AS_OF
            delay = (actual - planned).days
        deliveries.append([f"DL-{did:03d}", o[0], odate.isoformat(),
                           cust[2], rng.choice(VEHICLES),
                           planned.isoformat(), actual.isoformat()])
        did += 1
    write_csv("deliveries.csv",
              ["delivery_id", "order_id", "order_date", "route", "vehicle",
               "planned_date", "actual_date"], deliveries)

    calls = []
    cid = 1
    week = datetime.date(2026, 10, 12)
    while week <= AS_OF:
        for cust in CUSTOMERS:
            cdate = week + datetime.timedelta(days=rng.randint(0, 4))
            if cdate > AS_OF:
                continue
            outcome = rng.choices(["Order placed", "Follow-up", "No order"],
                                  weights=[45, 30, 25])[0]
            value = (rng.randint(5000, 60000) if outcome == "Order placed"
                     else 0)
            calls.append([f"CL-{cid:03d}", cdate.isoformat(), cust[0],
                          cust[1], outcome, value])
            cid += 1
        week += datetime.timedelta(days=7)
    write_csv("calls.csv",
              ["call_id", "date", "customer", "channel", "outcome",
               "order_value"], calls)

    receivables = []
    iid = 1
    credit_orders = [o for o in orders
                     if cust_by_name[o[2]][3] > 0]
    rng.shuffle(credit_orders)
    for o in credit_orders[:130]:
        cust = cust_by_name[o[2]]
        inv_date = datetime.date.fromisoformat(o[1])
        due = inv_date + datetime.timedelta(days=cust[3])
        amount = int(o[6]) * int(o[7]) * (100 - int(o[8])) // 100
        r = rng.random()
        if r < 0.70:
            paid = amount
        elif r < 0.90:
            paid = amount * rng.randint(40, 80) // 100
        else:
            paid = 0
        receivables.append([f"INV-{iid:03d}", o[0], o[2], inv_date.isoformat(),
                            due.isoformat(), amount, paid])
        iid += 1
    receivables.sort(key=lambda x: x[3])
    write_csv("receivables.csv",
              ["invoice_id", "order_id", "customer", "invoice_date",
               "due_date", "amount", "paid_amount"], receivables)

    with open(DATA / "targets.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "target"])
        for m, t in MONTHLY_TARGET.items():
            w.writerow([m, t])
    print("  targets.csv: 8 rows")
    print("DONE")


if __name__ == "__main__":
    main()
