#!/usr/bin/env python3
"""Generate deterministic fictional billing/sales/inventory data for AquaPure Distributors.

A packaged-drinking-water distributor in Patna (fictional). Outputs CSVs under data/:
  skus.csv, customers.csv, invoices.csv, invoice_lines.csv, purchases.csv

All money math uses Decimal with ROUND_HALF_UP (matches Excel ROUND).
Run: python3 generate.py [--outdir data] [--seed 20261013]
"""
import argparse
import csv
import random
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP

SEED = 20261013
START = date(2026, 4, 1)
END = date(2026, 9, 30)
GST_RATE = Decimal("0.18")


def r2(x):
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


SKUS = [
    # sku, rate ex-GST, opening stock, reorder level, qty range per line
    ("20L Water Jar", 60, 800, 200, (10, 80)),
    ("1L Bottle Case (12 pc)", 240, 500, 150, (3, 40)),
    ("500ml Bottle Case (24 pc)", 220, 600, 150, (3, 40)),
    ("5L Water Jar", 45, 400, 100, (10, 60)),
    ("Paper Cups Pack (100 pc)", 90, 1000, 300, (5, 30)),
    ("Hot & Cold Water Dispenser", 8500, 25, 5, (1, 2)),
]
SKU_WEIGHTS = [0.35, 0.22, 0.18, 0.12, 0.08, 0.05]

CUSTOMERS = [
    "Sharma Tea Stall", "Hotel Rajwada", "City Gym", "Ganga Boys Hostel",
    "Cafe Brewlane", "OfficeHub Cowork", "MediCare Clinic",
    "Bright School Canteen", "Station Food Plaza", "Wedding Palace Events",
    "FreshMart Store", "TechPark Canteen",
]


def gen(rng):
    sku_map = {s[0]: {"rate": s[1], "opening": s[2], "reorder": s[3], "qrange": s[4]}
               for s in SKUS}

    invoices, lines = [], []
    inv_seq = 0
    day = START
    while day <= END:
        n = rng.choices([0, 1, 2, 3], weights=[0.12, 0.60, 0.22, 0.06])[0]
        if day.weekday() == 6 and n > 0 and rng.random() < 0.5:
            n -= 1  # quieter Sundays
        for _ in range(n):
            inv_seq += 1
            inv_no = f"INV-2026-{inv_seq:04d}"
            customer = rng.choice(CUSTOMERS)
            nlines = rng.choices([1, 2, 3], weights=[0.5, 0.35, 0.15])[0]
            subtotal = Decimal("0")
            for ln in range(1, nlines + 1):
                sku = rng.choices([s[0] for s in SKUS], weights=SKU_WEIGHTS)[0]
                qlo, qhi = sku_map[sku]["qrange"]
                qty = rng.randint(qlo, qhi)
                rate = sku_map[sku]["rate"]
                line_total = r2(Decimal(qty) * Decimal(rate))
                subtotal += line_total
                lines.append({"inv_no": inv_no, "date": day.isoformat(),
                              "line_no": ln, "sku": sku, "qty": qty,
                              "rate": rate})
            subtotal = r2(subtotal)
            gst = r2(subtotal * GST_RATE)
            grand = subtotal + gst
            u = rng.random()
            if u < 0.70:
                paid = grand
            elif u < 0.90:
                paid = r2(grand * Decimal(str(round(rng.uniform(0.4, 0.9), 2))))
            else:
                paid = Decimal("0.00")
            invoices.append({"inv_no": inv_no, "date": day.isoformat(),
                             "customer": customer,
                             "subtotal": str(subtotal), "gst": str(gst),
                             "grand": str(grand), "paid": str(paid)})
        day += timedelta(days=1)

    # Stock-in lots: generous, then fix up any negative running stock.
    purchases = []
    lot_seq = 0
    day = START
    while day <= END:
        if rng.random() < 0.30:
            for sku in rng.sample([s[0] for s in SKUS], rng.randint(1, 3)):
                lot_seq += 1
                qlo, qhi = sku_map[sku]["qrange"]
                qty = rng.randint(qlo, qhi) * rng.randint(3, 6)
                purchases.append({"date": day.isoformat(), "sku": sku,
                                  "lot_no": f"LOT-{lot_seq:03d}", "qty": qty,
                                  "rate": sku_map[sku]["rate"]})
        day += timedelta(days=6)

    def running_stock():
        stock = {s[0]: s[2] for s in SKUS}
        events = []
        for p in purchases:
            events.append((p["date"], "in", p["sku"], p["qty"]))
        for l in lines:
            events.append((l["date"], "out", l["sku"], l["qty"]))
        events.sort()
        worst = {}
        for d, kind, sku, qty in events:
            stock[sku] += qty if kind == "in" else -qty
            worst[sku] = min(worst.get(sku, 10 ** 9), stock[sku])
        return stock, worst

    _, worst = running_stock()
    # Fix-up: for any SKU that would go negative, insert a lot 7 days before
    # the earliest trouble. Deterministic; loop until clean.
    for _ in range(20):
        _, worst = running_stock()
        bad = {s: w for s, w in worst.items() if w < 0}
        if not bad:
            break
        for sku, w in bad.items():
            lot_seq += 1
            purchases.append({"date": (START - timedelta(days=1)).isoformat(),
                              "sku": sku, "lot_no": f"LOT-{lot_seq:03d}",
                              "qty": abs(int(w)) + sku_map[sku]["reorder"],
                              "rate": sku_map[sku]["rate"]})
    purchases.sort(key=lambda p: (p["date"], p["lot_no"]))
    _, worst = running_stock()
    assert all(w >= 0 for w in worst.values()), f"negative stock: {worst}"
    return sku_map, invoices, lines, purchases


def write_csv(path, rows, fields):
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--outdir", default="data")
    ap.add_argument("--seed", type=int, default=SEED)
    args = ap.parse_args()
    rng = random.Random(args.seed)
    sku_map, invoices, lines, purchases = gen(rng)
    import os
    os.makedirs(args.outdir, exist_ok=True)
    write_csv(f"{args.outdir}/skus.csv",
              [{"sku": s[0], "rate": s[1], "opening_stock": s[2],
                "reorder_level": s[3]} for s in SKUS],
              ["sku", "rate", "opening_stock", "reorder_level"])
    write_csv(f"{args.outdir}/customers.csv",
              [{"customer": c} for c in CUSTOMERS], ["customer"])
    write_csv(f"{args.outdir}/invoices.csv", invoices,
              ["inv_no", "date", "customer", "subtotal", "gst", "grand", "paid"])
    write_csv(f"{args.outdir}/invoice_lines.csv", lines,
              ["inv_no", "date", "line_no", "sku", "qty", "rate"])
    write_csv(f"{args.outdir}/purchases.csv", purchases,
              ["date", "sku", "lot_no", "qty", "rate"])
    print(f"invoices={len(invoices)} lines={len(lines)} purchases={len(purchases)}")


if __name__ == "__main__":
    main()
