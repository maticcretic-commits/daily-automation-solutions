#!/usr/bin/env python3
"""WarehousePulse — generate fictional warehouse data (seeded, deterministic).

Company: Sharma Wholesale Traders, Patna (FMCG wholesale warehouse, INR).
Period: opening stock as of 2026-09-30; receipts/issues/adjustments 2026-10-01..18.
Outputs data/*.csv consumed by build_workbook.py.
"""
import csv
import os
import random
from datetime import date, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
os.makedirs(DATA, exist_ok=True)

rng = random.Random(20261019)

PERIOD_START = date(2026, 10, 1)
PERIOD_DAYS = 18  # Oct 1..18

# sku, name, category, uom, reorder_level, std_cost (INR)
SKUS = [
    # Staples
    ("ST-001", "Basmati Rice 5kg", "Staples", "bag", 120, 485),
    ("ST-002", "Wheat Flour 10kg", "Staples", "bag", 150, 425),
    ("ST-003", "Toor Dal 1kg", "Staples", "pkt", 200, 165),
    ("ST-004", "Sugar 5kg", "Staples", "bag", 100, 230),
    ("ST-005", "Mustard Oil 1L", "Staples", "btl", 180, 198),
    ("ST-006", "Salt 1kg", "Staples", "pkt", 250, 32),
    ("ST-007", "Poha 500g", "Staples", "pkt", 160, 45),
    ("ST-008", "Sona Masoori Rice 5kg", "Staples", "bag", 110, 340),
    # Beverages
    ("BV-001", "Masala Chai 250g", "Beverages", "pkt", 140, 145),
    ("BV-002", "Instant Coffee 100g", "Beverages", "jar", 90, 285),
    ("BV-003", "Mango Juice 1L", "Beverages", "btl", 200, 115),
    ("BV-004", "Cola 750ml x12 case", "Beverages", "case", 80, 420),
    ("BV-005", "Lassi 200ml x24 case", "Beverages", "case", 110, 360),
    ("BV-006", "Green Tea 100 bags", "Beverages", "box", 70, 195),
    ("BV-007", "Energy Drink 250ml x24", "Beverages", "case", 60, 1080),
    ("BV-008", "Cold Coffee 300ml", "Beverages", "btl", 90, 95),
    # Snacks
    ("SN-001", "Potato Chips 150g x20", "Snacks", "case", 90, 520),
    ("SN-002", "Namkeen Mix 400g", "Snacks", "pkt", 170, 95),
    ("SN-003", "Biscuits 300g x12", "Snacks", "case", 100, 300),
    ("SN-004", "Dry Fruit Mix 500g", "Snacks", "pkt", 60, 445),
    ("SN-005", "Instant Noodles 70g x12", "Snacks", "case", 130, 168),
    ("SN-006", "Roasted Peanuts 200g", "Snacks", "pkt", 150, 75),
    ("SN-007", "Chocolate Bar 40g x30", "Snacks", "case", 70, 750),
    ("SN-008", "Popcorn Kernels 1kg", "Snacks", "pkt", 120, 88),
    # Personal Care
    ("PC-001", "Bath Soap 125g x4", "Personal Care", "pack", 160, 132),
    ("PC-002", "Shampoo 340ml", "Personal Care", "btl", 120, 210),
    ("PC-003", "Toothpaste 150g", "Personal Care", "tube", 200, 105),
    ("PC-004", "Detergent Powder 1kg", "Personal Care", "pkt", 180, 128),
    ("PC-005", "Hair Oil 200ml", "Personal Care", "btl", 100, 165),
    ("PC-006", "Face Wash 150ml", "Personal Care", "tube", 90, 175),
    # Household
    ("HH-001", "Dish Bar 200g x3", "Household", "pack", 170, 96),
    ("HH-002", "Floor Cleaner 1L", "Household", "btl", 110, 145),
    ("HH-003", "LED Bulb 9W x10", "Household", "box", 80, 690),
    ("HH-004", "Garbage Bags 30pc", "Household", "roll", 140, 120),
    ("HH-005", "Broom + Dustpan set", "Household", "set", 90, 185),
    ("HH-006", "Storage Container 5L", "Household", "pc", 100, 240),
]

BINS = [f"{z}{i}" for z in "ABC" for i in range(1, 5)]  # A1..C4
BIN_CAPACITY = 2000  # units per bin

# Popularity weights: fast movers get more issue lines
POPULAR = {"ST-001", "ST-005", "BV-003", "SN-001", "SN-005", "PC-003", "ST-002"}


def main():
    # Assign each SKU a bin (deterministic, round-robin with rng shuffle)
    order = list(range(len(SKUS)))
    rng.shuffle(order)
    sku_bin = {}
    for i, idx in enumerate(order):
        sku_bin[SKUS[idx][0]] = BINS[i % len(BINS)]

    # Opening stock as of 2026-09-30
    opening = {}
    for sku, *_rest in SKUS:
        q = rng.randint(40, 420)
        if sku in ("SN-004", "PC-006"):  # deliberately thin openers
            q = rng.randint(8, 20)
        opening[sku] = q

    receipts, issues, adjustments = [], [], []
    grn_n = do_n = adj_n = 0
    avail = dict(opening)  # running stock; issues/adjustments can never
    # drive it negative (a negative close would trip the Checks sheet)
    for d in range(PERIOD_DAYS):
        day = PERIOD_START + timedelta(days=d)
        # receipts: 5-9 lines/day (BV-004 cola gets none all period:
        # it drains toward reorder, lighting up the alert)
        for _ in range(rng.randint(5, 9)):
            grn_n += 1
            sku = rng.choice([s for s in SKUS if s[0] != "BV-004"])[0]
            cost = next(s[5] for s in SKUS if s[0] == sku)
            # small cost jitter vs std cost (purchase price variance)
            unit_cost = round(cost * rng.uniform(0.97, 1.04), 2)
            qty = rng.randint(20, 220)
            receipts.append([day.isoformat(), f"GRN-{grn_n:04d}", sku, qty,
                             unit_cost, sku_bin[sku]])
            avail[sku] += qty
        # issues: 11-18 lines/day, popular SKUs more often; capped at stock
        for _ in range(rng.randint(11, 18)):
            pool = list(POPULAR) * 3 + [s[0] for s in SKUS]
            sku = rng.choice(pool)
            qty = min(rng.randint(4, 60), avail[sku])
            if qty <= 0:
                continue
            do_n += 1
            issues.append([day.isoformat(), f"DO-{do_n:04d}", sku, qty,
                           sku_bin[sku]])
            avail[sku] -= qty
        # adjustments: ~1-2/day, never below zero stock
        for _ in range(rng.randint(0, 2)):
            sku = rng.choice(SKUS)[0]
            qty = max(-rng.randint(2, 25), -avail[sku])
            if qty == 0:
                continue
            adj_n += 1
            reason = rng.choice(["Damaged in handling", "Expired / quality hold",
                                 "Cycle-count correction", "Pilferage"])
            adjustments.append([day.isoformat(), f"ADJ-{adj_n:04d}", sku, qty,
                                reason])
            avail[sku] += qty

    # Force a genuine STOCKOUT on SN-005 (noodles): one final dispatch that
    # takes the remainder to exactly zero on the last day.
    if avail["SN-005"] > 0:
        do_n += 1
        q = avail["SN-005"]
        issues.append([date(2026, 10, 18).isoformat(), f"DO-{do_n:04d}",
                       "SN-005", q, sku_bin["SN-005"]])
        avail["SN-005"] = 0

    def write(name, header, rows):
        with open(os.path.join(DATA, name), "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)

    write("skus.csv",
          ["sku", "name", "category", "uom", "reorder_level", "std_cost", "bin"],
          [[s[0], s[1], s[2], s[3], s[4], s[5], sku_bin[s[0]]] for s in SKUS])
    write("opening.csv", ["sku", "opening_qty"],
          [[s[0], opening[s[0]]] for s in SKUS])
    write("receipts.csv",
          ["date", "grn", "sku", "qty", "unit_cost", "bin"], receipts)
    write("issues.csv", ["date", "do_no", "sku", "qty", "bin"], issues)
    write("adjustments.csv",
          ["date", "adj_id", "sku", "qty", "reason"], adjustments)
    write("bins.csv", ["bin", "capacity"],
          [[b, BIN_CAPACITY] for b in BINS])

    print(f"wrote data/: {len(SKUS)} skus, {len(receipts)} receipts, "
          f"{len(issues)} issues, {len(adjustments)} adjustments")


if __name__ == "__main__":
    main()
