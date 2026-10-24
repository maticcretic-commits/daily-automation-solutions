#!/usr/bin/env python3
"""RefreshForge — deterministic synthetic source data.

Fictional demo company: "Sharma Foods Processing, Patna" (food processing
plant). Four departmental daily exports, 01-Oct-2026 .. 23-Oct-2026, written
to drop/ (the "watch folder" a Power Automate flow would poll):

  drop/sales.csv      — distributor orders (order_id natural key)
  drop/production.csv — line batches (batch_id natural key)
  drop/inventory.csv  — stock movements (txn_id natural key)
  drop/issues.csv     — quality/ops issues (issue_id natural key)

data/skus.csv — product master (12 SKUs, unit price + opening stock).

Everything is seeded (random.Random(20261024)) so the output is byte-stable.
Deliberate wrinkle: 14-Oct has a machine-breakdown day (production dip,
issue spike) — the Daily sheet's conditional formatting flags it.
"""
import csv
import os
import random

BASE = os.path.dirname(os.path.abspath(__file__))
DROP = os.path.join(BASE, "drop")
DATA = os.path.join(BASE, "data")

SEED = 20261024
START = (2026, 10, 1)
DAYS = 23  # Oct 1..23

SKUS = [
    ("ATTA-5KG", "Whole Wheat Atta 5kg", "Staples", "bag", 235.0, 420),
    ("RICE-25KG", "Parboiled Rice 25kg", "Staples", "bag", 1150.0, 260),
    ("DAL-1KG", "Arhar Dal 1kg", "Pulses", "pkt", 148.0, 900),
    ("OIL-1L", "Mustard Oil 1L", "Oils", "btl", 195.0, 700),
    ("SUGAR-1KG", "Sugar 1kg", "Staples", "pkt", 46.0, 1500),
    ("SALT-1KG", "Iodised Salt 1kg", "Staples", "pkt", 28.0, 2200),
    ("BESAN-1KG", "Besan 1kg", "Pulses", "pkt", 92.0, 800),
    ("SUJI-1KG", "Suji 1kg", "Staples", "pkt", 54.0, 950),
    ("MAIDA-1KG", "Maida 1kg", "Staples", "pkt", 48.0, 1000),
    ("TEA-250G", "CTC Tea 250g", "Beverages", "pkt", 135.0, 1100),
    ("BISC-400G", "Glucose Biscuits 400g", "Snacks", "pkt", 38.0, 1600),
    ("NOODLE-280G", "Instant Noodles 280g", "Snacks", "pkt", 56.0, 1300),
]

CUSTOMERS = [
    "Verma Distributors", "Gupta Traders", "Singh & Sons", "Patna Agro Mart",
    "Mithila Foods", "Kosi Marketing", "Saran Supply Co.", "Bhojpur Traders",
    "Nalanda Foods", "Magadh Distributors",
]

LINES = ["Line-A", "Line-B"]
ISSUE_CATS = ["Quality", "Machine", "Material", "Safety"]


def day(n):
    return f"2026-10-{n:02d}"


def write(name, header, rows, folder=DROP):
    path = os.path.join(folder, name)
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"wrote {path}: {len(rows)} rows")
    return path


def generate():
    rng = random.Random(SEED)

    write("skus.csv",
          ["sku", "name", "category", "uom", "unit_price", "opening_qty"],
          [[s[0], s[1], s[2], s[3], f"{s[4]:.2f}", s[5]] for s in SKUS],
          folder=DATA)

    prices = {s[0]: s[4] for s in SKUS}
    sku_ids = [s[0] for s in SKUS]

    sales, production, inventory, issues = [], [], [], []
    on = 1
    bn = 1
    tn = 1
    qn = 1
    for d in range(1, DAYS + 1):
        dt = day(d)
        breakdown = (d == 14)  # machine-breakdown day

        # ---- sales orders ----
        n_orders = rng.randint(6, 14)
        for _ in range(n_orders):
            sku = rng.choice(sku_ids)
            qty = rng.randint(5, 60)
            sales.append([f"S{on:04d}", dt, rng.choice(CUSTOMERS),
                          sku, qty, f"{prices[sku]:.2f}"])
            on += 1

        # ---- production batches ----
        for line in LINES:
            if breakdown and line == "Line-B":
                continue  # Line-B down all day on the 14th
            n_batch = rng.randint(1, 3)
            for _ in range(n_batch):
                sku = rng.choice(sku_ids)
                produced = rng.randint(200, 900)
                rejected = (rng.randint(30, 90) if breakdown
                            else rng.randint(2, 18))
                production.append([f"B{bn:04d}", dt, line, sku,
                                   produced, rejected])
                bn += 1

        # ---- inventory movements ----
        for _ in range(rng.randint(3, 6)):
            sku = rng.choice(sku_ids)
            movement = rng.choice(["IN", "IN", "OUT", "OUT", "OUT"])
            qty = rng.randint(50, 400)
            inventory.append([f"T{tn:04d}", dt, sku, movement, qty])
            tn += 1

        # ---- issues ----
        n_issues = rng.randint(0, 2) + (5 if breakdown else 0)
        for _ in range(n_issues):
            cat = ("Machine" if breakdown and rng.random() < 0.7
                   else rng.choice(ISSUE_CATS))
            sev = rng.choice(["P1", "P2", "P2", "P3", "P3"])
            status = rng.choice(["Open", "Closed", "Closed"])
            issues.append([f"Q{qn:04d}", dt, cat, sev, status,
                           f"{cat} issue on {dt} ({sev})"])
            qn += 1

    write("sales.csv",
          ["order_id", "date", "customer", "sku", "qty", "unit_price"], sales)
    write("production.csv",
          ["batch_id", "date", "line", "sku", "qty_produced",
           "qty_rejected"], production)
    write("inventory.csv",
          ["txn_id", "date", "sku", "movement", "qty"], inventory)
    write("issues.csv",
          ["issue_id", "date", "category", "severity", "status",
           "description"], issues)
    print(f"totals: sales={len(sales)} production={len(production)} "
          f"inventory={len(inventory)} issues={len(issues)}")


if __name__ == "__main__":
    generate()
