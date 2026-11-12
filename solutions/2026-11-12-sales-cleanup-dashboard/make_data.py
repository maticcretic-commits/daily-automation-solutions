#!/usr/bin/env python3
"""Generate the fictional dirty sales data for SalesClean (build-2026-11-12).

Fictional business: Gupta Traders, Patna — FMCG wholesaler.
Period: 2024-11-01 .. 2026-10-31 (24 months). Deterministic (seed 20261112).

Writes data/raw_sales_dirty.csv (raw rows as "received", dirt included)
and data/targets.csv (monthly revenue targets).

Dirt injected (every injected row has a D-xxxx order id so the cleaning
ground truth stays exact):
- 25 exact duplicate rows (copies of clean rows)
- 20 near-duplicate rows (copies with mangled customer/product name
  case/spacing — they normalise back to the original and must be dropped
  as DUPLICATE)
- 10 valid rows with mangled names only (unique dates — must be KEPT,
  names normalised)
- 6 rows with valid DD/MM/YYYY text dates (must be KEPT, dates normalised)
- 12 blank qty, 10 blank unit_price, 6 qty <= 0, 5 price <= 0
- 8 bad dates (4 impossible like 2025-02-30, 2 "not-a-date", 2 empty)
- 4 unknown customer, 4 unknown product
"""
import csv
import random
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)
SEED = 20261112

PRODUCTS = [
    ("Rice 25kg", "Staples", 1250),
    ("Wheat Flour 10kg", "Staples", 480),
    ("Sugar 5kg", "Staples", 240),
    ("Chana Dal 5kg", "Pulses", 310),
    ("Moong Dal 2kg", "Pulses", 260),
    ("Mustard Oil 5L", "Edible Oil", 720),
    ("Sunflower Oil 5L", "Edible Oil", 640),
    ("Chilli Powder 1kg", "Spices", 280),
    ("Turmeric Powder 1kg", "Spices", 190),
    ("Coriander Powder 1kg", "Spices", 160),
]

CUSTOMERS = [
    ("Sharma Distributors", "Patna", "Distributor"),
    ("Kumar & Sons", "Patna", "Distributor"),
    ("Verma Trading Co.", "Patna", "Wholesale"),
    ("City Hospital Canteen", "Patna", "Institutional"),
    ("Anand General Store", "Muzaffarpur", "Retail"),
    ("Gupta Kirana", "Muzaffarpur", "Retail"),
    ("Singh Traders", "Gaya", "Distributor"),
    ("Rani Sati Stores", "Gaya", "Retail"),
    ("Maa Durga Mart", "Bhagalpur", "Retail"),
    ("Bhagalpur Wholesale Hub", "Bhagalpur", "Wholesale"),
    ("New Bharat Foods", "Patna", "Institutional"),
    ("Lakshmi Traders", "Muzaffarpur", "Distributor"),
    ("Jai Mata Di Store", "Gaya", "Retail"),
    ("Sahu Brothers", "Bhagalpur", "Distributor"),
    ("College Canteen (Magadh Univ.)", "Gaya", "Institutional"),
    ("Railway Canteen (PNBE)", "Patna", "Institutional"),
]

MONTHS = [(y, m) for y in (2024, 2025, 2026) for m in range(1, 13)
          if (y, m) >= (2024, 11) and (y, m) <= (2026, 10)]


def days_in_month(y, m):
    if m == 12:
        return 31
    import datetime
    return (datetime.date(y, m + 1, 1) - datetime.date(y, m, 1)).days


def mangle_name(name, rng):
    """Case/spacing variant of a name that normalises back to itself."""
    style = rng.choice(["upper", "lower", "spaces", "title-mess"])
    if style == "upper":
        return name.upper()
    if style == "lower":
        return name.lower()
    if style == "spaces":
        return "  " + name.replace(" ", "  ") + " "
    return name.title().replace("  ", " ")


def main():
    rng = random.Random(SEED)
    rows = []
    seq = 0

    for (y, m) in MONTHS:
        n = 55  # ~55 order lines per month -> 1,320 clean rows
        for _ in range(n):
            seq += 1
            day = rng.randint(1, days_in_month(y, m))
            cust, region, channel = rng.choice(CUSTOMERS)
            prod, cat, base = rng.choice(PRODUCTS)
            qty = rng.randint(5, 120)
            price = round(base * rng.uniform(0.95, 1.05), 2)
            rows.append({
                "order_id": f"SO-{seq:04d}",
                "date": f"{y:04d}-{m:02d}-{day:02d}",
                "customer": cust, "region": region, "channel": channel,
                "product": prod, "category": cat,
                "qty": qty, "unit_price": price,
            })

    clean_rows = list(rows)
    dirty = list(rows)

    # 10 valid rows with mangled names only (unique dates -> KEPT, normalised)
    for i in range(10):
        r = dict(rng.choice(clean_rows))
        r["order_id"] = f"D-K{i+1:02d}"
        r["date"] = f"2026-10-{10 + i:02d}"  # unique late-Oct dates
        r["customer"] = mangle_name(r["customer"], rng)
        r["product"] = mangle_name(r["product"], rng)
        dirty.append(r)

    # 6 rows with valid DD/MM/YYYY text dates (KEPT, dates normalised).
    # Unique Oct 2026 days so they are NOT duplicates of the copied rows.
    for i in range(6):
        r = dict(rng.choice(clean_rows))
        r["order_id"] = f"D-T{i+1:02d}"
        r["date"] = f"{20 + i}/10/2026"
        dirty.append(r)

    # 25 exact duplicates
    for i, r in enumerate(rng.sample(clean_rows, 25)):
        d = dict(r)
        d["order_id"] = f"D-E{i+1:02d}"
        dirty.append(d)

    # 20 near-duplicates (mangled names, same key after normalisation)
    for i, r in enumerate(rng.sample(clean_rows, 20)):
        d = dict(r)
        d["order_id"] = f"D-N{i+1:02d}"
        d["customer"] = mangle_name(d["customer"], rng)
        d["product"] = mangle_name(d["product"], rng)
        dirty.append(d)

    # Bad-value rows (each a fresh row, distinct order id)
    def bad_row(oid, **kw):
        r = dict(rng.choice(clean_rows))
        r["order_id"] = oid
        r.update(kw)
        dirty.append(r)

    for i in range(12):
        bad_row(f"D-Q{i+1:02d}", qty="")
    for i in range(10):
        bad_row(f"D-P{i+1:02d}", unit_price="")
    for i in range(6):
        bad_row(f"D-Z{i+1:02d}", qty=rng.choice([0, -5, -20]))
    for i in range(5):
        bad_row(f"D-R{i+1:02d}", unit_price=rng.choice([0, -50]))
    bad_dates = ["2025-02-30", "2025-02-30", "2025-13-01", "2024-11-31",
                 "not-a-date", "12-03-2025", "", ""]
    for i, bd in enumerate(bad_dates):
        bad_row(f"D-B{i+1:02d}", date=bd)
    for i in range(4):
        bad_row(f"D-C{i+1:02d}", customer=f"Unknown Buyer {i+1}")
    for i in range(4):
        bad_row(f"D-U{i+1:02d}", product=f"Mystery Item {i+1}")

    rng.shuffle(dirty)

    with open(DATA / "raw_sales_dirty.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["order_id", "date", "customer",
                                          "region", "channel", "product",
                                          "category", "qty", "unit_price"])
        w.writeheader()
        w.writerows(dirty)

    # Monthly targets with festive-season uplift (Oct/Nov higher).
    # Expected monthly revenue is ~Rs 1.7M (55 lines x ~62 units x ~Rs 500),
    # so targets sit just under that for 85-115% attainment.
    with open(DATA / "targets.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "target"])
        for (y, m) in MONTHS:
            base_target = 1_600_000
            if m in (10, 11):
                base_target = 2_100_000
            elif m in (3, 4):
                base_target = 1_800_000
            w.writerow([f"{y:04d}-{m:02d}", base_target])

    print(f"raw rows: {len(dirty)} (clean base {len(clean_rows)})")
    print(f"months: {len(MONTHS)}, targets written")


if __name__ == "__main__":
    main()
