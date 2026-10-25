#!/usr/bin/env python3
"""Generate deterministic SYNTHETIC QuickBooks-style exports for a fictional
construction firm (Maple Ridge Builders Ltd., Victoria BC, CAD).

Outputs (data/):
  qbo_pl.csv   - one row per category per month (May-Oct 2026), signed amounts
                 (revenue positive, costs negative), as a QBO P&L export looks.
  ar.csv       - open/paid customer invoices: InvoiceID, Customer, IssueDate,
                 DueDate, Amount, PaidAmount.
  ap.csv       - vendor bills: BillID, Vendor, DueDate, Amount, PaidAmount.
  cash.csv     - single row: opening bank balance.

All figures are invented demo data (seed 20261025), labelled SYNTHETIC in the
workbook Guide sheet. Nothing here is a real client's data.
"""
import csv
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

MONTHS = ["2026-05", "2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]

# (category, low, high) - monthly amounts, CAD. Costs stored negative.
CATEGORIES = [
    ("Revenue - Residential", 90000, 160000),
    ("Revenue - Commercial", 40000, 90000),
    ("Cost - Labour", -75000, -45000),
    ("Cost - Materials", -70000, -35000),
    ("Cost - Subcontractors", -45000, -20000),
    ("Overhead - Office", -10000, -6000),
    ("Overhead - Insurance", -6000, -4000),
    ("Overhead - Vehicle", -9000, -5000),
]

CUSTOMERS = [
    "Harbourview Developments", "Cedar Ridge Homes", "Westshore Properties",
    "Peninsula Contracting", "Fairfield Renovations", "Gordon Head Builders",
    "Sooke Bay Estates", "Langford Living", "Oak Bay Holdings",
    "Colwood Construction",
]
VENDORS = [
    "Island Lumber Supply", "Pacific Concrete Ltd.", "West Coast Electrical",
    "Salish Plumbing Co.", "Victoria Roofing Supply", "Capital Drywall",
    "Peninsula Scaffolding", "Harbour Paint & Finishes",
]

OPENING_CASH = 84500


def main():
    rng = random.Random(20261025)
    os.makedirs(DATA, exist_ok=True)

    with open(os.path.join(DATA, "qbo_pl.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Month", "Category", "Amount"])
        for m in MONTHS:
            for cat, lo, hi in CATEGORIES:
                w.writerow([m, cat, rng.randint(lo, hi)])

    with open(os.path.join(DATA, "ar.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["InvoiceID", "Customer", "IssueDate", "DueDate",
                    "Amount", "PaidAmount"])
        for i in range(1, 31):
            cust = CUSTOMERS[rng.randrange(len(CUSTOMERS))]
            # Issue dates spread Aug-Oct 2026; Net 30 terms.
            day = rng.randint(1, 28)
            mon = rng.choice([8, 9, 10])
            issue = f"2026-{mon:02d}-{day:02d}"
            due_mon, due_day = (mon + 1, day) if mon < 10 else (11, day)
            if due_mon == 11 and day > 30:
                due_day = 30
            due = f"2026-{due_mon:02d}-{due_day:02d}"
            amt = rng.randint(2000, 28000)
            roll = rng.random()
            if roll < 0.45:
                paid = amt  # paid in full
            elif roll < 0.70:
                paid = rng.randint(500, amt - 500)  # partial
            else:
                paid = 0  # open
            w.writerow([f"INV-{1000 + i}", cust, issue, due, amt, paid])

    with open(os.path.join(DATA, "ap.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["BillID", "Vendor", "DueDate", "Amount", "PaidAmount"])
        for i in range(1, 25):
            vend = VENDORS[rng.randrange(len(VENDORS))]
            mon = rng.choice([8, 9, 10])
            day = rng.randint(1, 28)
            due_mon = mon + 1 if mon < 10 else 11
            amt = rng.randint(1500, 22000)
            roll = rng.random()
            if roll < 0.55:
                paid = amt
            elif roll < 0.80:
                paid = rng.randint(400, amt - 400)
            else:
                paid = 0
            w.writerow([f"BILL-{2000 + i}", vend,
                        f"2026-{due_mon:02d}-{day:02d}", amt, paid])

    with open(os.path.join(DATA, "cash.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Item", "Amount"])
        w.writerow(["Opening bank balance (2026-05-01)", OPENING_CASH])

    print("generated:",
          "qbo_pl.csv (48 rows), ar.csv (30 rows), ap.csv (24 rows), cash.csv")


if __name__ == "__main__":
    main()
