#!/usr/bin/env python3
"""Generate deterministic synthetic data for the PaymentTracker demo.

Fictional wholesaler: Sharma Traders, New Delhi — a building-materials
wholesaler selling on 30-day credit to local retailers. Period:
Apr-Oct 2026, invoices due 30 days after issue, as-of date 3 Nov 2026
(the Assumptions sheet date). All figures invented for this demo; no
real client data. Seed fixed so re-runs are identical. Rows are sorted
by due date, so the Overdue sheet lists the earliest-due invoices first.
"""
import csv
import datetime
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
SEED = 20261103
AS_OF = datetime.date(2026, 11, 3)

CUSTOMERS = [
    "Aggarwal Hardware", "Bansal Traders", "Chawla Building Supplies",
    "Delhi Doors & Ply", "Gupta Sanitary House", "Jain Electricals",
    "Kapoor Paint Centre", "Khanna Timber Depot", "Kumar & Sons Traders",
    "Mahajan Steel Traders", "Malhotra Cement Store", "Mehta Glass House",
    "Mittal Hardware Mart", "Nanda Plywood Centre", "Oberoi Bath Fittings",
    "Punjab Iron Store", "Rajdhani Tiles Hub", "Rana Construction Supply",
    "Sachdeva Electricals", "Sethi Sanitaryware", "Sharma Marble House",
    "Singh Hardware Store", "Soni Paints & Hardware", "Taneja Timber Traders",
    "Verma Building Materials", "Vohra Steel & Cement", "Yadav Traders",
    "Anand Ceramics",
]

# (month, invoice count) — Apr-Oct 2026
PLAN = [(4, 16), (5, 18), (6, 17), (7, 18), (8, 17), (9, 18), (10, 16)]


def round100(x):
    return int(round(x / 100.0) * 100)


def main():
    rng = random.Random(SEED)
    rows = []
    seq = 0
    for month, count in PLAN:
        for _ in range(count):
            seq += 1
            customer = rng.choice(CUSTOMERS)
            day = rng.randint(1, 28)
            inv_date = datetime.date(2026, month, day)
            due = inv_date + datetime.timedelta(days=30)
            amount = round100(rng.randint(8000, 180000))
            roll = rng.random()
            overdue = due < AS_OF
            if roll < 0.45:
                paid = amount
            elif roll < 0.65:
                paid = round100(amount * rng.uniform(0.40, 0.80))
            elif roll < 0.80:
                # Unpaid: some already overdue, some still current.
                paid = 0
            else:
                # Partial and still outstanding.
                paid = round100(amount * rng.uniform(0.20, 0.60)) if overdue else 0
                if not overdue and rng.random() < 0.5:
                    paid = round100(amount * rng.uniform(0.30, 0.70))
            paid = min(paid, amount)
            rows.append({
                "customer": customer,
                "invoice_no": f"INV-2026-{seq:03d}",
                "invoice_date": inv_date.isoformat(),
                "due_date": due.isoformat(),
                "amount": amount,
                "paid": paid,
            })
    # Due date order == invoice date order (fixed 30-day terms), but sort
    # explicitly so the Overdue sheet ordering is by due date by construction.
    rows.sort(key=lambda r: (r["due_date"], r["invoice_no"]))
    os.makedirs(DATA, exist_ok=True)
    with open(os.path.join(DATA, "payments.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["customer", "invoice_no",
                                          "invoice_date", "due_date",
                                          "amount", "paid"])
        w.writeheader()
        w.writerows(rows)
    print(f"wrote {len(rows)} invoices, {len(CUSTOMERS)} customers")


if __name__ == "__main__":
    main()
