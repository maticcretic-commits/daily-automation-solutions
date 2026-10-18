#!/usr/bin/env python3
"""InvoiceLedger — fictional invoice/receipt dataset generator.

Produces data/invoices.csv and data/payments.csv for a fictional wholesale
garments trader ("Gupta Garments & Textiles", Patna) over Oct-2025..Sep-2026.

Every record is synthetic. Deterministic via seed so the workbook build and
the test ground truth always agree.
"""
import csv
import os
import random
from datetime import date, timedelta

SEED = 18102026
REPORT_DATE = date(2026, 9, 30)
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

SALES_PARTIES = [
    "Sharma Retail Store", "Patna Fashion House", "Kumkum Saree Sadan",
    "New Look Garments", "City Centre Apparel", "Bazaar Collection",
    "Trendy Threads", "Lakshmi Cloth Store", "Vishal Mega Mart (franchise)",
    "Gaya Textile Traders", "Muzaffarpur Cloth House", "Darbhanga Dress Point",
    "Bhagalpur Silk Centre", "Purnia Fashion Hub", "Arrah Garment Plaza",
]
PURCHASE_PARTIES = [
    "Surat Textile Mills", "Ludhiana Knitwear Co", "Tiruppur Garment Works",
    "Delhi Fabric Suppliers", "Kolkata Hosiery House", "Jaipur Print Works",
]

SALES_CATS = [("Fabrics", 0.42), ("Garments", 0.41), ("Accessories", 0.17)]
GST_RATES = {"Fabrics": 5, "Garments": 12, "Accessories": 12,
             "Inventory purchases": 5, "Rent": 18, "Salaries": 0,
             "Utilities": 18, "Transport": 5, "Marketing": 18,
             "Professional fees": 18, "Misc": 18}

# Seasonal multiplier per month index (0 = Oct-25 ... 11 = Sep-26)
SEASON = [1.25, 1.30, 1.05, 1.15, 1.20, 1.00, 0.90, 0.85, 0.80, 0.90, 1.00, 1.10]

MONTHS = [(2025, 10), (2025, 11), (2025, 12), (2026, 1), (2026, 2), (2026, 3),
          (2026, 4), (2026, 5), (2026, 6), (2026, 7), (2026, 8), (2026, 9)]

EXPENSE_PLAN = [  # (category, monthly_amount or None for ad-hoc, count_per_month)
    ("Rent", 25000, 1),
    ("Salaries", 85000, 1),
    ("Utilities", 6500, 1),
    ("Transport", None, 6),
    ("Marketing", None, 2),
    ("Professional fees", None, 1),
    ("Misc", None, 3),
]


def pick_cat(rng):
    x = rng.random()
    acc = 0.0
    for name, w in SALES_CATS:
        acc += w
        if x <= acc:
            return name
    return SALES_CATS[-1][0]


def month_days(y, m):
    if m == 12:
        nxt = date(y + 1, 1, 1)
    else:
        nxt = date(y, m + 1, 1)
    return (nxt - date(y, m, 1)).days


def main():
    rng = random.Random(SEED)
    invoices = []
    payments = []
    seq = {"S": 0, "P": 0, "E": 0}
    pay_seq = 0

    def add_invoice(inv_date, party, itype, category, amount, gst_rate,
                    pay_lag_days, unpaid_prob, partial_prob=0.0):
        nonlocal pay_seq
        pfx = {"SALES": "S", "PURCHASE": "P", "EXPENSE": "E"}[itype]
        seq[pfx] += 1
        inv_id = f"{pfx}-{seq[pfx]:05d}"
        gst_amt = round(amount * gst_rate / 100.0, 2)
        total = round(amount + gst_amt, 2)
        pay_date = inv_date + timedelta(days=pay_lag_days)
        if pay_date > REPORT_DATE or rng.random() < unpaid_prob:
            paid_amount, paid_date = 0.0, ""
        elif rng.random() < partial_prob:
            paid_amount = round(total * rng.uniform(0.4, 0.8), 2)
            paid_date = pay_date.isoformat()
        else:
            paid_amount, paid_date = total, pay_date.isoformat()
        invoices.append({
            "inv_id": inv_id, "date": inv_date.isoformat(), "party": party,
            "type": itype, "category": category, "amount": f"{amount:.2f}",
            "gst_rate": gst_rate, "gst_amt": f"{gst_amt:.2f}",
            "total": f"{total:.2f}", "paid_amount": f"{paid_amount:.2f}",
            "paid_date": paid_date,
        })
        if paid_amount > 0:
            pay_seq += 1
            payments.append({
                "pay_id": f"PAY-{pay_seq:05d}", "inv_id": inv_id,
                "type": itype, "category": category,
                "date": paid_date, "amount": f"{paid_amount:.2f}",
            })

    for mi, (y, m) in enumerate(MONTHS):
        ndays = month_days(y, m)
        # --- sales invoices: ~32/month scaled by season ---
        n_sales = max(18, int(round(32 * SEASON[mi] * rng.uniform(0.9, 1.1))))
        for _ in range(n_sales):
            d = date(y, m, rng.randint(1, ndays))
            cat = pick_cat(rng)
            amt = round(rng.uniform(8000, 95000), 2)
            lag = int(rng.triangular(5, 75, 22))
            add_invoice(d, rng.choice(SALES_PARTIES), "SALES", cat, amt,
                        GST_RATES[cat], lag, unpaid_prob=0.12, partial_prob=0.10)
        # --- purchase invoices: ~18/month ---
        n_pur = max(10, int(round(18 * rng.uniform(0.85, 1.1))))
        for _ in range(n_pur):
            d = date(y, m, rng.randint(1, ndays))
            amt = round(rng.uniform(10000, 120000), 2)
            lag = int(rng.triangular(5, 35, 14))
            add_invoice(d, rng.choice(PURCHASE_PARTIES), "PURCHASE",
                        "Inventory purchases", amt,
                        GST_RATES["Inventory purchases"], lag,
                        unpaid_prob=0.06, partial_prob=0.05)
        # --- expenses ---
        for cat, fixed, n in EXPENSE_PLAN:
            for _ in range(n if fixed is None else 1):
                d = date(y, m, rng.randint(1, ndays))
                if fixed is not None:
                    amt = round(fixed * rng.uniform(0.98, 1.02), 2)
                elif cat == "Transport":
                    amt = round(rng.uniform(1200, 8500), 2)
                elif cat == "Marketing":
                    amt = round(rng.uniform(5000, 25000), 2)
                elif cat == "Professional fees":
                    amt = round(rng.uniform(4000, 15000), 2)
                else:
                    amt = round(rng.uniform(800, 6000), 2)
                lag = int(rng.triangular(0, 10, 3))
                add_invoice(d, "Various", "EXPENSE", cat, amt,
                            GST_RATES[cat], lag, unpaid_prob=0.02)

    os.makedirs(OUT_DIR, exist_ok=True)
    with open(os.path.join(OUT_DIR, "invoices.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["inv_id", "date", "party", "type",
                                          "category", "amount", "gst_rate",
                                          "gst_amt", "total", "paid_amount",
                                          "paid_date"])
        w.writeheader()
        w.writerows(invoices)
    with open(os.path.join(OUT_DIR, "payments.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["pay_id", "inv_id", "type",
                                          "category", "date", "amount"])
        w.writeheader()
        w.writerows(payments)
    n_sales = sum(1 for i in invoices if i["type"] == "SALES")
    n_unpaid = sum(1 for i in invoices
                   if float(i["total"]) - float(i["paid_amount"]) > 0.005)
    print(f"invoices: {len(invoices)} (sales {n_sales}), "
          f"payments: {len(payments)}, with balance: {n_unpaid}")


if __name__ == "__main__":
    main()
