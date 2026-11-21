#!/usr/bin/env python3
"""Generate fictional bank-statement data for FinanceFlow AI.

Fictional firm: Kesarwani Traders, Patna (wholesale FMCG), Apr-Oct 2026.
Deterministic (seed 20261121). The narrations are deliberately messy,
like a real bank statement export: UPI/NEFT/IMPS prefixes, mixed case,
abbreviations, a couple of cryptic lines.

Injected defects (the pipeline must catch them):
  * one exact duplicate transaction (same date/amount/narration twice)
  * one stated-balance error (running balance column off by Rs 500)
  * three unclassifiable narrations (no keyword match)
  * one unusually large purchase (Rs 78,400) flagged as an outlier

Outputs (into data/):
  bank_statement.csv  date,narration,debit,credit,balance
  chart_of_accounts.csv  category,kind,keywords (semicolon-separated)
"""
import csv
import os
import random
from datetime import date, timedelta

SEED = 20261121
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(OUT, exist_ok=True)

rng = random.Random(SEED)

CUSTOMERS = ["ABC Retail", "Sharma Kirana", "Gupta Stores", "Verma Mart",
             "Singh Traders", "Patel General Store", "Kumar Traders",
             "Yadav Super Mart", "Mishra Provisions", "Jha Daily Needs"]
SUPPLIERS = ["Hindustan Foods Ltd", "Bihar Agro Supplies", "Patna Wholesale Co",
             "Eastern FMCG Distributors", "Ganga Trading House", "Magadh Foods Pvt"]
STAFF = ["Ramesh Kumar", "Sunita Devi", "Amit Verma", "Pooja Singh"]

COA = [
    ("Sales Income", "income", "SALES;PAYMENT"),
    ("Purchases (COGS)", "expense", "PURCHASE;STOCK;SUPPLIER"),
    ("Rent", "expense", "RENT"),
    ("Salaries", "expense", "SALARY"),
    ("Utilities", "expense", "ELECTRICITY;NBPDCL;BSNL;BROADBAND;WATER"),
    ("Transport & Fuel", "expense", "PETROL;DIESEL;OLA;DELIVERY;TRANSPORT"),
    ("GST Paid", "expense", "GST"),
    ("Bank Charges", "expense", "CHARGES;SMS;ATM"),
    ("Interest Income", "income", "INT CREDIT;INTEREST"),
    ("Owner Drawings", "expense", "SELF;TRANSFER"),
]


def daterange(start, end):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def main():
    rows = []  # (date, narration, debit, credit)
    # --- monthly recurring items -------------------------------------------
    for m in range(4, 11):  # Apr..Oct 2026
        y, mo = 2026, m
        rows.append((date(y, mo, 2), "NEFT/LANDLORD PATNA/RENT %s 2026" %
                     date(y, mo, 1).strftime("%b").upper(), 12000.0, 0.0))
        rows.append((date(y, mo, 5), "NEFT/SALARY - %s" % STAFF[0], 9500.0, 0.0))
        rows.append((date(y, mo, 5), "NEFT/SALARY - %s" % STAFF[1], 8500.0, 0.0))
        rows.append((date(y, mo, 10), "ELECTRICITY BILL - NBPDCL %s" %
                     date(y, mo, 1).strftime("%b").upper(),
                     round(rng.uniform(1800, 3400), 2), 0.0))
        rows.append((date(y, mo, 15), "UPI/BSNL/BROADBAND %s" %
                     date(y, mo, 1).strftime("%b").upper(),
                     round(rng.uniform(900, 1100), 2), 0.0))
        if m in (5, 8):  # quarterly-ish GST payments
            rows.append((date(y, mo, 20), "GST PMT - CGST/SGST %s 2026" %
                         date(y, mo, 1).strftime("%b").upper(),
                         round(rng.uniform(6000, 11000), 2), 0.0))
        rows.append((date(y, mo, 28), "INT CREDIT - SAVINGS A/C",
                     0.0, round(rng.uniform(250, 850), 2)))
    # --- trading activity ---------------------------------------------------
    for day in daterange(date(2026, 4, 1), date(2026, 10, 31)):
        if rng.random() < 0.55:  # a sale most days
            cust = rng.choice(CUSTOMERS)
            mode = rng.choice(["UPI", "NEFT", "IMPS"])
            amt = round(rng.uniform(8000, 38000), 2)
            rows.append((day, "%s/%d/%s/SALES" %
                         (mode, rng.randint(1000000000, 9999999999), cust.upper()),
                         0.0, amt))
        # Purchases on a regular restock cadence (~every 3 days) so monthly
        # totals stay realistic instead of clustering by chance.
        if day.day in (2, 5, 8, 11, 14, 17, 20, 23, 26, 29) and rng.random() < 0.85:
            sup = rng.choice(SUPPLIERS)
            mode = rng.choice(["NEFT", "UPI"])
            amt = round(rng.uniform(15000, 48000), 2)
            rows.append((day, "%s/%s/PURCHASE STOCK" % (mode, sup.upper()), amt, 0.0))
        if rng.random() < 0.12:
            rows.append((day, "PETROL - HP PETROL PUMP PATNA",
                         round(rng.uniform(800, 2600), 2), 0.0))
        if rng.random() < 0.08:
            rows.append((day, "UPI/OLA/DELIVERY %s" %
                         rng.choice(["GODOWN", "MARKET", "CUSTOMER"]),
                         round(rng.uniform(350, 1200), 2), 0.0))
        if rng.random() < 0.05:
            rows.append((day, "BANK CHARGES - SMS ALERT",
                         round(rng.uniform(118, 295), 2), 0.0))
    # --- injected defects ----------------------------------------------------
    dup = (date(2026, 6, 14), "UPI/9812345670/SHARMA KIRANA/SALES", 0.0, 21450.0)
    rows.append(dup)
    rows.append(dup)  # exact duplicate
    rows.append((date(2026, 7, 3), "UPI/9876543210/UNKNOWN", 4200.0, 0.0))
    rows.append((date(2026, 8, 19), "MISC CREDIT", 0.0, 1500.0))
    rows.append((date(2026, 9, 11), "NEFT/XXXX/XXXX", 0.0, 9800.0))
    rows.append((date(2026, 5, 22), "NEFT/EASTERN FMCG DISTRIBUTORS/PURCHASE STOCK",
                 78400.0, 0.0))  # outlier purchase
    rows.append((date(2026, 6, 30), "UPI/SELF/TRANSFER", 15000.0, 0.0))

    rows.sort(key=lambda r: (r[0], r[1], r[2], r[3]))

    # running balance; inject one stated-balance error on 2026-09-15-ish row
    balance = 85000.0
    out_rows = []
    error_done = False
    for (d, narr, debit, credit) in rows:
        balance = round(balance - debit + credit, 2)
        stated = balance
        if (not error_done) and d >= date(2026, 9, 15):
            stated = round(balance + 500.0, 2)  # stated balance wrong by +500
            error_done = True
        out_rows.append((d.isoformat(), narr, "%.2f" % debit, "%.2f" % credit,
                         "%.2f" % stated))

    with open(os.path.join(OUT, "bank_statement.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["date", "narration", "debit", "credit", "balance"])
        w.writerows(out_rows)

    with open(os.path.join(OUT, "chart_of_accounts.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["category", "kind", "keywords"])
        w.writerows(COA)

    print("bank_statement.csv: %d rows (opening balance 85000.00)" % len(out_rows))
    print("chart_of_accounts.csv: %d categories" % len(COA))
    print("closing balance (true): %.2f" % balance)


if __name__ == "__main__":
    main()
