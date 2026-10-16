#!/usr/bin/env python3
"""generate.py — Synthesise QuickBooks-style double-entry accounting data.

Fictional business: Agarwal Traders, a wholesale FMCG distributor (Patna, India).
Period: April 2025 .. September 2026 (18 months).
Output: data/*.csv — chart of accounts, customers, vendors, journal lines.

Every journal entry balances (total debits == total credits). All randomness is
seeded (seed=42) so the dataset is fully reproducible.
"""
import csv
import random
from datetime import date

random.seed(42)
OUT = "data"

ACCOUNTS = [
    ("1000", "Cash & Bank", "Asset"),
    ("1100", "Accounts Receivable", "Asset"),
    ("1200", "Inventory", "Asset"),
    ("1300", "GST Input Credit", "Asset"),
    ("1500", "Equipment", "Asset"),
    ("1510", "Accumulated Depreciation", "Asset"),
    ("2000", "Accounts Payable", "Liability"),
    ("2100", "GST Payable", "Liability"),
    ("2200", "Term Loan", "Liability"),
    ("3000", "Owner's Capital", "Equity"),
    ("3100", "Retained Earnings", "Equity"),
    ("4000", "Sales Revenue", "Income"),
    ("5000", "Cost of Goods Sold", "Expense"),
    ("5100", "Rent Expense", "Expense"),
    ("5200", "Salaries Expense", "Expense"),
    ("5300", "Utilities Expense", "Expense"),
    ("5400", "Transport & Logistics", "Expense"),
    ("5500", "Marketing Expense", "Expense"),
    ("5600", "Interest Expense", "Expense"),
    ("5700", "Depreciation Expense", "Expense"),
    ("5800", "Miscellaneous Expense", "Expense"),
]

CUSTOMERS = [
    "Sharma General Store", "Verma Kirana Mart", "Gupta Super Bazaar",
    "Patna Daily Needs", "Kumari Provision Store", "Singh Brothers Retail",
    "New City Mart", "Rajdhani Traders", "Maa Vaishno Stores", "Annapurna Retail",
]
VENDORS = [
    "Hindustan Foods Ltd", "Bihar Agro Supplies", "National FMCG Distributors",
    "Eastern Packaging Co", "Ganga Logistics Pvt Ltd",
]

ACCT = {no: name for no, name, _ in ACCOUNTS}

# ---------------------------------------------------------------- journals
_jno = 0
journals = []  # (date, jno, memo, [(acct, debit, credit, party)])


def add_entry(dt, memo, lines):
    """lines: list of (acct_no, debit, credit, party)."""
    global _jno
    _jno += 1
    jno = f"J-{_jno:04d}"
    dr = round(sum(l[1] for l in lines), 2)
    cr = round(sum(l[2] for l in lines), 2)
    assert abs(dr - cr) < 0.01, f"unbalanced {jno}: {dr} vs {cr}"
    journals.append((dt, jno, memo, lines))


def months():
    m, y = 4, 2025
    while (y, m) <= (2026, 9):
        yield y, m
        m += 1
        if m > 12:
            m, y = 1, y + 1


def day(y, m, d):
    import calendar
    return date(y, m, min(d, calendar.monthrange(y, m)[1]))


def inr(x):
    return round(float(x), 2)


# ------------------------------------------------------- opening balances
# Dr: Cash 8L, AR 12L, Inventory 15L, Equipment 10L = 45L
# Cr: Acc Depr 3L, AP 9L, Capital 25L, Retained Earnings 8L = 45L
add_entry(date(2025, 4, 1), "Opening balances 01-Apr-2025", [
    ("1000", 800000, 0, ""), ("1100", 1200000, 0, ""),
    ("1200", 1500000, 0, ""), ("1500", 1000000, 0, ""),
    ("1510", 0, 300000, ""), ("2000", 0, 900000, ""),
    ("3000", 0, 2500000, ""), ("3100", 0, 800000, ""),
])

# Term loan drawn April 2025: Dr Bank 20L / Cr Term Loan 20L
add_entry(date(2025, 4, 3), "Term loan disbursed (HDFC)", [
    ("1000", 2000000, 0, ""), ("2200", 0, 2000000, ""),
])

# ---------------------------------------------------------------- monthly
mi = 0
for y, m in months():
    mi += 1
    # --- sales: 7 invoices, revenue grows ~1.5%/mo + seasonality + noise
    base = 2800000 * (1.015 ** mi) * (1 + 0.12 * (1 if m in (10, 11, 3) else 0))
    month_rev = inr(base * random.uniform(0.94, 1.06))
    inv_total = 0.0
    for i in range(7):
        amt = inr(month_rev * random.uniform(0.10, 0.20))
        inv_total += amt
        cust = random.choice(CUSTOMERS)
        gst = inr(amt * 0.18)
        add_entry(day(y, m, random.randint(2, 24)),
                  f"Sales invoice INV-{y}{m:02d}-{i+1}", [
                      ("1100", inr(amt + gst), 0, cust),
                      ("4000", 0, amt, cust),
                      ("2100", 0, gst, "GST"),
                  ])
    # --- COGS at ~70% of revenue: Dr COGS / Cr Inventory
    cogs = inr(inv_total * random.uniform(0.68, 0.72))
    add_entry(day(y, m, 27), "Cost of goods sold - monthly", [
        ("5000", cogs, 0, ""), ("1200", 0, cogs, ""),
    ])
    # --- collections: ~85% of month AR collected
    coll = inr(inv_total * 1.18 * random.uniform(0.80, 0.90))
    add_entry(day(y, m, 28), "Customer receipts", [
        ("1000", coll, 0, ""), ("1100", 0, coll, ""),
    ])
    # --- purchases: 5 vendor bills (inventory + GST input)
    purch = inr(cogs * random.uniform(0.95, 1.05))
    for i in range(5):
        p = inr(purch / 5 * random.uniform(0.8, 1.2))
        gst_in = inr(p * 0.18)
        v = random.choice(VENDORS)
        add_entry(day(y, m, random.randint(3, 22)),
                  f"Purchase bill PB-{y}{m:02d}-{i+1}", [
                      ("1200", p, 0, v),
                      ("1300", gst_in, 0, v),
                      ("2000", 0, inr(p + gst_in), v),
                  ])
    # --- vendor payments
    pay = inr(purch * 1.18 * random.uniform(0.85, 0.95))
    add_entry(day(y, m, 29), "Vendor payments", [
        ("2000", pay, 0, ""), ("1000", 0, pay, ""),
    ])
    # --- GST net settlement (payable minus input credit)
    gst_net = inr(inv_total * 0.18 - purch * 0.18)
    if gst_net > 0:
        add_entry(day(y, m, 30), "GST paid for month", [
            ("2100", gst_net, 0, "GST"), ("1000", 0, gst_net, "GST"),
        ])
    # --- operating expenses
    salary = inr(110000 * (1.08 if (y, m) >= (2026, 4) else 1.0) * random.uniform(0.98, 1.02))
    add_entry(day(y, m, 5), "Office rent", [("5100", 25000, 0, ""), ("1000", 0, 25000, "")])
    add_entry(day(y, m, 7), "Staff salaries", [("5200", salary, 0, ""), ("1000", 0, salary, "")])
    util = inr(random.uniform(7000, 9500))
    add_entry(day(y, m, 12), "Utilities (power/water/phone)",
              [("5300", util, 0, ""), ("1000", 0, util, "")])
    transp = inr(random.uniform(15000, 22000))
    add_entry(day(y, m, 18), "Transport & logistics",
              [("5400", transp, 0, ""), ("1000", 0, transp, "")])
    mkt = inr(random.uniform(12000, 20000))
    add_entry(day(y, m, 20), "Marketing & promotions",
              [("5500", mkt, 0, ""), ("1000", 0, mkt, "")])
    # --- loan: interest + principal
    add_entry(day(y, m, 10), "Term loan interest", [("5600", 9000, 0, "HDFC"), ("1000", 0, 9000, "HDFC")])
    add_entry(day(y, m, 10), "Term loan principal repayment",
              [("2200", 55000, 0, "HDFC"), ("1000", 0, 55000, "HDFC")])
    # --- depreciation
    add_entry(day(y, m, 30), "Depreciation - equipment",
              [("5700", 6000, 0, ""), ("1510", 0, 6000, "")])
    # --- miscellaneous (some months)
    if random.random() < 0.6:
        mx = inr(random.uniform(3000, 9000))
        add_entry(day(y, m, random.randint(8, 26)), "Miscellaneous expenses",
                  [("5800", mx, 0, ""), ("1000", 0, mx, "")])

# ------------------------------------------------------------------ write
with open(f"{OUT}/chart_of_accounts.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["acct_no", "acct_name", "acct_type"])
    w.writerows(ACCOUNTS)

with open(f"{OUT}/customers.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["customer"])
    w.writerows([[c] for c in CUSTOMERS])

with open(f"{OUT}/vendors.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["vendor"])
    w.writerows([[v] for v in VENDORS])

with open(f"{OUT}/journals.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["date", "journal_no", "line_no", "acct_no", "acct_name",
                "debit", "credit", "memo", "party"])
    ln = 0
    for dt, jno, memo, lines in journals:
        for li, (acct, dr, cr, party) in enumerate(lines, 1):
            ln += 1
            w.writerow([dt.isoformat(), jno, li, acct, ACCT[acct],
                        f"{dr:.2f}", f"{cr:.2f}", memo, party])

print(f"accounts={len(ACCOUNTS)} entries={len(journals)} lines={ln}")
