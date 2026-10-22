"""Generate ReconPulse demo data: bank statement + book ledger for a fictional
wholesale distributor, with deliberate unmatched items on both sides.

Deterministic (seed 20261022). Outputs data/bank_statement.csv and
data/book_ledger.csv. No real persons or businesses.
"""
import csv
import os
import random

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data")
os.makedirs(DATA, exist_ok=True)

rng = random.Random(20261022)

OPENING_BANK = 45_00_000.00   # bank balance 30-Sep-2026
OPENING_BOOK = 45_00_000.00   # book bank-column balance 30-Sep-2026
ASOF = "2026-10-21"

CUSTOMERS = ["Sharma Retail Store", "Gupta General Store", "Verma Traders",
             "Singh Kirana", "Yadav Provision Store", "Mishra Mart",
             "Khan Departmental", "Patel Daily Needs", "Jain Super Bazaar",
             "Kumar & Sons"]
SUPPLIERS = ["Agrawal Foods Ltd", "Bansal Agro Products", "Chawla Distributors",
             "Dhillon FMCG", "Eagle Beverages", "FreshFarms Coop"]

bank_rows = []   # (date, desc, ref, debit, credit)
book_rows = []   # (date, narration, voucher, debit, credit)
used_keys = set()

# Reserve the deliberate unmatched amounts FIRST so matched items never
# share a (amount, side) key with them (keeps formula MATCH unambiguous).
for _amt, _side in [(236.00, "P"), (59.00, "P"), (1842.00, "R"), (12500.00, "R"),
                    (45000.00, "P"), (15000.00, "P"), (28000.00, "R")]:
    used_keys.add((_amt, _side))

def unique_amount(base, side):
    amt = round(base, 2)
    while (amt, side) in used_keys:
        amt = round(amt + 7.0, 2)
    used_keys.add((amt, side))
    return amt

def dstr(day):
    return f"2026-10-{day:02d}"

def add_matched_receipt(desc, ref_prefix, day, amount):
    """Money IN: bank credit / book debit. Bank clears 0-2 days after booking."""
    amount = unique_amount(amount, "R")
    lag = rng.choice([0, 0, 1, 1, 2])
    bday = min(day + lag, 21)
    book_rows.append((dstr(day), desc, f"{ref_prefix}{rng.randint(1000,9999)}",
                      amount, 0.0))
    bank_rows.append((dstr(bday), desc, f"UTR{rng.randint(100000,999999)}",
                      0.0, amount))

def add_matched_payment(desc, ref_prefix, day, amount):
    """Money OUT: bank debit / book credit."""
    amount = unique_amount(amount, "P")
    lag = rng.choice([0, 0, 1, 1, 2])
    bday = min(day + lag, 21)
    book_rows.append((dstr(day), desc, f"{ref_prefix}{rng.randint(1000,9999)}",
                      0.0, amount))
    bank_rows.append((dstr(bday), desc, f"UTR{rng.randint(100000,999999)}",
                      amount, 0.0))

# ---- matched receipts: customer collections ----
for i in range(42):
    c = rng.choice(CUSTOMERS)
    mode = rng.choice(["NEFT", "NEFT", "UPI", "RTGS", "cheque"])
    amt = rng.randint(8, 95) * 1000 + rng.choice([0, 50, 500])
    add_matched_receipt(f"{mode} - {c} (inv {rng.randint(1021,1188)})",
                        {"NEFT": "RV", "UPI": "RV", "RTGS": "RV", "cheque": "RV"}[mode],
                        rng.randint(1, 20), amt)
for i in range(30):
    amt = rng.randint(400, 9500)
    add_matched_receipt(f"UPI - retail counter sales {rng.randint(1,21)}-Oct",
                        "RV", rng.randint(1, 21), amt)

# ---- matched payments: suppliers, expenses, salaries, tax ----
for i in range(34):
    s = rng.choice(SUPPLIERS)
    amt = rng.randint(15, 180) * 1000 + rng.choice([0, 250])
    add_matched_payment(f"NEFT - {s} (PO {rng.randint(501,640)})", "PV",
                        rng.randint(1, 20), amt)
for i in range(12):
    desc = rng.choice(["Godown rent - Oct", "Electricity bill", "Transport - freight",
                       "Packing material", "Staff welfare", "Telephone/internet",
                       "Insurance premium", "Audit fee advance"])
    add_matched_payment(desc, "PV", rng.randint(1, 20),
                        rng.randint(2, 60) * 1000 + rng.choice([0, 500]))
for i in range(4):
    add_matched_payment(f"NEFT salary - staff batch {i+1} Oct", "PV",
                        rng.randint(3, 8), rng.randint(90, 140) * 1000)
add_matched_payment("GST payment - Sep 2026", "PV", 18, 86450)
add_matched_payment("TDS deposit Q2", "PV", 6, 23780)

# ---- deliberate UNMATCHED bank-only items ----
bank_rows.append((dstr(21), "Bank charges - Oct 2026", "CHG88231", 236.00, 0.0))
bank_rows.append((dstr(21), "SMS alert charges", "CHG88232", 59.00, 0.0))
bank_rows.append((dstr(21), "Interest credited", "INT1021", 0.0, 1842.00))
bank_rows.append((dstr(19), "UPI direct deposit - unidentified", "UPI99887766",
                  0.0, 12500.00))

# ---- deliberate UNMATCHED book-only items ----
book_rows.append((dstr(18), "Chq 452318 - Verma Enterprises (unpresented)", "PV7712",
                  0.0, 45000.00))
book_rows.append((dstr(20), "Chq 452320 - staff advance (unpresented)", "PV7740",
                  0.0, 15000.00))
book_rows.append((dstr(21), "Cash deposit in transit", "RV9901", 28000.00, 0.0))

bank_rows.sort(key=lambda r: (r[0], r[1]))
book_rows.sort(key=lambda r: (r[0], r[1]))

with open(os.path.join(DATA, "bank_statement.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Description", "Ref", "Debit", "Credit"])
    for r in bank_rows:
        w.writerow([r[0], r[1], r[2], f"{r[3]:.2f}", f"{r[4]:.2f}"])

with open(os.path.join(DATA, "book_ledger.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Narration", "Voucher", "Debit", "Credit"])
    for r in book_rows:
        w.writerow([r[0], r[1], r[2], f"{r[3]:.2f}", f"{r[4]:.2f}"])

print(f"bank rows: {len(bank_rows)}, book rows: {len(book_rows)}")
print(f"opening bank/book: {OPENING_BANK:,.2f} / {OPENING_BOOK:,.2f}, as-of {ASOF}")
