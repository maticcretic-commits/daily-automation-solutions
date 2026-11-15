#!/usr/bin/env python3
"""Generate fictional bank-statement CSVs for the Consolidation Pack demo.

Business: Sharma Traders, Patna (wholesale FMCG trader), Jan-Oct 2026.
Three accounts, deterministic (seed 20261115):
  - Canara Bank, current a/c ****4821, opening Rs 4,25,000
  - SBI, savings a/c ****7734, opening Rs 1,80,000
  - HDFC Bank, current a/c ****9156, opening Rs 6,10,000

Each CSV has Date, Description, Debit, Credit, Balance (running balance
per source, exact to the paise). Three own-account transfers appear in
two statements each (debit in the source, credit in the destination).
Debits are capped so no account ever goes negative.
"""
import csv
import random
from datetime import date
from pathlib import Path

SEED = 20261115
BASE = Path(__file__).resolve().parent
OUT = BASE / "data" / "statements"

CUSTOMERS = ["Motihari Distributors", "Gaya Traders", "Muzaffarpur Retail Mart",
             "Darbhanga Wholesale", "Bhagalpur General Store", "Siwan Traders",
             "Chapra Kirana House", "Hajipur Super Mart"]
SUPPLIERS = ["Patna Agro Foods Pvt Ltd", "Bihar Rice Mills", "Shree Edible Oils",
             "Ganga Spices & Co", "North Bihar Pulses Traders", "Magadh Sugar Works"]
STAFF = ["R. Prasad", "S. Verma", "A. Jha", "K. Singh"]

ACCOUNTS = [
    {"key": "canara", "bank": "Canara Bank", "type": "Current",
     "acct": "****4821", "opening": 425000.0, "n_txn": 58},
    {"key": "sbi", "bank": "State Bank of India", "type": "Savings",
     "acct": "****7734", "opening": 180000.0, "n_txn": 46},
    {"key": "hdfc", "bank": "HDFC Bank", "type": "Current",
     "acct": "****9156", "opening": 610000.0, "n_txn": 54},
]

# (from_key, to_key, date, amount) — own-account transfers, both legs.
TRANSFERS = [
    ("canara", "sbi", date(2026, 3, 10), 150000.0),
    ("sbi", "hdfc", date(2026, 6, 15), 80000.0),
    ("hdfc", "canara", date(2026, 9, 5), 200000.0),
]


def money(rng, lo, hi):
    return round(rng.uniform(lo, hi), 2)


def random_row(rng, acct):
    """One random transaction (date, description, debit, credit)."""
    month = rng.randint(1, 10)
    day = rng.randint(1, 28)
    d = date(2026, month, day)
    kind = rng.random()
    if kind < 0.42:  # sales receipt
        return (d, f"UPI/NEFT from {rng.choice(CUSTOMERS)}",
                0.0, money(rng, 12000, 185000))
    if kind < 0.66:  # supplier payment
        return (d, f"NEFT to {rng.choice(SUPPLIERS)}",
                money(rng, 20000, 240000), 0.0)
    if kind < 0.72:  # rent (1st of month, fixed amount)
        return (date(2026, month, 1), "Godown rent - Patna City",
                45000.0, 0.0)
    if kind < 0.78:  # salaries (5th of month)
        return (date(2026, month, 5), f"Salary - {rng.choice(STAFF)}",
                money(rng, 18000, 42000), 0.0)
    if kind < 0.83:  # GST
        return (date(2026, month, 20), "GST payment - GSTR-3B",
                money(rng, 15000, 95000), 0.0)
    if kind < 0.88:  # utilities
        return (d, rng.choice(
            ["Electricity bill - NBPDCL", "Mobile/internet - Jio",
             "Fuel - HP Petrol Pump", "Packing material - Patna Packaging"]),
            money(rng, 1500, 18000), 0.0)
    if kind < 0.93:  # bank charges / interest
        if acct["type"] == "Savings" and rng.random() < 0.5:
            return (date(2026, month, 28),
                    "Interest credited - savings a/c",
                    0.0, money(rng, 400, 1400))
        return (d, rng.choice(
            ["SMS charges", "Cheque book charges", "NEFT charges",
             "Cash handling charges"]),
            money(rng, 15, 450), 0.0)
    # misc receipt
    return (d, rng.choice(
        ["Cash deposit - counter sales", "Refund from supplier",
         "Scrap sale - packing material"]),
        money(rng, 2000, 25000), 0.0)


def build_account(rng, acct, transfers_out, transfers_in):
    """Return list of (date, description, debit, credit, balance)."""
    rows = []
    # Transfer legs first (fixed dates/amounts).
    for (d, amt) in transfers_out:
        rows.append((d, "Own account transfer out", amt, 0.0))
    for (d, amt) in transfers_in:
        rows.append((d, "Own account transfer in", 0.0, amt))
    # Random rows, skipping any exact repeat of an earlier row for this
    # account (fixed-amount rows like rent would otherwise collide).
    seen = {(r[0], r[1], r[2], r[3]) for r in rows}
    while len(rows) < acct["n_txn"]:
        cand = random_row(rng, acct)
        key = (cand[0], cand[1], cand[2], cand[3])
        if key in seen:
            continue
        seen.add(key)
        rows.append(cand)
    rows.sort(key=lambda r: (r[0], r[1]))
    # Walk in date order. Random debits are capped so the balance never
    # drops below Rs 5,000 PLUS every transfer-out still to come — so
    # transfer legs are never capped and stay exact to the rupee.
    bal = acct["opening"]
    fixed = []
    for (d, desc, dr, cr) in rows:
        if desc == "Own account transfer out":
            bal = round(bal + cr - dr, 2)
            fixed.append((d, desc, dr, cr, bal))
            continue
        reserve = sum(amt for (td, amt) in transfers_out if td > d)
        if dr > 0 and bal + cr - dr < 5000 + reserve:
            dr = round(max(0.0, bal + cr - 5000 - reserve), 2)
        bal = round(bal + cr - dr, 2)
        fixed.append((d, desc, dr, cr, bal))
    return fixed


def main():
    rng = random.Random(SEED)
    OUT.mkdir(parents=True, exist_ok=True)
    out_map = {a["key"]: a for a in ACCOUNTS}
    for acct in ACCOUNTS:
        t_out = [(d, amt) for (f, t, d, amt) in TRANSFERS if f == acct["key"]]
        t_in = [(d, amt) for (f, t, d, amt) in TRANSFERS if t == acct["key"]]
        rows = build_account(rng, acct, t_out, t_in)
        path = OUT / f"{acct['key']}.csv"
        with open(path, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Date", "Description", "Debit", "Credit", "Balance"])
            for (d, desc, dr, cr, bal) in rows:
                w.writerow([d.isoformat(), desc,
                            f"{dr:.2f}" if dr else "",
                            f"{cr:.2f}" if cr else "", f"{bal:.2f}"])
        credits = sum(r[3] for r in rows)
        debits = sum(r[2] for r in rows)
        print(f"{acct['key']:7s} {len(rows):3d} rows  credits "
              f"Rs {credits:,.2f}  debits Rs {debits:,.2f}  "
              f"closing Rs {rows[-1][4]:,.2f}")


if __name__ == "__main__":
    main()
