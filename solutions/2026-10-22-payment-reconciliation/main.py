"""Independent Python ground truth for ReconPulse.

Reads data/*.csv, performs the same matching the workbook formulas do
(exact amount+side key, |date diff| <= 3), and exposes expected KPIs for
tests/test_build.py to compare against the LibreOffice recalculation.
"""
import csv
import os
from datetime import date

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data")
OPENING = 4500000.00
ASOF = date(2026, 10, 21)


def load(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def key(amount, side):
    return (round(amount, 2), side)


def compute():
    bank = load("bank_statement.csv")
    book = load("book_ledger.csv")

    def amt(r):
        return float(r["Debit"]) + float(r["Credit"])

    def bside(r):
        return "R" if float(r["Credit"]) > 0 else "P"

    def kside(r):
        return "R" if float(r["Debit"]) > 0 else "P"

    book_by_key = {}
    for i, r in enumerate(book):
        book_by_key.setdefault(key(amt(r), kside(r)), []).append(i)

    matched = []          # (bank_idx, book_idx)
    unmatched_bank = []   # bank_idx
    claimed_book = set()
    for i, r in enumerate(bank):
        k = key(amt(r), bside(r))
        cands = book_by_key.get(k, [])
        hit = None
        for j in cands:
            if j in claimed_book:
                continue
            bd = date.fromisoformat(bank[i]["Date"])
            kd = date.fromisoformat(book[j]["Date"])
            if abs((bd - kd).days) <= 3:
                hit = j
                break
        if hit is None:
            unmatched_bank.append(i)
        else:
            matched.append((i, hit))
            claimed_book.add(hit)

    unmatched_book = [j for j in range(len(book)) if j not in claimed_book]

    bank_close = OPENING + sum(float(r["Credit"]) - float(r["Debit"]) for r in bank)
    book_close = OPENING + sum(float(r["Debit"]) - float(r["Credit"]) for r in book)

    def is_unpresented(j):
        return "unpresented" in book[j]["Narration"]

    def is_transit(j):
        return "in transit" in book[j]["Narration"]

    unpresented = sum(amt(book[j]) for j in unmatched_book if is_unpresented(j))
    transit = sum(amt(book[j]) for j in unmatched_book if is_transit(j))
    # Bank-side walk: strip bank-only items (not yet in books) to reach the
    # "fully recorded" balance.
    interest = sum(amt(bank[i]) for i in unmatched_bank
                   if "interest" in bank[i]["Description"].lower())
    unidentified = sum(amt(bank[i]) for i in unmatched_bank
                       if "unidentified" in bank[i]["Description"].lower())
    charges = sum(amt(bank[i]) for i in unmatched_bank
                  if "charge" in bank[i]["Description"].lower())
    adjusted_bank = bank_close - interest - unidentified + charges
    # Book-side walk: strip book-only items (not yet at bank).
    # Unpresented cheques are book CREDITS (recorded outflow) -> add back.
    # Transit deposits are book DEBITS (recorded inflow) -> subtract.
    adjusted_book = book_close + unpresented - transit
    recon_diff = round(adjusted_bank - adjusted_book, 2)

    oldest_days = 0
    for i in unmatched_bank:
        d = (ASOF - date.fromisoformat(bank[i]["Date"])).days
        oldest_days = max(oldest_days, d)
    for j in unmatched_book:
        d = (ASOF - date.fromisoformat(book[j]["Date"])).days
        oldest_days = max(oldest_days, d)

    total_receipts = sum(float(r["Credit"]) for r in bank)
    total_payments = sum(float(r["Debit"]) for r in bank)
    bank_charges = sum(amt(bank[i]) for i in unmatched_bank
                       if "charge" in bank[i]["Description"].lower())

    # daily EOD series 1..21 Oct
    daily = []
    for day in range(1, 22):
        d = date(2026, 10, day)
        bbal = OPENING + sum(float(r["Credit"]) - float(r["Debit"]) for r in bank
                             if date.fromisoformat(r["Date"]) <= d)
        kbal = OPENING + sum(float(r["Debit"]) - float(r["Credit"]) for r in book
                             if date.fromisoformat(r["Date"]) <= d)
        daily.append((d.isoformat(), round(bbal, 2), round(kbal, 2)))

    return {
        "bank_rows": len(bank),
        "book_rows": len(book),
        "matched": len(matched),
        "unmatched_bank": len(unmatched_bank),
        "unmatched_book": len(unmatched_book),
        "bank_close": round(bank_close, 2),
        "book_close": round(book_close, 2),
        "unpresented": round(unpresented, 2),
        "transit": round(transit, 2),
        "interest": round(interest, 2),
        "unidentified": round(unidentified, 2),
        "adjusted_bank": round(adjusted_bank, 2),
        "adjusted_book": round(adjusted_book, 2),
        "recon_diff": recon_diff,
        "oldest_days": oldest_days,
        "total_receipts": round(total_receipts, 2),
        "total_payments": round(total_payments, 2),
        "bank_charges": round(bank_charges, 2),
        "daily": daily,
        "unmatched_bank_desc": sorted(bank[i]["Description"] for i in unmatched_bank),
        "unmatched_book_narr": sorted(book[j]["Narration"] for j in unmatched_book),
    }


if __name__ == "__main__":
    import json
    print(json.dumps(compute(), indent=1, default=str))
