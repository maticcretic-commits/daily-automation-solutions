#!/usr/bin/env python3
"""Generate the fictional due-diligence data CSVs for DueDiligence Review Pack.

Target company (FICTIONAL — invented for the demo, not a real company):
Sutlej Agro Foods Pvt. Ltd., Ludhiana, Punjab — food processing,
FY24-FY26 (year ending 31 March). All figures in Rs lakh.

Design intent: the numbers must tell a story a reviewer would actually
flag — revenue grows 16% CAGR, but receivable days climb every year,
quick ratio stays under 1.0x, interest cover never reaches 2.0x,
contingent liabilities run to ~25% of net worth, a 0% director loan
sits repayable on demand, and a 45-day-old unreconciled cheque sits in
the bank reconciliation. All 9 red flags trigger by design, risk
score 31 (HIGH).

Bank reconciliation: monthly book balances come from receipts/payments
(opening cash = FY25 cash & bank, 165). Bank balances = book balance
minus the open-item difference at each month end, computed from the
item ledger (issue date, cleared date, signed effect on book - bank).
At 31 Mar 2026 the open items net to +1.55 (book 175, bank 173.45).
"""
import csv
import datetime as dt
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)

SEED = 20261119
COMPANY = "Sutlej Agro Foods Pvt. Ltd., Ludhiana"
AS_OF = dt.date(2026, 3, 31)

# ---------------- Financials (Rs lakh) ----------------
# P&L and balance-sheet lines; derived lines are formulas in the workbook.
FINANCIALS = [
    ("Revenue", 1840, 2150, 2480),
    ("COGS", 1196, 1398, 1636),
    ("Operating expenses", 420, 480, 545),
    ("Depreciation", 68, 74, 82),
    ("Interest", 92, 104, 118),
    ("Tax", 16, 24, 25),
    ("Share capital", 200, 200, 200),
    ("Reserves", 180, 250, 324),
    ("Long-term borrowings", 620, 580, 453),
    ("Short-term borrowings", 240, 280, 497),
    ("Trade payables", 310, 365, 430),
    ("Other current liabilities", 90, 110, 125),
    ("Fixed assets (net)", 720, 760, 820),
    ("Inventory", 380, 440, 510),
    ("Trade receivables", 290, 360, 445),
    ("Cash & bank", 150, 165, 175),
    ("Other current assets", 100, 60, 79),
]

# ---------------- Bank reconciliation ----------------
MONTHS = [(2025, 4), (2025, 5), (2025, 6), (2025, 7), (2025, 8),
          (2025, 9), (2025, 10), (2025, 11), (2025, 12), (2026, 1),
          (2026, 2), (2026, 3)]
RECEIPTS = [180, 195, 210, 185, 200, 215, 190, 205, 220, 195, 210, 225]
PAYMENTS = [178, 192, 208, 186, 198, 212, 192, 203, 218, 196, 208, 229]
OPENING_CASH = 165  # FY25 cash & bank; book must close at FY26 cash, 175

# Item ledger: (item, issue date, amount, signed effect on book-bank,
#               cleared date or None)
ITEMS = [
    ("Cheque issued - supplier payment", dt.date(2026, 2, 14), 4.20,
     -4.20, None),
    ("Cheque deposited - customer, not yet cleared",
     dt.date(2026, 3, 28), 8.40, 8.40, None),
    ("Bank charges debited, not in books", dt.date(2026, 3, 28), 0.15,
     0.15, None),
    ("UPI collection, not in books", dt.date(2026, 3, 19), 2.80, -2.80,
     None),
    ("Cheque issued - rent", dt.date(2026, 1, 10), 6.50, -6.50,
     dt.date(2026, 2, 5)),
    ("NEFT receipt, entered in books late", dt.date(2025, 12, 22), 3.10,
     -3.10, dt.date(2026, 1, 5)),
]


def month_end(y, m):
    if m == 12:
        return dt.date(y, 12, 31)
    return dt.date(y, m + 1, 1) - dt.timedelta(days=1)


def open_diff_at(end):
    """Sum of signed effects of items issued by `end` and not cleared."""
    total = 0.0
    for _, issue, _, effect, cleared in ITEMS:
        if issue <= end and (cleared is None or cleared > end):
            total += effect
    return round(total, 2)


# ---------------- Loans (Rs lakh) ----------------
LOANS = [
    # facility, lender, sanctioned, outstanding, rate, emi, overdue,
    # security, lt portion, current portion
    ("Term loan", "State Bank of India", 500, 420, 0.105, 9.80, 0,
     "Plant & machinery", 380, 40),
    ("Cash credit", "Canara Bank", 300, 285, 0.11, None, 0,
     "Stock & receivables", 0, 285),
    ("Vehicle loan", "HDFC Bank", 45, 18, 0.0925, 1.10, 0,
     "Hypothecation of vehicles", 6, 12),
    ("Unsecured loan", "Tata Capital (NBFC)", 150, 107, 0.15, 4.60, 0,
     "Unsecured", 67, 40),
    ("Director loan (related party)", "Promoter", 120, 120, 0.0, None, 0,
     "Unsecured, repayable on demand", 0, 120),
]

CONTINGENT = [
    ("GST demand under appeal", 42),
    ("Income-tax demand under appeal", 28),
    ("Bank guarantees issued", 60),
]

# ---------------- Red flags ----------------
# severity weights: High 5, Medium 3, Low 1. All 9 trigger by design.
REDFLAGS = [
    ("F1", "Working capital",
     "Receivable days rising three years running (57.5 -> 61.1 -> 65.5)",
     "High", 5),
    ("F2", "Contingent liabilities",
     "Contingent liabilities 24.8% of net worth (threshold 20%)",
     "High", 5),
    ("F3", "Related parties",
     "Director loan Rs 120 lakh at 0% interest, repayable on demand",
     "High", 5),
    ("F4", "Solvency",
     "Interest cover below 2.0x in all three years (1.70 / 1.90 / 1.84)",
     "Medium", 3),
    ("F5", "Liquidity",
     "Quick ratio below 1.0x in all three years (0.84 / 0.77 / 0.66)",
     "Medium", 3),
    ("F6", "Bank reconciliation",
     "Unreconciled cheque 45 days old at the as-of date (threshold 30)",
     "Medium", 3),
    ("F7", "Tax",
     "GST demand of Rs 42 lakh under appeal, unresolved",
     "Medium", 3),
    ("F8", "Cash conversion",
     "Receivables grew 53% in two years against revenue growth of 35% "
     "(cash conversion weak)",
     "Medium", 3),
    ("F9", "Audit",
     "Auditor emphasis of matter on the GST demand (FY26 audit report)",
     "Low", 1),
]


def main():
    # financials.csv
    with open(DATA / "financials.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["item", "fy24", "fy25", "fy26"])
        w.writerows(FINANCIALS)

    # bankrec.csv — book balances from receipts/payments; bank balances
    # from the item ledger.
    book = OPENING_CASH
    bank_rows = []
    for (y, m), rec, pay in zip(MONTHS, RECEIPTS, PAYMENTS):
        book = round(book + rec - pay, 2)
        end = month_end(y, m)
        diff = open_diff_at(end)
        bank = round(book - diff, 2)
        bank_rows.append({
            "month": f"{y}-{m:02d}",
            "receipts": rec,
            "payments": pay,
            "book_balance": book,
            "bank_balance": bank,
            "diff": diff,
        })
    with open(DATA / "bankrec.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["month", "receipts", "payments",
                                          "book_balance", "bank_balance",
                                          "diff"])
        w.writeheader()
        w.writerows(bank_rows)

    # bank_items.csv — the open items shown on the sheet are the four
    # with cleared_date empty; the ledger rows feed the test.
    with open(DATA / "bank_items.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["item", "date", "amount", "effect", "cleared"])
        for item, issue, amount, effect, cleared in ITEMS:
            w.writerow([item, issue.isoformat(), f"{amount:.2f}",
                        f"{effect:.2f}",
                        cleared.isoformat() if cleared else ""])

    # loans.csv
    with open(DATA / "loans.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["facility", "lender", "sanctioned", "outstanding",
                    "rate", "emi", "overdue", "security", "lt_portion",
                    "current_portion"])
        w.writerows(LOANS)

    # contingent.csv
    with open(DATA / "contingent.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["item", "amount"])
        w.writerows(CONTINGENT)

    # redflags.csv
    with open(DATA / "redflags.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["id", "area", "flag", "severity", "weight"])
        w.writerows(REDFLAGS)

    # meta.csv
    with open(DATA / "meta.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["key", "value"])
        w.writerow(["company", COMPANY])
        w.writerow(["as_of", AS_OF.isoformat()])
        w.writerow(["seed", SEED])

    # Sanity checks (fail loudly before the workbook is built).
    fin = {r[0]: r[1:] for r in FINANCIALS}
    for i, yr in enumerate(["FY24", "FY25", "FY26"]):
        nw = fin["Share capital"][i] + fin["Reserves"][i]
        le = (nw + fin["Long-term borrowings"][i]
              + fin["Short-term borrowings"][i] + fin["Trade payables"][i]
              + fin["Other current liabilities"][i])
        assets = (fin["Fixed assets (net)"][i] + fin["Inventory"][i]
                  + fin["Trade receivables"][i] + fin["Cash & bank"][i]
                  + fin["Other current assets"][i])
        assert le == assets, f"{yr} balance sheet off by {assets - le}"
    assert bank_rows[-1]["book_balance"] == fin["Cash & bank"][2], \
        "book close must equal FY26 cash & bank"
    assert bank_rows[-1]["diff"] == 1.55, \
        f"final open-item diff {bank_rows[-1]['diff']}, expected 1.55"
    loan_total = sum(r[3] for r in LOANS)
    assert loan_total == (fin["Long-term borrowings"][2]
                          + fin["Short-term borrowings"][2]), \
        f"loan total {loan_total} != FY26 borrowings"
    assert sum(r[8] for r in LOANS) == fin["Long-term borrowings"][2]
    assert sum(r[9] for r in LOANS) == fin["Short-term borrowings"][2]
    assert sum(a for _, a in CONTINGENT) == 130
    assert sum(r[4] for r in REDFLAGS) == 31, "risk score must be 31"
    print(f"wrote {len(FINANCIALS)} financial lines, "
          f"{len(bank_rows)} bank months, {len(ITEMS)} ledger items, "
          f"{len(LOANS)} loans, {len(REDFLAGS)} red flags")
    print(f"book closes {bank_rows[-1]['book_balance']}, "
          f"bank {bank_rows[-1]['bank_balance']}, "
          f"final diff {bank_rows[-1]['diff']}")
    print(f"loan total {loan_total}, contingent "
          f"{sum(a for _, a in CONTINGENT)}, risk score 31")


if __name__ == "__main__":
    main()
