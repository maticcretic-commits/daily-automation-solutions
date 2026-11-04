#!/usr/bin/env python3
"""Generate deterministic double-entry journal lines for MonthlyClosePack.

Fictional trader: Deshmukh Trading Co., Pune — FMCG wholesaler, Apr-Oct
2026. Every entry balances (total debits == total credits per entry
ID); the generator tracks running signed balances so GST payments never
exceed the GST Payable balance and cash never goes negative in this
seed. All figures invented for this demo; no real client data. Seed
fixed so re-runs are identical. Lines are sorted by date, then entry.
"""
import csv
import datetime
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
SEED = 20261104
GST_RATE = 0.18

# Signed balances: debit-normal accounts positive, credit-normal negative
# is avoided here — balances track the natural amount owed/held, with a
# separate sign map used only for the cash/GST sanity checks.
OPENING = {
    "Cash": 450000, "Accounts Receivable": 180000, "Inventory": 320000,
    "Accounts Payable": 210000, "GST Payable": 45000,
    "Short-term Loan": 300000,
}


def round100(x):
    return int(round(x / 100.0) * 100)


def main():
    rng = random.Random(SEED)
    lines = []
    seq = 0
    bal = dict(OPENING)

    def entry(date, legs, narration):
        """legs: list of (account, debit, credit, category)."""
        nonlocal seq
        seq += 1
        eid = f"JE-{seq:04d}"
        assert sum(d for _, d, _, _ in legs) == sum(c for _, _, c, _ in legs), \
            f"unbalanced entry {eid}: {legs}"
        for account, debit, credit, category in legs:
            lines.append({
                "entry_id": eid, "date": date.isoformat(),
                "account": account, "debit": debit, "credit": credit,
                "category": category, "narration": narration,
            })
        return eid

    for month in range(4, 11):
        def day():
            return datetime.date(2026, month, rng.randint(1, 28))

        # --- Sales: credit and cash, GST at 18% on the net sale value. ---
        for _ in range(14):
            sale = round100(rng.randint(40000, 160000))
            cogs = round100(sale * rng.uniform(0.58, 0.64))
            gst = round100(sale * GST_RATE)
            entry(day(), [
                ("Accounts Receivable", sale + gst, 0, "BalanceSheet"),
                ("Sales Revenue", 0, sale, "Revenue"),
                ("GST Payable", 0, gst, "BalanceSheet"),
                ("Cost of Goods Sold", cogs, 0, "COGS"),
                ("Inventory", 0, cogs, "BalanceSheet"),
            ], "Credit sale incl. GST; COGS at cost")
            bal["Accounts Receivable"] += sale + gst
            bal["GST Payable"] += gst
            bal["Inventory"] -= cogs
        for _ in range(8):
            sale = round100(rng.randint(30000, 120000))
            cogs = round100(sale * rng.uniform(0.58, 0.64))
            gst = round100(sale * GST_RATE)
            entry(day(), [
                ("Cash", sale + gst, 0, "BalanceSheet"),
                ("Sales Revenue", 0, sale, "Revenue"),
                ("GST Payable", 0, gst, "BalanceSheet"),
                ("Cost of Goods Sold", cogs, 0, "COGS"),
                ("Inventory", 0, cogs, "BalanceSheet"),
            ], "Cash sale incl. GST; COGS at cost")
            bal["Cash"] += sale + gst
            bal["GST Payable"] += gst
            bal["Inventory"] -= cogs

        # --- Collections from customers (kept below the AR balance). ---
        for _ in range(18):
            amt = round100(rng.randint(40000, 140000))
            amt = min(amt, max(bal["Accounts Receivable"] - 50000, 0))
            if amt <= 0:
                continue
            entry(day(), [
                ("Cash", amt, 0, "BalanceSheet"),
                ("Accounts Receivable", 0, amt, "BalanceSheet"),
            ], "Collection from customer")
            bal["Cash"] += amt
            bal["Accounts Receivable"] -= amt

        # --- Inventory purchases, on credit and for cash. ---
        for _ in range(11):
            amt = round100(rng.randint(50000, 160000))
            entry(day(), [
                ("Inventory", amt, 0, "BalanceSheet"),
                ("Accounts Payable", 0, amt, "BalanceSheet"),
            ], "Inventory purchase on credit")
            bal["Inventory"] += amt
            bal["Accounts Payable"] += amt
        for _ in range(4):
            amt = round100(rng.randint(30000, 90000))
            amt = min(amt, max(bal["Cash"] - 100000, 0))
            if amt <= 0:
                continue
            entry(day(), [
                ("Inventory", amt, 0, "BalanceSheet"),
                ("Cash", 0, amt, "BalanceSheet"),
            ], "Inventory purchase for cash")
            bal["Cash"] -= amt
            bal["Inventory"] += amt

        # --- Pay suppliers (kept below the AP balance). ---
        for _ in range(10):
            amt = round100(rng.randint(40000, 130000))
            amt = min(amt, max(bal["Accounts Payable"] - 50000, 0))
            if amt <= 0:
                continue
            entry(day(), [
                ("Accounts Payable", amt, 0, "BalanceSheet"),
                ("Cash", 0, amt, "BalanceSheet"),
            ], "Payment to supplier")
            bal["Accounts Payable"] -= amt
            bal["Cash"] -= amt

        # --- Fixed operating costs. ---
        entry(day(), [("Rent Expense", 35000, 0, "Opex"),
                      ("Cash", 0, 35000, "BalanceSheet")], "Monthly rent")
        bal["Cash"] -= 35000
        salary = round100(rng.randint(120000, 150000))
        entry(day(), [
            ("Salaries Expense", salary, 0, "Opex"),
            ("Cash", 0, salary, "BalanceSheet")], "Monthly salaries")
        bal["Cash"] -= salary
        for _ in range(2):
            amt = round100(rng.randint(8000, 18000))
            entry(day(), [("Utilities Expense", amt, 0, "Opex"),
                          ("Cash", 0, amt, "BalanceSheet")], "Utilities bill")
            bal["Cash"] -= amt
        for _ in range(2):
            amt = round100(rng.randint(10000, 30000))
            entry(day(), [("Marketing Expense", amt, 0, "Opex"),
                          ("Cash", 0, amt, "BalanceSheet")], "Marketing spend")
            bal["Cash"] -= amt
        for _ in range(2):
            amt = round100(rng.randint(1000, 4000))
            entry(day(), [("Bank Charges", amt, 0, "Opex"),
                          ("Cash", 0, amt, "BalanceSheet")], "Bank charges")
            bal["Cash"] -= amt

        # --- GST payment: a share of the outstanding GST balance. ---
        if bal["GST Payable"] > 20000:
            amt = round100(bal["GST Payable"] * rng.uniform(0.5, 0.8))
            amt = min(amt, bal["Cash"] - 100000)
            if amt > 0:
                entry(day(), [("GST Payable", amt, 0, "BalanceSheet"),
                              ("Cash", 0, amt, "BalanceSheet")],
                      "GST paid for the month")
                bal["GST Payable"] -= amt
                bal["Cash"] -= amt

        # --- Depreciation, drawings, occasional other income. ---
        entry(day(), [("Depreciation Expense", 12500, 0, "Opex"),
                      ("Accumulated Depreciation", 0, 12500, "BalanceSheet")],
              "Monthly depreciation")
        entry(day(), [("Drawings", 20000, 0, "BalanceSheet"),
                      ("Cash", 0, 20000, "BalanceSheet")], "Owner drawings")
        bal["Cash"] -= 20000
        if rng.random() < 0.7:
            amt = round100(rng.randint(3000, 8000))
            entry(day(), [("Cash", amt, 0, "BalanceSheet"),
                          ("Other Income", 0, amt, "Revenue")],
                  "Interest / other income")
            bal["Cash"] += amt

        # --- One-offs. ---
        if month == 4:
            entry(datetime.date(2026, 4, 2), [
                ("Cash", 200000, 0, "BalanceSheet"),
                ("Owner's Capital", 0, 200000, "BalanceSheet"),
            ], "Owner capital injection")
            bal["Cash"] += 200000
        if month == 5:
            entry(datetime.date(2026, 5, 5), [
                ("Cash", 150000, 0, "BalanceSheet"),
                ("Short-term Loan", 0, 150000, "BalanceSheet"),
            ], "Short-term loan drawdown")
            bal["Cash"] += 150000
            bal["Short-term Loan"] += 150000
        if month == 6:
            entry(datetime.date(2026, 6, 10), [
                ("Equipment", 180000, 0, "BalanceSheet"),
                ("Cash", 0, 180000, "BalanceSheet"),
            ], "Equipment purchase")
            bal["Cash"] -= 180000
        if month in (9, 10):
            entry(day(), [("Short-term Loan", 50000, 0, "BalanceSheet"),
                          ("Cash", 0, 50000, "BalanceSheet")],
                  "Loan repayment")
            bal["Cash"] -= 50000
            bal["Short-term Loan"] -= 50000

    lines.sort(key=lambda r: (r["date"], r["entry_id"]))
    # Re-number entry IDs in date order so JE numbers follow the journal.
    renumber = {}
    for r in lines:
        if r["entry_id"] not in renumber:
            renumber[r["entry_id"]] = f"JE-{len(renumber) + 1:04d}"
        r["entry_id"] = renumber[r["entry_id"]]
    # Re-check balance per entry after renumbering (amounts untouched).
    from collections import defaultdict
    totals = defaultdict(lambda: [0, 0])
    for r in lines:
        totals[r["entry_id"]][0] += r["debit"]
        totals[r["entry_id"]][1] += r["credit"]
    assert all(d == c for d, c in totals.values()), "unbalanced entry found"
    assert bal["Cash"] > 0, f"cash went negative: {bal['Cash']}"
    assert bal["GST Payable"] >= 0, f"GST payable negative: {bal['GST Payable']}"
    assert bal["Inventory"] > 0, f"inventory negative: {bal['Inventory']}"

    os.makedirs(DATA, exist_ok=True)
    with open(os.path.join(DATA, "journal.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["entry_id", "date", "account",
                                          "debit", "credit", "category",
                                          "narration"])
        w.writeheader()
        w.writerows(lines)
    print(f"wrote {len(lines)} journal lines, {len(totals)} entries; "
          f"cash Rs {bal['Cash']:,}, AR Rs {bal['Accounts Receivable']:,}, "
          f"AP Rs {bal['Accounts Payable']:,}, GST payable "
          f"Rs {bal['GST Payable']:,}, inventory Rs {bal['Inventory']:,}")


if __name__ == "__main__":
    main()
