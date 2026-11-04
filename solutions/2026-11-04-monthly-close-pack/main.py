#!/usr/bin/env python3
"""Independent Python ground truth for the MonthlyClosePack workbook.

Recomputes every checked figure straight from data/journal.csv,
data/accounts.csv and data/budget.csv. Deliberately duplicates the
accounting definitions used by build_workbook.py (the test also
asserts the recalculated workbook matches these numbers, so a shared
bug would need to appear in both places with the same wrong value).
"""
import csv
import datetime
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
AS_OF = datetime.date(2026, 11, 4)
MONTHS = [(202604, "Apr 2026"), (202605, "May 2026"), (202606, "Jun 2026"),
          (202607, "Jul 2026"), (202608, "Aug 2026"), (202609, "Sep 2026"),
          (202610, "Oct 2026")]


def load_accounts():
    accts = list(csv.DictReader(open(os.path.join(DATA, "accounts.csv"))))
    for a in accts:
        a["opening_debit"] = int(a["opening_debit"])
        a["opening_credit"] = int(a["opening_credit"])
    return accts


def compute():
    accts = load_accounts()
    opening = {a["account"]: a["opening_debit"] - a["opening_credit"]
               for a in accts}  # signed: + debit balance, - credit balance
    lines = list(csv.DictReader(open(os.path.join(DATA, "journal.csv"))))
    for r in lines:
        r["debit"] = int(r["debit"])
        r["credit"] = int(r["credit"])
        d = datetime.date.fromisoformat(r["date"])
        r["month_key"] = d.year * 100 + d.month

    assert sum(r["debit"] for r in lines) == sum(r["credit"] for r in lines)
    per_entry = defaultdict(lambda: [0, 0])
    for r in lines:
        per_entry[r["entry_id"]][0] += r["debit"]
        per_entry[r["entry_id"]][1] += r["credit"]
    assert all(d == c for d, c in per_entry.values())

    # Signed movements per account: debit adds, credit subtracts.
    move = defaultdict(int)
    dr_move = defaultdict(int)
    cr_move = defaultdict(int)
    for r in lines:
        move[r["account"]] += r["debit"] - r["credit"]
        dr_move[r["account"]] += r["debit"]
        cr_move[r["account"]] += r["credit"]
    closing = {a["account"]: opening[a["account"]] + move[a["account"]]
               for a in accts}
    # Trial balance totals: debit balances vs credit balances.
    tb_dr = sum(v for v in closing.values() if v > 0)
    tb_cr = sum(-v for v in closing.values() if v < 0)
    # Opening trial balance must also balance.
    op_dr = sum(v for v in opening.values() if v > 0)
    op_cr = sum(-v for v in opening.values() if v < 0)

    by_type = defaultdict(int)
    for a in accts:
        by_type[a["type"]] += closing[a["account"]]
    revenue = -by_type["Revenue"]          # credit balances are negative
    cogs = by_type["COGS"]
    expenses = by_type["Expense"]
    gross = revenue - cogs
    net_income = gross - expenses

    # Balance sheet.
    assets = (closing["Cash"] + closing["Accounts Receivable"]
              + closing["Inventory"] + closing["Prepaid Expenses"]
              + closing["Equipment"] + closing["Accumulated Depreciation"])
    liabilities = -(closing["Accounts Payable"] + closing["GST Payable"]
                    + closing["Short-term Loan"])
    capital = -closing["Owner's Capital"]
    retained = -closing["Retained Earnings"] + net_income
    drawings = closing["Drawings"]
    equity = capital + retained - drawings
    bs_diff = assets - (liabilities + equity)

    # Cash flow (indirect): must equal the change in Cash.
    cash_change = closing["Cash"] - opening["Cash"]
    cf_operating = (net_income
                    + move["Accumulated Depreciation"] * -1  # credit movement
                    - (closing["Accounts Receivable"]
                       - opening["Accounts Receivable"])
                    - (closing["Inventory"] - opening["Inventory"])
                    - (closing["Prepaid Expenses"]
                       - opening["Prepaid Expenses"])
                    + (-(closing["Accounts Payable"]
                         - opening["Accounts Payable"]))
                    + (-(closing["GST Payable"] - opening["GST Payable"])))
    # move["Accumulated Depreciation"] is (debit - credit) = -credit movement,
    # so depreciation add-back = -move. The expression above uses that.
    cf_investing = -(closing["Equipment"] - opening["Equipment"])
    cf_financing = (-(closing["Short-term Loan"] - opening["Short-term Loan"])
                    + (-(closing["Owner's Capital"]
                         - opening["Owner's Capital"]))
                    - closing["Drawings"])
    cf_total = cf_operating + cf_investing + cf_financing

    # Monthly P&L from journal lines (category + month key).
    monthly = {}
    for key, label in MONTHS:
        sel = [r for r in lines if r["month_key"] == key]
        rev = sum(r["credit"] for r in sel if r["category"] == "Revenue")
        cg = sum(r["debit"] for r in sel if r["category"] == "COGS")
        opex = sum(r["debit"] for r in sel if r["category"] == "Opex")
        monthly[key] = {"revenue": rev, "cogs": cg, "gross": rev - cg,
                        "opex": opex, "net": rev - cg - opex,
                        "entries": len({r["entry_id"] for r in sel})}
    # Budget comparison.
    budget = {int(b["month_key"]): b
              for b in csv.DictReader(open(os.path.join(DATA, "budget.csv")))}
    budget_rows = {}
    for key, _ in MONTHS:
        b = budget[key]
        budget_rows[key] = {
            "sales_budget": int(b["sales_budget"]),
            "net_budget": int(b["net_budget"]),
            "sales_actual": monthly[key]["revenue"],
            "net_actual": monthly[key]["net"],
        }

    return {
        "lines": len(lines), "entries": len(per_entry),
        "journal_dr": sum(r["debit"] for r in lines),
        "opening_dr": op_dr, "opening_cr": op_cr,
        "tb_dr": tb_dr, "tb_cr": tb_cr,
        "closing": closing, "dr_move": dict(dr_move),
        "cr_move": dict(cr_move),
        "revenue": revenue, "cogs": cogs, "expenses": expenses,
        "gross": gross, "net_income": net_income,
        "assets": assets, "liabilities": liabilities, "equity": equity,
        "retained": retained, "bs_diff": bs_diff,
        "cash_change": cash_change,
        "cf_operating": cf_operating, "cf_investing": cf_investing,
        "cf_financing": cf_financing, "cf_total": cf_total,
        "monthly": monthly, "budget_rows": budget_rows,
        "depreciation": -move["Accumulated Depreciation"],
    }


if __name__ == "__main__":
    t = compute()
    print(f"{t['lines']} lines, {t['entries']} entries; "
          f"revenue Rs {t['revenue']:,}, COGS Rs {t['cogs']:,}, "
          f"net income Rs {t['net_income']:,}; assets Rs {t['assets']:,}, "
          f"L+E Rs {t['liabilities'] + t['equity']:,}, BS diff {t['bs_diff']}; "
          f"cash change Rs {t['cash_change']:,} = CF total "
          f"Rs {t['cf_total']:,}")
