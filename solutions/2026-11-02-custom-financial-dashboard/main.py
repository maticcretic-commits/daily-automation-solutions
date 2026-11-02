#!/usr/bin/env python3
"""Independent Python ground truth for the FinancePulse dashboard.

Recomputes every checked figure straight from data/*.csv plus the
Assumptions constant (duplicated here deliberately — the test also
asserts the workbook's Assumptions sheet carries the same value).
"""
import csv
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

OPEN_CASH = 125000


def compute():
    income = list(csv.DictReader(open(os.path.join(DATA, "income.csv"))))
    expenses = list(csv.DictReader(open(os.path.join(DATA, "expenses.csv"))))
    sales = list(csv.DictReader(open(os.path.join(DATA, "sales.csv"))))
    investments = list(csv.DictReader(open(os.path.join(DATA, "investments.csv"))))

    tot_income = sum(int(r["amount"]) for r in income)
    tot_expenses = sum(int(r["amount"]) for r in expenses)
    net = tot_income - tot_expenses
    cash_close = OPEN_CASH + net
    savings = net / tot_income if tot_income else 0

    sales_rev = sum(int(r["qty"]) * int(r["unit_price"]) for r in sales)
    sales_cost = sum(int(r["qty"]) * int(r["cost_per_unit"]) for r in sales)
    gross = sales_rev - sales_cost

    inv_invested = sum(float(r["units"]) * float(r["buy_price"]) for r in investments)
    inv_value = sum(float(r["units"]) * float(r["current_price"]) for r in investments)
    inv_gain = inv_value - inv_invested
    inv_gain_pct = inv_gain / inv_invested if inv_invested else 0

    monthly = {}
    for m in range(1, 11):
        key = 202600 + m
        inc = sum(int(r["amount"]) for r in income if int(r["date"][5:7]) == m)
        exp = sum(int(r["amount"]) for r in expenses if int(r["date"][5:7]) == m)
        rev = sum(int(r["qty"]) * int(r["unit_price"]) for r in sales
                  if int(r["date"][5:7]) == m)
        cost = sum(int(r["qty"]) * int(r["cost_per_unit"]) for r in sales
                   if int(r["date"][5:7]) == m)
        monthly[key] = {"income": inc, "expenses": exp, "net": inc - exp,
                        "sales_rev": rev, "sales_cost": cost, "gross": rev - cost}

    exp_by_cat = defaultdict(int)
    for r in expenses:
        exp_by_cat[r["category"]] += int(r["amount"])

    return {
        "income_count": len(income), "expense_count": len(expenses),
        "sales_count": len(sales),
        "tot_income": tot_income, "tot_expenses": tot_expenses, "net": net,
        "cash_close": cash_close, "savings": savings,
        "sales_rev": sales_rev, "sales_cost": sales_cost, "gross": gross,
        "inv_invested": inv_invested, "inv_value": inv_value,
        "inv_gain": inv_gain, "inv_gain_pct": inv_gain_pct,
        "monthly": monthly, "exp_by_cat": dict(exp_by_cat),
    }


if __name__ == "__main__":
    t = compute()
    print(t)
