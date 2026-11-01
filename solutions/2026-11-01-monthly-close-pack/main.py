#!/usr/bin/env python3
"""Independent Python ground truth for the monthly close pack.

Recomputes every checked figure straight from data/*.csv plus the
Assumptions constants (duplicated here deliberately — the test also
asserts the workbook's Assumptions sheet carries the same values).
"""
import csv
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

OPEN_CASH, OPEN_AR, OPEN_EQUIP = 850000, 120000, 400000
OPEN_AP, OPEN_EQUITY = 60000, 1310000
DEPR_MONTHLY, CAPEX, DRAWINGS = 5000, 75000, 200000
AR_SHARE, AP_SHARE = 0.20, 0.10


def compute():
    txns = list(csv.DictReader(open(os.path.join(DATA, "transactions.csv"))))
    budgets = list(csv.DictReader(open(os.path.join(DATA, "budget.csv"))))
    monthly = {}
    for m in range(1, 11):
        rows = [t for t in txns if int(t["date"][5:7]) == m]
        rev = sum(int(t["amount"]) for t in rows if t["type"] == "Revenue")
        direct = sum(int(t["amount"]) for t in rows if t["type"] == "Direct")
        opex = sum(int(t["amount"]) for t in rows if t["type"] == "Operating")
        monthly[m] = {"revenue": rev, "direct": direct, "opex": opex,
                      "net": rev - direct - opex - DEPR_MONTHLY}
    revenue = sum(v["revenue"] for v in monthly.values())
    direct = sum(v["direct"] for v in monthly.values())
    opex = sum(v["opex"] for v in monthly.values())
    depr = DEPR_MONTHLY * 10
    net = revenue - direct - opex - depr
    oct_rev, oct_exp = monthly[10]["revenue"], monthly[10]["direct"] + monthly[10]["opex"]
    ar_close = AR_SHARE * oct_rev
    ap_close = AP_SHARE * oct_exp
    equip_close = OPEN_EQUIP + CAPEX - depr
    collections = revenue + OPEN_AR - ar_close
    payments = (direct + opex) + OPEN_AP - ap_close
    cash_close = OPEN_CASH + collections - payments - CAPEX - DRAWINGS
    equity_close = OPEN_EQUITY + net - DRAWINGS
    assets = cash_close + ar_close + equip_close
    bva = {}
    for b in budgets:
        bva.setdefault(b["category"], {"type": b["type"], "budget": 0, "actual": 0})
        bva[b["category"]]["budget"] += int(b["budget"])
    for t in txns:
        bva[t["category"]]["actual"] += int(t["amount"])
    return {
        "txn_count": len(txns), "revenue": revenue, "direct": direct,
        "opex": opex, "depreciation": depr, "net_income": net,
        "margin": net / revenue if revenue else 0,
        "cash_close": cash_close, "ar_close": ar_close,
        "ap_close": ap_close, "equip_close": equip_close,
        "assets": assets, "equity_close": equity_close,
        "balance_gap": assets - (ap_close + equity_close),
        "monthly": monthly, "bva": bva,
        "budget_total": sum(v["budget"] for v in bva.values()),
        "actual_total": sum(v["actual"] for v in bva.values()),
    }


if __name__ == "__main__":
    t = compute()
    for k in ("txn_count", "revenue", "direct", "opex", "depreciation",
              "net_income", "cash_close", "ar_close", "ap_close",
              "equip_close", "assets", "equity_close", "balance_gap"):
        print(k, t[k])
    print("margin", round(t["margin"], 4))
