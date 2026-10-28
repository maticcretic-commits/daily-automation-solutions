#!/usr/bin/env python3
"""Independent ground truth: recompute every workbook figure from data/*.csv.

Used by tests/test_build.py; never reads the workbook, so a formula bug in
build_workbook.py cannot hide behind its own numbers.
"""
import csv
import json
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
YEARS = ["FY2024", "FY2025", "FY2026"]


def compute():
    st = {}
    with open(os.path.join(DATA, "statements.csv"), newline="") as f:
        for r in csv.DictReader(f):
            st[(r["LineItem"], r["Year"])] = int(r["Amount"])
    opening = {k: v for (k, y), v in st.items() if y == "Opening"}

    def v(item, year):
        return st[(item, year)]

    out = {"years": {}}
    prev = dict(opening)
    prev_cash = None
    for year in YEARS:
        rev = v("Revenue", year)
        cogs = v("Cost of goods sold", year)
        gross = rev - cogs
        ebitda = gross - v("Operating expenses", year)
        ebit = ebitda - v("Depreciation", year)
        pbt = ebit - v("Interest expense", year)
        pat = pbt - v("Tax expense", year)
        ca = (v("Cash and bank", year) + v("Trade receivables", year)
              + v("Inventories", year) + v("Other current assets", year))
        nca = (v("Property plant and equipment (net)", year)
               + v("Other non-current assets", year))
        cl = (v("Trade payables", year) + v("Short-term borrowings", year)
              + v("Other current liabilities", year))
        equity = v("Share capital", year) + v("Retained earnings", year)
        dwc = ((v("Trade receivables", year) - prev["Trade receivables"])
               + (v("Inventories", year) - prev["Inventories"])
               - (v("Trade payables", year) - prev["Trade payables"])
               + (v("Other current assets", year) - prev["Other current assets"])
               - (v("Other current liabilities", year)
                  - prev["Other current liabilities"]))
        cfo = (pat + v("Depreciation", year) - dwc)
        out["years"][year] = {
            "revenue": rev, "cogs": cogs, "gross": gross, "ebitda": ebitda,
            "ebit": ebit, "pbt": pbt, "pat": pat,
            "total_assets": ca + nca, "total_le": cl + v("Long-term borrowings", year) + equity,
            "balance_gap": (ca + nca) - (cl + v("Long-term borrowings", year) + equity),
            "current_ratio": ca / cl,
            "quick_ratio": (ca - v("Inventories", year)) / cl,
            "debt_equity": (v("Short-term borrowings", year)
                            + v("Long-term borrowings", year)) / equity,
            "interest_coverage": ebit / v("Interest expense", year),
            "gross_margin": gross / rev,
            "net_margin": pat / rev,
            "roa": pat / (ca + nca),
            "roe": pat / equity,
            "recv_days": v("Trade receivables", year) / rev * 365,
            "inv_days": v("Inventories", year) / cogs * 365,
            "pay_days": v("Trade payables", year) / cogs * 365,
            "cash_conversion": cfo / pat,
            "working_capital": ca - cl,
            "cfo": cfo,
            "closing_cash": v("Cash and bank", year),
        }
        prev = {item: v(item, year) for item in
                ("Trade receivables", "Inventories", "Trade payables",
                 "Other current assets", "Other current liabilities",
                 "Short-term borrowings", "Long-term borrowings")}
        prev_cash = v("Cash and bank", year)

    monthly = []
    with open(os.path.join(DATA, "monthly.csv"), newline="") as f:
        for r in csv.DictReader(f):
            monthly.append({k: (r[k] if k == "Month" else int(r[k]))
                            for k in r})
    out["monthly"] = monthly
    out["trend_total_revenue"] = sum(m["Revenue"] for m in monthly)
    out["trend_total_cogs"] = sum(m["COGS"] for m in monthly)
    out["dashboard"] = {
        "revenue": out["years"]["FY2026"]["revenue"],
        "pat": out["years"]["FY2026"]["pat"],
        "net_margin": out["years"]["FY2026"]["net_margin"],
        "current_ratio": out["years"]["FY2026"]["current_ratio"],
        "recv_days": out["years"]["FY2026"]["recv_days"],
        "cash_conversion": out["years"]["FY2026"]["cash_conversion"],
    }
    return out


if __name__ == "__main__":
    print(json.dumps(compute(), indent=2))
