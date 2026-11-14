#!/usr/bin/env python3
"""Make the input data for AcquisitionModel.

Fictional target: Sharma Precision Components Pvt. Ltd., Patna — a
precision auto-parts manufacturer being acquired. All figures are
invented for the demo; no real business data is used.

Writes three deterministic CSVs (seed 20261114):
- data/assumptions.csv — deal + valuation inputs (asking price, net debt,
  shares, tax, WACC, terminal growth, depreciation, capex, working
  capital % of revenue, annual interest).
- data/historical.csv — actuals 2023-2025 (revenue, COGS, opex, other
  income/expense). 2025 is the base year for the projections.
- data/scenarios.csv — Base / Bull / Bear scenario drivers: revenue
  growth, gross margin, opex growth.

The workbook (generate.py) reads these files; every projected figure is
a live formula off them, so changing an input and re-running rebuilds
the whole model.
"""
import csv
import random
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
SEED = 20261114

ASSUMPTIONS = [
    ("asking_price_usd", 2400000),
    ("net_debt_usd", 650000),
    ("shares_outstanding", 100000),
    ("tax_rate", 0.25),
    ("wacc", 0.11),
    ("terminal_growth", 0.025),
    ("depreciation_usd_per_year", 220000),
    ("capex_usd_per_year", 180000),
    ("working_capital_pct_revenue", 0.12),
    ("interest_usd_per_year", 90000),
]

# Historical actuals. Revenue grows ~15%/yr; margins improve as the
# plant scales. 2025 opex is the base for projected opex.
HISTORICAL = [
    {"year": 2023, "revenue": 3100000, "cogs": 2015000, "opex": 1180000,
     "other_income": 12000, "other_expense": 8000},
    {"year": 2024, "revenue": 3600000, "cogs": 2304000, "opex": 1290000,
     "other_income": 15000, "other_expense": 9000},
    {"year": 2025, "revenue": 4200000, "cogs": 2604000, "opex": 1450000,
     "other_income": 18000, "other_expense": 10000},
]

SCENARIOS = [
    {"scenario": "Base", "revenue_growth": 0.08, "gross_margin": 0.38,
     "opex_growth": 0.05},
    {"scenario": "Bull", "revenue_growth": 0.12, "gross_margin": 0.41,
     "opex_growth": 0.04},
    {"scenario": "Bear", "revenue_growth": 0.03, "gross_margin": 0.34,
     "opex_growth": 0.07},
]


def main():
    random.seed(SEED)  # data is fixed; seed documents determinism
    DATA.mkdir(exist_ok=True)
    with open(DATA / "assumptions.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["key", "value"])
        w.writerows(ASSUMPTIONS)
    with open(DATA / "historical.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["year", "revenue", "cogs",
                                          "opex", "other_income",
                                          "other_expense"])
        w.writeheader()
        w.writerows(HISTORICAL)
    with open(DATA / "scenarios.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["scenario", "revenue_growth",
                                          "gross_margin", "opex_growth"])
        w.writeheader()
        w.writerows(SCENARIOS)
    print(f"wrote {len(ASSUMPTIONS)} assumptions, "
          f"{len(HISTORICAL)} historical years, "
          f"{len(SCENARIOS)} scenarios")


if __name__ == "__main__":
    main()
