#!/usr/bin/env python3
"""FlowForge data generator — deterministic fictional data for a 13-week
cash-flow forecast + budget-variance dashboard.

Scenario: 'Patna Traders Pvt. Ltd.', a fictional FMCG wholesale distributor.
26 weeks of history (W-25 .. W0) of weekly cash inflows/outflows by category,
plus a budget plan for those weeks, plus assumption values for the forecast.

Deterministic: fixed seed, no network, no secrets. Values in INR.
"""
import csv
import random
from pathlib import Path

DATA = Path(__file__).parent / "data"
DATA.mkdir(exist_ok=True)
random.seed(20261015)

WEEKS_HIST = 26  # W-25 .. W0
OPENING_BALANCE = 42_50_000

INFLOW_CATS = ["sales_receipts", "gst_refund", "other_income"]
OUTFLOW_CATS = ["supplier_payments", "payroll", "rent", "logistics", "gst_tax", "loan_emi"]

# --- History: actuals -----------------------------------------------------
hist_inflows, hist_outflows = [], []
trend = 0.0
for w in range(-WEEKS_HIST + 1, 1):
    trend += random.uniform(0.000, 0.012)  # mild growth drift
    sales = int(900_000 * (1 + trend) * random.uniform(0.88, 1.12))
    refund = int(random.choice([0, 0, 0, 45_000, 72_000, 30_000]))
    other = int(random.uniform(2_000, 15_000))
    hist_inflows.append({"week": w, "sales_receipts": sales,
                         "gst_refund": refund, "other_income": other})
    supp = int(sales * random.uniform(0.68, 0.80))
    payroll = int(random.uniform(116_000, 124_000))
    rent = 25_000
    logistics = int(random.uniform(38_000, 52_000))
    gst = 0 if (w % 4 != 0) else int(supp * 0.18 * random.uniform(0.85, 1.1))  # monthly GST
    emi = 78_000 if (w % 4 == 1) else 0  # monthly EMI slice
    hist_outflows.append({"week": w, "supplier_payments": supp, "payroll": payroll,
                          "rent": rent, "logistics": logistics,
                          "gst_tax": gst, "loan_emi": emi})

# --- Budget plan for the historical weeks (planned, vs actuals above) ------
hist_budget = []
for w in range(-WEEKS_HIST + 1, 1):
    hist_budget.append({
        "week": w,
        "budget_inflow": int(930_000 * random.uniform(0.97, 1.03)),
        "budget_outflow": int(810_000 * random.uniform(0.97, 1.03)),
    })

# --- Forecast assumptions -------------------------------------------------
assumptions = [
    ("opening_balance", 0, "Projected opening bank balance for Week 1 of forecast"),
    ("inflow_growth", 0.02, "Expected weekly inflow growth vs trailing 4-week average"),
    ("outflow_growth", 0.015, "Expected weekly outflow growth vs trailing 4-week average"),
    ("cash_floor", 5_00_000, "Minimum safe cash balance — forecast weeks below this are flagged"),
    ("payroll_weekly", 120_000, "Fixed weekly payroll baseline inside supplier mix"),
    ("gst_monthly", 1, "GST week of each 4-week cycle (1..4)"),
    ("emi_monthly", 78_000, "Monthly loan EMI charged on gst-week + 1"),
    ("avg_weeks_basis", 4, "Trailing weeks used for the run-rate forecast"),
    ("variance_alert_pct", 10, "Budget-vs-actual variance % that triggers an alert"),
]


def write_csv(name, rows, fields):
    with open(DATA / name, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def main():
    write_csv("history_inflows.csv", hist_inflows, ["week"] + INFLOW_CATS)
    write_csv("history_outflows.csv", hist_outflows, ["week"] + OUTFLOW_CATS)
    write_csv("history_budget.csv", hist_budget, ["week", "budget_inflow", "budget_outflow"])
    with open(DATA / "assumptions.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["key", "value", "description"])
        w.writerows(assumptions)

    tin = sum(sum(r[c] for c in INFLOW_CATS) for r in hist_inflows)
    tout = sum(sum(r[c] for c in OUTFLOW_CATS) for r in hist_outflows)
    print(f"weeks={WEEKS_HIST} total_inflow={tin:,} total_outflow={tout:,} "
          f"closing={OPENING_BALANCE + tin - tout:,}")

    # Persist expected closing for the test harness
    (DATA / "expected.txt").write_text(
        f"closing={OPENING_BALANCE + tin - tout}\n")


if __name__ == "__main__":
    main()
