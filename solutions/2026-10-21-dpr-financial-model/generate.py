"""Write the deterministic input CSVs for the DPRPulse model.

Everything here is an INPUT (blue cells in the workbook); every computed
figure in dprpulse.xlsx is a live formula over these inputs.
"""
import csv
import os

from main import CAPEX, UTIL, N_YEARS

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(DATA, exist_ok=True)

ASSUMPTIONS = [
    ("capacity_mt_per_year", 1200, "", "rated turmeric-processing capacity"),
    ("price_y1_rs_per_kg", 150, "Rs/kg", "ex-factory realisation, year 1"),
    ("price_escalation", 0.03, "/yr", "annual price increase"),
    ("raw_material_pct_revenue", 0.76, "%", "raw turmeric, % of revenue"),
    ("packing_pct_revenue", 0.06, "%", "packing material, % of revenue"),
    ("admin_selling_pct_revenue", 0.05, "%", "admin & selling, % of revenue"),
    ("power_fuel_y1_lakh", 22.0, "Rs lakh", "dryers/grinders, year 1"),
    ("power_escalation", 0.04, "/yr", ""),
    ("salaries_y1_lakh", 34.0, "Rs lakh", "year 1 wage bill"),
    ("salary_escalation", 0.05, "/yr", ""),
    ("repairs_lakh_per_year", 2.2, "Rs lakh", "2% of P&M cost, flat"),
    ("insurance_y1_lakh", 1.5, "Rs lakh", "year 1"),
    ("insurance_escalation", 0.03, "/yr", ""),
    ("depr_building_lakh_per_year", 2.75, "Rs lakh", "55 / 20 yrs SLM"),
    ("depr_pm_lakh_per_year", 12.0, "Rs lakh", "120 (incl. contingency) / 10 yrs"),
    ("depr_electrical_lakh_per_year", 1.8, "Rs lakh", "18 / 10 yrs SLM"),
    ("prelim_amort_lakh_per_year", 2.4, "Rs lakh", "12 / 5 yrs, years 1-5"),
    ("tax_rate", 0.25, "%", "corporate tax, no loss carry-forward modelled"),
    ("term_loan_lakh", 180.0, "Rs lakh", ""),
    ("interest_rate", 0.10, "/yr", ""),
    ("moratorium_years", 1, "yrs", "interest serviced, no principal"),
    ("repayment_years", 6, "yrs", "annual instalments, years 2-7"),
    ("promoter_equity_lakh", 60.0, "Rs lakh", "25% margin"),
    ("discount_rate", 0.12, "/yr", "NPV/IRR hurdle"),
    ("project_cost_check_lakh", 240.0, "Rs lakh", "must equal capex total"),
]

with open(os.path.join(DATA, "assumptions.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["param", "value", "unit", "note"])
    w.writerows(ASSUMPTIONS)

with open(os.path.join(DATA, "project_cost.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["head", "amount_lakh"])
    w.writerows(CAPEX)

with open(os.path.join(DATA, "utilisation.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["year", "utilisation_pct"])
    for t in range(1, N_YEARS + 1):
        w.writerow([t, UTIL[t - 1]])

print("wrote", len(ASSUMPTIONS), "assumptions,",
      len(CAPEX), "capex heads,", N_YEARS, "utilisation years ->", DATA)
