#!/usr/bin/env python3
"""ForecastForge — independent Python model (ground truth).

Fictional company: Ganga Foods Pvt. Ltd., Bodh Gaya (packaged foods).
36-month forecast, Apr 2026 – Mar 2029 (FY2027 / FY2028 / FY2029).

This module is the INDEPENDENT computation. generate.py builds the
workbook purely from data/assumptions.csv with live Excel formulas;
tests compare the LibreOffice recalc of that workbook against the
figures computed here. The two must agree — that is the verification.

Model rules (all mirrored by workbook formulas):
  revenue(m)  = rev_base * (1 + growth) ** (m - 1)
  cogs(m)     = revenue(m) * cogs_pct
  salaries(m) = salary_base * (1 + salary_growth) ** (m - 1)
  marketing(m)= revenue(m) * marketing_pct   (variable cost)
  rent / utilities / other opex are flat per month
  depreciation = 35000 + 5000 from month 7 (capex 600k, Oct 2026)
                      + 10000/3 from month 13 (capex 400k, Apr 2027)
  interest(m) = opening loan balance * 11% / 12; principal 50k/month
  tax(m)      = 25% * max(0, pbt(m))   (monthly provision)
  receivables = month-end revenue * 20/30 days
  inventory   = month-end COGS * 15/30 days
  payables    = month-end COGS * 25/30 days
  break-even monthly revenue = fixed costs / (1 - cogs% - marketing%)
    where fixed costs = salaries + rent + utilities + other
                        + depreciation + interest (month-end values)
"""
import json

# --- Drivers (Base scenario) -------------------------------------------
DRIVERS = {
    "company": "Ganga Foods Pvt. Ltd., Bodh Gaya",
    "rev_base": 950000.0,          # month-1 revenue (Apr 2026)
    "growth": {"Base": 0.022, "Optimistic": 0.032, "Pessimistic": 0.012},
    "cogs_pct": {"Base": 0.58, "Optimistic": 0.54, "Pessimistic": 0.62},
    "salary_base": 185000.0,
    "salary_growth": 0.006,        # per month
    "rent": 45000.0,
    "utilities": 28000.0,
    "other_opex": 22000.0,
    "marketing_pct": 0.03,         # of revenue
    "loan_opening": 2500000.0,
    "loan_rate": 0.11,             # per year
    "loan_principal": 50000.0,     # per month
    "tax_rate": 0.25,
    "plant_opening": 4200000.0,
    "plant_life_months": 120,
    "capex": {7: 600000.0, 13: 400000.0},   # month -> amount
    "wc_recv_days": 20,
    "wc_inv_days": 15,
    "wc_pay_days": 25,
    # Opening balance sheet, 1 Apr 2026 (must balance: 6,070,000)
    "cash_opening": 800000.0,
    "recv_opening": 650000.0,
    "inv_opening": 420000.0,
    "pay_opening": 310000.0,
    "equity": 2500000.0,
    "retained_opening": 760000.0,
}

MONTHS = 36
MONTH_LABELS = [
    f"{m:02d}-{y}" for y in (2026, 2027, 2028, 2029)
    for m in range(4, 13) for _ in [0]
][:0]  # placeholder, replaced below
# Apr 2026 .. Mar 2029
_LABELS = []
_y, _m = 2026, 4
for _ in range(MONTHS):
    _LABELS.append(f"{_y}-{_m:02d}")
    _m += 1
    if _m == 13:
        _m, _y = 1, _y + 1
MONTH_LABELS = _LABELS
FY_LABELS = ["FY2027", "FY2028", "FY2029"]


def monthly(scenario="Base"):
    """Return list of 36 monthly dicts for the given scenario."""
    d = DRIVERS
    g = d["growth"][scenario]
    cogs_pct = d["cogs_pct"][scenario]
    dep_base = d["plant_opening"] / d["plant_life_months"]
    rows = []
    for m in range(1, MONTHS + 1):
        rev = d["rev_base"] * (1 + g) ** (m - 1)
        cogs = rev * cogs_pct
        gross = rev - cogs
        salaries = d["salary_base"] * (1 + d["salary_growth"]) ** (m - 1)
        marketing = rev * d["marketing_pct"]
        opex = salaries + d["rent"] + d["utilities"] + d["other_opex"] + marketing
        ebitda = gross - opex
        dep = dep_base
        if m >= 7:
            dep += 600000.0 / 120
        if m >= 13:
            dep += 400000.0 / 120
        loan_open = max(0.0, d["loan_opening"] - d["loan_principal"] * (m - 1))
        interest = loan_open * d["loan_rate"] / 12
        pbt = ebitda - dep - interest
        tax = max(0.0, pbt) * d["tax_rate"]
        net = pbt - tax
        rows.append({
            "month": m, "label": MONTH_LABELS[m - 1],
            "revenue": rev, "cogs": cogs, "gross": gross,
            "salaries": salaries, "rent": d["rent"],
            "utilities": d["utilities"], "marketing": marketing,
            "other": d["other_opex"], "opex": opex, "ebitda": ebitda,
            "depreciation": dep, "interest": interest,
            "pbt": pbt, "tax": tax, "net": net,
            "capex": d["capex"].get(m, 0.0),
            "loan_open": loan_open,
        })
    return rows


def annual(rows):
    """Aggregate the 36 months into 3 fiscal years (Apr-Mar)."""
    out = []
    for yi in range(3):
        chunk = rows[yi * 12:(yi + 1) * 12]
        s = lambda k: sum(r[k] for r in chunk)
        out.append({
            "fy": FY_LABELS[yi],
            "revenue": s("revenue"), "cogs": s("cogs"),
            "gross": s("gross"), "opex": s("opex"),
            "ebitda": s("ebitda"), "depreciation": s("depreciation"),
            "interest": s("interest"), "pbt": s("pbt"),
            "tax": s("tax"), "net": s("net"), "capex": s("capex"),
        })
    return out


def balance_sheet(rows, years):
    """Year-end balance sheets + cash flow. Returns (bs, cf) lists."""
    d = DRIVERS
    bs, cf = [], []
    prev = {"recv": d["recv_opening"], "inv": d["inv_opening"],
            "pay": d["pay_opening"], "cash": d["cash_opening"]}
    cum_net = 0.0
    cum_dep = 0.0
    cum_capex = 0.0
    for yi in range(3):
        last = rows[(yi + 1) * 12 - 1]      # Mar month of the FY
        rec = last["revenue"] * d["wc_recv_days"] / 30
        inv = last["cogs"] * d["wc_inv_days"] / 30
        pay = last["cogs"] * d["wc_pay_days"] / 30
        y = years[yi]
        cum_net += y["net"]
        cum_dep += y["depreciation"]
        cum_capex += y["capex"]
        net_fa = d["plant_opening"] + cum_capex - cum_dep
        loan = d["loan_opening"] - d["loan_principal"] * 12 * (yi + 1)
        re_ = d["retained_opening"] + cum_net
        # Indirect cash flow for the year
        d_rec = rec - prev["recv"]
        d_inv = inv - prev["inv"]
        d_pay = pay - prev["pay"]
        principal = d["loan_principal"] * 12
        cf_year = (y["net"] + y["depreciation"] - d_rec - d_inv + d_pay
                   - y["capex"] - principal)
        cash = prev["cash"] + cf_year
        assets = cash + rec + inv + net_fa
        liab = pay + loan + d["equity"] + re_
        bs.append({
            "fy": y["fy"], "cash": cash, "receivables": rec,
            "inventory": inv, "net_fa": net_fa, "assets": assets,
            "payables": pay, "loan": loan, "equity": d["equity"],
            "retained": re_, "liab": liab,
            "balance_check": assets - liab,
        })
        cf.append({
            "fy": y["fy"], "net": y["net"],
            "depreciation": y["depreciation"],
            "d_recv": d_rec, "d_inv": d_inv, "d_pay": d_pay,
            "capex": y["capex"], "principal": principal,
            "net_cash": cf_year, "opening_cash": prev["cash"],
            "closing_cash": cash,
        })
        prev = {"recv": rec, "inv": inv, "pay": pay, "cash": cash}
    return bs, cf


def ratios(rows, years, bs):
    """Per-FY ratios + break-even monthly revenue (year-end month)."""
    d = DRIVERS
    out = []
    for yi in range(3):
        y, b = years[yi], bs[yi]
        last = rows[(yi + 1) * 12 - 1]
        fixed = (last["salaries"] + d["rent"] + d["utilities"]
                 + d["other_opex"] + last["depreciation"] + last["interest"])
        cm_ratio = 1 - d["cogs_pct"]["Base"] - d["marketing_pct"]
        out.append({
            "fy": y["fy"],
            "gross_margin": y["gross"] / y["revenue"],
            "net_margin": y["net"] / y["revenue"],
            "current_ratio": (b["cash"] + b["receivables"] + b["inventory"])
                             / b["payables"],
            "debt_equity": b["loan"] / (b["equity"] + b["retained"]),
            "roe": y["net"] / (b["equity"] + b["retained"]),
            "breakeven_revenue": fixed / cm_ratio,
            "fixed_costs_month": fixed,
        })
    return out


def sensitivity_centre(years):
    """Year-3 net profit — the sensitivity grid's centre cell must equal this."""
    return years[2]["net"]


def sensitivity_cell(rows, years, dg, dc):
    """Year-3 net profit under growth offset dg and COGS% offset dc.

    Closed form, mirroring the monthly model exactly for revenue
    (geometric series over months 25..36) and reusing the fixed
    Year-3 cost blocks from the annual model.
    """
    d = DRIVERS
    g = d["growth"]["Base"] + dg
    r = 1 + g
    # sum_{k=24}^{35} r^k = r^24 * (r^12 - 1) / (r - 1)
    rev3 = d["rev_base"] * r ** 24 * (r ** 12 - 1) / (r - 1)
    cogs3 = rev3 * (d["cogs_pct"]["Base"] + dc)
    gross3 = rev3 - cogs3
    y3 = years[2]
    marketing3_base = 0.03 * sum(rows[m]["revenue"] for m in range(24, 36))
    opex3 = (y3["opex"] - marketing3_base) + 0.03 * rev3
    pbt3 = gross3 - opex3 - y3["depreciation"] - y3["interest"]
    return pbt3 - max(0.0, pbt3) * d["tax_rate"]


def build_ground_truth(scenario="Base"):
    rows = monthly(scenario)
    years = annual(rows)
    bs, cf = balance_sheet(rows, years)
    rat = ratios(rows, years, bs)
    return {
        "company": DRIVERS["company"],
        "scenario": scenario,
        "months": rows,
        "years": years,
        "balance_sheet": bs,
        "cash_flow": cf,
        "ratios": rat,
        "sensitivity_centre": sensitivity_centre(years),
        "opening_check": (
            DRIVERS["cash_opening"] + DRIVERS["recv_opening"]
            + DRIVERS["inv_opening"] + DRIVERS["plant_opening"]
            - (DRIVERS["pay_opening"] + DRIVERS["loan_opening"]
               + DRIVERS["equity"] + DRIVERS["retained_opening"])),
    }


if __name__ == "__main__":
    gt = build_ground_truth()
    print(f"Opening BS check: {gt['opening_check']:.2f} (must be 0)")
    for y in gt["years"]:
        print(f"{y['fy']}: revenue {y['revenue']:,.0f}  net {y['net']:,.0f}  "
              f"net margin {y['net']/y['revenue']:.1%}")
    for b in gt["balance_sheet"]:
        print(f"{b['fy']}: cash {b['cash']:,.0f}  balance check "
              f"{b['balance_check']:.2f}")
    for r in gt["ratios"]:
        print(f"{r['fy']}: break-even monthly revenue "
              f"{r['breakeven_revenue']:,.0f}")
    print(f"Sensitivity centre (Y3 net): {gt['sensitivity_centre']:,.0f}")
