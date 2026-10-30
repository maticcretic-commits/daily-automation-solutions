#!/usr/bin/env python3
"""Independent ground truth: recompute every workbook figure from data/*.csv.

Used by tests/test_build.py; never reads the workbook, so a formula bug in
build_workbook.py cannot hide behind its own numbers. Valuation: monthly
annuity in arrears at a single annual yield of 8.5% (monthly rate = 8.5%/12).
PV per lease = monthly disparity * (1 - (1+r)^-n) / r.
"""
import csv
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
ANNUAL_YIELD = 0.085
MONTHLY_RATE = ANNUAL_YIELD / 12.0
STATUS_TOL = 25.0


def compute():
    leases = []
    with open(os.path.join(DATA, "leases.csv"), newline="") as f:
        for r in csv.DictReader(f):
            contract = float(r["ContractRentMonthly"])
            market = float(r["MarketRentMonthly"])
            months = int(r["RemainingMonths"])
            disp = market - contract
            factor = (1 - (1 + MONTHLY_RATE) ** (-months)) / MONTHLY_RATE
            leases.append({
                "id": r["LeaseID"], "contract": contract, "market": market,
                "months": months, "disparity": disp,
                "annual": disp * 12,
                "undiscounted": disp * months,
                "pv": disp * factor,
                "status": ("At market" if abs(disp) <= STATUS_TOL
                           else "Under-rented" if disp > 0 else "Over-rented"),
            })
    total_contract = sum(l["contract"] for l in leases)
    total_market = sum(l["market"] for l in leases)
    total_disp = sum(l["disparity"] for l in leases)
    schedule = []
    cumulative = 0.0
    for m in range(1, 61):
        monthly = sum(l["disparity"] for l in leases if l["months"] >= m)
        discounted = monthly * (1 + MONTHLY_RATE) ** (-m)
        cumulative += discounted
        schedule.append({"month": m, "monthly": monthly,
                         "factor": (1 + MONTHLY_RATE) ** (-m),
                         "discounted": discounted, "cumulative": cumulative})
    return {
        "count": len(leases),
        "leases": leases,
        "total_contract": total_contract,
        "total_market": total_market,
        "total_disparity": total_disp,
        "total_annual": total_disp * 12,
        "total_undiscounted": sum(l["undiscounted"] for l in leases),
        "total_pv": sum(l["pv"] for l in leases),
        "under": sum(1 for l in leases if l["status"] == "Under-rented"),
        "over": sum(1 for l in leases if l["status"] == "Over-rented"),
        "at_market": sum(1 for l in leases if l["status"] == "At market"),
        "avg_months": sum(l["months"] for l in leases) / len(leases),
        "annual_yield": ANNUAL_YIELD,
        "monthly_rate": MONTHLY_RATE,
        "schedule": schedule,
        "schedule_pv": cumulative,
        "max_months": max(l["months"] for l in leases),
    }


if __name__ == "__main__":
    t = compute()
    print(f"leases={t['count']} contract=${t['total_contract']:,.0f}/mo "
          f"market=${t['total_market']:,.0f}/mo disparity=${t['total_disparity']:,.0f}/mo "
          f"PV=${t['total_pv']:,.2f} undiscounted=${t['total_undiscounted']:,.0f} "
          f"schedulePV=${t['schedule_pv']:,.2f} under={t['under']} over={t['over']} "
          f"at_market={t['at_market']} avg_months={t['avg_months']:.1f}")
    for l in t["leases"]:
        print(f"  {l['id']} disp=${l['disparity']:,.0f} PV=${l['pv']:,.2f} {l['status']}")
