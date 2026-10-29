#!/usr/bin/env python3
"""Generate deterministic synthetic data for the RenewableMIS demo.

Fictional portfolio: 17 solar/wind/hybrid/rooftop sites across nine Indian
states plus two international rooftop sites, Apr-Oct 2026 (7 months).
Seed 20261029. All figures are invented for the demo; no real client data.
Planned generation = capacity x monthly capacity-utilisation factor x hours.
Actual generation = planned x a seeded performance factor (0.84-1.05), with
two injected events: heavy curtailment at Kutch Wind in Jul 2026 and a long
outage at Kamuthi Solar in Sep 2026. Revenue is in Rs lakh:
revenue_lakh = MWh x tariff(Rs/kWh) / 100.
"""
import csv
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
SEED = 20261029

# (SiteID, Site, State, Technology, CapacityMW, TariffRsPerKWh)
SITES = [
    ("S01", "Jodhpur Solar Park", "Rajasthan", "Solar", 50, 2.95),
    ("S02", "Kutch Solar Park", "Gujarat", "Solar", 40, 2.85),
    ("S03", "Solapur Solar Park", "Maharashtra", "Solar", 25, 3.10),
    ("S04", "Pavagada Solar Park", "Karnataka", "Solar", 30, 3.05),
    ("S05", "Kamuthi Solar Park", "Tamil Nadu", "Solar", 20, 3.20),
    ("S06", "Mahbubnagar Solar Park", "Telangana", "Solar", 15, 3.00),
    ("S07", "Anantapur Solar Park", "Andhra Pradesh", "Solar", 35, 2.90),
    ("S08", "Rewa Solar Park", "Madhya Pradesh", "Solar", 45, 2.80),
    ("S09", "Jhansi Solar Park", "Uttar Pradesh", "Solar", 10, 3.30),
    ("W01", "Kutch Wind Farm", "Gujarat", "Wind", 30, 3.60),
    ("W02", "Tirunelveli Wind Farm", "Tamil Nadu", "Wind", 25, 3.75),
    ("W03", "Satara Wind Farm", "Maharashtra", "Wind", 20, 3.85),
    ("W04", "Chitradurga Wind Farm", "Karnataka", "Wind", 15, 3.70),
    ("H01", "Bikaner Hybrid Park", "Rajasthan", "Hybrid", 60, 3.15),
    ("H02", "Kutch Hybrid Park", "Gujarat", "Hybrid", 40, 3.10),
    ("R01", "Dubai Rooftop Solar", "UAE", "Rooftop", 2, 3.40),
    ("R02", "Singapore Rooftop Solar", "Singapore", "Rooftop", 1.5, 3.50),
]

# Monthly capacity-utilisation factors by technology (Apr-Oct).
CUF = {
    "Solar":   [0.21, 0.22, 0.19, 0.18, 0.19, 0.21, 0.22],
    "Wind":    [0.30, 0.32, 0.36, 0.38, 0.37, 0.33, 0.30],
    "Hybrid":  [0.26, 0.27, 0.28, 0.28, 0.27, 0.27, 0.26],
    "Rooftop": [0.18, 0.18, 0.17, 0.16, 0.17, 0.18, 0.18],
}

MONTHS = ["2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
DAYS = [30, 31, 30, 31, 31, 30, 31]


def main():
    os.makedirs(DATA, exist_ok=True)
    rng = random.Random(SEED)
    with open(os.path.join(DATA, "sites.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["SiteID", "Site", "State", "Technology", "CapacityMW", "TariffRsPerKWh"])
        for s in SITES:
            w.writerow(s)
    plan_rows, actual_rows = [], []
    for site_id, site, state, tech, cap, tariff in SITES:
        for i, month in enumerate(MONTHS):
            hours = DAYS[i] * 24
            planned = round(cap * CUF[tech][i] * hours)
            planned_rev = round(planned * tariff / 100.0, 2)
            factor = round(rng.uniform(0.84, 1.05), 3)
            curt_rate = round(rng.uniform(0.0, 0.06), 3)
            outage = rng.randint(0, 40)
            # Injected events (kept deterministic and documented in README).
            if site_id == "W01" and month == "2026-07":
                curt_rate = 0.18
                factor = 0.86
            if site_id == "S05" and month == "2026-09":
                outage = 120
                factor = 0.78
            actual = round(planned * factor)
            actual_rev = round(actual * tariff / 100.0, 2)
            curt = round(planned * curt_rate)
            plan_rows.append([site_id, site, state, tech, month, planned, planned_rev])
            actual_rows.append([site_id, site, state, tech, month, actual, actual_rev, curt, outage])
    with open(os.path.join(DATA, "plan.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["SiteID", "Site", "State", "Technology", "Month", "PlannedMWh", "PlannedRevenueLakh"])
        w.writerows(plan_rows)
    with open(os.path.join(DATA, "actuals.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["SiteID", "Site", "State", "Technology", "Month", "ActualMWh",
                    "ActualRevenueLakh", "CurtailmentMWh", "OutageHours"])
        w.writerows(actual_rows)
    print(f"sites={len(SITES)} plan_rows={len(plan_rows)} actual_rows={len(actual_rows)}")


if __name__ == "__main__":
    main()
