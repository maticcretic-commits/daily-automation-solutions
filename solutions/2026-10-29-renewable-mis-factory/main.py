#!/usr/bin/env python3
"""Independent ground truth: recompute every workbook figure from data/*.csv.

Used by tests/test_build.py; never reads the workbook, so a formula bug in
build_workbook.py cannot hide behind its own numbers.
"""
import csv
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
MONTHS = ["2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
DAYS = [30, 31, 30, 31, 31, 30, 31]


def compute():
    sites = {}
    with open(os.path.join(DATA, "sites.csv"), newline="") as f:
        for r in csv.DictReader(f):
            sites[r["SiteID"]] = r
    plan = defaultdict(lambda: [0, 0.0])
    actual = defaultdict(lambda: [0, 0.0, 0, 0])
    with open(os.path.join(DATA, "plan.csv"), newline="") as f:
        for r in csv.DictReader(f):
            plan[r["Month"]][0] += int(r["PlannedMWh"])
            plan[r["Month"]][1] += float(r["PlannedRevenueLakh"])
    with open(os.path.join(DATA, "actuals.csv"), newline="") as f:
        for r in csv.DictReader(f):
            a = actual[r["Month"]]
            a[0] += int(r["ActualMWh"])
            a[1] += float(r["ActualRevenueLakh"])
            a[2] += int(r["CurtailmentMWh"])
            a[3] += int(r["OutageHours"])

    monthly = {}
    for i, m in enumerate(MONTHS):
        p_mwh, p_rev = plan[m]
        a_mwh, a_rev, curt, outage = actual[m]
        hours = DAYS[i] * 24 * len(sites)
        monthly[m] = {
            "planned": p_mwh, "actual": a_mwh,
            "achievement": a_mwh / p_mwh if p_mwh else 0,
            "planned_rev": p_rev, "actual_rev": a_rev,
            "rev_achievement": a_rev / p_rev if p_rev else 0,
            "curtailment": curt, "outage": outage,
            "availability": 1 - outage / hours if hours else 0,
        }
    tot_plan = sum(v["planned"] for v in monthly.values())
    tot_act = sum(v["actual"] for v in monthly.values())
    tot_p_rev = sum(v["planned_rev"] for v in monthly.values())
    tot_a_rev = sum(v["actual_rev"] for v in monthly.values())
    tot_curt = sum(v["curtailment"] for v in monthly.values())
    tot_out = sum(v["outage"] for v in monthly.values())
    tot_hours = sum(DAYS) * 24 * len(sites)
    out = {
        "sites": len(sites),
        "rows": len(MONTHS) * len(sites),
        "planned_mwh": tot_plan,
        "actual_mwh": tot_act,
        "achievement": tot_act / tot_plan if tot_plan else 0,
        "planned_rev": tot_p_rev,
        "actual_rev": tot_a_rev,
        "rev_achievement": tot_a_rev / tot_p_rev if tot_p_rev else 0,
        "curtailment": tot_curt,
        "curtailment_pct": tot_curt / tot_plan if tot_plan else 0,
        "outage": tot_out,
        "availability": 1 - tot_out / tot_hours if tot_hours else 0,
        "monthly": monthly,
    }

    # State and technology rollups (recomputed from the raw files directly).
    state = defaultdict(lambda: [0, 0, 0.0, 0.0, 0, 0])
    tech = defaultdict(lambda: [0, 0, 0.0, 0.0, 0, 0])
    with open(os.path.join(DATA, "plan.csv"), newline="") as f:
        for r in csv.DictReader(f):
            s = state[r["State"]]
            s[0] += int(r["PlannedMWh"])
            s[2] += float(r["PlannedRevenueLakh"])
            t = tech[r["Technology"]]
            t[0] += int(r["PlannedMWh"])
            t[2] += float(r["PlannedRevenueLakh"])
    with open(os.path.join(DATA, "actuals.csv"), newline="") as f:
        for r in csv.DictReader(f):
            s = state[r["State"]]
            s[1] += int(r["ActualMWh"])
            s[3] += float(r["ActualRevenueLakh"])
            s[4] += int(r["CurtailmentMWh"])
            s[5] += int(r["OutageHours"])
            t = tech[r["Technology"]]
            t[1] += int(r["ActualMWh"])
            t[3] += float(r["ActualRevenueLakh"])
            t[4] += int(r["CurtailmentMWh"])
            t[5] += int(r["OutageHours"])
    out["state"] = {k: {"planned": v[0], "actual": v[1],
                         "achievement": v[1] / v[0] if v[0] else 0,
                         "planned_rev": v[2], "actual_rev": v[3],
                         "curtailment": v[4], "outage": v[5]}
                    for k, v in state.items()}
    out["technology"] = {k: {"planned": v[0], "actual": v[1],
                             "achievement": v[1] / v[0] if v[0] else 0,
                             "planned_rev": v[2], "actual_rev": v[3],
                             "curtailment": v[4], "outage": v[5]}
                        for k, v in tech.items()}
    return out
