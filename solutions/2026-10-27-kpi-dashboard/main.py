#!/usr/bin/env python3
"""Independent ground truth: recompute every dashboard figure from data/*.csv.

Used by tests/test_build.py; never reads the workbook, so a formula bug in
build_workbook.py cannot hide behind its own numbers.
"""
import csv
import json
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")


def compute():
    monthly = defaultdict(lambda: {"revenue": 0, "orders": 0, "cost": 0,
                                   "ontime": 0, "csat": 0.0})
    channel = defaultdict(lambda: {"revenue": 0, "orders": 0, "cost": 0})
    with open(os.path.join(DATA, "raw_transactions.csv"), newline="") as f:
        for r in csv.DictReader(f):
            m = monthly[r["Month"]]
            m["revenue"] += int(r["Revenue"])
            m["orders"] += int(r["Orders"])
            m["cost"] += int(r["Cost"])
            m["ontime"] += int(r["OnTimeOrders"])
            m["csat"] += float(r["CSATSum"])
            c = channel[r["Channel"]]
            c["revenue"] += int(r["Revenue"])
            c["orders"] += int(r["Orders"])
            c["cost"] += int(r["Cost"])
    targets = {}
    with open(os.path.join(DATA, "monthly_targets.csv"), newline="") as f:
        for r in csv.DictReader(f):
            targets[r["Month"]] = int(r["RevenueTarget"])

    out_months = {}
    for month in sorted(monthly):
        m = monthly[month]
        out_months[month] = {
            "revenue": m["revenue"],
            "target": targets[month],
            "variance": m["revenue"] - targets[month],
            "variance_pct": (m["revenue"] - targets[month]) / targets[month],
            "orders": m["orders"],
            "aov": m["revenue"] / m["orders"],
            "cost": m["cost"],
            "gross_profit": m["revenue"] - m["cost"],
            "margin_pct": (m["revenue"] - m["cost"]) / m["revenue"],
            "ontime_pct": m["ontime"] / m["orders"],
            "csat": m["csat"] / m["orders"],
        }
    ytd_rev = sum(m["revenue"] for m in monthly.values())
    ytd_cost = sum(m["cost"] for m in monthly.values())
    out = {
        "months": out_months,
        "channels": {k: {"revenue": v["revenue"], "orders": v["orders"],
                         "cost": v["cost"],
                         "margin_pct": (v["revenue"] - v["cost"]) / v["revenue"],
                         "share": v["revenue"] / ytd_rev}
                     for k, v in sorted(channel.items())},
        "dashboard": {
            "latest_revenue": out_months["2026-09"]["revenue"],
            "ytd_revenue": ytd_rev,
            "ytd_margin": (ytd_rev - ytd_cost) / ytd_rev,
            "latest_ontime": out_months["2026-09"]["ontime_pct"],
            "latest_csat": out_months["2026-09"]["csat"],
            "target_attainment": ytd_rev / sum(targets.values()),
        },
    }
    return out


if __name__ == "__main__":
    print(json.dumps(compute(), indent=2))
