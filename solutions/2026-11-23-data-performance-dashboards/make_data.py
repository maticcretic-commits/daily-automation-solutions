#!/usr/bin/env python3
"""Generate fictional daily performance data for PerfDash.

Fictional business: Ganga Retail & Services Pvt. Ltd., Patna — an FMCG
wholesaler/retailer selling across five channels, Apr–Oct 2026.
Deterministic (seed 20261123): every run produces identical CSVs.

Outputs (into data/):
  performance_daily.csv — one row per channel-day (date, channel,
    product, campaign, impressions, clicks, orders, revenue, ad spend)
  targets.csv           — monthly revenue target per month
  ground_truth.json     — independent Python aggregates the tests and
    the workbook's Checks sheet are verified against
"""
import csv
import datetime as dt
import json
import os
import random

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
SEED = 20261123

CHANNELS = {
    # base impressions/day, CTR, conversion, average order value (Rs).
    # Ad channels bill per click (cpc); warm channels bill per order (cpo)
    # so every channel's ROAS lands in a realistic 2-9x band.
    "Online Store":    {"imp": 6000, "ctr": 0.032, "conv": 0.045, "aov": 850, "cpc": 6},
    "Marketplace":     {"imp": 8000, "ctr": 0.021, "conv": 0.038, "aov": 720, "cpc": 4},
    "WhatsApp Orders": {"imp": 400,  "ctr": 0.150, "conv": 0.060, "aov": 640, "cpo": 90},
    "Walk-in":         {"imp": 50,   "ctr": 0.500, "conv": 0.450, "aov": 980, "cpo": 110},
    "Paid Campaigns":  {"imp": 8500, "ctr": 0.028, "conv": 0.030, "aov": 790, "cpc": 9},
}
PRODUCTS = ["Atta 10kg", "Rice 25kg", "Mustard Oil 5L", "Sugar 5kg",
            "Tea 1kg", "Biscuits Family Pack", "Detergent 4kg", "Shampoo 650ml"]
CAMPAIGNS = {  # month -> campaign (15% of rows use Referral Boost instead)
    "2026-04": "Summer Refresh", "2026-05": "Summer Refresh",
    "2026-06": "Monsoon Sale", "2026-07": "Monsoon Sale",
    "2026-08": "Independence Fest", "2026-09": "Back to School",
    "2026-10": "Diwali Early-Bird",
}
# Festive build-up into Oct; Jun dips with the monsoon.
SEASON = {"2026-04": 1.00, "2026-05": 1.05, "2026-06": 0.95,
          "2026-07": 1.00, "2026-08": 1.08, "2026-09": 1.12, "2026-10": 1.35}
# Monthly revenue target = actual x factor (a mix of beats and misses).
TARGET_FACTOR = {"2026-04": 0.97, "2026-05": 1.03, "2026-06": 0.99,
                 "2026-07": 1.05, "2026-08": 0.96, "2026-09": 1.02,
                 "2026-10": 1.04}
START = dt.date(2026, 4, 1)
END = dt.date(2026, 10, 31)


def main():
    rng = random.Random(SEED)
    rows = []
    day = START
    while day <= END:
        month = day.strftime("%Y-%m")
        weekend = 1.15 if day.weekday() >= 5 else 1.0
        # Two channels report on a typical day, three on busy days.
        n_channels = 3 if rng.random() < 0.35 else 2
        for channel in rng.sample(sorted(CHANNELS), n_channels):
            cfg = CHANNELS[channel]
            noise = rng.uniform(0.85, 1.15)
            impressions = int(cfg["imp"] * SEASON[month] * weekend * noise)
            clicks = int(impressions * cfg["ctr"] * rng.uniform(0.9, 1.1))
            orders = int(clicks * cfg["conv"] * rng.uniform(0.9, 1.1))
            revenue = round(orders * cfg["aov"] * rng.uniform(0.92, 1.08), 2)
            if "cpc" in cfg:
                spend = round(clicks * cfg["cpc"] * rng.uniform(0.95, 1.05), 2)
            else:  # warm channel: cost per completed order
                spend = round(orders * cfg["cpo"] * rng.uniform(0.95, 1.05), 2)
            campaign = ("Referral Boost" if rng.random() < 0.15
                        else CAMPAIGNS[month])
            rows.append({
                "date": day.isoformat(),
                "channel": channel,
                "product": rng.choice(PRODUCTS),
                "campaign": campaign,
                "impressions": impressions,
                "clicks": clicks,
                "orders": orders,
                "revenue": revenue,
                "ad_spend": spend,
            })
        day += dt.timedelta(days=1)

    rows.sort(key=lambda r: (r["date"], r["channel"]))
    with open(os.path.join(DATA, "performance_daily.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    def agg(keyfn):
        out = {}
        for r in rows:
            k = keyfn(r)
            d = out.setdefault(k, {"impressions": 0, "clicks": 0, "orders": 0,
                                   "revenue": 0.0, "ad_spend": 0.0})
            d["impressions"] += r["impressions"]
            d["clicks"] += r["clicks"]
            d["orders"] += r["orders"]
            d["revenue"] += r["revenue"]
            d["ad_spend"] += r["ad_spend"]
        for d in out.values():
            d["revenue"] = round(d["revenue"], 2)
            d["ad_spend"] = round(d["ad_spend"], 2)
            d["ctr"] = d["clicks"] / d["impressions"] if d["impressions"] else 0.0
            d["cpc"] = d["ad_spend"] / d["clicks"] if d["clicks"] else 0.0
            d["roas"] = d["revenue"] / d["ad_spend"] if d["ad_spend"] else 0.0
            d["aov"] = d["revenue"] / d["orders"] if d["orders"] else 0.0
        return out

    months = agg(lambda r: r["date"][:7])
    channels = agg(lambda r: r["channel"])
    campaigns = agg(lambda r: r["campaign"])

    targets = {}
    for m, d in sorted(months.items()):
        targets[m] = round(d["revenue"] * TARGET_FACTOR[m], 2)
    with open(os.path.join(DATA, "targets.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "target_revenue"])
        for m in sorted(targets):
            w.writerow([m, f"{targets[m]:.2f}"])

    totals = {"impressions": sum(r["impressions"] for r in rows),
              "clicks": sum(r["clicks"] for r in rows),
              "orders": sum(r["orders"] for r in rows),
              "revenue": round(sum(r["revenue"] for r in rows), 2),
              "ad_spend": round(sum(r["ad_spend"] for r in rows), 2)}
    totals["ctr"] = totals["clicks"] / totals["impressions"]
    totals["cpc"] = totals["ad_spend"] / totals["clicks"]
    totals["roas"] = totals["revenue"] / totals["ad_spend"]
    totals["aov"] = totals["revenue"] / totals["orders"]

    focus = "2026-10"
    fm = dict(months[focus])
    fm["month"] = focus
    fm["target"] = targets[focus]
    fm["variance"] = round(fm["revenue"] - fm["target"], 2)
    fm["variance_pct"] = fm["variance"] / fm["target"]

    month_rows = []
    for m in sorted(months):
        d = dict(months[m])
        d["month"] = m
        d["target"] = targets[m]
        d["variance"] = round(d["revenue"] - d["target"], 2)
        d["variance_pct"] = d["variance"] / d["target"]
        month_rows.append(d)

    best_channel = max(channels, key=lambda c: channels[c]["revenue"])

    gt = {
        "rows": len(rows),
        "months": month_rows,
        "channels": [{"channel": c, **channels[c]} for c in sorted(channels)],
        "campaigns": [{"campaign": c, **campaigns[c]} for c in sorted(campaigns)],
        "totals": totals,
        "focus_month": fm,
        "best_channel": best_channel,
        "targets": targets,
    }
    with open(os.path.join(DATA, "ground_truth.json"), "w") as f:
        json.dump(gt, f, indent=2)

    print(f"rows={len(rows)} total_revenue={totals['revenue']:,.2f} "
          f"total_spend={totals['ad_spend']:,.2f} roas={totals['roas']:.2f} "
          f"focus_oct_revenue={fm['revenue']:,.2f} best={best_channel}")


if __name__ == "__main__":
    main()
