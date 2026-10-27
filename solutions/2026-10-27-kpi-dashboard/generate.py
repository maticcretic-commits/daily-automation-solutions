#!/usr/bin/env python3
"""Generate deterministic synthetic data for the KPI dashboard demo.

Fictional company: Nagpur Retail Co. (fashion/home retail, Nagpur, India).
Daily rows, Jan-Sep 2026, three channels. All figures are SYNTHETIC and
labelled as such in the workbook; nothing here is client data.
Seeded (20261027) so re-runs produce identical CSVs.
"""
import csv
import datetime
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
SEED = 20261027

CHANNELS = {
    "Online": {"unit_price": 850, "margin": 0.34},
    "Store": {"unit_price": 1200, "margin": 0.30},
    "Wholesale": {"unit_price": 950, "margin": 0.26},
}

# Monthly revenue targets (INR) -- deliberately ambitious in festival months.
TARGETS = {
    "2026-01": 4200000, "2026-02": 3900000, "2026-03": 4500000,
    "2026-04": 4300000, "2026-05": 4700000, "2026-06": 4600000,
    "2026-07": 5000000, "2026-08": 5400000, "2026-09": 5800000,
}
ORDERS_TARGETS = {
    "2026-01": 4800, "2026-02": 4500, "2026-03": 5100,
    "2026-04": 4900, "2026-05": 5300, "2026-06": 5200,
    "2026-07": 5600, "2026-08": 6000, "2026-09": 6400,
}


def main():
    rng = random.Random(SEED)
    rows = []
    day = datetime.date(2026, 1, 1)
    end = datetime.date(2026, 9, 30)
    while day <= end:
        # Weekend uplift for Store, month-end push for Wholesale.
        for channel, spec in CHANNELS.items():
            base = rng.randint(6, 30)
            if channel == "Store" and day.weekday() >= 5:
                base = int(base * 1.4)
            if channel == "Wholesale" and day.day >= 26:
                base = int(base * 1.3)
            orders = base
            price = spec["unit_price"] * rng.uniform(0.92, 1.10)
            revenue = round(orders * price)
            cost = round(revenue * (1 - spec["margin"]) * rng.uniform(0.97, 1.03))
            ontime = int(round(orders * rng.uniform(0.88, 0.99)))
            csat_sum = round(orders * rng.uniform(4.1, 4.8), 2)
            rows.append([
                day.isoformat(), day.strftime("%Y-%m"), channel, orders,
                revenue, cost, ontime, csat_sum,
            ])
        day += datetime.timedelta(days=1)

    os.makedirs(DATA, exist_ok=True)
    with open(os.path.join(DATA, "raw_transactions.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Month", "Channel", "Orders", "Revenue",
                    "Cost", "OnTimeOrders", "CSATSum"])
        w.writerows(rows)
    with open(os.path.join(DATA, "monthly_targets.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Month", "RevenueTarget", "OrdersTarget"])
        for month in sorted(TARGETS):
            w.writerow([month, TARGETS[month], ORDERS_TARGETS[month]])
    print(f"raw rows: {len(rows)} (plus header)")
    print(f"months: {len(TARGETS)}")


if __name__ == "__main__":
    main()
