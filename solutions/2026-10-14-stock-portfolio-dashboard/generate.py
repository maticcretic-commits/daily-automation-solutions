#!/usr/bin/env python3
"""StockPulse data generator — fictional 30-stock Indian portfolio, seeded/deterministic.

Outputs (under data/):
  companies.csv  ticker,company,sector
  holdings.csv   ticker,qty,avg_cost
  prices.csv     ticker,w0,w1,...,w32   (33 weekly closes)
  weeks.csv      week_index,week_label,date
  index.csv      week_index,nifty_like   (fictional benchmark)

All values are FICTIONAL and generated from a fixed seed — a demo, not market data.
"""
import csv
import os
import random
from datetime import date, timedelta

SEED = 20261014
N_WEEKS = 33  # W0..W32

COMPANIES = [
    # (ticker, company, sector, start_price)
    ("HDFCBANK", "HDFC Bank Ltd", "Banking", 1620.0),
    ("ICICIBANK", "ICICI Bank Ltd", "Banking", 1180.0),
    ("SBIN", "State Bank of India", "Banking", 810.0),
    ("KOTAKBANK", "Kotak Mahindra Bank", "Banking", 1960.0),
    ("AXISBANK", "Axis Bank Ltd", "Banking", 1185.0),
    ("INFY", "Infosys Ltd", "IT", 1840.0),
    ("TCS", "Tata Consultancy Services", "IT", 4150.0),
    ("HCLTECH", "HCL Technologies Ltd", "IT", 1780.0),
    ("TECHM", "Tech Mahindra Ltd", "IT", 1560.0),
    ("WIPRO", "Wipro Ltd", "IT", 590.0),
    ("RELIANCE", "Reliance Industries Ltd", "Energy", 2980.0),
    ("ONGC", "Oil & Natural Gas Corp", "Energy", 285.0),
    ("NTPC", "NTPC Ltd", "Energy", 395.0),
    ("POWERGRID", "Power Grid Corp", "Energy", 320.0),
    ("MARUTI", "Maruti Suzuki India Ltd", "Auto", 12800.0),
    ("TATAMOTORS", "Tata Motors Ltd", "Auto", 980.0),
    ("M&M", "Mahindra & Mahindra Ltd", "Auto", 2850.0),
    ("HEROMOTOCO", "Hero MotoCorp Ltd", "Auto", 5230.0),
    ("HINDUNILVR", "Hindustan Unilever Ltd", "FMCG", 2560.0),
    ("ITC", "ITC Ltd", "FMCG", 505.0),
    ("NESTLEIND", "Nestle India Ltd", "FMCG", 2450.0),
    ("DABUR", "Dabur India Ltd", "FMCG", 620.0),
    ("SUNPHARMA", "Sun Pharmaceutical Ind", "Pharma", 1740.0),
    ("DRREDDY", "Dr Reddy's Laboratories", "Pharma", 1280.0),
    ("CIPLA", "Cipla Ltd", "Pharma", 1540.0),
    ("ULTRACEMCO", "UltraTech Cement Ltd", "Cement & Infra", 11500.0),
    ("LT", "Larsen & Toubro Ltd", "Cement & Infra", 3640.0),
    ("TATASTEEL", "Tata Steel Ltd", "Metals", 168.0),
    ("HINDALCO", "Hindalco Industries Ltd", "Metals", 690.0),
    ("TITAN", "Titan Company Ltd", "Consumer Goods", 3480.0),
]

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


def main():
    rng = random.Random(SEED)
    os.makedirs(DATA_DIR, exist_ok=True)

    # --- companies ---
    with open(os.path.join(DATA_DIR, "companies.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ticker", "company", "sector"])
        for t, c, s, _ in COMPANIES:
            w.writerow([t, c, s])

    # --- weekly prices: seeded geometric random walk, 33 closes ---
    start = date(2026, 3, 7)
    prices = {}
    for t, _c, _s, p0 in COMPANIES:
        drift = rng.uniform(-0.002, 0.004)
        vol = rng.uniform(0.015, 0.045)
        series = [p0]
        for _ in range(1, N_WEEKS):
            shock = rng.gauss(0, 1) * vol
            series.append(round(series[-1] * (1 + drift + shock), 2))
        prices[t] = series
    with open(os.path.join(DATA_DIR, "prices.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ticker"] + [f"w{i}" for i in range(N_WEEKS)])
        for t, _c, _s, _p in COMPANIES:
            w.writerow([t] + prices[t])

    # --- holdings: qty + avg cost (seeded around the start price) ---
    holdings = {}
    with open(os.path.join(DATA_DIR, "holdings.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ticker", "qty", "avg_cost"])
        for t, _c, _s, p0 in COMPANIES:
            qty = rng.randint(2, 40) if p0 > 3000 else rng.randint(10, 220)
            avg_cost = round(p0 * rng.uniform(0.84, 1.12), 2)
            holdings[t] = (qty, avg_cost)
            w.writerow([t, qty, avg_cost])

    # --- weeks + fictional benchmark ---
    bench = 24000.0
    with open(os.path.join(DATA_DIR, "weeks.csv"), "w", newline="") as f, \
         open(os.path.join(DATA_DIR, "index.csv"), "w", newline="") as g:
        w = csv.writer(f)
        gwriter = csv.writer(g)
        w.writerow(["week_index", "week_label", "date"])
        gwriter.writerow(["week_index", "nifty_like"])
        for i in range(N_WEEKS):
            d = start + timedelta(weeks=i)
            w.writerow([i, f"W{i}", d.isoformat()])
            gwriter.writerow([i, round(bench, 2)])
            bench = round(bench * (1 + rng.gauss(0.0018, 0.014)), 2)

    total_invested = sum(q * c for q, c in holdings.values())
    print(f"30 tickers, {N_WEEKS} weekly closes each, invested total = Rs {total_invested:,.2f}")
    print("seed:", SEED)


if __name__ == "__main__":
    main()
