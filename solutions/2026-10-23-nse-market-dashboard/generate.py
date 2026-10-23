#!/usr/bin/env python3
"""Generate deterministic SYNTHETIC NSE-style snapshot data for NSEPulse.

Two snapshots (T-1 = previous close, T = today) for 60 F&O-style tickers
across 8 sectors. All figures are fictional demo data — the workbook's Guide
sheet and README say so explicitly. Deterministic: same seed, same CSVs.
"""
import csv
import os
import random

SEED = 20261023
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

SECTORS = {
    "Banking": ["HDFCBANK", "ICICIBANK", "SBIN", "KOTAKBANK", "AXISBANK",
                "INDUSINDBK", "FEDERALBNK", "BANDHANBNK", "IDFCFIRSTB"],
    "IT": ["TCS", "INFY", "WIPRO", "HCLTECH", "TECHM", "LTIM", "MPHASIS", "COFORGE"],
    "Energy": ["RELIANCE", "ONGC", "NTPC", "POWERGRID", "BPCL", "IOC", "GAIL", "TATAPOWER"],
    "Auto": ["MARUTI", "TATAMOTORS", "M&M", "BAJAJ-AUTO", "HEROMOTOCO", "EICHERMOT", "ASHOKLEY"],
    "Pharma": ["SUNPHARMA", "DRREDDY", "CIPLA", "DIVISLAB", "BIOCON", "LUPIN", "AUROPHARMA"],
    "FMCG": ["HINDUNILVR", "ITC", "NESTLEIND", "BRITANNIA", "DABUR", "MARICO", "GODREJCP"],
    "Metals": ["TATASTEEL", "JSWSTEEL", "HINDALCO", "VEDL", "COALINDIA", "NMDC", "SAIL"],
    "Realty": ["DLF", "GODREJPROP", "OBEROIRLTY", "PRESTIGE", "SOBHA", "BRIGADE", "PHOENIXLTD"],
}


def generate():
    rng = random.Random(SEED)
    t1_rows, t_rows = [], []
    for sector, tickers in SECTORS.items():
        for ticker in tickers:
            base = round(10 ** rng.uniform(2.2, 3.65), 2)      # ~160 - 4500
            ltp_t1 = base
            ltp_t = round(base * (1 + rng.gauss(0, 0.02)), 2)   # daily move
            vol20 = int(rng.uniform(50_000, 8_000_000))         # 20-day avg volume
            volume = int(vol20 * rng.uniform(0.4, 2.8))         # today's volume
            oi_t1 = int(rng.uniform(1_000_000, 60_000_000))      # open interest
            oi_t = int(oi_t1 * (1 + rng.gauss(0, 0.06)))
            t1_rows.append([ticker, sector, f"{ltp_t1:.2f}", vol20, oi_t1])
            t_rows.append([ticker, sector, f"{ltp_t:.2f}", volume, oi_t])
    return t1_rows, t_rows


def write_csv(path, header, rows):
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    t1, t = generate()
    write_csv(os.path.join(OUT_DIR, "snapshots_t1.csv"),
              ["Ticker", "Sector", "LTP", "Vol20d_Avg", "OI"], t1)
    write_csv(os.path.join(OUT_DIR, "snapshots_t.csv"),
              ["Ticker", "Sector", "LTP", "Volume", "OI"], t)
    print(f"wrote {len(t1)} tickers x 2 snapshots -> {OUT_DIR}")


if __name__ == "__main__":
    main()
