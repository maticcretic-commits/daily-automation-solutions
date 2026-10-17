#!/usr/bin/env python3
"""ProfitPulse — fictional company financials generator.

Builds 5 years (FY21–FY25) of consistent financial statements for the
fictional "Shakti Consumer Goods Pvt. Ltd." (FMCG manufacturer, INR lakh,
seeded/deterministic). Cash is the balance-sheet plug so A = L + E holds
exactly every year. Output: data/*.csv inputs for build_workbook.py.
"""
import csv, os
from pathlib import Path

SEED_COMPANY = "Shakti Consumer Goods Pvt. Ltd."
YEARS = ["FY21", "FY22", "FY23", "FY24", "FY25"]          # P&L / computed years
BS_YEARS = ["FY20"] + YEARS                               # balance sheet incl. opening

# ---- P&L inputs (INR lakh) -------------------------------------------------
REV   = [8450, 9240, 10180, 11090, 12150]
COGS  = [5220, 5650, 6180, 6720, 7360]
SGA   = [2010, 2140, 2290, 2430, 2610]
DEP   = [320, 345, 370, 395, 420]
INT   = [180, 195, 170, 150, 135]          # deleveraging story
TAX_RATE = 0.25

# ---- Balance-sheet raw items (INR lakh); FY20 included for averages -------
BS_RAW = {
    "Receivables":          [1290, 1410, 1560, 1780, 1980, 2240],
    "Inventory":            [1040, 1120, 1230, 1380, 1520, 1690],
    "Other current assets": [240, 260, 280, 300, 320, 340],
    "PPE (net)":            [2840, 2980, 3120, 3260, 3380, 3500],
    "Other non-current assets": [170, 180, 190, 200, 210, 220],
    "Payables":             [900, 980, 1080, 1190, 1300, 1420],
    "Short-term debt":      [500, 520, 560, 480, 420, 380],
    "Other current liabilities": [290, 310, 340, 370, 400, 430],
    "Long-term debt":       [1520, 1420, 1300, 1080, 880, 680],
    "Other non-current liabilities": [230, 240, 250, 260, 270, 280],
    "Share capital":        [2000, 2000, 2000, 2000, 2000, 2000],
}

SHARE_CAPITAL = 2000
RESERVES_FY20 = 2460  # = FY21 opening reserves


def pnl_series():
    ebitda = [r - c - s for r, c, s in zip(REV, COGS, SGA)]
    ebit = [e - d for e, d in zip(ebitda, DEP)]
    ebt = [e - i for e, i in zip(ebit, INT)]
    tax = [round(e * TAX_RATE) for e in ebt]
    net = [e - t for e, t in zip(ebt, tax)]
    return {
        "Revenue": REV, "COGS": COGS, "SG&A": SGA,
        "EBITDA": ebitda, "Depreciation": DEP, "EBIT": ebit,
        "Interest": INT, "EBT": ebt, "Tax": tax, "Net income": net,
    }


def reserves_series(net):
    out, run = [], RESERVES_FY20
    for n in net:
        run += n
        out.append(run)
    return [RESERVES_FY20] + out  # FY20..FY25


def write_csv(path, header, rows):
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def main():
    here = Path(__file__).resolve().parent
    data = here / "data"
    data.mkdir(exist_ok=True)

    pnl = pnl_series()
    net = pnl["Net income"]
    reserves = reserves_series(net)

    # --- pnl.csv : item, FY21..FY25 -----------------------------------------
    write_csv(data / "pnl.csv",
              ["Item"] + YEARS,
              [[k] + v for k, v in pnl.items()])

    # --- bs.csv : item, FY20..FY25 (Cash computed as plug) -------------------
    liab_cl = ["Payables", "Short-term debt", "Other current liabilities"]
    liab_ncl = ["Long-term debt", "Other non-current liabilities"]
    assets_noncash = ["Receivables", "Inventory", "Other current assets",
                      "PPE (net)", "Other non-current assets"]
    cash = []
    for i in range(6):
        tot_cl = sum(BS_RAW[k][i] for k in liab_cl)
        tot_ncl = sum(BS_RAW[k][i] for k in liab_ncl)
        equity = SHARE_CAPITAL + reserves[i]
        total_assets = tot_cl + tot_ncl + equity
        noncash = sum(BS_RAW[k][i] for k in assets_noncash)
        cash.append(round(total_assets - noncash, 2))
    assert all(c > 0 for c in cash), f"cash plug went non-positive: {cash}"
    rows = [["Cash & bank (plug)"] + cash]
    rows += [[k] + BS_RAW[k] for k in assets_noncash]
    rows += [[k] + BS_RAW[k] for k in liab_cl + liab_ncl]
    rows += [["Share capital"] + [SHARE_CAPITAL] * 6]
    rows += [["Reserves & surplus"] + reserves]
    write_csv(data / "bs.csv", ["Item"] + BS_YEARS, rows)

    # --- meta.csv ------------------------------------------------------------
    write_csv(data / "meta.csv",
              ["key", "value"],
              [["company", SEED_COMPANY],
               ["currency", "INR lakh"],
               ["years", ",".join(YEARS)],
               ["bs_years", ",".join(BS_YEARS)]])
    print("data written:", sorted(p.name for p in data.glob("*.csv")))
    print("FY25 net income (lakh):", net[-1], "| FY25 cash plug:", cash[-1])


if __name__ == "__main__":
    main()
