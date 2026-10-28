#!/usr/bin/env python3
"""Generate deterministic synthetic data for the StatementPulse demo.

Fictional company: Sharma Retail Pvt. Ltd., Patna (fashion/home retail).
Three financial years (FY2024-FY2026) plus monthly FY2026 detail.
All figures are SYNTHETIC, labelled as such in the workbook; nothing here
is client data. Seeded (20261028) so re-runs produce identical CSVs.

Construction (accounting identities hold by design):
- P&L: Revenue -> COGS -> Gross profit -> Opex -> EBITDA -> Depreciation
  -> EBIT -> Interest -> PBT -> Tax (25%) -> PAT.
- Balance sheet: receivables/inventory/payables from days ratios;
  PPE rolls forward (opening + capex - depreciation); retained earnings
  roll forward (opening + PAT - dividends); long-term borrowings are a
  chosen funding schedule; CASH is the balancing plug
  (Total L+E minus non-cash assets), so the balance sheet balances
  exactly up to rupee rounding on the days-based items.
- Cash flow (indirect): CFO = PAT + depreciation - change in working
  capital; CFI = -capex; CFF = change in borrowings - dividends.
  By the balance-sheet identity, CFO+CFI+CFF equals the change in the
  cash plug (within a few rupees of rounding); the workbook Checks
  sheet tolerates < Rs 100 on that tie-out.
- Monthly FY2026 revenue/COGS/opex are allocated from the annual
  totals (largest-remainder apportionment), so monthly sums tie
  exactly to the annual figures. March month-end balances equal the
  FY2026 balance-sheet values exactly.
"""
import csv
import datetime
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
SEED = 20261028

YEARS = ["FY2024", "FY2025", "FY2026"]
REVENUE = {"FY2024": 42000000, "FY2025": 49560000, "FY2026": 60463200}
COGS_RATIO = {"FY2024": 0.62, "FY2025": 0.605, "FY2026": 0.59}
OPEX_RATIO = {"FY2024": 0.22, "FY2025": 0.215, "FY2026": 0.21}
DEPRECIATION = {"FY2024": 1800000, "FY2025": 2100000, "FY2026": 2400000}
INTEREST = {"FY2024": 2200000, "FY2025": 2500000, "FY2026": 2300000}
TAX_RATE = 0.25
CAPEX = {"FY2024": 3000000, "FY2025": 3500000, "FY2026": 2800000}
DIVIDENDS = {"FY2024": 1000000, "FY2025": 1200000, "FY2026": 1500000}
RECV_DAYS = {"FY2024": 48, "FY2025": 52, "FY2026": 58}
INV_DAYS = {"FY2024": 42, "FY2025": 45, "FY2026": 50}
PAY_DAYS = {"FY2024": 38, "FY2025": 40, "FY2026": 36}
ST_DEBT = {"FY2024": 4000000, "FY2025": 5000000, "FY2026": 3500000}
LT_DEBT = {"FY2024": 12000000, "FY2025": 13500000, "FY2026": 11000000}
SHARE_CAPITAL = 10000000
OTHER_CA = {"FY2024": 800000, "FY2025": 900000, "FY2026": 1000000}
OTHER_NCA = 1500000
OTHER_CL = {"FY2024": 600000, "FY2025": 700000, "FY2026": 650000}

# Opening (start of FY2024) balance-sheet values.
OPENING = {
    "Trade receivables": 5000000,
    "Inventories": 6000000,
    "Other current assets": 700000,
    "Property plant and equipment (net)": 14000000,
    "Other non-current assets": OTHER_NCA,
    "Trade payables": 4000000,
    "Short-term borrowings": 3500000,
    "Other current liabilities": 500000,
    "Long-term borrowings": 11000000,
    "Share capital": SHARE_CAPITAL,
    "Retained earnings": 6000000,
}


def allocate(total, weights):
    """Split integer total across weights; parts sum exactly to total."""
    raw = [total * w / sum(weights) for w in weights]
    parts = [int(x) for x in raw]
    remainder = total - sum(parts)
    order = sorted(range(len(raw)), key=lambda i: raw[i] - parts[i], reverse=True)
    i = 0
    while remainder > 0:
        parts[order[i % len(order)]] += 1
        remainder -= 1
        i += 1
    return parts


def main():
    rng = random.Random(SEED)
    pl, bs = {}, {}
    prev = dict(OPENING)
    prev_re = OPENING["Retained earnings"]
    prev_ppe = OPENING["Property plant and equipment (net)"]
    for year in YEARS:
        rev = REVENUE[year]
        cogs = round(rev * COGS_RATIO[year])
        opex = round(rev * OPEX_RATIO[year])
        dep = DEPRECIATION[year]
        interest = INTEREST[year]
        gross = rev - cogs
        ebitda = gross - opex
        ebit = ebitda - dep
        pbt = ebit - interest
        tax = round(pbt * TAX_RATE)
        pat = pbt - tax
        pl[year] = {
            "Revenue": rev, "Cost of goods sold": cogs,
            "Operating expenses": opex, "Depreciation": dep,
            "Interest expense": interest, "Tax expense": tax,
        }
        recv = round(rev * RECV_DAYS[year] / 365)
        inv = round(cogs * INV_DAYS[year] / 365)
        pay = round(cogs * PAY_DAYS[year] / 365)
        ppe = prev_ppe + CAPEX[year] - dep
        re_ = prev_re + pat - DIVIDENDS[year]
        non_cash_assets = recv + inv + OTHER_CA[year] + ppe + OTHER_NCA
        total_le = (pay + ST_DEBT[year] + OTHER_CL[year] + LT_DEBT[year]
                    + SHARE_CAPITAL + re_)
        cash = total_le - non_cash_assets
        assert cash > 0, f"cash plug non-positive in {year}: {cash}"
        bs[year] = {
            "Cash and bank": cash, "Trade receivables": recv,
            "Inventories": inv, "Other current assets": OTHER_CA[year],
            "Property plant and equipment (net)": ppe,
            "Other non-current assets": OTHER_NCA,
            "Trade payables": pay, "Short-term borrowings": ST_DEBT[year],
            "Other current liabilities": OTHER_CL[year],
            "Long-term borrowings": LT_DEBT[year],
            "Share capital": SHARE_CAPITAL, "Retained earnings": re_,
        }
        prev_re, prev_ppe = re_, ppe

    # Cash-flow check: indirect-method net change vs cash-plug movement.
    prev_bs = dict(OPENING)
    opening_cash = (sum(v for k, v in OPENING.items()
                        if k in ("Trade payables", "Short-term borrowings",
                                 "Other current liabilities",
                                 "Long-term borrowings", "Share capital",
                                 "Retained earnings"))
                    - sum(v for k, v in OPENING.items()
                          if k in ("Trade receivables", "Inventories",
                                   "Other current assets",
                                   "Property plant and equipment (net)",
                                   "Other non-current assets")))
    prev_cash = opening_cash
    cf = {}
    for year in YEARS:
        b, p = bs[year], pl[year]
        pat = (p["Revenue"] - p["Cost of goods sold"] - p["Operating expenses"]
               - p["Depreciation"] - p["Interest expense"] - p["Tax expense"])
        dwc = ((b["Trade receivables"] - prev_bs["Trade receivables"])
               + (b["Inventories"] - prev_bs["Inventories"])
               - (b["Trade payables"] - prev_bs["Trade payables"])
               + (b["Other current assets"] - prev_bs["Other current assets"])
               - (b["Other current liabilities"] - prev_bs["Other current liabilities"]))
        cfo = pat + p["Depreciation"] - dwc
        cfi = -CAPEX[year]
        cff = ((b["Short-term borrowings"] - prev_bs["Short-term borrowings"])
               + (b["Long-term borrowings"] - prev_bs["Long-term borrowings"])
               - DIVIDENDS[year])
        net = cfo + cfi + cff
        cf[year] = {"CFO": cfo, "CFI": cfi, "CFF": cff, "Net change": net,
                    "Opening cash": prev_cash,
                    "Closing cash": prev_cash + net}
        print(f"{year}: cash plug {b['Cash and bank']:,} vs CF closing "
              f"{prev_cash + net:,} (diff {b['Cash and bank'] - (prev_cash + net):,})")
        prev_bs = dict(b)
        prev_cash = b["Cash and bank"]

    # ---- Write statements CSV (long format) ----
    os.makedirs(DATA, exist_ok=True)
    rows = []
    for year in YEARS:
        for item, amount in pl[year].items():
            rows.append(["Profit & loss", item, year, amount])
        for item, amount in bs[year].items():
            rows.append(["Balance sheet", item, year, amount])
        rows.append(["Cash flow", "Capital expenditure", year, CAPEX[year]])
        rows.append(["Cash flow", "Dividends paid", year, DIVIDENDS[year]])
    for item, amount in OPENING.items():
        rows.append(["Balance sheet", item, "Opening", amount])
    # Opening cash is the plug of the opening balance sheet (computed
    # above); the workbook's FY2024 opening-cash formula reads it here.
    rows.append(["Balance sheet", "Cash and bank", "Opening", opening_cash])
    with open(os.path.join(DATA, "statements.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Section", "LineItem", "Year", "Amount"])
        w.writerows(rows)

    # ---- Monthly FY2026 detail ----
    months = []
    day = datetime.date(2025, 4, 1)
    while day <= datetime.date(2026, 3, 1):
        months.append(day.strftime("%Y-%m"))
        day = (day.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
    # Seasonal weights: festive Oct/Nov uplift, lean Feb.
    weights = [1.0, 1.0, 1.05, 0.95, 1.0, 1.05, 1.25, 1.30, 1.10, 0.95, 0.85, 1.05]
    revs = allocate(REVENUE["FY2026"], weights)
    cogss = allocate(round(REVENUE["FY2026"] * COGS_RATIO["FY2026"]), weights)
    opexes = allocate(round(REVENUE["FY2026"] * OPEX_RATIO["FY2026"]), weights)
    b26 = bs["FY2026"]
    mrows = []
    for i, month in enumerate(months):
        frac = (i + 1) / 12
        if month == "2026-03":
            recv, inv, pay = (b26["Trade receivables"], b26["Inventories"],
                              b26["Trade payables"])
        else:
            recv = round(b26["Trade receivables"] * (0.72 + 0.28 * frac)
                         * rng.uniform(0.96, 1.04))
            inv = round(b26["Inventories"] * (0.75 + 0.25 * frac)
                        * rng.uniform(0.96, 1.04))
            pay = round(b26["Trade payables"] * (0.80 + 0.20 * frac)
                        * rng.uniform(0.96, 1.04))
        mrows.append([month, revs[i], cogss[i], opexes[i], recv, inv, pay])
    with open(os.path.join(DATA, "monthly.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Month", "Revenue", "COGS", "Opex", "Receivables",
                    "Inventory", "Payables"])
        w.writerows(mrows)
    print(f"statement rows: {len(rows)}; monthly rows: {len(mrows)}")
    print(f"FY2026 PAT: {pat:,} (last computed year)")


if __name__ == "__main__":
    main()
