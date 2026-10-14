# StockPulse — AI-Enabled Stock Portfolio Excel Dashboard

**Demo for:** *"AI Excel Dashboard Developer for Stock Analysis"* — an AI-enabled,
user-friendly Excel dashboard tracking 30 Indian stock holdings with weekly AI reporting.
Built as a working reference implementation (fictional data), not client work.

## What it does

`generate.py` builds fictional, seeded, deterministic market data for a 30-stock
Indian portfolio (HDFC Bank … Titan, 9 sectors, 33 weekly closes W0–W32 plus a
fictional benchmark). `build_workbook.py` turns it into a genuine 8-sheet
OOXML `.xlsx` where **every figure is a live formula** (plain A1 ranges, no Table
objects, no macros):

| Sheet | Contents |
|---|---|
| Guide | What it is, data notes, refresh procedure, handover |
| Holdings | 30 positions — qty, avg cost, last price via `INDEX/MATCH`, invested, market value, P&L, P&L %, allocation % (all formulas) |
| Prices | 33 weekly closes per stock + last-week and 4-week return formula columns |
| Weekly | Portfolio value per week (`SUMPRODUCT` of holdings × prices), weekly return, drawdown, benchmark |
| Sector | Sector allocation, P&L and P&L % via `SUMIF`/`COUNTIF` |
| Dashboard | 13 live KPI cards (invested, MV, P&L, weekly return, best/worst performer, largest holding, positive weeks) + 3 native charts: portfolio trend (line), sector allocation (pie), P&L % by holding (bar) |
| AI_Report | Weekly report written by a **rule-based generator** from the workbook's own numbers — movers, allocation drift, benchmark comparison, watchlist alerts (demo of the AI weekly-reporting step; production = scheduled LLM call over the same KPI block) |
| Checks | 8 integrity checks (allocation = 100 %, MV reconciliation, P&L = MV − invested, price lookups, rollup ties, no negative qty) — all PASS |

## Run it

```bash
python3 generate.py          # data/*.csv  (seed 20261014 — deterministic)
python3 build_workbook.py    # stock-portfolio-dashboard.xlsx
python3 -m unittest discover -s tests   # 20/20 incl. LibreOffice headless recalc
```

Requires Python 3 + `openpyxl`; LibreOffice needed only for the recalc test.
To go live on real data: replace `Prices` closes and `Holdings` qty/avg-cost with
yours, repoint the Last-Price column, add a week row on `Weekly` — everything
else recalculates, and the report generator re-runs on the new numbers.

**All data is fictional (seeded random walk). Not market data, not advice.**
