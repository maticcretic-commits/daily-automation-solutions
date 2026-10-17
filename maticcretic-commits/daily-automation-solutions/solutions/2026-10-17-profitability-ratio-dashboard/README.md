# ProfitPulse — Profitability & Financial-Ratio Dashboard Factory

A stdlib-plus-openpyxl generator that turns 5 years of financial statements
into a genuine, formula-driven Excel profitability dashboard: **22 ratios**,
**DuPont ROE decomposition**, **8 live red flags**, and a **KPI dashboard
with 3 native charts** — every computed figure a live formula, nothing
hard-coded.

## The client problem it solves

Ratio-analysis gigs ("profitability analysis report", "5-year ratio trends
with red-flag commentary", "Excel financial dashboards") usually arrive as a
pile of statements and a deadline. ProfitPulse takes raw P&L + balance-sheet
inputs and produces an investor-grade pack: statements, ratios, trend charts,
and a red-flag checklist — refreshable next year by pasting new numbers.

Reference: fictional company "Shakti Consumer Goods Pvt. Ltd." (FMCG),
FY21–FY25, INR lakh. All data is synthetic and deterministic.

## What's inside the workbook

| Sheet | Contents |
|---|---|
| `Guide` | Map of sheets, refresh instructions, ratio glossary, red-flag thresholds |
| `Raw_PNL` | P&L inputs FY21–FY25; gross / EBITDA / EBIT / EBT / net are formulas |
| `Raw_BS` | Balance-sheet inputs FY20–FY25; totals are formulas; **cash is the balancing plug** so A = L + E holds exactly |
| `Ratios` | 22 live ratios: liquidity (current/quick/cash, working capital), profitability (4 margins, ROE, ROA, ROCE), leverage (D/E, D/A, interest coverage, equity ratio), efficiency (receivable/inventory/payable days, CCC, asset turnover), growth (revenue & net-income YoY) |
| `DuPont` | ROE = net margin × asset turnover × equity multiplier, with a live tie-out to direct ROE |
| `Dashboard` | 8 FY25 KPI cards + FY21 deltas, 3 native charts (margin trends, revenue vs net income, working-capital days) |
| `AI_Report` | Rule-based narrative + 8-item red-flag checklist (all statuses live formulas, ⚠/✓ conditional formatting) |
| `Checks` | 8 integrity checks: balance-sheet balance, reserves rollforward, P&L foot, DuPont tie-out, cash > 0, margin bounds, working-capital identity — all must read TRUE / 0 |

## Run it

```bash
python3 generate.py        # builds data/*.csv (fictional, seeded)
python3 build_workbook.py  # writes profitpulse-ratio-dashboard.xlsx
python3 -m pytest tests/ -q
```

Needs `openpyxl` + `pytest`; the recalc test needs LibreOffice (`soffice`).

## Refresh for a real client

1. Overwrite the raw (non-formula) values in `Raw_PNL` / `Raw_BS`.
2. Everything else recalculates — keep the FY20 balance-sheet column (averages and FY21 ratios need it).

## Verification

`tests/test_build.py` (11 tests): workbook builds; all 8 sheets present;
every Ratios cell is a formula; invalid range literals and blank-guards are
rejected by regex scan; spot-checked formula wiring; **headless LibreOffice
recalc matched an independent Python ground truth on all 15 key ratios × 5
years** (current/quick ratio, margins, ROE/ROA/ROCE, D/E, interest coverage,
receivable/inventory/payable days, CCC, asset turnover); zero formula errors;
all 8 integrity checks pass post-recalc; Dashboard carries 3 native charts.

## Files

- `generate.py` — synthetic 5-year statements (cash as balance plug)
- `build_workbook.py` — the workbook factory (openpyxl)
- `tests/test_build.py` — build + wiring + LO-recalc tests
- `data/pnl.csv`, `data/bs.csv`, `data/meta.csv` — inputs
- `profitpulse-ratio-dashboard.xlsx` — the built workbook
