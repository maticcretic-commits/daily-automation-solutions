# ForecastForge — 3-Year Financial Forecast Model

A formula-driven 3-year financial forecast model for a fictional wholesale
FMCG distributor, **Kosi Traders Pvt. Ltd., Patna** (FY26 actuals,
FY27–FY29 forecast). Built as a working demo for the Freelancer project
**“3 Year Financial Forecast Model”** (ID 40821252, Bodhgaya client):
a flexible 3-year Excel forecast model — revenue, EBITDA, net income,
cash flow, assumptions sheet, scenario toggles — plus a concise dashboard
with charts.

## What it does

- **Assumptions sheet** — FY26 actuals and the opening balance-sheet
  position, a scenario table (Base / Optimistic / Pessimistic: revenue
  growth, gross margin, opex growth), and an active-scenario block that
  follows the Dashboard selector.
- **P&L** — Revenue, COGS, Gross Profit, four opex lines, EBITDA,
  Depreciation, Interest, PBT, Tax, Net Income. Every forecast cell is
  a live formula; FY26 is the actuals column.
- **CashFlow** — Net Income + Depreciation, working-capital moves
  (receivables 8% of revenue, inventory 6% of COGS, payables 10% of
  COGS), capex, debt repayment, opening/closing cash.
- **BalanceSheet** — Cash, Receivables, Inventory, Net Fixed Assets /
  Payables, Debt, Share Capital, Retained Earnings, with a live balance
  check (must be 0 in every year).
- **Dashboard** — scenario selector (B3), FY29 KPIs (revenue, EBITDA,
  net income, closing cash, EBITDA margin, revenue CAGR, 3-year total
  net income), a 4-year summary table, and two native charts (revenue
  vs EBITDA columns; closing-cash line).
- **Checks** — 10 integrity checks; all must read PASS in every
  scenario.

## How to use it

1. Open `forecast-model.xlsx`.
2. Pick a scenario on the Dashboard (cell **B3**). Everything updates.
3. Change any assumption on the Assumptions sheet; the P&L, cash flow,
   balance sheet and dashboard follow.
4. Read the Checks sheet after any change.

## What the scenarios show (Base, verified by recalc)

| | FY27 | FY28 | FY29 |
|---|---|---|---|
| Revenue (Rs) | 14,160,000 | 16,284,000 | 18,238,080 |
| EBITDA (Rs) | 492,800 | 775,720 | 960,766 |
| Net income (Rs) | −27,200 | 205,290 | 358,925 |
| Closing cash (Rs) | 1,877,888 | 1,508,181 | 1,317,349 |

**Pessimistic is the interesting one:** growth of 8/6/5% at a 28%
margin does not cover capex plus debt repayment, and the model goes
cash-negative in FY28 — a funding gap of roughly Rs 3.5 lakh. That is
the model doing its job, not a bug: switch the scenario and the gap
shows up on the Dashboard.

## Files

- `generate.py` — builds `forecast-model.xlsx` from `data/*.csv`
  (openpyxl; genuine OOXML, no pandas).
- `data/base_inputs.csv` — FY26 actuals, opening position, rates.
- `data/scenarios.csv` — the three scenarios.
- `forecast-model.xlsx` — the built workbook.
- `tests/test_build.py` — 1 test module: formula presence, LibreOffice
  headless recalc vs an independent Python model for the Base scenario
  **and** for Pessimistic (scenario switch via the Dashboard selector),
  zero formula errors, all 10 Checks PASS.

## Run

```bash
python3 generate.py
python3 -m pytest tests/ -v        # or: python3 tests/test_build.py
```

All data is fictional. Kosi Traders is invented; the FY26 figures are
plausible demo inputs, not a real company's books.
