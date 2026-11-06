# StraticaMIS — 3-Statement Financial Model + MIS Dashboard

A working demo built around a real Freelancer brief: **"Stratica OS –
Financial Planning, Budgeting & Investor-Ready Financial Model"**
(Freelancer project 40804245, posted 5 Nov 2026, Hyderabad) — *"Develop
financial dashboards, MIS reports"* plus a 3-statement model with budgets,
forecasts and cash-flow management.

**Demo company:** Shree Traders, a fictional wholesale trading SME.
**Period:** Apr 2025 – Mar 2026 (FY 2026-27), monthly.

## What it does

- **Inputs** — one place to paste monthly revenue, COGS, five opex lines
  and capex. Everything else is a live formula off this sheet.
- **PL** — monthly profit & loss: gross profit, EBITDA, depreciation,
  interest, profit before tax, tax (25% of positive PBT), net profit,
  plus an FY total row.
- **BS** — monthly balance sheet with opening balances at 31 Mar 2025
  (cash ₹12,00,000; receivables ₹18,00,000; inventory ₹15,00,000;
  fixed assets ₹40,00,000; payables ₹10,00,000; term loan ₹25,00,000;
  capital ₹30,00,000; retained earnings ₹20,00,000).
- **CF** — cash-flow statement by the indirect method, so the balance
  sheet balances by construction.
- **Budget** — budget vs actual for revenue, gross profit, EBITDA and
  net profit, with rupee and % variances. Budget: revenue from
  ₹19,00,000 growing 4% a month, 45% gross margin, ₹5,20,000/month
  operating expense.
- **Dashboard** — 10 KPIs plus three native charts (revenue budget vs
  actual, EBITDA trend, closing cash balance).
- **Checks** — 8 self-tests, all PASS (balance sheet balances every
  month, cash ties between BS and CF, retained earnings move by net
  profit, FY totals agree, tax never negative, cash always positive,
  loan and fixed-asset closing balances agree with the schedule).

## Model rules

- Receivables close at 50% of monthly revenue (15 days), inventory at
  80% of monthly COGS (24 days), payables at 45% of monthly COGS.
- Depreciation is a flat ₹60,000/month; the term loan repays
  ₹1,00,000/month at 1% monthly interest on the opening balance.
- The model carries no overdraft or short-term borrowings, so the
  current ratio runs high — noted on the Guide sheet, not hidden.

## Verified results (FY 2026-27, demo data)

- Revenue ₹2,72,64,000 · Net profit ₹24,76,500 · Net margin 9.1% ·
  EBITDA margin 15.6%
- Closing cash ₹27,06,250 · Current ratio 7.63 · Receivables 15.0 days ·
  Inventory 24.0 days
- Revenue budget attainment 95.5% · Operating cash flow ₹35,06,250

Verified by `tests/test_build.py`: the generator rebuilds the workbook,
every formula is lint-checked (no short-form ranges, no structured
references), and a LibreOffice headless recalc is compared cell-by-cell
against an independent Python ground truth — zero formula errors, 8/8
checks PASS.

## How to run it

```bash
python3 generate.py            # rebuilds stratica-mis-model.xlsx + data/inputs.csv
python3 tests/test_build.py    # rebuilds, recalcs in LibreOffice, checks vs ground truth
```

Requires Python 3 with `openpyxl`; LibreOffice (`soffice`) for the
recalc test. No macros, no external data. Demo data is fictional and
deterministic (seed 20261106) — this is a reporting template, not
financial advice.

## Files

- `generate.py` — builds the workbook
- `data/inputs.csv` — the 12 months of demo inputs
- `tests/test_build.py` — generator + recalc verification
- `stratica-mis-model.xlsx` — the finished workbook
