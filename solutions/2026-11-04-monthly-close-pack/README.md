# MonthlyClosePack — Monthly Bookkeeping & Reporting Pack

Built around a real Freelancer brief: turn a month's bookkeeping into a
reconciled monthly reporting pack — profit & loss, balance sheet and
cash-flow statements, in Excel, with every figure formula-linked to the
underlying journal so the pack ties out by construction. One `.xlsx`,
no macros, nothing that triggers a macro warning.

Demo data is fictional: Deshmukh Trading Co., a Pune FMCG wholesaler,
Apr–Oct 2026, 1,542 double-entry journal lines across 540 balanced
entries and 21 accounts, as of 4 Nov 2026. The generator is seeded and
deterministic — re-running it reproduces the same data byte for byte.

## What the workbook does

- **Journal** — the entry table. Columns A–E and G–H are typed/pasted
  values (entry ID, date, account, debit, credit, category,
  narration). Column F (Month Key) is a formula, pre-filled for 150
  headroom rows past the data (to row 1693): paste a new balanced
  entry inside the headroom and every statement updates; blank rows
  stay blank, never errors.
- **TrialBalance** — one row per account: opening debit/credit,
  debit/credit movements via `SUMIF` over the Journal, and closing
  debit/credit via a signed-balance `MAX` split. Totals must match:
  Rs 17,729,300 debit = Rs 17,729,300 credit.
- **IncomeStatement** — revenue, cost of goods sold, gross profit,
  six operating-expense lines and net income, each a live `SUMIFS`
  off the trial balance.
- **BalanceSheet** — assets, liabilities and equity, with retained
  earnings carried as opening retained earnings plus this period's
  net income, less drawings. Balance check reads 0.
- **CashFlow** — indirect method: net income plus depreciation, working
  capital movements, equipment purchases, loan/capital/drawings
  movements. Ties exactly to the Rs 3,991,000 change in Cash.
- **Monthly** — revenue / COGS / gross / opex / net by month, plus
  journal-line counts and cumulative net income.
- **Budget** — monthly sales and net-income budgets (typed values)
  against actuals, with variance and variance %.
- **Dashboard** — 13 KPIs, an expense-mix table, and 3 native Excel
  charts: revenue vs COGS by month, operating expenses by type,
  cumulative net income.
- **Checks** — 12 integrity checks (journal balances, trial balance
  balances, opening balances, P&L tie-out, balance-sheet balance,
  cash-flow tie-out, monthly/budget tie-outs, non-negative cash, net
  movements sum to zero). All read PASS.
- **Assumptions** — the as-of date (4 Nov 2026) and period.
- **Guide** — how to add entries and read the pack, in the workbook.

## Verified figures (LibreOffice headless recalc vs independent Python)

- Revenue Rs 13,711,300 · COGS Rs 8,341,100 · gross profit Rs 5,370,200
- Operating expenses Rs 1,766,800 · net income Rs 3,603,400 (26.3% margin)
- Total assets Rs 7,333,900 = liabilities Rs 2,705,500 + equity Rs 4,628,400
- Cash Rs 4,441,000 (up Rs 3,991,000): operating +Rs 4,061,000,
  investing −Rs 180,000, financing +Rs 110,000
- Monthly net income: Apr 513,900 · May 453,100 · Jun 439,900 ·
  Jul 517,400 · Aug 624,600 · Sep 540,200 · Oct 514,300
- Zero formula errors in the recalculated workbook; all 12 Checks PASS.

## Files

- `generate.py` — deterministic synthetic data → `data/journal.csv`
  (every entry asserted balanced; cash/GST/inventory sanity-guarded)
- `build_workbook.py` — builds `monthly-close-pack.xlsx` (openpyxl,
  live A1-style formulas only, no Table objects, no macros)
- `main.py` — independent ground truth recomputed from the CSVs
- `tests/test_build.py` — builds, recalcs in LibreOffice, and compares
  every checked figure; also scans the builder for short-form range
  literals (`$H$2:$5000`) and asserts headroom formulas and 3 charts
  are present
- `powerbi/MEASURES.md` — DAX mapping for rebuilding the same model
  as a `.pbix`
- `data/accounts.csv` — chart of accounts with opening balances
  (opening trial balance: Rs 1,580,000 debit = Rs 1,580,000 credit)
- `data/budget.csv` — monthly sales and net-income budgets
- `data/journal.csv` — the 1,542-line source journal
- `monthly-close-pack.xlsx` — the finished workbook

## How to run

```sh
python3 generate.py          # regenerate data/journal.csv (seeded)
python3 build_workbook.py    # build monthly-close-pack.xlsx
python3 tests/test_build.py  # full verification (needs LibreOffice)
```

Requires Python 3.11+ and `openpyxl`. LibreOffice is needed only for
the recalc verification step.

## Reference

Freelancer brief: "Monthly Bookkeeping & Reporting" (Project ID
40801320, Pune) — monthly P&L, balance sheet and cash-flow report pack
in Excel/PDF with reconciliation, posted ~7 hours before the 4 Nov
2026 scout, 1 proposal at scout time. This is a reference
implementation built from fictional data, not client work. The
project did not pass the visibility gate at scout time (posted over 4
hours before the scout), so no bid was drafted for it.
