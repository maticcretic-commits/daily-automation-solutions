# Monthly Close Pack — demo

A monthly financial reporting pack for a service business, built around
a real Freelancer brief: *"Monthly Financial Reports Preparation"* —
pull Excel/Google Sheets data each month, verify accuracy, and deliver
a P&L, a Balance Sheet with an integrated cash-flow view, and a
budget-versus-actual report by the 5th business day, with every figure
transparent and auditable in Excel.

This is a **demonstration built from invented data** (a fictional
business, "Verma Consulting Services Pvt. Ltd., Vadodara", Jan–Oct
2026), not client work. It shows the exact pack the brief asks for,
with the reconciliation discipline the brief cares about: the balance
sheet balances, the cash-flow walk ties to the balance-sheet cash
movement, and every number traces to a raw transaction or a budget row.

## What it contains

`monthly-close-pack.xlsx` — 10 sheets, every figure a live formula:

- **Guide** — how to run the monthly close with this pack.
- **Dashboard** — 8 KPIs plus 3 native charts (revenue vs budget by
  month, cumulative net income, expense mix).
- **PL_Monthly** — monthly P&L, every line a `SUMIFS` over
  `Raw_Transactions`; depreciation from Assumptions.
- **Balance_Sheet** — opening / movement / closing, with a balance
  check that must read 0. Receivables close at 20% of October revenue
  (uncollected); payables at 10% of October expenses (unpaid).
- **CashFlow** — indirect walk (net income + depreciation − change in
  receivables + change in payables − equipment − drawings) tying
  exactly to the balance-sheet cash movement.
- **Budget_vs_Actual** — 12 categories, actual vs budget, variance and
  variance %, with an ON TRACK / OVER BUDGET / UNDER BUDGET status.
- **Raw_Transactions** (361 rows), **Raw_Budget** (120 rows) —
  the only hard-coded inputs, with AutoFilter and frozen panes.
- **Assumptions** — opening balances, depreciation, purchases,
  drawings, and the receivables/payables shares.
- **Checks** — 10 checks, all PASS after a headless LibreOffice recalc.

Headline demo figures: revenue Rs 6,021,016 · net income
Rs 1,475,921 (24.5% margin) · cash Rs 2,073,319 · balance gap 0.

## How to run it

```sh
python3 generate.py          # rebuild data/*.csv (seeded, deterministic)
python3 build_workbook.py    # rebuild monthly-close-pack.xlsx
python3 tests/test_build.py  # rebuild + LibreOffice recalc + ground-truth checks
```

Requirements: Python 3 + openpyxl; LibreOffice (`soffice`) for the
recalc test only. `main.py` is the independent Python ground truth the
tests compare against — it reads the same CSVs and re-derives every
checked figure without touching the workbook.

## Monthly close routine (the deliverable's real value)

1. Drop the month's transactions into `Raw_Transactions` (same columns)
   and the month's budgets into `Raw_Budget`; extend the `SUMIFS`
   ranges to cover the new rows.
2. Update the Assumptions shares only if the collection/payment
   pattern genuinely changed.
3. Open the workbook: Dashboard, P&L, balance sheet, cash flow and
   budget-vs-actual all recompute. Read **Checks** — all 10 must PASS
   before the pack goes to leadership.
4. Write the variance note from `Budget_vs_Actual` (the Status column
   flags what needs explaining) and deliver by the 5th business day.

## Power BI version

`powerbi/MEASURES.md` maps every Excel KPI and chart to DAX measures
and a star-schema model, for the same pack as a `.pbix`.

## Limits, stated plainly

- Synthetic data only; collection/payment shares (20%/10%) are stated
  assumptions, not observed behaviour — a real close replaces them
  with the client's actual ageing.
- One company, one currency (INR), no GST/tax computation — the brief
  asks for reporting, not filing.
- No macros, no external links, no Table objects: plain A1 formulas
  throughout, so the file recalculates identically in Excel and
  LibreOffice.
