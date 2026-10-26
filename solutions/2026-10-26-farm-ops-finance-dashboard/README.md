# FarmPulse — Farm Ops & Finance Dashboard (demo)

**Status: demo, not client work.** Built 2026-10-26 around the shape of a
real Freelancer brief (single interactive farm dashboard from Excel/CSV
data: income by activity, expenses by category, profit by enterprise over
time, with drill-through and season/paddock/herd/cost-centre filters).
All data is synthetic and fictional (seed 20261026) — no client data.

## What it does

- `farmpulse.xlsx` — 8 sheets, every figure a live formula:
  - **Guide** — refresh + drill instructions.
  - **Dashboard** — 6 KPIs (total income, total expenses, net profit,
    margin, best enterprise, expense ratio), a quarterly income-vs-
    expense table, and 3 native charts (profit by enterprise, income
    vs expense by quarter, expenses by category).
  - **Enterprise_Profit** — profit by enterprise × quarter, with total
    income and margin; margin is green above 20%, red below 0%.
  - **Income_Activity** — income by activity × quarter.
  - **Expense_Category** — expenses by category × quarter.
  - **Raw_Income / Raw_Expense** — 172 income + 164 expense transactions
    (date, quarter, enterprise, activity/category, paddock/herd, cost
    centre, amount), AutoFilter + frozen panes for drill-through.
  - **Checks** — 8 live checks (totals tie across sheets, row counts,
    no negative amounts); all must read TRUE.
- `powerbi/MEASURES.md` — the DAX measures that reproduce this workbook
  in Power BI from the same two fact tables.

## Refresh

Paste new rows into `Raw_Income` / `Raw_Expense` (same columns, rows
2–1000) and every summary, KPI and chart updates by formula. Keep the
quarter label format `YYYY-Qn`.

## Run it

```bash
python3 generate.py        # writes data/income.csv, data/expense.csv
python3 build_workbook.py  # writes farmpulse.xlsx
python3 tests/test_build.py  # LibreOffice recalc vs ground truth
```

## How to use it for a real farm

Swap the synthetic CSVs for the farm's real Excel/CSV exports, keep the
columns, and the workbook becomes the live dashboard. For Power BI,
import the two Raw tables and apply `powerbi/MEASURES.md`.
