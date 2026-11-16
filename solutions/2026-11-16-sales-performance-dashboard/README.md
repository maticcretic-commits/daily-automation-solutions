# Sales Performance Dashboard

A formula-driven Excel sales dashboard for a wholesale FMCG distributor —
built as a working demo for the Freelancer project **Interactive Sales
Performance BI Dashboard** (revenue, targets, % to goal, product / order /
customer views, filterable, CSV export).

## The problem it solves

Sales numbers live in CSV exports: revenue in one file, targets in another,
and no single page that says whether the month is on track. This pack turns
those exports into one interactive dashboard — pick a month and a category
and every KPI, table and chart updates from live formulas.

## What is inside

| File | What it is |
| --- | --- |
| `sales-performance-dashboard.xlsx` | The dashboard workbook (6 sheets, 2 native charts) |
| `data/orders.csv` | 246 fictional orders, Jan–Oct 2026 (the CSV export the dashboard reads) |
| `data/targets.csv` | Monthly revenue target per category (50 rows) |
| `make_data.py` | Generates the fictional data (seed `20261116`, deterministic) |
| `generate.py` | Builds the workbook from the CSVs |
| `tests/test_build.py` | Test suite (LibreOffice recalc vs independent Python ground truth) |

## The workbook

- **Guide** — how to use it, sheet by sheet.
- **Data** — one row per order; Revenue = Qty × UnitPrice (live formula),
  Month = LEFT(Date, 7). Paste a fresh CSV export here and extend the
  formulas down.
- **Targets** — one row per category × month. Change a target and
  % to goal updates everywhere.
- **Monthly** — revenue / target / % to goal / orders per category × month
  (SUMIFS over Data and Targets).
- **Dashboard** — focus-month + focus-category dropdowns driving 7 KPIs
  (Revenue, Target, % to Goal, Orders, Avg Order Value, Top Product
  Revenue, Top Product), a monthly revenue-vs-target table, a by-category
  table, a by-product table, and 2 native charts.
- **Checks** — 10 integrity checks, all PASS (they assume the Dashboard
  filters are at their defaults, All / All).

## Verified totals (fictional demo data)

- 246 orders, revenue **₹1,26,17,720** vs target **₹1,29,75,000**
  — **97.2% to goal**
- Top product: **Sugar 1kg — ₹21,19,905**
- 10/10 integrity checks PASS; zero formula errors under LibreOffice
  headless recalc; every KPI, table and chart cell matches an independent
  Python recomputation from the CSVs.

## How to run

```bash
python3 make_data.py    # regenerate data/orders.csv + data/targets.csv
python3 generate.py     # rebuild sales-performance-dashboard.xlsx
python3 -m pytest tests/test_build.py -q   # ALL TESTS PASS
```

Requires Python 3 + openpyxl; the test suite also uses LibreOffice
(`soffice --headless`) for the recalculation check.

## Notes

- All data is fictional (Ganga Traders, Patna — an invented wholesale FMCG
  distributor), generated deterministically so the pack rebuilds identically.
- "All"-filter formulas use the SUMPRODUCT `((range=crit)+(crit="All"))`
  pattern: safe because no data row ever equals "All", so each factor is
  exactly 0 or 1.
- The workbook is an Excel demo of the dashboard pattern; the same model
  maps directly to Power BI (star schema + DAX measures) or Tableau.
