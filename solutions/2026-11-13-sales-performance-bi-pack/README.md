# SalesPerf BI Pack — Interactive Sales Performance Dashboard

A working demo of a real Freelancer brief: **"Interactive Sales Performance
BI Dashboard"** (Freelancer, project ID 40815164, posted ~12 hours before
the 13 Nov 2026 scout, 1 proposal, average bid ₹25,000) — build a
brand-new real-time sales/business-performance dashboard from scratch:
revenue, monthly/quarterly targets, % to goal, products, orders and
customer purchasing, with date/category filters and CSV export; Power BI,
Tableau or Looker Studio, with refresh instructions.

This pack answers the brief with a formula-driven Excel workbook (the
same model a Power BI dataset would consume) plus the CSV exports the
brief asks for. It is a demo built on fictional data — not client work.

## Fictional business

Agarwal Traders, Patna — FMCG wholesaler, 1 Jan – 31 Oct 2026.
1,249 orders, 14 products in 4 categories, 26 customers in 12 regions.
Total revenue **₹1,39,01,817.90** against a ₹1,41,92,000 target (**98.0%**
attainment). Best month: **2026-09**; top product: **Rice 25kg**; top
customer: **Motihari Distributors**. All data is deterministic
(`make_data.py`, seed 20261113) and invented — no real business data.

## What is in this folder

- `generate.py` — builds the workbook and the two CSV exports.
- `make_data.py` — generates the fictional data in `data/`.
- `data/orders.csv` — 1,249 order rows.
- `data/targets.csv` — monthly revenue targets (demo targets: a fixed
  percentage of each month's actual revenue, factors 0.92–1.30, so the
  red/amber/green attainment bands all show; in production the client
  supplies real targets).
- `data/products.csv`, `data/customers.csv` — the masters the workbook
  looks up against.
- `data/monthly_export.csv` — the Monthly summary as a flat file
  (month, orders, revenue, target, % to goal, avg order, quarter).
- `data/category_export.csv` — the By_Category summary as a flat file.
- `sales-performance-bi-pack.xlsx` — the dashboard workbook.
- `tests/test_build.py` — the test suite (see below).

## The workbook (13 sheets)

| Sheet | Contents |
|---|---|
| Guide | What this is, how to refresh, attainment bands, CSV export |
| Orders | Raw rows; Category (INDEX/MATCH to Products), Amount, Month and Quarter are live formulas |
| Products / Customers | Master lists |
| Targets | Monthly revenue targets |
| Monthly | COUNTIFS/SUMIFS orders + revenue vs target, % to goal (green ≥100%, amber 80–99%, red <80%), avg order, quarter |
| Quarterly | SUMIFS over Monthly + months-on-target per quarter |
| By_Product / By_Customer / By_Region / By_Category | COUNTIFS/SUMIFS + revenue share |
| Dashboard | Working Focus-month + Focus-category dropdowns driving three live SUMIFS filter KPIs, 13 KPIs, 3 native charts (monthly revenue vs target, % to goal by month, revenue share by category) |
| Checks | 10 integrity checks, all PASS |

## How to run

```bash
python3 make_data.py    # regenerate data/*.csv (deterministic)
python3 generate.py     # rebuild the workbook + CSV exports
python3 tests/test_build.py   # full suite incl. LibreOffice recalc
```

Requires Python 3 + `openpyxl`; the test suite also needs LibreOffice
(`soffice --headless`) for the recalculation check.

## How to refresh (for the client)

1. Append the new month's rows to `data/orders.csv` and its target to
   `data/targets.csv`.
2. Re-run `python3 generate.py` — it rebuilds the workbook and
   re-exports both CSVs.
3. Every figure on Monthly / Quarterly / By_* / Dashboard is a live
   formula, so the dashboard updates from the rebuilt Orders sheet.

## Tests

`tests/test_build.py` rebuilds everything fresh, regex-scans every
formula for short-form range literals (invalid in LibreOffice),
independently re-aggregates `data/*.csv`, and compares every Monthly /
Quarterly / By_* row, every Dashboard KPI (including the live filter
cells at their defaults: month 2026-10, category Edible Oil), both CSV
exports, and all 10 Checks against a LibreOffice headless recalculation
of the workbook. **All tests pass** — zero formula errors, all Checks
PASS.
