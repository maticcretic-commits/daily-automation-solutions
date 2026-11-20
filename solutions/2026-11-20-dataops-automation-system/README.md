# DataOps Automation System

A working demo of the system in the Freelancer lead **"Data Analysis &
Automation System (Excel + Python + SQL + Power BI)"**: raw business CSVs
→ a SQL layer → a formula-driven Excel workbook → Power BI-ready exports.
Built for Narmada Traders, Indore — a fictional wholesale FMCG
distributor, Apr–Oct 2026. All data is invented for the demo (seed
20261120); nothing here is a real client's books.

## What it does

1. **`make_data.py`** generates the source data: 252 sales orders,
   126 purchase orders, 42 expenses, and a 10-product master.
2. **`pipeline.py`** loads the CSVs into SQLite and runs the auditable
   queries in **`queries.sql`** (monthly aggregates, revenue by category
   and channel, monthly expenses and purchases).
3. The same script builds **`dataops-automation-system.xlsx`**:
   - **Sales / Purchases / Expenses** — raw data with live derived
     columns (Revenue, Cost, Margin, Month).
   - **Monthly** — one row per month, every figure a live `SUMIFS` /
     `COUNTIFS` formula, plus a **SQL cross-check**: the SQL aggregates
     sit beside the Excel formulas and the diff columns must read 0.
   - **Dashboard** — pick a focus month in B2; the KPIs, monthly trend
     and category mix update. Three native charts: monthly revenue,
     monthly net profit, revenue by category.
   - **Checks** — 10 integrity checks (row counts, totals, SQL-vs-Excel
     diffs, Dashboard tie-out), all PASS.
4. **Power BI-ready star schema**: `data/fact_sales.csv` (one row per
   order line, with revenue / cost / margin) + `data/dim_products.csv`
   (product master). Load both into Power BI, relate on `product`, and
   the measures `Revenue = SUM(fact_sales[revenue])` and
   `Margin % = SUM(fact_sales[margin]) / SUM(fact_sales[revenue])`
   reproduce the Dashboard KPIs. `data/sql_summary.csv` is the SQL audit
   trail behind the cross-check.

## Verified results (Apr–Oct 2026)

| Figure | Value |
| --- | --- |
| Revenue | ₹25,03,873 |
| COGS | ₹20,76,754 |
| Gross profit | ₹4,27,119 (17.1% margin) |
| Expenses | ₹3,36,106 |
| **Net profit** | **₹91,013** |
| Best month (Oct) | Revenue ₹4,08,214 · net ₹21,877 · 46 orders |
| Weakest month (Apr) | Net −₹9,928 (the only loss month) |
| Top category | Grains ₹16,40,490 (65.5% of revenue) |
| Top channel | Wholesale ₹15,19,690 (60.7% of revenue) |

## How to run

```bash
python3 generate.py        # regenerates data + workbook
python3 tests/test_build.py  # full test suite (see below)
```

Requires Python 3.8+ and `openpyxl` (`pip install openpyxl`). No other
dependencies — the SQL layer is the standard-library `sqlite3`.

## Tests

`tests/test_build.py` (all pass):

- SQL aggregates in `sql_summary.csv` match an independent Python
  re-computation from the CSVs, month by month.
- LibreOffice headless recalc of the workbook against the same ground
  truth: Monthly revenue / COGS / gross profit / purchases / expenses /
  net profit / orders, SQL cross-check diffs all 0, Dashboard KPIs for
  the default focus month (2026-10), category totals — and all 10
  Checks read PASS, with zero formula errors.

## Files

| File | Purpose |
| --- | --- |
| `README.md` | This file |
| `make_data.py` | Deterministic fictional data generator |
| `pipeline.py` | SQLite loader + SQL runner + workbook builder |
| `generate.py` | Runs make_data then pipeline |
| `queries.sql` | The five auditable SQL queries |
| `data/sales.csv`, `data/purchases.csv`, `data/expenses.csv`, `data/products.csv`, `data/meta.csv` | Source data |
| `data/sql_summary.csv` | SQL monthly aggregates (audit trail) |
| `data/fact_sales.csv`, `data/dim_products.csv` | Power BI star schema |
| `dataops-automation-system.xlsx` | The workbook |
| `tests/test_build.py` | Test suite |
