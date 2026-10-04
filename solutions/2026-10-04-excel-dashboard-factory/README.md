# DashboardForge — Automated Excel Sales Dashboard Factory

**Built:** 2026-10-04 · **Stack:** Python 3 (standard library only — no dependencies)
**Demo for:** *"Interactive Monthly Sales Excel Dashboard"* style gigs — turn a raw
monthly sales CSV into a client-ready, fully formatted multi-sheet Excel workbook.

## The problem

Small businesses track sales in raw CSV exports (order id, date, region, category,
units, price) but can't answer basic questions: *Which month performed best? What's
the average order value? Where is revenue trending?* Hiring a freelancer to hand-build
a dashboard every month is slow and expensive.

## The solution

`dashboard_factory.py` reads any sales CSV with the columns
`OrderID, Date, Region, Category, Product, Units, UnitPrice` and generates a real
`.xlsx` workbook (genuine Office Open XML, written with the stdlib `zipfile` module —
no `openpyxl`, no pandas, no Excel required to build it) containing:

| Sheet | Contents |
|---|---|
| **Raw Data** | Full sales table · styled headers · AutoFilter · frozen top row · currency/date/number formats |
| **Monthly Summary** | Pivot-style table (Month / Revenue / Units / Orders / Avg Order Value) built from **real Excel `SUMIFS`/`COUNTIFS` formulas** over the raw rows · TOTAL row · conditional formatting highlights the **top-3 revenue months** |
| **Dashboard** | Merged title · 4 KPI cells (Total Revenue, Total Units, Total Orders, Avg Order Value) wired by formula to the summary totals · **3 native Excel charts** — clustered column (monthly revenue), line (orders trend), pie (revenue share) |

Because the summary uses live Excel formulas (not hard-coded values), the workbook
**recalculates automatically** if the client edits the raw data — and no VBA/macros
are needed, so there are no macro-security warnings on open.

## How to run

```bash
# 1. Generate the deterministic sample dataset (12 months, 551 orders)
python dashboard_factory.py --generate-sample sample_sales_2026.csv

# 2. Build the dashboard workbook from any compatible CSV
python dashboard_factory.py --input sample_sales_2026.csv --output SalesDashboard.xlsx
# rows: 551 | months: 12 | total revenue: $423,400.44
# dashboard written: SalesDashboard.xlsx

# 3. Open SalesDashboard.xlsx in Excel / LibreOffice / Google Sheets
```

Input CSV columns: `OrderID, Date (YYYY-MM-DD), Region, Category, Product, Units, UnitPrice`
(extra columns such as `SalesRep` are accepted and ignored). Bad dates, non-numeric
quantities, or missing columns fail fast with a clear error naming the row.

## Tests

```bash
python test_dashboard_factory.py   # 31 tests, all passing
```

Coverage: input validation (missing columns, bad dates, bad numbers, empty file),
package structure (all 16 OOXML parts, 3 sheet names, number formats), raw sheet
(row count, headers, AutoFilter, frozen panes, revenue total), summary sheet
(`SUMIFS`/`COUNTIFS` ranges cover every raw row, month criteria use `DATE`/`EDATE`,
AOV guards division by zero, TOTAL row, top-3 conditional formatting), dashboard
(KPI formulas point at the TOTAL row, merged title), charts (bar/line/pie present,
titles correct, series reference the summary ranges), and byte-identical
deterministic rebuilds.

**Independent formula verification:** the generated workbook was opened in
LibreOffice Calc (headless recalculation) and every formula result was checked
against an independent Python computation — January revenue `$26,141.90` and the
`$423,400.44` total matched exactly, as did all four dashboard KPIs.

## Files

- `dashboard_factory.py` — the factory: CSV validation, OOXML workbook builder, charts, CLI
- `test_dashboard_factory.py` — 31-test suite (stdlib `unittest`)
- `sample_sales_2026.csv` — deterministic 12-month sample dataset (seed `20261004`)
- `README.md` — this file

## Customizing for a client

- Change brand colors: edit the `fills`/`fonts` in `styles_xml()` (header blue is `2E75B6`).
- Add a region/category slicer-style sheet: add a builder like `build_summary_sheet()`
  grouping by that column — the `SUMIFS` pattern generalizes directly.
- Feed Power BI: the **Monthly Summary** sheet (or the raw CSV) loads straight into
  Power BI Desktop as a clean, typed table — the same measures become DAX
  (`Total Revenue = SUM('Raw Data'[Revenue])`, `AOV = DIVIDE([Total Revenue],[Orders])`).
