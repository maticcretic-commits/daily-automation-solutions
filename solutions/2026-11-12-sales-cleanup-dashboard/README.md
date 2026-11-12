# SalesClean Dashboard Factory — Excel Sales Dashboard & Cleanup

A working demo of a real Freelancer brief: **"Excel Sales Dashboard &
Cleanup"** (Freelancer, posted ~17 hours before the 2026-11-12 scout,
39 proposals, average bid ₹16,683): clean two years of sales data
(duplicates, blanks, inconsistent names) and build a refreshable Excel
dashboard — sales by product / customer / region against target.

## The fictional business

Gupta Traders, Patna — FMCG wholesaler. Two years of sales,
2024-11-01 .. 2026-10-31, deterministic (`data/*.csv`, seed 20261112,
`make_data.py`):

- 10 products across Staples / Pulses / Edible Oil / Spices
- 16 customers across Patna / Muzaffarpur / Gaya / Bhagalpur
- 1,320 clean order lines, plus 110 injected dirty rows = **1,430 raw rows**
- Monthly revenue targets in `data/targets.csv` (festive Oct/Nov higher)

## What the cleaner does

`generate.py` applies six rules in a fixed order — a row is dropped at
the **first** rule it fails, so the per-reason counts are exact:

| # | Rule | Dropped | What it catches |
|---|------|---------|-----------------|
| 1 | BAD_DATE | 8 | Empty or unparseable dates (ISO or DD/MM/YYYY only; dash DD-MM-YYYY rejected as ambiguous) |
| 2 | UNKNOWN_CUSTOMER | 4 | Customer not in the master list after normalisation |
| 3 | UNKNOWN_PRODUCT | 4 | Product not in the master list after normalisation |
| 4 | BAD_QTY | 18 | Quantity blank, non-numeric, or ≤ 0 |
| 5 | BAD_PRICE | 15 | Unit price blank, non-numeric, or ≤ 0 |
| 6 | DUPLICATE | 45 | Same (date, customer, product, qty, price) already kept — exact copies and case/spacing near-duplicates |
| | **Total dropped** | **94** | |
| | **Kept** | **1,336 (93.4%)** | |

Normalisation: names match on a canonical key (lowercase, alphanumeric
only), so "SHARMA DISTRIBUTORS ", "Sharma Distributors" and
"sharma  distributors" all resolve to the master name; 10 valid rows
with mangled names and 6 rows with DD/MM/YYYY text dates are kept with
normalised values. Dates are stored as real date values, never text.

## Files

| File | What it is |
|---|---|
| `generate.py` | Cleaning engine + workbook builder |
| `make_data.py` | Regenerates the deterministic fictional data (incl. the dirt) |
| `data/raw_sales_dirty.csv` | The 1,430 raw rows as "received" |
| `data/cleaned_sales.csv` | The 1,336 kept rows (machine-readable output) |
| `data/targets.csv` | Monthly revenue targets |
| `sales-cleanup-dashboard.xlsx` | Formula-driven workbook (see below) |
| `tests/test_build.py` | Full test suite (see below) |

## The workbook

1. **Guide** — what it is, how to refresh, the rules, the sheet map.
2. **Raw_Dirty** — all 1,430 rows as received, with a Disposition per row
   (`KEEP` or `DROP:<reason>`).
3. **Cleaning_Log** — drop counts per reason + kept / dropped / received totals.
4. **Cleaned_Sales** — the 1,336 kept rows; Amount, Month and Quarter are
   live formulas.
5. **Monthly** — 24 months: orders (COUNTIFS), revenue (SUMIFS), target,
   attainment, avg order, plus a TOTAL row.
6. **By_Product** — 10 products: orders, revenue, share of total.
7. **By_Customer** — 16 customers: orders, revenue, share of total.
8. **By_Region** — 4 regions: orders, revenue, share of total.
9. **Dashboard** — 11 live KPIs + 3 native charts (monthly revenue column,
   revenue by product, revenue share by region).
10. **Checks** — 10 integrity checks, all PASS (raw = kept + dropped,
    revenue tie-outs, no blanks, no non-positive amounts, 24 months covered,
    attainment within 0–200%, cleaning-log total = rows dropped).

## Verified results (LibreOffice headless recalc vs independent Python ground truth)

- Total revenue **₹3,82,36,794.73** across **1,336 orders**, 24 months
- Overall attainment vs target **92.8%**; clean rate **93.4%**
- Top product: **Rice 25kg** (₹1,05,56,080.28, 130 orders)
- Top customer: **Bhagalpur Wholesale Hub** (₹30,36,151.92, 92 orders)
- Top region: **Patna** (₹1,48,08,119.91, 508 orders)
- Best month: **2026-07** (₹21,61,015.57); weakest: 2026-05 (₹13,40,808.04)

## How to refresh it for a real client

1. Paste the new month's raw rows at the bottom of **Raw_Dirty**.
2. Export that sheet to `data/raw_sales_dirty.csv` (or drop the new CSV in
   `data/`).
3. Re-run `python3 generate.py` — it re-cleans everything and rebuilds the
   workbook. Every figure on Monthly / By_* / Dashboard is a live formula,
   so the dashboard updates from the rebuilt Cleaned_Sales sheet.
4. Re-run `python3 tests/test_build.py` — it rebuilds from scratch and
   re-verifies every figure against an independent ground truth.

## Tests

`python3 tests/test_build.py` — **ALL TESTS PASS**:

- Independent re-implementation of the cleaning spec: kept/dropped counts,
  per-reason counts, and every kept row's normalised fields match.
- Formula hygiene: every formula regex-scanned for short-form range
  literals (`$H$2:$5000` — invalid in LibreOffice).
- LibreOffice headless recalc: every Monthly row, every By_Product /
  By_Customer / By_Region row, every Dashboard KPI and every Check
  compared against the same ground truth; all 10 Checks read PASS and no
  cell carries a formula error.
