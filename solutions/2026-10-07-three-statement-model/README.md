# StatementsForge — Three-Statement Financial Model Factory

**Built:** 2026-10-07 · **Wedge:** DashboardForge (Excel/Power BI dashboards & reporting automation)

## The client problem

A "three-statement Excel model" engagement fails in one place: a model whose
numbers are hard-coded. One assumption change and the NPV, margins, and cash
position silently go stale — and nobody notices until the investor review.
The fix is a model where **every number traces to a formula**.

## What this does

Reads a trial-balance CSV (`Account, Category, Subcategory, Opening, Closing`)
and builds a genuine 6-sheet `.xlsx` workbook (real OOXML, **stdlib only** —
no openpyxl, no pandas):

1. **Trial Balance** — raw input, formula `Change` column, AutoFilter, frozen panes
2. **Income Statement** — every line a live `SUMIFS` on the Trial Balance;
   Total Revenue, Gross Profit, EBITDA, EBIT, Profit Before Tax, Net Income
3. **Balance Sheet** — Assets = Liabilities + Equity with a live balance check
   (green/red); Retained Earnings articulates as *Opening + Net Income*
   linked to the Income Statement
4. **Cash Flow (indirect)** — movements formula-linked to the Trial Balance
   Change column; ending cash tied back to the Balance Sheet with a live
   tie-out check (green/red)
5. **Ratios** — current/quick ratio, debt/equity, gross/EBITDA/net margins,
   ROE, ROA, common-size income statement — all live formulas
6. **Dashboard** — 5 formula-wired KPIs + 3 native Excel charts
   (revenue vs cost structure, asset mix, cash walk)

Plus `audit.csv`: an independent Python-side computation of the full model —
the numbers every formula must reproduce.

Sample data: **Sharma Textiles Pvt Ltd, FY 2025-26** (₹) — Net Income ₹12.48L,
Total Assets ₹158.98L, Closing Cash ₹27.83L, balance check 0.

## Run it

```bash
# 1. Generate the deterministic sample trial balance
python3 statements_forge.py generate --out sample_company_2026.csv

# 2. Build the workbook + audit CSV
python3 statements_forge.py build --csv sample_company_2026.csv \
    --name "Sharma Textiles Pvt Ltd" \
    --out three_statement_model.xlsx --audit-csv audit.csv

# 3. Run the tests
python3 test_statements_forge.py
```

Bring your own data: any CSV with the same header works. Category values the
sheets expect: `Asset`, `Liability`, `Equity`, `Revenue`, `Expense`;
subcategories: `Current`, `Non-current`, `Share Capital`, `Retained`, `Sales`,
`Other`, `COGS`, `Operating`, `Finance`, `Tax`.

## Verification

- **29/29 unit tests pass** — sample determinism, opening/closing identity,
  plug arithmetic, formula wiring on all six sheets, no hard-coded totals,
  chart ranges, audit-CSV agreement, no AI-assistant branding.
- **LibreOffice Calc headless recalc**: all 23 checked cells match the
  independent Python model exactly (revenue ₹1,87,40,000 · net income ₹12,48,000 ·
  total assets ₹1,58,98,000 · balance check 0 · cash tie-out 0 ·
  current ratio 2.44 · net margin 6.66%).

## Files

| File | What |
|---|---|
| `statements_forge.py` | The factory (stdlib only) |
| `test_statements_forge.py` | 29 tests |
| `sample_company_2026.csv` | Deterministic sample trial balance |
| `README.md` | This file |
