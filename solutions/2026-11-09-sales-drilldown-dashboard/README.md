# SalesDrill — Interactive Sales Dashboard

A formula-driven Excel sales dashboard built from raw order history —
the reference implementation for Freelancer project 40808588,
"Interactive Sales Dashboard Builder" (Delhi; ₹100–400/hour, average
bid ₹250/hour): turn Excel/CSV order history into a single interactive
sales/revenue dashboard with drill-down by product/category/customer,
period comparison and a refresh workflow.

## The demo business

Delhi Traders, a fictional Delhi wholesaler. 646 orders, April–September
2026, 16 products in 4 categories, 30 customers. All data is fictional
and deterministic — see `data/orders.csv`, `data/products.csv`,
`data/customers.csv` (seeded generation; re-running `generate.py`
reproduces the same workbook).

## What's inside `sales-drilldown-dashboard.xlsx`

| Sheet | Contents |
|---|---|
| Guide | What the pack is, how to refresh it, how to read the KPIs |
| Orders | All 646 order lines. Amount, Cost, Margin and Month are live formulas (Cost via VLOOKUP into the Products cost master) |
| Products | Price & cost master (16 products, 4 categories) |
| Monthly | Revenue, orders, units, avg order, margin, margin % per month, plus prior-month revenue and MoM change |
| Category | Drill-down by category: revenue, orders, units, margin, share of revenue, avg order |
| Customers | Drill-down by customer (ranked by revenue): orders, revenue, share, avg order |
| Dashboard | 10 formula-wired KPIs + 3 native charts (monthly revenue, category share, top-10 customers) |
| Checks | 8 integrity checks — all must read PASS |

Key numbers (verified by the test suite against an independent Python
ground truth): total revenue ₹1,14,01,067 · total margin ₹25,13,202
(22.0%) · best month Jul 2026 · top category Electronics · top customer
Bhatia General Store · 100% of customers ordered more than once.

## How to run

```bash
python3 generate.py          # rebuilds sales-drilldown-dashboard.xlsx
python3 tests/test_build.py  # rebuilds, recalcs in LibreOffice, checks
                             # every rollup and KPI against ground truth
```

Requires Python 3 + openpyxl, and LibreOffice (`soffice`) for the
recalc check. No macros, no structured table references, plain A1 ranges
(LibreOffice-safe), real date values, percentages as 0–1 fractions.

## Refresh workflow

1. Paste new order rows at the bottom of Orders (columns A–H).
2. Copy the formulas in I, J, K, L down to the new rows.
3. Extend the range ends in Monthly, Category, Customers and Checks
   from row 647 to the new last row.
4. Re-check the Checks sheet — all 8 must read PASS.

## Note

This is a demo reference implementation built from a public job brief,
not client work. The DashboardForge reference implementations it builds
on: [CollectionsPulse](../2026-10-05-collections-pulse-dashboard) (NBFC
collections dashboard, 30/30 tests) and
[ExecPortfolio](../2026-11-07-exec-portfolio) (executive portfolio pack,
10/10 checks).
