# SalesInv — Sales & Inventory Dashboard

A working reference implementation built around a real Freelancer brief:
**"Sales Insights & Power BI Dashboard"** (Siuri, India, ₹750–1,250/hour,
3 proposals, avg bid ₹784/hr, posted ~10h before the 2026-11-10 scout).
The brief asked for a sales and inventory Power BI dashboard (profit,
customer, product and monthly trends) with **reusable SQL reporting
tables** and **Excel validation**. This is that build, as a demo: a
formula-driven workbook plus a SQLite reporting database.

> Demo only. The business ("Siuri Traders") and all data are fictional
> and deterministic. This is not client work and no client data is used.

## What is in this folder

| File | What it is |
| --- | --- |
| `sales-inventory-dashboard.xlsx` | The dashboard workbook (10 sheets, all figures live formulas) |
| `sales_inventory.db` | SQLite database: raw tables + 6 reusable reporting views |
| `generate.py` | Builds both outputs from `data/*.csv` (run it to rebuild) |
| `make_data.py` | Generates the fictional `data/*.csv` (seed 20261110) |
| `data/products.csv` | 14 products, 4 categories, prices, opening stock, reorder levels |
| `data/customers.csv` | 24 customers with regions |
| `data/sales.csv` | 520 sales orders, May–Oct 2026 |
| `data/purchases.csv` | 93 restock lines, May–Oct 2026 |
| `tests/test_build.py` | Full test suite (see "Tests" below) |

## The workbook (10 sheets)

- **Guide** — what each sheet does, how to refresh, how to read the KPIs.
- **Purchases** — stock-in lines; Total Cost is a live formula.
- **Sales** — order lines; Amount / Cost / Margin / Month are live formulas.
  Product and Region are **validated dropdowns** (Product from the Products
  master, Region from North/South/East/West) — the Excel validation from
  the brief.
- **Products** — price & cost master, opening stock, reorder level.
- **Inventory** — closing stock, stock value, OUT/LOW/OK status and days
  of cover, all live formulas over Sales + Purchases. One product is out
  of stock, 7 are at/below their reorder level.
- **Monthly** — revenue, orders, units, margin + prior-month comparison.
- **Category** / **Customers** — drill-downs by category and by customer
  (ranked by revenue).
- **Dashboard** — 14 KPIs + 3 native charts (monthly revenue, category
  share, top 10 customers).
- **Checks** — 10 integrity checks, all must read PASS (including a stock
  conservation check: opening + purchased − sold = closing).

## The database (reusable SQL reporting tables)

`sales_inventory.db` holds `products`, `customers`, `purchases`, `sales`
plus these views — query them from Power BI, Excel, or any SQL client:

- `v_monthly_sales` — revenue, orders, units, margin per month
- `v_profit_by_product` / `v_profit_by_category` — profit rollups
- `v_customer_ranking` — customers by revenue
- `v_inventory_status` — opening, purchased, sold, closing, stock value,
  OUT/LOW/OK status per product
- `v_reorder_alerts` — the 8 products needing a reorder

Example: `SELECT * FROM v_reorder_alerts ORDER BY closing;`

## Verified numbers (test output, 2026-11-10)

- Total revenue **₹1,796,235** · margin **₹519,118 (28.9%)**
- Best month **Aug 2026** · top category **Electronics** ·
  top customer **Singh Provisions**
- Closing stock value **₹159,060** · reorder alerts **8** · stock-out **1** ·
  average days of cover **25.7 days**

## Tests

```sh
python3 tests/test_build.py
```

What it checks: rebuilds both outputs fresh; scans every formula for
short-form range literals (invalid in LibreOffice); verifies all six SQL
views against an independent Python ground truth; recalculates the
workbook with LibreOffice 24.2.7.2 headless and compares every Monthly,
Category, Customer, Inventory and Dashboard KPI cell against the same
ground truth; asserts all 10 Checks read PASS, the Sales validations
exist, and no cell carries a formula error.

Result on 2026-11-10: **ALL PASS** — Monthly 6/6, Category 4/4,
Customers 24/24, Inventory 14/14, Dashboard KPIs 14/14, Checks 10/10,
SQL views verified, validations present, no formula errors.

## How to refresh (for a real client)

1. Paste new purchase rows at the bottom of Purchases (columns A–E) and
   copy the Total Cost formula down.
2. Paste new order rows at the bottom of Sales (columns A–H) and copy the
   formulas in I, J, K, L down.
3. Extend the range ends in Inventory, Monthly, Category, Customers and
   Checks to the new last rows.
4. Re-check the Checks sheet.
5. Append the new rows to `data/*.csv` and re-run `generate.py` to
   rebuild the `.db`.

## Limits (honest)

- Fictional data; no real client, no Power BI file (`.pbix` needs
  Power BI Desktop — the workbook + SQL views are the portable equivalent
  a client can open in Excel and point Power BI at).
- The workbook is formula-driven, not a live database link: refresh is
  the paste-and-extend workflow above.
