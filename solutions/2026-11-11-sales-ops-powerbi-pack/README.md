# SalesOpsPulse — Sales & Operations Power BI Pack

A working demo of the 5-page Power BI set from a real Freelancer brief:
**"Power BI Developer for Sales and Operations Dashboards (SAP Business One
data)"** (Freelancer ID 40812525, posted ~1 hour before the 2026-11-11
scout, 2 proposals, average bid $90): sales overview (YTD, month pacing,
channels), inventory/reorder, delivery routes, calling plan, and
receivables & ageing — from SAP Business One + Excel, DAX / Power Query,
star schema, published to Power BI Service with scheduled refresh.

Power BI Desktop is not available on this machine, so the pack ships as
the two things a Power BI build is made of: (1) a SQLite database with
the same facts as 9 reusable SQL views — the star-schema tables a .pbix
would consume — plus a DAX measures file, and (2) a formula-driven Excel
workbook that mirrors all 5 pages with live formulas, so every number
can be checked without Power BI.

## The fictional business

Sharma Trading Co., Patna — FMCG wholesaler. SAP Business One-style
data, 2026-04-01 .. 2026-11-10 (as-of 2026-11-10), deterministic
(`data/*.csv`, seed 20261111, `make_data.py`):

- 12 SKUs across Staples / Pulses / Edible Oil / Spices
- 20 customers across Distributor / Retail / Institutional / Online
- 513 sales-order lines, 93 stock-in lines, 167 deliveries (orders from
  2026-09-01), 90 weekly calling-plan rows (Oct 12 – Nov 10),
  130 invoices with payment status
- Monthly revenue targets in `data/targets.csv` (November target
  ₹4,00,000; pacing annualises 10 days of November sales over 30 days)

## Files

| File | What it is |
|---|---|
| `generate.py` | Builds the database and the workbook from `data/*.csv` |
| `make_data.py` | Regenerates the deterministic fictional data |
| `sales_ops.db` | SQLite: 7 raw tables + 9 reporting views (the star schema) |
| `sales-ops-powerbi-pack.xlsx` | 16-sheet formula-driven workbook (the 5 pages + drill-downs) |
| `powerbi/MEASURES.md` | DAX measures for each of the 5 Power BI pages |
| `tests/test_build.py` | Full test suite (see below) |
| `data/*.csv` | The fictional source data |

## The workbook — the 5 pages

1. **Sales_Overview** — monthly orders / units / revenue / margin,
   margin %, target, attainment, November pacing, and revenue split by
   channel (Distributor / Retail / Institutional / Online).
2. **Inventory_Reorder** — extracted reorder list (OUT / LOW SKUs only)
   with closing stock, reorder level and stock value.
3. **Delivery routes** — `Raw_Deliveries` (planned vs actual, on-time
   flag, delay days) with the route rollup on the Dashboard.
4. **Calling plan** — `Raw_Calls` (weekly plan, outcome, order value)
   with conversion on the Dashboard.
5. **Receivables_Ageing** — ageing buckets (Current / 0-30 / 31-60 /
   61-90 / 90+ / Paid) plus the top 5 debtors by outstanding amount.

Plus `Category` and `Customers` drill-downs, a `Dashboard` with 19 KPIs
and 3 native charts (monthly revenue, revenue by category, receivables
ageing), and a `Checks` sheet with 10 integrity checks.

## Demo state (as of 2026-11-10)

- Revenue ₹57,33,874 across 513 orders (avg order ₹11,177); margin
  ₹6,44,434 (11.2%); best month May 2026 (₹11,06,118); top customer
  Sharma Distributors
- November revenue so far ₹1,87,281 — pacing 140.5% of target
- Stock value ₹12,25,910; 3 reorder alerts (Sugar, Moong Dal, Coriander
  Powder — all LOW at 30 units); 0 stockouts; avg cover 54.2 days
- Deliveries 67.7% on time, avg delay 0.44 days
- 90 calls, 44.4% converted to orders
- Receivables outstanding ₹3,20,392; 90+ days overdue ₹1,37,887;
  top debtor Verma Wholesale (₹63,800, 63 days)

## How to run

```bash
python3 make_data.py   # regenerate data/*.csv (optional)
python3 generate.py    # rebuild sales_ops.db + sales-ops-powerbi-pack.xlsx
python3 tests/test_build.py
```

## Tests

`tests/test_build.py` — all pass:

- Rebuilds the workbook + DB fresh; regex-scans every formula for
  short-form range literals (`$H$2:$5000` — invalid in LibreOffice).
- Verifies all 9 SQLite views against an independent Python ground
  truth computed from `data/*.csv`.
- Recalculates the workbook with LibreOffice headless and compares
  every Sales_Overview row, every Raw_Inventory row, every ageing
  bucket, every Category row, the top-5 debtors, every Customers row
  and all 19 Dashboard KPIs against the same ground truth.
- Asserts all 10 Checks read PASS and no cell carries a formula error.

## Build notes

- SQLite does integer division on integer operands, so the views divide
  by `100.0`, not `100` — otherwise each order line's revenue is
  truncated to whole rupees and the views disagree with the workbook
  by a few rupees per month.
- The reorder extraction uses a per-row numeric helper
  (`ROW()-4`, because the range starts at row 5) + SMALL/INDEX.
  `SMALL(IF(...))` silently mis-evaluates in LibreOffice, and a
  `ROW()-1` helper points past the end of the range for rows below 13.
- Balancing receipts: the generator tops up any SKU that would close
  negative (Sugar / Moong Dal / Coriander), so the "no negative
  closing stock" check holds while the LOW alerts stay live.
