# OpsPulse — Weekly Operations Dashboard

A working demo built around a real freelance brief: turn weekly CSV
exports for product movements, shift schedules and order fulfilment into
one refreshable operations dashboard — in Power BI if preferred, in Excel
if that suits the model better.

The brief: **Power BI Operations Dashboard Development** (Freelancer
project ID 40794911, ₹750–1,250 INR/hour, Mumbai, 1 proposal, posted
about 11 hours before the 31 October 2026 scout). Managers need to track
inventory levels, staff utilisation and overall workflow efficiency at a
glance, from several weekly-refreshed CSV files, with no hard-coded paths,
working slicers/drill-throughs/filters, and simple instructions for
dropping in new files and refreshing.

## The problem

Three feeds, three questions, no single view: which items are approaching
low stock, which teams are running overtime, and how quickly orders move
through each stage? Rebuilt by hand each week, the answers drift — a
stock figure copied from last week, an overtime total that quietly
double-counts, a fulfilment average that includes a shipment dated
before its order.

## The approach

- **One workbook, three feeds.** `Movements` (168 rows: receipts, sales,
  adjustments), `Raw_Shifts` (120 shifts, 10 staff, 4 teams) and `Orders`
  (96 orders) feed summary sheets; nothing is pasted as values.
- **Every figure is a live formula.** Stock on hand is opening stock plus
  receipts minus sales plus adjustments (`SUMIFS` by product and type);
  reorder flags fire when stock is at or below the reorder level; overtime
  is `MAX(0, worked − 8)` per shift, summed by employee and by week;
  fulfilment days are ship date minus order date, and on-time means
  3 days or fewer — both standards live on the `Assumptions` sheet, so
  changing them reprices the whole workbook.
- **Weekly view that ties out.** The `Weekly` sheet groups all three feeds
  by Monday week-start with `COUNTIF`/`SUMIF`/`COUNTIFS`/`AVERAGEIF`, and
  the `Checks` sheet proves the dashboard, weekly and detail totals agree.
- **Refresh-proof by construction.** New weeks are new rows in the same
  columns (extend the formula ranges); dates are real date values, ranges
  are plain A1 ranges, and there are no Table objects or hard-coded paths.
  The Power BI DAX measures in `powerbi/MEASURES.md` rebuild the same
  dashboard as a `.pbix`.

## Verified figures (fictional demo data, 7 Sep – 31 Oct 2026, 8 weeks)

- 24 SKUs; **6 need reorder** (25%): Mustard Oil, Tea Powder, USB Cable,
  Mixer Grinder, Cotton Saree, Pressure Cooker.
- Stock on hand: **1,317 units, Rs 655,690** at cost.
- Orders: **96 orders, Rs 894,721**; **61.5% on time**, average
  fulfilment **2.98 days**.
- Staff: **132.0 overtime hours** across 120 shifts; utilisation
  **110.6%** of scheduled hours (worked hours exceed the 8-hour standard
  on most shifts — the overtime story is the headline, not a footnote).
- Last week (26 Oct) covers orders placed 26–27 Oct only, so they could
  all ship by 31 Oct; that is why its order count (3) is lower.

Verification: `tests/test_build.py` rebuilds the workbook from
`data/*.csv`, recalculates it headlessly in LibreOffice, and compares
every checked figure against an independent Python recomputation
(`main.py`, which never reads the workbook). All checks passed, including
all 10 Checks-sheet rows reading PASS and zero formula errors.

## How to run it

```bash
python3 generate.py          # (re)create data/*.csv — deterministic demo data
python3 build_workbook.py    # build ops-pulse.xlsx
python3 tests/test_build.py  # rebuild + LibreOffice recalc + ground-truth comparison
```

Requires Python 3.10+ and `openpyxl`; the test additionally needs
LibreOffice (`soffice`) for the headless recalculation.

To use it with real data: replace the rows in `Movements`, `Raw_Shifts`
and `Orders` with your own weekly CSV exports (same columns; extend the
formula ranges if you add rows) and set your standards on the
`Assumptions` sheet. The Power BI DAX measures in `powerbi/MEASURES.md`
rebuild the same dashboard as a `.pbix`.

## Files

- `ops-pulse.xlsx` — the workbook (Guide, Dashboard, Inventory,
  Movements, Raw_Shifts, Staff, Orders, Weekly, Assumptions, Checks;
  3 charts on the Dashboard)
- `generate.py` — deterministic synthetic operations data (fictional;
  no real client data)
- `build_workbook.py` — the workbook builder (all formulas, A1 ranges only)
- `main.py` — independent ground-truth recomputation
- `data/products.csv`, `data/movements.csv`, `data/shifts.csv`,
  `data/orders.csv` — the demo datasets
- `tests/test_build.py` — build + LibreOffice recalc + ground-truth tests
- `powerbi/MEASURES.md` — DAX measures for a Power BI version

## Reference

Built for the DashboardForge wedge (Excel/Power BI dashboards &
reporting/workflow automation). This is a reference implementation
(a demo, not client work), built from a public freelance brief.
