# WarehousePulse — warehouse stock tracker + Power BI-ready model

**Demo built for a real freelance brief:** an "Excel & Power BI Warehouse Tracker"
(Excel data-entry system + automation + interactive Power BI dashboard with
drill-through). This is a working prototype of exactly that deliverable —
not a mock-up.

**Fictional company:** Sharma Wholesale Traders, Patna (FMCG wholesale warehouse,
INR). Opening stock as of 30-Sep-2026; 127 receipts, 223 issues and
17 adjustments across 01–18 Oct 2026; 36 SKUs in 5 categories, 12 bins.
All data is seeded and deterministic (`generate.py`, seed `20261019`).

## What it does

`warehouse-tracker-dashboard.xlsx` — every computed cell is a **live Excel
formula** (SUMIFS / VLOOKUP / IF / LARGE / INDEX-MATCH …), nothing pasted as
values:

| Sheet | Contents |
|---|---|
| `Guide` | Daily routine, Power BI connection steps, formula philosophy |
| `SKU_Master` | 36 SKUs: category, UOM, reorder level, std cost, bin, opening qty |
| `Raw_Receipts` / `Raw_Issues` / `Raw_Adjustments` | Movement tables with AutoFilter; paste new rows here |
| `Stock_Ledger` | Per-SKU closing, weighted-average cost, stock value, avg daily issue, days-of-cover, status (`OK` / `REORDER` / `STOCKOUT` / `OVERSTOCK`), suggested reorder qty — all formulas |
| `Bin_Map` | Per-bin qty on hand, SKU count, utilisation %, `FULL` / `EMPTY` / `OK` flag |
| `Entry` | Data-entry form with dropdowns (type, SKU) + auto-filled name/bin; copy the row into the matching `Raw_*` sheet |
| `Dashboard` | 8 formula-linked KPI cards, 3 native charts (stock value by category, daily receipts vs issues, status mix) and a **category drill-through**: pick a category, see its top-5 SKUs by stock value |
| `AI_Report` | Rule-based narrative + ⚠/✓/ℹ flags that recalculate with the data |
| `Checks` | 8 integrity checks — all must read `TRUE` after any recalculation |
| `PowerBI_Model` | Star-schema documentation, DAX measures and the 2-minute Power Query refresh routine |

`powerbi/` — the same model as a star schema for Power BI:
`dim_sku.csv`, `dim_bin.csv`, `dim_date.csv` (Oct–Dec 2026),
`fact_inventory_daily.csv` (one row per SKU per day) and `MEASURES.md`
with 9 ready DAX measures (Closing Qty, Stock Value, Days of Cover,
Reorder/Stockout SKUs, Overstock Value, Bin Utilisation %, …).

## How to run

```bash
python3 generate.py          # rebuild data/*.csv (deterministic)
python3 powerbi_export.py    # rebuild powerbi/*.csv + MEASURES.md
python3 build_workbook.py    # rebuild warehouse-tracker-dashboard.xlsx
python3 -m pytest tests/ -q  # 12 tests (needs openpyxl + LibreOffice)
```

Requirements: Python 3.10+, `openpyxl`; LibreOffice for the headless-recalc test.

## Verification

12/12 tests pass, including a **headless LibreOffice recalculation** of the
built workbook compared cell-by-cell against an independent Python ground
truth: all 36 SKU closings, weighted-average costs, stock values, statuses,
reorder quantities and days-of-cover match; all 8 `Checks` read `TRUE`;
zero formula errors (`#DIV/0!`, `#VALUE!`, `#NAME?`, …); the Power BI fact
table ties to the ledger; the drill-through resolves.

Demo state: total stock value **₹39,75,438.48**; 2 stockouts (`ST-007`,
`SN-005`), 5 SKUs on reorder, 16 overstock flags — the alerts are live, not
staged screenshots.

## Files

- `README.md` — this file
- `generate.py` — seeded fictional data generator → `data/*.csv`
- `build_workbook.py` — `openpyxl` workbook builder (formulas, charts, validation)
- `powerbi_export.py` — star-schema CSVs + DAX measures
- `tests/test_build.py` — 12 tests incl. LO recalc vs ground truth
- `data/` — `skus.csv`, `opening.csv`, `receipts.csv`, `issues.csv`, `adjustments.csv`, `bins.csv`
- `powerbi/` — `dim_sku.csv`, `dim_bin.csv`, `dim_date.csv`, `fact_inventory_daily.csv`, `MEASURES.md`
- `warehouse-tracker-dashboard.xlsx` — the deliverable workbook
