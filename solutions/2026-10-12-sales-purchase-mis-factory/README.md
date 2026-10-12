# SalesPulse MIS Factory

A working demo for the classic "Excel & MIS Reporting" client brief: several
workbooks of daily sales/purchase activity → one consolidated, refreshable
MIS workbook with trend analysis and a management dashboard.

Built around a real Freelancer brief (project ID 40762180): turn raw daily
sales & purchase workbooks into meaningful MIS outputs — clean the data,
analyse trends (seasonality, growth, anomalies), and present interactive
dashboards that refresh without errors when new rows are pasted in.

## What it does

`generate.py` simulates the client's source files (fictional data for
**Verma Trading Co., Patna** — a wholesale FMCG distributor, Apr–Sep 2026)
and builds `sales-purchase-mis.xlsx`:

| Sheet | Contents (all formula-driven) |
|---|---|
| `Dashboard` | 8 KPI cards, sales-by-category chart, top-5-products chart, monthly trend chart |
| `Sales` / `Purchases` | Raw data as styled ranges with formula headroom — paste new rows, Amount/COGS pre-filled formulas calculate immediately |
| `Products` | Product master; `StdCost` = average purchase cost via `AVERAGEIF` |
| `Monthly` | Sales / Purchases / COGS / Gross Profit / Margin % / seasonality index via `SUMIFS` |
| `Daily` | Daily sales with avg ± 2σ anomaly flags |
| `Checks` | 6 integrity checks (reconciliation, negatives, date range, master coverage) — must all read OK |
| `Guide` | Handover notes: how to keep the reports updated |

Every derived figure is a **live Excel formula** — nothing hard-coded. COGS is
computed per-row with `XLOOKUP` into the product master, so cost changes flow
through automatically.

## Run it

```bash
python3 generate.py            # writes data/*.csv + sales-purchase-mis.xlsx
python3 -m pytest tests/ -q    # 14 tests, incl. LibreOffice recalc vs ground truth
```

Requires: Python 3.10+, `openpyxl`, `pytest` (LibreOffice optional — used by
one test to recalculate the workbook headlessly).

## Design notes

- **Refresh-proof**: plain A1-style formulas over bounded ranges with headroom
  (Sales rows 2–5000, Purchases 2–2000), so pasted rows are picked up
  automatically — the exact acceptance criterion in the brief. Formula columns
  are pre-filled into the headroom, so new rows calculate on paste.
- **Trust-first**: the `Checks` sheet reconciles every amount back to
  `Qty × Rate`; a single broken formula surfaces as FAIL before the weekly
  review, not after.
- **Anomaly detection**: daily sales flagged outside avg ± 2σ, plus an
  explicit zero-sales-day rule (a flat-zero day when the business normally
  trades is exactly the dip a manager wants surfaced). The demo data contains
  one bulk-order spike and one near-zero day — both get flagged.
- Deterministic data (`SEED = 20261012`): the same inputs always produce the
  same workbook, so results are reproducible.

*Demo implementation built for evaluation — fictional company and figures.*
