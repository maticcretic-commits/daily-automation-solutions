# KPI Pulse — Monthly KPI Dashboard (demo)

**Problem (real listing):** "Power BI KPI Dashboard Development" —
Freelancer, Nagpur client, ₹1,500–12,500 fixed, posted ~11 hours before
the 27 Oct 2026 scout, just 1 proposal on its detail page. A business
needs one dashboard where monthly KPIs — revenue vs target, margin,
on-time delivery, customer satisfaction — stay consistent across every
view, instead of five spreadsheets that disagree.

**What this builds:** a formula-driven Excel KPI dashboard plus the
Power BI measure mapping for the same model. Paste new daily rows into
`Raw` and Monthly, Channel, Dashboard and Checks all update by formula.
Nothing is typed twice; nothing is a pasted value pretending to be a
calculation.

## Contents

- `generate.py` — deterministic synthetic data (fictional Nagpur Retail
  Co., Jan–Sep 2026, 3 channels, 819 daily rows, seed 20261027).
- `build_workbook.py` — builds `kpi-dashboard.xlsx` (Guide, Dashboard,
  Monthly, Channel, Raw, Targets, Checks).
- `main.py` — independent ground truth recomputed from the CSVs.
- `tests/test_build.py` — builds, recalcs in LibreOffice, and compares
  every checked figure to ground truth.
- `data/raw_transactions.csv`, `data/monthly_targets.csv` — the inputs.
- `powerbi/MEASURES.md` — the DAX mapping for the Power BI version.

## Dashboard

Six KPI cards (latest-month revenue, YTD revenue, YTD gross margin,
latest on-time delivery, latest CSAT, YTD target attainment) plus three
native charts: revenue vs target by month, gross margin % trend, and
revenue share by channel. Monthly carries a RAG status per month
(On track / Watch / At risk) driven by variance against target.
Checks must all read TRUE — any FALSE means the model broke.

## Run it

```bash
python3 tests/test_build.py   # generates data, builds, recalcs, verifies
```

All figures are synthetic demo data, labelled as such in the workbook.
This is a reference implementation, not client work.
