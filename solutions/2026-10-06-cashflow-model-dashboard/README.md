# CashFlowForge — Dynamic Project Cash-Flow Model + Investor Dashboard

A stdlib-only Python factory that turns a monthly project cash-flow CSV into a
genuine multi-sheet Excel workbook (real OOXML, no openpyxl/pandas): a dynamic
cash-flow model with live formulas plus an investor summary dashboard.

Built as a working demo for the kind of engagement finance teams and lenders
ask for: *"dynamic Excel cash-flow model + summary dashboard — monthly
inflows/outflows, IRR, payback, investor-ready summary."*

## The problem it solves

Investment and lending decisions die in static spreadsheets: someone
hard-codes the NPV into a cell, the assumptions change, and nobody notices the
dashboard is stale. CashFlowForge makes staleness impossible — every number on
every sheet is a live Excel formula linked back to one editable Assumptions
sheet.

## What's inside the workbook

| Sheet | Contents |
|---|---|
| **Assumptions** | Editable inputs: project name, initial investment, annual discount rate (monthly rate is a formula), horizon, currency |
| **Cash Flow** | M0..M24 schedule. M0 outflow is formula-linked to the investment. Net CF, Cumulative CF, Discount Factor (`1/(1+r)^t`), Discounted CF are all live formulas. Red fill flags negative months |
| **Metrics** | Total Inflows / Outflows, Net Cash Position, **NPV**, **IRR** (Excel `IRR()` over the net-CF series, plus an annualized `(1+r)^12−1` view), **Payback month** (`MATCH` on first non-negative cumulative), Benefit–Cost Ratio (`SUMPRODUCT`) — each with a "how it is computed" note |
| **Dashboard** | 6 formula-wired investor KPIs + 3 native Excel charts: monthly net cash flow (column), cumulative J-curve (line), inflows vs outflows (pie) |

Change any blue input on Assumptions and the entire workbook recalculates.

## Run it

```bash
# 1. generate the deterministic sample project (aquaculture, 24 months)
python3 cashflow_forge.py generate --out sample_project_2026.csv

# 2. build the workbook + the independent audit CSV
python3 cashflow_forge.py build --csv sample_project_2026.csv \
    --name "Coastal Aquaculture - Phase 1" \
    --investment 2000000 --rate 0.12 \
    --out cashflow_model.xlsx --audit-csv monthly_cashflow.csv

# 3. run the tests
python3 test_cashflow_forge.py
```

Bring your own data: any CSV with `Month,Inflows,Outflows` for months 1–24
works, plus your own `--investment` and `--rate`.

## Verification (not just "tests pass")

- **23/23 unit tests pass** — writer primitives, sample determinism, workbook
  package validity (all 4 sheets, 3 charts, drawing, styles, shared strings),
  exact formula-string wiring on every sheet, and "no hard-coded metric"
  (every Metrics value cell is a formula cell).
- **Independent Python recomputation** of all 8 metrics (totals, NPV,
  bisection IRR, payback, BCR) matches the workbook's design.
- **Real spreadsheet-engine recalc**: the generated `.xlsx` was recalculated
  headless in LibreOffice Calc and all 8 cached metric values matched the
  independent Python computation exactly —
  NPV ₹932,070 · IRR 33.0% annualized · payback month 20 · BCR 1.08.
- `monthly_cashflow.csv` is computed independently in Python as an audit
  trail for the M0..M24 schedule.

## Files

- `cashflow_forge.py` — the factory (stdlib only)
- `test_cashflow_forge.py` — 23 tests
- `sample_project_2026.csv` — deterministic 24-month sample project
- `README.md` — this file
