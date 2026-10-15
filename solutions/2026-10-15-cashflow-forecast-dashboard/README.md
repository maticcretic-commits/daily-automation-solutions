# FlowForge — 13-Week Cash Flow Forecast + Budget-Variance Dashboard

A working demo of the kind of cash-flow reporting automation finance teams ask
for on freelance marketplaces: a weekly cash model where **every money figure
is a live formula** — one assumption change re-flows the whole forecast.

Demo scenario (fictional): *Patna Traders Pvt. Ltd.*, an FMCG wholesale
distributor. 26 weeks of history, 13-week forward forecast. All values in INR.

## What it does

- **History** — 26 weeks of actuals by category (3 inflow, 6 outflow), with
  totals, net, opening/closing balance, budget plan, budget-vs-actual variance
  %, and automatic `⚠ Review` flags past the 10% threshold
- **Forecast** — 13-week projection from the trailing-4-week run rate plus
  editable growth assumptions; weeks under the cash floor are flagged
  `▼ BELOW FLOOR` with red conditional formatting
- **Dashboard** — 8 live KPI cards (projected closing, minimum cash, weeks
  below floor, avg weekly net, …) plus 3 native charts: cash-balance line
  (history vs forecast), weekly inflow-vs-outflow columns, outflow-mix pie
- **Checks** — 8 automated PASS/FAIL integrity checks (reconciliation,
  continuity, alert logic, zero formula errors)
- **Report** — rule-based weekly narrative generated from the data
- **Assumptions** — the only editable inputs (yellow cells): growth rates,
  cash floor, openings

## Run it

```bash
python3 generate.py          # deterministic fictional data -> data/*.csv
python3 build_workbook.py   # -> cashflow-forecast.xlsx
python3 tests/test_build.py # 10 tests incl. LibreOffice headless recalc
```

Requires `openpyxl` and a headless LibreOffice (`soffice`) for the recalc test.

## Weekly routine (the actual client workflow)

1. Paste new weekly actuals into History columns B–D / F–K (values only)
2. Read Report for the narrative; act on any `BELOW FLOOR` / `⚠ Review` flags
3. Re-check Checks — all 8 must read PASS

## Verification

10/10 tests pass. LibreOffice headless recalc of the built workbook matches the
independent Python ground truth to the rupee (26 history closings, 13 forecast
closings, all 8 integrity checks PASS, dashboard KPIs exact).

*Demo, not client work. Built as a portfolio reference implementation for
Excel/finance dashboard gigs.*
