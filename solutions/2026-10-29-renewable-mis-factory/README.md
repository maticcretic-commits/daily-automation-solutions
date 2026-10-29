# RenewableMIS — Monthly MIS: Actual vs Annual Operating Plan

A working demo built around a real Freelancer brief: **"Automate Renewable
Energy MIS Reporting"** (Freelancer project 40790612, Hyderabad) — automate
the monthly management report comparing generation and revenue actually
achieved against the Annual Operating Plan (AOP) budget for a solar, wind
and hybrid portfolio across nine Indian states plus international rooftop
sites. Today the report is assembled manually in Excel by copying and
linking cells. This demo proves the automated version: two raw tables in,
one formula-driven workbook out, with every figure traceable to source.

This is a **demonstration with fictional data**, not client work. The
portfolio (17 sites, Apr–Oct 2026) is invented and generated
deterministically (seed 20261029 in `generate.py`).

## What it builds

`renewable-mis.xlsx` — 9 sheets, every summary figure a live formula:

| Sheet | Contents |
|---|---|
| Guide | How to refresh, sheet map, data note |
| Dashboard | 10 headline KPIs + 3 native charts (technology column, generation-mix pie, monthly achievement line) |
| Monthly | Planned vs actual by month, variance, achievement %, revenue achievement, curtailment, availability from outage hours |
| State | Planned vs actual by state (9 Indian states + UAE + Singapore) |
| Technology | Planned vs actual by Solar / Wind / Hybrid / Rooftop |
| Sites | The 17-site portfolio: capacity (MW), tariff (Rs/kWh) |
| Raw_Plan | AOP baseline, one row per site-month |
| Raw_Actual | Actuals + formula derived columns (variance, achievement %, curtailment %) |
| Checks | 8 integrity checks — every Status must read PASS |

Two documented events are baked into the synthetic data so the checks and
commentary have something real to catch: heavy curtailment at Kutch Wind
Farm in Jul 2026 (18% of planned generation curtailed) and a 120-hour
outage at Kamuthi Solar Park in Sep 2026.

`powerbi/MEASURES.md` maps every Excel KPI to its DAX equivalent and lays
out the three Power BI pages (Overview / Monthly / State & Technology)
plus the scheduled-refresh story — the same two raw tables feed both.

## How to run

```sh
python3 generate.py          # writes data/*.csv (deterministic, seed 20261029)
python3 build_workbook.py    # builds renewable-mis.xlsx
python3 tests/test_build.py  # rebuilds, recalcs headlessly in LibreOffice,
                             # compares every checked figure to the independent
                             # Python ground truth in main.py
```

Requirements: Python 3 + openpyxl; LibreOffice (`soffice`) for the recalc
test. No other dependencies.

## How the automation works (the pitch in one paragraph)

The client's monthly report stops being a copy-and-link exercise: the
generation/revenue export lands in the two raw tables, the AOP baseline
stays untouched as its own table, and Monthly / State / Technology /
Dashboard all recalculate from SUMIFS formulas — a wrong number is
traceable to its site-month row in seconds, and next month is an append,
not a rebuild. The Checks sheet proves the tie-outs (monthly, state and
technology totals all reconcile to the raw tables) before anyone reads
the dashboard.

## Files

- `generate.py` — deterministic synthetic data (seed 20261029)
- `build_workbook.py` — openpyxl workbook builder (plain A1 ranges, no Table objects)
- `main.py` — independent ground truth (never reads the workbook)
- `tests/test_build.py` — LibreOffice recalc vs ground truth + formula-error scan
- `data/sites.csv`, `data/plan.csv`, `data/actuals.csv` — generated inputs
- `powerbi/MEASURES.md` — DAX mapping + Power BI page plan
- `renewable-mis.xlsx` — the built workbook
