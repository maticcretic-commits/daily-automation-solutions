# ExecPortfolio — Executive Portfolio Reporting (Excel foundation for an M365 / Power BI pack)

Reference implementation (demo, not client work) built around a real
Freelancer brief: **"Executive Portfolio Power BI Dashboards -- 2"**
(project 40806608) — end-to-end M365 executive portfolio reporting:
Power BI dashboards fed from Excel, Planner and SharePoint via Power Query,
with Power Automate refresh and alerts.

## What this is

The Excel foundation that feeds that Power BI model. Source tables mimic
the three M365 feeds, and every reporting figure is a live Excel formula —
nothing is hard-coded except the source rows and the as-of date.

| Sheet | Contents |
|---|---|
| Guide | How to use the workbook and refresh it |
| Projects | Excel project master — 12 projects, 4 programmes (inputs) |
| Tasks | Planner-style task export — 60 tasks (inputs) |
| Risks | SharePoint-style risk list — 24 risks (inputs) |
| Spend | Monthly actuals staging, Jun–Nov 2026, programme denormalised (inputs) |
| Portfolio | Per-project rollup — budget, spent, variance, % complete, task counts, overdue, RAG, top risk (all formulas) |
| Monthly | Spend by programme per month (all formulas) |
| Dashboard | 9 KPIs + 3 native charts |
| Checks | 10 integrity checks — all must read PASS before the pack goes out |

Demo portfolio: **Meridian Group**, a fictional company. All data is
fictional and deterministic (seed 20261107).

## Verified results (LibreOffice headless recalc vs independent Python ground truth)

- Total portfolio budget **$1,745,000**; spent to date **$1,740,200** (−0.3% vs budget)
- Projects on track (Green): **2**; at risk (Red): **0**; Amber: 10
- Overdue tasks: **14** of 60; tasks completed: **26 / 60** (43.3%)
- Average project completion: **65.3%**; highest risk score: **20** (of 25)
- Checks: **10/10 PASS**, zero formula errors

## Model rules

- **% complete** per project = AVERAGEIF of its tasks' % complete.
- A task is **overdue** when its due date is before the as-of date
  (Portfolio!M1, currently 2026-11-07) and its status is not Completed.
- **RAG:** Green when spend is within 10% over budget AND no task is
  overdue; Red when spend is over 25% above budget OR 5+ tasks overdue;
  Amber otherwise.
- **Risk score** = Likelihood (1–5) × Impact (1–5).
- Monthly spend is SUMIFS over the Spend staging table, which carries the
  programme denormalised — exactly what a Power Query merge of the Excel
  master into the spend extract produces.

## How to use it

1. Paste your own rows into Projects / Tasks / Risks / Spend and extend
   the formula ranges on Portfolio and Monthly to match.
2. Change the as-of date (Portfolio!M1) to re-cut the report for any date.
3. In Power BI, load one query per sheet through Power Query; the
   Portfolio sheet is the fact table and the Dashboard KPIs map 1:1 to
   DAX measures. Refresh discipline mirrors the Power Automate flow:
   refresh the extracts, paste them in, check the Checks sheet.
4. Run the tests: `python3 tests/test_build.py` — rebuilds the workbook,
   scans formulas for invalid range literals, recalculates with
   LibreOffice headless, and compares every KPI, rollup, monthly and
   programme total against an independent Python ground truth.

## Files

- `generate.py` — builds the workbook and the CSVs in `data/`
- `data/projects.csv`, `data/tasks.csv`, `data/risks.csv`, `data/spend.csv`
- `tests/test_build.py` — full verification suite
- `exec-portfolio.xlsx` — the built workbook

No macros. Plain A1 ranges with the column letter repeated on both ends
(LibreOffice-safe). Nothing here is financial advice; it is a reporting
template.
