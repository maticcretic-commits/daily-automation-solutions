# AccountingPulse — Excel Accounting Automation Factory

A working demo for the **Excel Accounting Automation Specialist** brief:
automate A/P, A/R, payroll, bank reconciliation, aging schedules, and
financial-statement reporting in Excel — with dashboards that refresh from
the source data and documentation your team can actually follow.

**What it does:** raw accounting CSVs → a 9-sheet Excel workbook where
*every* figure is a live formula. Nothing is pasted as values.

## The sheets

| Sheet | Contents |
|---|---|
| `Raw_AP` | 39 vendor bills + formula columns: days overdue, aging bucket, outstanding |
| `Raw_AR` | 36 customer invoices + same formula columns (+ in-transit flag) |
| `Payroll` | 48 salary rows (8 staff × 6 months), net pay = gross − deductions |
| `Bank` | 107-line bank statement with running net column |
| `Bank_Rec` | Bank reconciliation: adjusted bank balance vs adjusted book balance, with a live `OK - reconciled` / `OUT OF BALANCE` check. Handles deposits in transit, outstanding cheques, unrecorded bank charges, and interest credited |
| `AP_Aging` / `AR_Aging` | Aging buckets (Current / 1–30 / 31–60 / 61–90 / 90+) via `SUMIFS`, totals, counts, and an integrity check that ties back to the raw ledgers. 90+ flagged red |
| `P&L` | Monthly P&L Apr–Sep 2026, every line a live `SUMIFS` |
| `Dashboard` | Inputs (as-of date, opening cash), 8 KPIs, and 3 native Excel charts: AP-vs-AR aging bars, monthly revenue-vs-expenses line, payroll-by-department pie |

## How to run

```bash
python3 generate.py                 # builds accounting-pulse.xlsx from data/
python3 -m pytest tests/ -q         # 14 tests, incl. LibreOffice recalc
```

The sample data covers fictional **Rajesh Kumar & Sons, wholesale traders,
Patna** (INR), Apr–Sep 2026, as of 30-Sep-2026.

## Refresh routine (the automation part)

1. Append new bills / invoices / payroll / statement rows to the CSVs in
   `data/` (same columns).
2. Re-run `python3 generate.py` — aging, reconciliation, P&L, KPIs and
   charts rebuild from the fresh data.
3. Or paste new rows straight into the `Raw_*` sheets and copy the formula
   columns down — everything downstream is formula-driven, so it recalculates
   in place.

## Verification

`tests/test_generate.py` (14/14 pass) does a LibreOffice headless recalc of
the built workbook and checks every figure against ground truth computed
independently from the CSVs: aging buckets, the full bank reconciliation
(statement balance, deposits in transit, outstanding cheques, adjusted bank
vs adjusted book, difference = 0.00), P&L lines, KPIs, and department
payroll totals. No `#VALUE!` / `#DIV/0!` errors anywhere.

Part of the [daily-automation-solutions](https://github.com/maticcretic-commits/daily-automation-solutions)
demo series — reference implementations, not client work.
