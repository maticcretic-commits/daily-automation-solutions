# Excel Finance Tracker

A small-business finance dashboard generator: sample transactions, invoices,
and budgets go in as CSVs, and a formatted Excel workbook comes out with a KPI
dashboard, budget-vs-actual analysis, invoice aging, a tax-prep summary, and
three native charts — all driven by live formulas.

## The problem

Small businesses run their finances in scattered spreadsheets: one tab for
expenses, a separate invoice list, budgets in someone's head. Nothing ties
together — no budget-vs-actual view, no invoice aging, no cash-flow picture,
and month-end means rebuilding everything by hand.

## The approach

`finance_tracker.py` builds a single workbook where every money cell is a
live formula:

- **Dashboard** — Total Income, Total Expenses, Net Profit, Cash on Hand,
  Outstanding Receivables, Overdue Amount, plus 3 native charts
  (monthly income vs expenses, cumulative net cash, expenses by category).
- **Transactions** — raw data with `Month` (`TEXT`) and `TaxCat` (`VLOOKUP`)
  helper columns; new rows just get pasted in.
- **Budget vs Actual** — 9-month budget vs `SUMIFS` actuals per category,
  variance and variance %, red conditional formatting when over budget.
- **Invoices** — Days Outstanding and a Paid / Overdue / Outstanding flag
  computed from the editable "As of" date; paste new invoices in.
- **Tax Prep** — deductible totals grouped by tax category via the
  `TaxCat` helper, so year-end categories are one glance.
- **ChartData** — helper series feeding the charts.

## How to run it

```bash
python3 make_sample_data.py   # regenerate the sample CSVs (deterministic)
python3 finance_tracker.py    # builds Excel_Finance_Tracker.xlsx
python3 -m unittest test_finance_tracker -v   # 30 tests, incl. LibreOffice recalc
```

Requirements: Python 3, `openpyxl`. The recalc tests need LibreOffice
(`soffice`) for the headless formula roundtrip.

## Files

| File | What it is |
|---|---|
| `finance_tracker.py` | Workbook generator |
| `make_sample_data.py` | Deterministic sample-data generator |
| `sample_transactions.csv` | 87 transactions, Jan–Sep 2026 |
| `sample_invoices.csv` | 20 invoices (14 paid, 3 overdue, 3 outstanding) |
| `sample_budgets.csv` | Monthly budgets for 9 expense categories |
| `Excel_Finance_Tracker.xlsx` | Generated workbook (open in Excel / Sheets / LibreOffice) |
| `test_finance_tracker.py` | 30 tests: data integrity, formula structure, recalculated values |
