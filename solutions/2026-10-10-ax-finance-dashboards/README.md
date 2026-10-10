# AX-Style Finance Dashboards

**The client problem.** A finance team running Microsoft Dynamics AX 2012 needs
proper financial dashboards — revenue, expenses, and cash-flow views built from
their ERP's general-ledger data — instead of hand-rebuilt spreadsheets every
month-end. This is a working reference implementation of that exact job: it
takes AX-style GL exports (a chart of accounts plus dated debit/credit
postings) and produces a finance dashboard workbook where **every statement
figure is a live Excel formula** — no hard-coded numbers anywhere.

## What it builds

`ax-finance-dashboards.xlsx` — 6 sheets, sample data for fictional
"Desert Rose Trading LLC", Apr–Sep 2026:

| Sheet | Contents |
|---|---|
| `GL_Transactions` | Raw GL export: entry, date, account, debit, credit, net |
| `Chart_of_Accounts` | Account master with groups + cash-flow sections |
| `Trial_Balance` | Debit / credit / balance per account (SUMIFS), with a balanced? check |
| `P&L` | Monthly profit & loss: revenue, COGS, gross profit, opex lines, operating income, net income |
| `Cash_Flow` | Indirect-method cash flow: operating / investing / financing, with a "ties to trial balance" check |
| `Dashboard` | 6 KPIs (revenue, expenses, net income, gross margin %, cash change, closing cash) + 3 native Excel charts |

Swap in your own export and every statement recalculates — the month columns
key off the date headers, so extending the period is just more rows in the GL.

## How to run

```bash
pip install openpyxl pytest
python generate.py                      # reads data/, writes ax-finance-dashboards.xlsx
python generate.py --data my_export/ --out my_dashboards.xlsx
python -m pytest tests/ -q              # 12 tests, incl. LibreOffice recalc check
```

`data/` holds the sample AX-style exports:

- `chart_of_accounts.csv` — account_code, account_name, account_group,
  normal_balance, cashflow_section
- `gl_transactions.csv` — entry_id, date, account_code, description, debit,
  credit (every entry double-entry balanced)

## Verification

The test suite checks the inputs (every journal entry balances), the workbook
structure (6 sheets, 5 named ranges, 3 charts, formulas not hard-coded values),
then runs the workbook through a headless LibreOffice recalculation and
compares each statement figure against ground truth computed independently
from the CSVs: P&L totals, monthly add-ups, all three cash-flow sections,
closing cash vs the trial balance, and dashboard KPIs. 12/12 pass, zero
formula errors.

## Notes

- Currency is labeled SAR in the sample data; replace with your own — the
  formulas don't care.
- Depreciation is modeled as a non-cash add-back (posted to accumulated
  depreciation, account 1510), so the cash-flow statement reconciles exactly.
- Balance-sheet deltas respect each account's normal balance (credit-normal
  accounts like payables/loans are negated when computing economic increases).
