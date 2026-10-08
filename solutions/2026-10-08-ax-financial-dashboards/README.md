# AX 2012 Financial Dashboards

Turn a raw Dynamics AX 2012 general-ledger trial-balance export into a
self-contained Excel financial dashboard workbook — income statement, balance
sheet, KPI dashboard, and native Excel charts — all driven by live formulas.

## The problem it solves

Finance teams exporting trial balances from Dynamics AX 2012 get hundreds of
flat GL rows (fiscal period, main account, debit/credit). Turning that into a
month-end pack — P&L by period, a balance sheet that actually balances, margin
and liquidity KPIs, trend charts — is normally a day of manual Excel work, and
it breaks the moment next month's export lands.

This generator does it in seconds and keeps working: every statement cell is a
`SUMIFS` formula against the `Data` sheet, so dropping in a new period's export
rows refreshes the whole pack.

## What's inside

| Sheet | Contents |
|---|---|
| `Data` | Normalized export + a `Signed` amount column (credit-normal accounts flipped) |
| `Income Statement` | Revenue / expense lines per period, totals, net income — all `SUMIFS` |
| `Balance Sheet` | Assets, liabilities, equity per period; retained earnings rolls in the current period's profit (pre-close presentation); a check row proves Assets − (L&E) = 0 |
| `KPI Dashboard` | Annual revenue, net income, net margin, current assets/liabilities, current ratio, debt, debt-to-equity |
| `ChartData` | Helper series feeding the charts |
| `Charts` | 3 native Excel charts: monthly revenue vs expenses, net-income trend, annual expense breakdown |

## How to run

```bash
python3 make_sample_data.py     # regenerate the sample AX export (deterministic)
python3 ax_dashboards.py        # builds AX_2012_Financial_Dashboards.xlsx
python3 -m unittest test_ax_dashboards -v   # 33 tests, incl. LibreOffice recalc
```

Requirements: Python 3.8+, `openpyxl`. The recalc tests need LibreOffice
(`soffice`) for the headless formula-verification roundtrip.

To use your own export: keep the CSV columns
`FiscalYear, Period, MainAccount, AccountName, AccountCategory, Debit, Credit`
with `AccountCategory` in {Revenue, Expense, Asset, Asset-Contra, Liability,
Equity}, then run `python3 ax_dashboards.py your_export.csv out.xlsx`.
Edit the `SIGN` map in `ax_dashboards.py` if your chart of accounts uses
different normal balances.

## Design notes

- **Double-entry by construction.** The sample data is built like a real
  ledger: Cash & Bank is the residual account, so every period's debits equal
  credits exactly — no fudge factors. Retained earnings carries the opening
  balance; current profit rolls in on the balance sheet, which is why the
  check row reads zero.
- **No macros, no plugins.** Plain formulas + native charts, so the workbook
  opens in Excel, Google Sheets, and LibreOffice.
- **Deterministic.** Same input → byte-identical CSV; tests assert the trial
  balance proves out and recompute every KPI independently.

## Limitations

- Snapshot model: one workbook per fiscal year; multi-year comparison would
  need an extra `FiscalYear` dimension in the `SUMIFS` criteria.
- Charts are static ranges (12 periods); extending beyond a year means
  extending the `ChartData` references.
- This is a demo built on synthetic data, not a client engagement.
