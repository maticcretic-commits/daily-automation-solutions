# QB2BI — QuickBooks Data → Relational DB → Live Financial Reports

**Built:** 2026-10-16 · **Pattern:** the "QuickBooks Data Integration for PowerBI" freelance gig —
a client wants their QuickBooks data flowing into live financial reports instead of
manual month-end spreadsheets.

## What it does

1. **Normalise** — QuickBooks-style double-entry journals land in a relational
   SQLite database (`accounts` + `journal_lines`), the same shape a Power BI
   model or data warehouse would consume.
2. **Report** — `qb2bi-finance-dashboard.xlsx` renders trial balance, monthly P&L,
   balance sheet, a 9-KPI dashboard with 3 native charts, a rule-based management
   narrative, and 8 automated integrity checks — **every summary cell is a live
   Excel formula** (`SUMIFS` over the raw journals), so new months flow through
   with zero manual work.

## The demo data

Fictional business: **Agarwal Traders**, a wholesale FMCG distributor (Patna, India).
18 months (Apr-2025 → Sep-2026): 447 balanced journal entries, 1,116 lines —
sales invoices with GST, collections, purchase bills, vendor payments, payroll,
rent, loan interest + principal, depreciation. All seeded (`seed=42`), fully
reproducible. **Synthetic data — for demo purposes only.**

## How to run

```bash
python3 generate.py        # synthesise data/*.csv
python3 build_db.py        # load into qb2bi.sqlite (asserts debits == credits)
python3 build_workbook.py  # render the live-formula workbook
pytest -q                  # 10 tests incl. LibreOffice headless recalc
```

Requires: Python 3.10+, `openpyxl`, `pytest`, LibreOffice (`soffice`) for the
recalc test.

## Sheet map

| Sheet | Contents |
|---|---|
| Guide | Pipeline overview + run instructions |
| Raw_Journals | Every journal line + `Period` helper column (formula) |
| TrialBalance | Per-account debit/credit/net via `SUMIFS` |
| PL | Monthly P&L Apr-25…Sep-26 + YTD column |
| BalanceSheet | As of Mar-26 and Sep-26, with Assets = Liabilities + Equity check |
| Dashboard | 9 live KPI cards + 3 native charts (revenue vs profit, expense mix, balance-sheet compare) |
| AI_Report | Rule-based management narrative (trend, drivers, cash, recommendations) |
| Checks | 8 automated integrity checks — all read **PASS** |

## Verification

`tests/test_build.py` re-saves the workbook through LibreOffice headless and
compares every key figure against an independent Python ground truth computed
directly from the CSVs: **10/10 tests pass** — YTD revenue, COGS, expenses and
net profit match to the rupee; total assets = total liabilities + equity;
all 8 checks PASS; trial balance ties to the ledger.

## Files

- `generate.py` — synthetic data generator
- `build_db.py` — CSV → SQLite relational model
- `build_workbook.py` — workbook renderer (all formulas, no hard-coded totals)
- `tests/test_build.py` — pytest suite
- `data/` — chart_of_accounts.csv, journals.csv, customers.csv, vendors.csv
- `qb2bi.sqlite` — the relational database
- `qb2bi-finance-dashboard.xlsx` — the deliverable
