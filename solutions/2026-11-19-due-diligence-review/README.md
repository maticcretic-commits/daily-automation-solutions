# DueDiligence Review Pack — Pre-Investment Due Diligence Review

A formula-driven pre-investment due-diligence review workbook for a
fictional Indian private limited company, **Sutlej Agro Foods Pvt. Ltd.,
Ludhiana** (food processing), FY24–FY26, as of 31 Mar 2026.

Built as a working demo for the Freelancer project **"Pre-Investment
Due Diligence Review"** — early-stage due diligence on an Indian private
limited company before a term sheet: ROC/MCA/shareholding/FEMA-FDI
compliance plus 3 years' audited financials, GST/income-tax filings,
bank statements, loans and contingent liabilities, red-flag memo in
3 working days, full report in 7.

**ALL DATA IS FICTIONAL** — invented for the demo. No real company, no
real filings, no real bank statements.

## What it does

- **Financials** — 3-year P&L and balance sheet (Rs lakh). Gross profit,
  EBITDA, EBIT, PBT, net income, net worth and both totals are live
  formulas. Revenue grows 1,840 → 2,150 → 2,480.
- **Ratios** — 12 ratios x 3 years plus a growth block (revenue CAGR,
  cash growth, contingent liabilities as % of net worth, receivables
  vs revenue growth). Every cell is a live formula.
- **BankRec** — 12-month book-vs-bank reconciliation from the cash
  book, plus the open items at the as-of date with days outstanding
  (live). Book closes at 175; bank at 173.45; the 1.55 difference is
  exactly the open items (cheque issued 4.20, cheque deposited 8.40,
  bank charges 0.15, UPI collection 2.80).
- **Loans** — 5 loan facilities (term loan, cash credit, vehicle loan,
  NBFC unsecured loan, 0% director loan repayable on demand) with
  long-term / current split totalling 950, equal to FY26 borrowings;
  plus contingent liabilities of 130 (24.8% of net worth).
- **RedFlags** — 9 scored red flags. Every status is a live trigger
  formula reading the other sheets — change a number and the flags
  move. Risk score **31 = HIGH**.
- **Dashboard** — target, as-of date, 15 KPIs, severity table and
  3 native charts (revenue vs EBITDA; current vs quick ratio;
  flags by severity).
- **Checks** — 10 integrity checks vs the independent Python ground
  truth; all read PASS.

## Headline verdict (live formula)

Revenue CAGR **16.1%** over FY24–FY26, but receivable days climb every
year (**57.5 → 61.1 → 65.5**), quick ratio stays under 1.0x
(**0.84 / 0.77 / 0.66**), interest cover never reaches 2.0x
(**1.70 / 1.90 / 1.84**), contingent liabilities run to **24.8%** of
net worth, and a **Rs 120 lakh director loan at 0%** sits repayable on
demand. Risk score **31 — HIGH: proceed only with price protection
and conditions**.

## Files

- `due-diligence-review.xlsx` — the workbook (8 sheets)
- `generate.py` — builds the workbook from `data/`
- `make_data.py` — generates the fictional CSVs (deterministic,
  seed 20261119)
- `data/` — `financials.csv`, `bankrec.csv`, `bank_items.csv`,
  `loans.csv`, `contingent.csv`, `redflags.csv`, `meta.csv`
- `tests/test_build.py` — rebuilds, then LibreOffice headless recalc
  vs the Python ground truth

## Tests

```bash
python3 tests/test_build.py
```

Verifies, against an independent Python re-implementation from
`data/`: derived P&L/balance-sheet lines, all 12 ratios plus the
growth block, the 12-month bank reconciliation (book balances from
receipts/payments, bank balances from the item ledger), loan totals
and LT/current split, contingent liabilities, all 9 red-flag triggers,
the risk score and verdict — zero formula errors, all 10 Checks PASS.
