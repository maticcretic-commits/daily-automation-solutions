# Bank Statement Consolidation Pack

Demo built for the Freelancer project **“Bank Statement Excel Consolidation”**
(Nagpur client, ₹12,500–37,500 fixed): merge a stack of bank statements
into one clean unified worksheet — every transaction exactly as in the
source, balances to the cent.

## The problem

A business with several bank accounts receives one statement file per
account. Reconciling them by hand means retyping, lost rows, and closing
balances that never tie out. The brief asks for a single tab with every
transaction exactly as in the source, and balances that reconcile to the
cent.

## The demo

Fictional business **Sharma Traders, Patna** (wholesale FMCG trader),
Jan–Oct 2026, three accounts:

| Account | Rows | Credits | Debits | Closing |
|---|---|---|---|---|
| Canara Bank, current ****4821 | 58 | ₹30,46,374.97 | ₹28,22,825.23 | ₹6,48,549.74 |
| State Bank of India, savings ****7734 | 46 | ₹15,70,310.62 | ₹15,18,946.98 | ₹2,31,363.64 |
| HDFC Bank, current ****9156 | 54 | ₹19,34,713.16 | ₹14,36,880.61 | ₹11,07,832.55 |

Totals: **158 transactions, credits ₹65,51,398.75, debits ₹57,78,652.82,
closing ₹19,87,745.93.** (Figures are from the verified test run; they
move if the data is regenerated — the workbook always ties to its own
source files.)

Three own-account transfers (Canara→SBI ₹1,50,000; SBI→HDFC ₹80,000;
HDFC→Canara ₹2,00,000) appear in two statements each and net to zero.

## What is in the workbook (`bank-statement-consolidation-pack.xlsx`)

- **Guide** — what each sheet does and how to refresh the pack.
- **Sources** — register of the three source files: rows, opening,
  credits, debits and computed closing per account, all live
  SUMIFS/COUNTIFS over Transactions, with an OK/MISMATCH check against
  each file's own closing balance.
- **Transactions** — the unified ledger: all 158 transactions,
  date-sorted, with category, per-account sequence number and a live
  running balance per account (opening + credits − debits up to that
  row). “Final row” marks each account's last transaction; “Dup flag”
  marks any exact repeat within the same account.
- **Monthly** — opening / credits / debits / closing per account per
  month, all SUMIFS; each month's closing feeds the next, and October
  closings tie to Sources.
- **Dashboard** — focus-account and focus-month dropdowns drive the
  KPIs (credits, debits, net flow, closing balance, transaction count,
  largest receipt/payment), a monthly net-flow table, a by-category
  table, and three native charts (monthly net flow, credits vs debits
  by account, debits by category).
- **Checks** — 10 live reconciliation checks, all PASS: row counts,
  credit/debit totals, final running balance vs each source closing,
  transfers netting to zero, zero duplicates, zero uncategorised rows,
  October monthly closings = total closing.

## Files

- `make_data.py` — generates the three fictional statements in
  `data/statements/` (deterministic, seed 20261115). Random debits are
  capped so a balance never drops below ₹5,000 plus any transfer-out
  still to come, so transfer legs stay exact to the rupee.
- `generate.py` — builds the workbook from `data/statements/*.csv` and
  writes `data/consolidated.csv`, the merged ledger as a flat export.
- `data/statements/canara.csv`, `sbi.csv`, `hdfc.csv` — the source
  statements (Date, Description, Debit, Credit, Balance).
- `data/consolidated.csv` — the merged ledger export.
- `tests/test_build.py` — rebuilds everything, then verifies with a
  LibreOffice headless recalc against an independent Python
  recomputation from the CSVs.

## How to run

```bash
python3 make_data.py
python3 generate.py
python3 tests/test_build.py   # ALL TESTS PASS
```

Requires Python 3 with `openpyxl`, and LibreOffice (`soffice`) for the
recalc test.

## Notes

- All money cells are live formulas; the only hard-coded values are the
  source CSV figures themselves.
- Fictional data for demonstration only — no real account numbers, no
  real transactions.
- Largest receipt in the demo is a ₹2,00,000 own-account transfer-in;
  transfers are listed as their own category so income and expense
  totals are never inflated by internal movement.
