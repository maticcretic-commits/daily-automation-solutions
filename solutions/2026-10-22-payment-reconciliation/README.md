# ReconPulse — Bank Reconciliation Workbench

A formula-driven Excel bank reconciliation for a fictional wholesale
distributor (*Mehta Trading Co., Ludhiana* — demo data, 1–21 Oct 2026).

## The problem

Month-end bank reconciliation is usually a manual, error-prone hunt:
a 128-line bank statement vs a 127-line book ledger, with clearing lags
(bank dates trail booking dates by 0–2 days), bank charges and interest
the books never recorded, cheques issued but not yet presented, and the
odd unidentified deposit. This workbook automates the matching and walks
both sides to one agreed balance — every figure a live formula.

## How it works

1. **Raw_Bank / Raw_Book** — paste a fresh statement and ledger here
   (running balances are formulas off the opening balance on Summary).
2. **Recon** — one row per bank line. Column F builds a match key
   `amount|side` (`R` = receipt, `P` = payment); column G `MATCH`es it in
   the ledger; column K marks **MATCHED** only when the booked date is
   within ±3 days of the bank date (the clearing window).
3. **Unmatched** — bank-only items (charges, interest, unidentified
   deposits) and book-only items (unpresented cheques, deposits in
   transit), each with a reason, a suggested action, and days outstanding.
4. **Summary** — the two-sided reconciliation walk: bank balance
   ± bank-only items and book balance ± book-only items must converge to
   the same adjusted balance (difference = ₹0), plus KPIs.
5. **Dashboard** — 8 KPI cards and 3 native Excel charts (match pie,
   unmatched-by-reason bar, daily bank-vs-book balance line).
6. **Checks** — 8 self-tests; all must read TRUE.

Demo state: 124 of 128 bank lines auto-matched; 4 bank-only and 3
book-only items deliberately unmatched; adjusted balance ₹25,78,990 on
both sides; oldest outstanding item 3 days.

## Repo layout

- `generate.py` — deterministic demo-data generator (seed 20261022) →
  `data/bank_statement.csv`, `data/book_ledger.csv`
- `build_workbook.py` — builds `reconpulse.xlsx` (openpyxl; stdlib + openpyxl only)
- `main.py` — independent Python ground truth (matching + all KPIs)
- `tests/test_build.py` — 11 tests, incl. LibreOffice headless
  recalculation of the workbook compared cell-by-cell against `main.py`
- `data/` — the two input CSVs

## Run it

```bash
python3 generate.py        # regenerate demo data
python3 build_workbook.py  # build reconpulse.xlsx
python3 main.py            # print ground-truth KPIs
python3 -m pytest tests/ -q
```

## Compatibility notes

- No `XLOOKUP`/`MAXIFS`/`MINIFS` — `MATCH`/`INDEX`/`COUNTIFS`/`SUMIFS` only.
- Extraction lists use a numeric helper column + `SMALL()`; the
  `SMALL(IF(range="x",…))` array pattern silently mis-evaluates in
  LibreOffice (implicit intersection) — see build notes.
- No macros, no external links. Opens in Excel, LibreOffice, and
  Google Sheets.
