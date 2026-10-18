# InvoiceLedger — invoice-to-reports workbook

A working demo for the "Accurate Finance Data & Reports" brief: turn a pile of
invoices and receipts into management-ready **P&L**, **cash-flow** and
**receivables-aging** reports in one formula-driven Excel workbook.

## The problem it solves

Small trading businesses live in invoice PDFs and spreadsheets, but month-end
still means a day of manual copy-paste: revenue by product line, what was
actually collected vs merely billed, which customers are overdue, and how much
GST is owed. One wrong paste and the P&L silently misstates the month.

## The approach

`generate.py` creates a realistic fictional dataset (787 invoices, 657
payments, Oct-2025 → Sep-2026) for "Gupta Garments & Textiles", a wholesale
garments trader — with seasonal demand, late payers, and partial payments.

`build_workbook.py` then builds `invoice-ledger-reports.xlsx` where **every
computed cell is a live Excel formula** — nothing is hard-coded:

- **Guide** — how to use it, data dictionary, report date (aging pivots on it)
- **Raw_Invoices / Raw_Payments** — paste new rows here; everything recalculates
- **P&L** — monthly revenue by line (Fabrics/Garments/Accessories), COGS,
  expenses, margins, plus a monthly GST output/input/net-payable block
- **CashFlow** — receipts vs payments on cash dates, net and cumulative flow
- **Receivables** — outstanding invoices with days-outstanding and 0–30/31–60/
  61–90/90+ aging buckets, avg collection days
- **Dashboard** — 6 KPI cards and 3 native charts (revenue vs expenses,
  category mix, monthly net cash flow)
- **AI_Report** — rule-based narrative and ✓/⚠ flags (receivables ratio,
  margin health, overdue share, cash dips, collection speed, concentration)
- **Checks** — 8 self-auditing formulas (revenue tie-out, profit arithmetic,
  cash tie-out, aging tie-out, no negative balances, no over-payments)

## How to run it

```bash
python3 generate.py        # -> data/invoices.csv, data/payments.csv
python3 build_workbook.py  # -> invoice-ledger-reports.xlsx
python3 -m unittest discover -s tests   # 14 tests, incl. LibreOffice recalc
```

Tests verify CSV integrity, workbook structure, and — via a headless
LibreOffice recalculation — that every formula result matches an independent
Python ground truth (FY/monthly P&L, cash flow, aging buckets, GST, KPIs;
all 8 checks TRUE).

Requirements: Python 3.10+, `openpyxl`. LibreOffice only needed for the
recalc test.

## Files

| File | Purpose |
|---|---|
| `generate.py` | Fictional invoice/payment dataset generator (seeded, reproducible) |
| `build_workbook.py` | Builds the formula-driven `.xlsx` |
| `tests/test_build.py` | 14 tests incl. headless-LO recalc vs ground truth |
| `data/invoices.csv`, `data/payments.csv` | Generated demo data |
| `invoice-ledger-reports.xlsx` | The workbook |

All data is fictional and generated for demonstration.
