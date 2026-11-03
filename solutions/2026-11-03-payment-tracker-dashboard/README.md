# PaymentTracker — Automated Excel Payment Tracker

Built around a real Freelancer brief: turn a raw customer-payment table
into one clean Excel workbook — formulas for the pending amount and for
paid / partial / unpaid status, conditional-format colour cues, a
real-time summary of total due / paid / pending, and new rows addable
without breaking any formula. One `.xlsx`, no macros, nothing that
triggers a macro warning.

Demo data is fictional: Sharma Traders, a building-materials wholesaler
in New Delhi selling on 30-day credit, Apr–Oct 2026, 120 invoices across
28 customers, as of 3 Nov 2026. The generator is seeded and
deterministic — re-running it reproduces the same data byte for byte.

## What the workbook does

- **Payments** — the entry table. Columns A–F are typed/pasted values
  (customer, invoice no, invoice date, due date, invoice amount,
  amount paid). Columns G–L are formulas, pre-filled for 200 headroom
  rows past the data (to row 321): pending amount, status, days
  overdue, month key, an overdue-row helper, and an age bucket.
  Paste a new row into A–F inside the headroom and every summary
  updates; blank rows stay blank, never errors.
- **Dashboard** — 11 KPIs (total invoiced, collected, pending,
  collection rate, invoice counts, overdue count/amount, customers,
  average days overdue, partially-paid count), a status-mix table, and
  3 native Excel charts: invoiced vs collected by month, invoices by
  status, cumulative collected.
- **Customers** — one row per customer: invoices, invoiced, paid,
  pending, overdue amount/count, collection %, plus a TOTAL row.
- **Monthly** — invoiced / collected / pending / collection % /
  cumulative collected / invoice count by invoice month, plus TOTAL.
- **Overdue** — a live list of the 50 earliest-due outstanding
  invoices, pulled with `SMALL` + `INDEX` over a per-row helper column
  (the LibreOffice-safe pattern — no array `IF` inside `SMALL`).
  It re-sorts itself as rows change; 57 invoices are overdue in this
  data, so the last 7 fall outside the 50-row list by design.
- **Assumptions** — the as-of date (3 Nov 2026). Change it and every
  status, overdue count and overdue amount re-ages.
- **Checks** — 10 integrity checks (dashboard/monthly/customers
  tie-outs, status counts, no negative pending amounts). All read PASS.
- **Guide** — how to add rows and the status rules, in the workbook.

Status rules: Paid = nothing pending. Partial = something paid, balance
outstanding, due date not passed. Unpaid = nothing paid, due date not
passed. Partial – Overdue / Unpaid – Overdue = balance outstanding and
the due date has passed the as-of date.

## Verified figures (LibreOffice headless recalc vs independent Python)

- Invoiced Rs 11,426,700 · collected Rs 6,916,500 · pending Rs 4,510,200
- Collection rate 60.5% · 120 invoices · 28 customers
- Overdue: 57 invoices, Rs 4,197,300 · average 85.1 days overdue
- Status mix: 53 Paid · 7 Partial · 3 Unpaid · 38 Partial – Overdue ·
  19 Unpaid – Overdue
- Zero formula errors in the recalculated workbook; all 10 Checks PASS.

## Files

- `generate.py` — deterministic synthetic data → `data/payments.csv`
- `build_workbook.py` — builds `payment-tracker.xlsx` (openpyxl, live
  A1-style formulas only, no Table objects, no macros)
- `main.py` — independent ground truth recomputed from the CSV
- `tests/test_build.py` — builds, recalcs in LibreOffice, and compares
  every checked figure; also scans the builder for short-form range
  literals (`$H$2:$5000`) and asserts headroom formulas, 3 charts and
  conditional-formatting rules are present
- `powerbi/MEASURES.md` — DAX mapping for rebuilding the same model
  as a `.pbix`
- `data/payments.csv` — the 120-invoice source table
- `payment-tracker.xlsx` — the finished workbook

## How to run

```sh
python3 generate.py          # regenerate data/payments.csv (seeded)
python3 build_workbook.py    # build payment-tracker.xlsx
python3 tests/test_build.py  # full verification (needs LibreOffice)
```

Requires Python 3.11+ and `openpyxl`. LibreOffice is needed only for
the recalc verification step.

## Reference

Freelancer brief: "Automated Excel Payment Tracker" (Project ID
40796333, New Delhi) — Excel workbook with formulas for pending amount
and paid/partial/unpaid status, conditional formatting, real-time
summary, new rows addable without breaking formulas, one `.xlsx`, no
macros. This is a reference implementation built from fictional data,
not client work.
