# AquaBill — Automated Excel Billing, Sales & Inventory System

**Built for:** a water-bottle distribution business that needs billing, sales
tracking and inventory in one refresh-proof Excel workbook
([live demo](https://github.com/maticcretic-commits/daily-automation-solutions/tree/main/solutions/2026-10-13-water-bottle-billing-inventory)).

## The problem

A packaged-drinking-water distributor bills from notebooks, tracks stock by
hand, and discovers stock-outs when a customer complains. Invoices are
re-typed, GST is calculated on a phone calculator, and nobody can answer
"what did we sell last month?" without a day of manual work.

## The approach

One workbook where **every figure is a live formula** over the raw tables —
nothing hard-coded, nothing to re-type:

- **InvoiceLines** — paste one row per line item; line totals compute themselves
- **Invoices** — the register: subtotal, 18% GST, grand total, paid, balance and
  Paid / Partial / Unpaid status all derive from the lines
- **Inventory** — per-SKU opening stock, stock-in, sales-out, closing stock,
  reorder level, and an OK / REORDER / OUT OF STOCK flag
- **Purchases** — stock-in lots feed the inventory automatically
- **GST_Summary** — monthly taxable value with CGST 9% + SGST 9% split
- **Daily** — daily units and revenue trend
- **Dashboard** — 8 KPI cards + monthly sales chart + units-by-SKU chart
- **Checks** — 8 integrity checks that must all read TRUE (register math,
  line pricing, tax, stock, duplicate invoices, receivables, GST tie-out,
  daily tie-out)
- **Guide** — operator handover: how to add invoices, record stock and
  payments, change the GST rate

Paste new rows and the dashboard, tax summary and stock levels update —
the Checks sheet proves it still ties out.

## How to run

```bash
python3 generate.py            # builds deterministic demo data in data/
python3 build_workbook.py      # writes water-bottle-billing-inventory.xlsx
python3 -m pytest tests/ -q    # 15 tests, incl. LibreOffice recalc check
```

Requires Python 3, `openpyxl`, `pytest`, and LibreOffice (`soffice`) for the
recalc test. The data generator is seeded (`--seed 20261013`), so every run
produces identical data.

## Verification

15/15 tests pass, including a LibreOffice headless recalculation of the real
`.xlsx` compared against an independent Python ground truth: all 8 dashboard
KPIs, monthly GST, per-SKU closing stock and statuses match, zero formula
errors, and all 8 integrity checks read TRUE. A static scan also guards the
known `$H$2:$5000`-style invalid range-literal bug class across every formula
cell.

## Notes

- All customers, transactions and figures are **fictional demo data** for
  evaluation (AquaPure Distributors, Patna; Apr–Sep 2026).
- GST treatment shown is 18% intra-state (CGST 9% + SGST 9%); change the rate
  with find/replace per the Guide sheet.
- Reference implementation (demo, not client work). Part of
  [DashboardForge](https://github.com/maticcretic-commits/dashboardforge).
