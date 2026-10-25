# QBO Owner Dashboard — Construction Financial Dashboard demo

**The client problem (live freelance brief, 2026-10-25):** a construction
business's outgoing CFO kept a monthly custom dashboard combining QuickBooks
Online exports and spreadsheets. The owner needs a replacement that is simple
to read, stays in sync with the books, and survives the CFO's departure —
no hard-coded numbers that quietly go stale.

**What this builds:** a formula-driven Excel workbook for a fictional
construction firm (Maple Ridge Builders Ltd., Victoria BC, CAD) from
QuickBooks-style exports:

- `Guide` — what the workbook is, how to refresh it, the rules it follows.
- `Dashboard` — 6 KPI cards (YTD revenue, gross margin %, cash on hand,
  AR/AP outstanding, AR overdue 61+ days) + 3 native charts (monthly
  revenue vs costs, gross-margin trend, AR aging mix).
- `Monthly_PL` — May–Oct 2026 P&L; every figure a live SUMIFS over Raw_QBO.
- `Raw_QBO` — the pasted QBO P&L export (48 rows).
- `AR_Aging` / `AP_Aging` — invoice/bill rows with formula outstanding and
  aging buckets (Current / 1-30 / 31-60 / 61+) computed off the As-of date,
  plus bucket summaries.
- `Cash` — opening balance + AR collected − AP paid = cash on hand.
- `Checks` — 8 live tie-out checks (PL vs raw, bucket totals vs ledgers,
  dashboard vs sources, margin bounds, no negative revenue month).

**Verified figures (LibreOffice headless recalc vs independent Python ground
truth, 6/6 tests pass):** YTD revenue CAD 1,150,273; gross profit CAD
153,694 (13.4% margin); cash on hand CAD 164,858; AR outstanding CAD
173,005; AP outstanding CAD 65,464; all 8 Checks TRUE; zero formula errors.

**How to run:**

```bash
python3 main.py            # regenerates data/*.csv, builds qbo-owner-dashboard.xlsx
python3 -m unittest discover -s tests   # 6 tests, incl. LibreOffice recalc
```

**Refresh contract:** paste a new month's QBO export into `Raw_QBO`, new
invoices/bills into `AR_Aging`/`AP_Aging`, update the As-of date — the
dashboard follows. All totals are SUM/SUMIFS formulas (plain A1 ranges with
headroom, no Table objects, no structured references) so the workbook
recalculates identically in Excel and LibreOffice.

**Data note:** every figure is SYNTHETIC demo data (seed 20261025), invented
for this demo — not a real client's books.
