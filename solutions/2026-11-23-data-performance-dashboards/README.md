# PerfDash — Data Performance Dashboards

A working demo built around the Freelancer brief **"Excel Data Performance
Dashboards"** (posted 23 Nov 2026, $15–25/hr, 6 proposals at scout time):
Excel/CSV files turned into KPI dashboards with PivotTables-style summaries,
fully editable, with a walkthrough. Fictional data: Ganga Retail &
Services Pvt. Ltd., Patna — an FMCG wholesaler/retailer selling across
five channels, Apr–Oct 2026.

## What it does

`performance-dashboards.xlsx` is a genuine formula-driven workbook — every
computed cell is a live Excel formula, so changing any raw value refreshes
the whole book:

- **Guide** — the walkthrough: how to paste new daily rows, how the
  focus-month dropdown works, a sheet map, metric definitions, and what
  the data says.
- **Raw** — 495 daily channel rows (date, channel, product, campaign,
  impressions, clicks, orders, revenue, ad spend) plus live derived
  columns: Month, CTR, CPC, ROAS, AOV — and a TOTAL row.
- **Monthly** — revenue / orders / spend by month (SUMIFS), with target
  vs actual, variance (₹) and variance %.
- **Channel** — the same cut by sales channel, with revenue share.
- **Campaign** — the same cut by campaign, with orders share.
- **Targets** — monthly revenue targets.
- **Dashboard** — a focus-month dropdown (data validation) driving 7 KPIs
  (revenue, orders, ad spend, ROAS, CTR, AOV, best channel) plus a
  monthly table and 3 native charts: monthly revenue, revenue share by
  channel, ROAS by campaign.
- **Checks** — 10 integrity checks against the independent Python ground
  truth; all read PASS.

## Verified figures (LibreOffice headless recalc vs ground truth)

- Total revenue Apr–Oct 2026: **₹31,21,779**; ad spend ₹5,89,468;
  blended ROAS **5.30x**.
- Best channel: **Walk-in** (₹10,05,024 revenue, ROAS 8.96) — about a
  third of revenue with almost no ad spend. Paid Campaigns return the
  least per rupee (ROAS 2.51) and deserve the first budget review.
- Festive peak: **October 2026, ₹5,97,659** (Diwali Early-Bird) — but
  still 3.8% under target; the miss is a demand-planning note, not a
  dashboard error. Three of seven months beat target, four missed.
- All 10 Checks PASS; zero formula errors on recalculation.

## Files

- `generate.py` — builds the workbook from the CSVs (openpyxl).
- `make_data.py` — generates the fictional data (seed 20261123) and the
  independent ground truth.
- `data/performance_daily.csv` — the 495 daily rows.
- `data/targets.csv` — monthly revenue targets.
- `data/ground_truth.json` — independent Python aggregates.
- `tests/test_build.py` — 19 tests: data sanity, formula hygiene
  (no short range literals, no structured references), and a LibreOffice
  headless recalc of every KPI, monthly/channel/campaign row and check
  against the ground truth.

## Run it

```bash
python3 make_data.py      # regenerate data + ground truth
python3 generate.py       # rebuild performance-dashboards.xlsx
python3 -m unittest discover -s tests -v
```

No third-party packages beyond openpyxl are needed to build; the recalc
test needs LibreOffice (`soffice`).

## Reference implementations (DashboardForge)

- [Excel Dashboard Factory](../2026-10-04-excel-dashboard-factory/) —
  raw sales CSV to a multi-sheet dashboard with live formulas and charts.
- [Collections Pulse Dashboard](../2026-10-05-collections-pulse-dashboard/) —
  NBFC loan-ledger collections dashboard with a priority follow-up queue.
- [Sales-Purchase MIS Factory](../2026-10-12-sales-purchase-mis-factory/) —
  multi-workbook sales/purchase consolidation with anomaly flags.
