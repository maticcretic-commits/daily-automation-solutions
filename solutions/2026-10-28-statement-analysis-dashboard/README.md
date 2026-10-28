# StatementPulse — Financial Statement Analysis Dashboard (Demo)

Built around the Freelancer lead **"Financial Statement Analysis Needed ASAP"**
(Brisbane client): decision-ready analysis from recent financial statements —
ratios, trends, cash-flow health, and red flags for investors/management.

## The problem

A set of financial statements on its own does not answer the questions a
review meeting asks: is the business profitable, liquid, and actually
generating cash — and where are the warning signs? The usual failure is a
polished summary whose numbers cannot be traced back to the underlying
statements, so the first hard question kills it.

## The approach

One Excel workbook where every figure is a live formula pointing at raw
line items — nothing hard-coded:

- **Statements** — P&L, balance sheet, and indirect cash flow rebuilt from
  the raw line items. Subtotals, the balance check (assets − liabilities −
  equity = 0), and the cash-flow tie-out (closing cash = balance-sheet cash)
  are all formulas. FY2024 working-capital movements compare against the
  stored opening balance sheet.
- **Ratios** — liquidity, leverage, profitability, efficiency, and cash
  conversion, each benchmarked with an FY2026 OK/FLAG status (e.g. current
  ratio ≥ 1.5, receivables days ≤ 60, CFO/PAT ≥ 1.0).
- **Trend** — FY2026 monthly build-up (revenue, gross profit, margin,
  working capital); monthly totals tie exactly to the annual statements.
- **Dashboard** — six FY2026 KPI cards, a revenue-vs-PAT chart, a monthly
  revenue chart, and an FY2026 cost-mix chart.
- **Checks** — eight integrity checks, all TRUE (balance-sheet identity and
  cash-flow tie-out for all three years, monthly tie-out, retained-earnings
  roll-forward).

The demo data is **synthetic**: fictional Sharma Retail Pvt. Ltd., Patna,
FY2024–FY2026, deterministically generated (seed 20261028). The balance
sheet is constructed so the accounting identities hold — retained earnings
roll forward, PPE rolls forward, and cash is the balancing plug, which the
indirect cash flow then reproduces exactly.

## How to run it

```sh
python3 generate.py          # rebuild data/*.csv (deterministic)
python3 build_workbook.py    # rebuild statementpulse.xlsx
python3 tests/test_build.py  # rebuild + LibreOffice headless recalc vs
                             # independent Python ground truth (main.py)
```

Requires Python 3 with `openpyxl`, plus LibreOffice (`soffice`) for the
recalc test. The test asserts zero formula errors and that every checked
figure matches the independent ground truth.

## Verified results (FY2026)

Revenue ₹60,463,200 · PAT ₹5,544,480 · net margin 9.2% · current ratio 3.04 ·
receivables 58 days · CFO/PAT 0.77. The cash-conversion flag (below 1.0) and
receivables days (58, near the 60-day benchmark) are the two watch items the
dashboard surfaces — exactly the kind of red flag the lead asks for.

## Files

- `generate.py` — deterministic synthetic data generator
- `build_workbook.py` — workbook builder (all formulas, A1-style ranges)
- `main.py` — independent ground truth (never reads the workbook)
- `tests/test_build.py` — LibreOffice recalc verification
- `data/statements.csv`, `data/monthly.csv` — source data
- `powerbi/MEASURES.md` — DAX mapping for a Power BI version
- `statementpulse.xlsx` — the finished workbook
