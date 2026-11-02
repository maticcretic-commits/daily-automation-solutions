# FinancePulse — custom financial dashboard (demo)

A manual-entry personal financial dashboard, built around a real
Freelancer brief: *“Custom Financial Dashboard Build”* — an interactive
dashboard (Excel, Google Sheets, Power BI, Tableau or a lightweight web
app) showing income / expenses / cash balance, sales / costs / profit,
and investments / returns / portfolio value, with unlocked formulas and
a hand-off guide.

This is a **demonstration built from invented data** (a fictional
household, “A. Sharma, Patna”, Jan–Oct 2026), not client work. It shows
the exact workbook the brief asks for, with the discipline the brief
cares about: every figure is a live formula, nothing is pasted as a
value, no sheet is protected, and no macro is used — so the owner can
inspect, trust and extend every number.

## What it contains

`finance-dashboard.xlsx` — 9 sheets, every figure a live formula:

- **Guide** — how to use the workbook: what to type, what never to
  type over, and how the monthly routine works.
- **Dashboard** — 12 KPIs plus 3 native charts (income vs expenses by
  month, cumulative cash balance, expense mix).
- **Monthly** — month-by-month income, expenses, net, sales revenue,
  sales cost, gross profit, cumulative cash and savings rate, all
  `SUMIFS` over the entry sheets; cumulative cash starts from the
  opening balance in Assumptions.
- **Income** (32 rows), **Expenses** (160 rows), **Sales** (60 rows) —
  the only sheets you type into. Amount, revenue, cost, profit and
  month-key columns are formulas, pre-filled for 200 extra headroom
  rows, so new manual entries pasted inside the headroom are picked
  up with no formula editing.
- **Investments** — 8 holdings (ETFs, shares, deposits, funds) with
  invested / value / gain / gain % all formula-driven, plus a TOTAL row.
- **Assumptions** — opening cash Rs 125,000, currency, period, headroom.
- **Checks** — 10 live checks, all PASS after a headless LibreOffice
  recalc (dashboard ties to raw sums, cash articulates from opening +
  net, sales and portfolio totals tie, no negative amounts, expense
  mix reconciles).

Headline demo figures: income Rs 1,067,200 · expenses Rs 558,766 ·
net Rs 508,434 (47.6% savings rate) · closing cash Rs 633,434 ·
sales revenue Rs 241,074, cost Rs 75,680, gross profit Rs 165,394 ·
portfolio value Rs 522,485 on Rs 501,066 invested (gain 4.3%).

## How to run it

```sh
python3 generate.py          # rebuild data/*.csv (seeded, deterministic)
python3 build_workbook.py    # rebuild finance-dashboard.xlsx
python3 tests/test_build.py  # rebuild + LibreOffice recalc + ground-truth checks
```

Requirements: Python 3 + openpyxl; LibreOffice (`soffice`) for the
recalc test only. `main.py` is the independent Python ground truth the
tests compare against — it reads the same CSVs and re-derives every
checked figure without touching the workbook.

`powerbi/MEASURES.md` maps the same KPIs to DAX measures for the Power
BI version of this dashboard, since the brief lists Power BI as an
accepted delivery format.

## Manual-entry routine (the deliverable's real value)

1. Type each new income, expense or sale into the next free row of the
   matching entry sheet, inside the pre-filled headroom. Only type the
   Date / Source / Category / Amount (or Product / Qty / prices) —
   never type over a formula column.
2. On Investments, edit Units / Buy price / Current price in place;
   insert a new holding above the TOTAL row.
3. Read Monthly and Dashboard; if any Checks row reads FAIL, find the
   broken entry — never paste values over a formula to force it green.
4. The workbook stays macro-free and unprotected, so the owner (or
   their accountant) can audit any figure in the formula bar.
