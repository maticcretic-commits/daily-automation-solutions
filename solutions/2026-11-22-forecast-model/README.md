# ForecastForge — 3-Year Financial Forecast Model

A formula-driven 3-year financial forecast workbook, built around the
Freelancer lead **"3 Year Financial Forecast Model"** (Bodh Gaya client,
₹12,500–37,500 fixed, avg bid ₹19,278): three-year P&L, balance sheet
and cash-flow forecast in Excel, with assumptions, ratios, break-even
and sensitivity analysis.

Fictional company: **Ganga Foods Pvt. Ltd., Bodh Gaya** — packaged
foods, Apr 2026 – Mar 2029 (FY2027 / FY2028 / FY2029). All figures are
deterministic and reproducible from `data/assumptions.csv`.

## What it builds

`forecast-model.xlsx` — 9 sheets, no macros, every figure a live
A1-style formula (LibreOffice-safe):

| Sheet | Contents |
|---|---|
| Guide | How to use it, what each sheet does, honest limits |
| Assumptions | Every driver as a plain value + Scenario switch (Base / Optimistic / Pessimistic dropdown in B2) |
| Monthly | 36 months of P&L: revenue, COGS, salaries, opex, EBITDA, depreciation, interest, PBT, tax, net profit, capex, loan balance |
| Annual_PL | 3 fiscal years by SUMIFS over Monthly (Apr–Mar) |
| Balance_Sheet | Year-end, with a live balance check (must read 0) |
| Cash_Flow | Indirect method; closing cash feeds the Balance Sheet |
| Ratios | Gross / net margin, current ratio, debt/equity, ROE, break-even monthly revenue |
| Sensitivity | 5×5 Year-3 net-profit grid: revenue-growth offset × COGS% offset; centre cell equals Annual Year-3 net profit |
| Dashboard | Scenario readout, 6 KPIs, 3 native charts (monthly revenue vs net profit, annual revenue vs net profit, cash walk) |

## Verified figures (Base scenario, LibreOffice recalc)

| | FY2027 | FY2028 | FY2029 |
|---|---|---|---|
| Revenue | ₹1,28,85,744 | ₹1,67,30,937 | ₹2,17,23,560 |
| Net profit | ₹6,71,959 | ₹16,65,589 | ₹30,37,810 |
| Net margin | 5.2% | 10.0% | 14.0% |
| Closing cash | ₹9,10,677 | ₹19,25,792 | ₹46,62,257 |
| Break-even monthly revenue | ₹8,98,610 | ₹9,30,760 | ₹9,57,168 |

Balance check is 0 in every column; cash flow ties out to the balance
sheet every year; the sensitivity centre cell equals Year-3 net profit
(₹30,37,810) exactly.

## Model rules

- Revenue: month-1 ₹9,50,000, compounding monthly (Base 2.2%,
  Optimistic 3.2%, Pessimistic 1.2%).
- COGS: 58% of revenue (Base), 54% / 62% in the other scenarios.
- Salaries ₹1,85,000/month growing 0.6%/month; rent ₹45,000,
  utilities ₹28,000, other opex ₹22,000, marketing 3% of revenue.
- Depreciation: plant ₹42,00,000 over 10 years, stepping up when capex
  lands (₹6,00,000 Oct 2026, ₹4,00,000 Apr 2027).
- Loan ₹25,00,000 at 11% p.a., ₹50,000 principal repaid monthly.
- Tax: 25% provision on positive monthly PBT.
- Working capital: receivables 20 days of month-end revenue,
  inventory 15 days of month-end COGS, payables 25 days of month-end
  COGS. Opening balance sheet balances exactly (₹60,70,000).

## How to run

```bash
python3 make_data.py   # writes data/assumptions.csv + data/ground_truth.json
python3 generate.py    # writes forecast-model.xlsx
python3 -m unittest discover -s tests -v   # 23 tests, incl. LibreOffice recalc
```

Requires Python 3 + openpyxl; the recalc test needs LibreOffice
(`soffice --headless`).

## Files

- `model.py` — independent Python model (ground truth for the tests)
- `make_data.py` — writes the assumptions CSV and ground-truth JSON
- `generate.py` — builds the workbook with openpyxl
- `tests/test_build.py` — model math, formula hygiene, LibreOffice recalc
- `data/assumptions.csv`, `data/ground_truth.json` — inputs + ground truth
- `forecast-model.xlsx` — the built workbook

## Honest limits

- Tax is a flat 25% provision on positive monthly PBT, not a full
  Indian income-tax computation (no MAT, no loss carry-forward).
- No GST, no inventory valuation choice, no debtor ageing — working
  capital is a simple days-of-sales rule.
- A real engagement would start from the client's actual trial balance
  and bank statements, not these fictional drivers.

## Reference

- Lead: https://www.freelancer.com/projects/financial-modeling/year-financial-forecast-model
- Daily routine: `maticcretic-commits/daily-automation-solutions`,
  solution for 2026-11-22.
