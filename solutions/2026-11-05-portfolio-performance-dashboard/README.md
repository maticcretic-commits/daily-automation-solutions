# PortfolioPulse — Risk-Adjusted Portfolio Performance Dashboard

Built around a real Freelancer brief: a dynamic Excel spreadsheet /
dashboard for portfolio performance analysis — risk-adjusted returns
(alpha, beta, Sharpe ratio, drawdowns, volatility) with a dashboard
the client can refresh by pasting new monthly returns.

Demo data is fictional and deterministic (seeded): a diversified
portfolio vs a market benchmark, monthly returns Jan 2025–Oct 2026
(22 months). One `.xlsx`, no macros, nothing that triggers a macro warning.

## What the workbook does

- **Guide** — how to paste your own returns, add a month, and read the methods.
- **Returns** — Month, Portfolio return, Benchmark return, Risk-free
  (monthly, 0.45%). Every other column is a live formula: excess
  returns, active return, wealth curves (growth of 1.00), and
  drawdown from the running peak.
- **Metrics** — 21 live metrics: cumulative and annualised returns,
  monthly and annualised volatility, Sharpe ratios, beta (SLOPE),
  monthly and annualised alpha, correlation, max drawdowns, best /
  worst month, win rate, tracking error, information ratio.
- **Dashboard** — 10 KPIs linked to Metrics, plus 3 native Excel
  charts: growth of 1.00 invested, monthly returns, drawdown from peak.
- **Checks** — 8 integrity checks. All read PASS.

## Verified figures (LibreOffice headless recalc vs independent Python)

- Portfolio cumulative +20.81% · benchmark +14.77%
- Portfolio annualised +10.86% · volatility (annualised) 17.42% (benchmark 15.28%)
- Sharpe (portfolio) 0.37 · beta 1.09 · alpha (annualised) +2.84%
- Portfolio max drawdown −19.12% · win rate 59.1%
- Zero formula errors in the recalculated workbook; all 8 Checks PASS.

## Methods (stated in the workbook, not hidden)

- Sharpe = mean excess return / return volatility × √12.
- Beta = SLOPE of portfolio returns on benchmark returns.
- Alpha = mean portfolio excess − beta × mean benchmark excess.
- Drawdown = wealth / running peak wealth − 1.
- This is a reporting template with fictional demo data, not investment advice.

## Files

- `generate.py` — seeded generator + workbook builder (openpyxl).
- `data/returns.csv` — the 22 monthly returns the workbook is built from.
- `tests/test_build.py` — generator run + LibreOffice recalc comparison
  against an independent Python ground truth (21 metrics, 8 checks, zero errors).
- `portfolio-performance.xlsx` — the deliverable workbook.

## Run it

```sh
python3 generate.py
python3 tests/test_build.py
```
