# AcquisitionModel — Financial Model for a Business Acquisition

A working demo of the Freelancer brief **"Comprehensive Financial Model
for Business Acquisition"** (posted ~7h before the 2026-11-14 scout, 19
proposals, avg bid $366 USD, Cupertino US): an Excel financial model to
assess buying / expanding a business, with cash-flow projections,
scenario analysis, defined assumptions and a user-friendly interface.

## The fictional deal

**Sharma Precision Components Pvt. Ltd., Patna** — a precision
auto-parts manufacturer. Asking price **$2,400,000**. All figures are
invented for the demo; no real business data is used.

## What the workbook does

- **Historical** — 2023–2025 actuals (revenue $3.1M → $3.6M → $4.2M)
  with every derived line a live formula. 2025 is the base year.
- **Projections** — 2026–2030 P&L and free cash flow. Pick **Base /
  Bull / Bear** in `Projections!B3` and revenue growth, gross margin
  and opex growth all switch; every year, the DCF and the dashboard
  follow. Free cash flow = NOPAT + depreciation − capex − change in
  working capital (12% of revenue).
- **Valuation** — discounted cash flow at WACC 11%: present value of
  2026–2030 free cash flow + present value of terminal value (Gordon
  growth 2.5%) = enterprise value; minus net debt $650,000 = equity
  value; divided by 100,000 shares = implied price/share, compared
  with the $24.00/share asking price. Plus a **sensitivity grid**:
  enterprise value by WACC (9–13%) × terminal growth (1.5–3.5%).
- **Dashboard** — 13 live KPIs (scenario, asking price, enterprise and
  equity value, implied price/share, premium/(discount), verdict, 2030
  revenue/EBITDA/margin, average and cumulative FCF) and 3 native
  charts: revenue vs EBITDA, free cash flow, 2030 revenue by scenario.
- **Checks** — 10 integrity checks, all must read PASS.

## The honest outcome

On the demo numbers the base case comes out at roughly **$1.63M
enterprise value / $0.98M equity value — about $9.83/share against a
$24.00 asking price — so the verdict reads PASS**. The model is the
deliverable, not a flattering answer: change the assumptions (growth,
margin, WACC, asking price) and the verdict follows.

## Files

| File | What it is |
| --- | --- |
| `generate.py` | Builds `acquisition-model.xlsx` from `data/*.csv` |
| `make_data.py` | Writes the deterministic input CSVs (seed 20261114) |
| `data/assumptions.csv` | Deal + valuation inputs (asking price, net debt, shares, tax, WACC, terminal growth, depreciation, capex, working capital %, interest) |
| `data/historical.csv` | 2023–2025 actuals |
| `data/scenarios.csv` | Base / Bull / Bear drivers |
| `acquisition-model.xlsx` | The built workbook |
| `tests/test_build.py` | Tests (see below) |

## How to run

```bash
python3 make_data.py
python3 generate.py
python3 tests/test_build.py
```

Requires Python 3 + openpyxl. The test re-implements the whole model
independently in Python and compares it against a LibreOffice headless
recalculation of the workbook (revenue, EBITDA and free cash flow per
year; PV of FCF; terminal value; enterprise and equity value; implied
price/share; premium vs asking; sensitivity-grid base cell; all 10
Checks PASS; no formula errors anywhere).

## How to refresh

1. Edit `data/assumptions.csv`, `data/historical.csv` or
   `data/scenarios.csv`.
2. Re-run `python3 generate.py` (and `tests/test_build.py` to verify).
3. Every projected figure is a live formula, so the model rebuilds
   from the new inputs. In production the client supplies the real
   historicals and the real asking price.
