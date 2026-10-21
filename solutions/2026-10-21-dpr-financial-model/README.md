# DPRPulse — Bank-Ready DPR Financial Model (Excel)

A formula-driven Detailed Project Report (DPR) financial model for a
fictional turmeric-processing unit — **Sharma Agro Foods, Patna (Bihar)**.
Built as a reference implementation for bank-loan DPR work: project cost,
means of finance, loan amortisation, 7-year P&L, cash flow, DSCR, projected
balance sheet, key ratios, and a 4-scenario sensitivity analysis — every
computed cell is a live Excel formula, nothing is pasted as a value.

> **Demo case.** All names, places and numbers are fictional and
> deterministic. Change the blue input cells and the whole model
> recalculates.

## The model

| Item | Value |
|---|---|
| Project cost | Rs 240 lakh (land 30, building 55, P&M 110, contingency 10, electricals 18, preliminary 12, WC margin 5) |
| Means of finance | Promoter Rs 60 lakh (25% margin) + term loan Rs 180 lakh |
| Term loan | 10% p.a., 7 years, 1-year moratorium (interest serviced), 6 annual instalments of Rs 41.33 lakh |
| Operations | 1,200 MT/yr turmeric processing; utilisation ramps 55% → 90%; Rs 150/kg growing 3%/yr |
| Base-case results | Min DSCR (repayment years) **1.50x**, avg DSCR 2.53x, NPV @12% Rs 241 lakh, equity IRR 82.7%, payback year 2, break-even at 41.3% utilisation |

## Sensitivity (all formula-driven)

| Scenario | Min DSCR | Verdict |
|---|---|---|
| S0 Base | 1.50x | services debt comfortably |
| S1 Revenue −5% | 1.36x | still above the 1.25x bank norm |
| S2 Raw material +3pp | 0.84x | **risk** — thin contribution margin; lock input prices |
| S3 Combined stress | 0.73x | **risk** — keep a contingency buffer |

## Sheets

`Guide` → `Assumptions` (blue = only inputs) → `Project_Cost` →
`Means_of_Finance` → `Loan_Schedule` → `Revenue_Cost` → `PL` →
`CashFlow_DSCR` → `BalanceSheet` → `Ratios` → `Sensitivity` →
`Dashboard` (8 KPIs, 3 charts, risk flags) → `Checks` (10 integrity checks,
all TRUE).

## How to run

```bash
python3 generate.py        # writes data/*.csv (the inputs)
python3 build_workbook.py # writes dprpulse.xlsx (formulas only)
python3 -m pytest tests/  # LO headless recalc vs independent Python model
```

## Verification

`tests/test_build.py` recalculates the workbook with LibreOffice headless and
compares every key figure against an independent Python implementation in
`main.py`: project cost, loan amortisation (fully repaid by Y7), revenue/PAT/
cash accruals for all 7 years, DSCR series, NPV @12%, equity IRR, payback,
balance-sheet identity (assets − liabilities = 0 every year), break-even
utilisation, and all four sensitivity scenarios — plus a scan for formula
errors and a check that all 10 `Checks` read TRUE.

Key accounting identities enforced by construction:

* `net cash to equity = cash accruals − principal repaid`
  (interest already sits inside PAT, so it is not subtracted twice)
* `Δcash = cash accruals − principal` ⇒ balance sheet balances every year
