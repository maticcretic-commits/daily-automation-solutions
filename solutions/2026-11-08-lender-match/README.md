# LenderMatch — Manufacturing Acquisition Lender Comparison Pack

A formula-driven Excel workbook that screens, prices, scores and ranks
lenders for a manufacturing acquisition — the deliverable behind the
Freelancer brief **"Loan Sourcing for Manufacturing Buyout"** (project
40804560): a curated list of non-traditional, asset-based and SBA 7(a)
lenders, with terms and a comparison matrix.

## The problem

A buyer acquiring a manufacturing business needs to compare lenders that
do not quote on the same terms: SBA 7(a) loans cap at $5M with 90% LTV
but close slowly; asset-based lenders move fast but cap LTV lower;
non-traditional lenders close in days at much higher rates. A side-by-side
list of rates does not answer "which lender can actually fund *this*
deal, at what monthly cost, and which trade-offs matter most?"

## The approach

One scenario drives everything. The buyer enters the acquisition price,
equity %, target revenue, monthly EBITDA, credit score and other monthly
debt on the **Scenario** sheet; every lender is then:

1. **Screened** — eligible only if the target loan fits its max loan, its
   LTV headroom, its minimum revenue and its minimum credit score.
2. **Priced** — monthly payment via `PMT` at the lender's mid rate over
   its own tenor, plus total interest, upfront origination fees and DSCR
   (monthly EBITDA / payment + other debt).
3. **Scored 0–100** — weighted: rate 35, closing speed 20, LTV 20,
   origination fee 15, tenor 10. Each component is normalised across all
   lenders in the table, so the score always reflects the actual field.
4. **Ranked** — the Dashboard names the best pick and the cheapest
   payment; the Checks sheet proves the arithmetic (10/10 PASS).

Demo scenario (fictional): $2,400,000 acquisition, 20% buyer equity,
$1,920,000 target loan, $6,500,000 target revenue, $145,000 monthly
EBITDA, 720 credit score. All 14 lenders are fictional; replace the
**Lenders** sheet rows with real quotes before deciding.

## Verified result (LibreOffice 24.2.7.2 headless recalc vs independent Python ground truth)

- Eligible: **7 of 14** lenders
- Best pick: **BluePeak Commercial — score 75.2** (8.50% mid rate, 10-yr tenor)
- Lowest monthly payment: **$23,805.25** (BluePeak Commercial)
- Best-pick DSCR: **4.05**; total interest on the best pick: **$936,630**
- Fastest eligible close: **18 days** (Foundry Equipment Finance)
- Highest eligible LTV: **90%** (SBA 7(a) lenders)
- Checks: **10/10 PASS**, zero formula errors

## How to run

```bash
python3 generate.py            # rebuilds lendermatch.xlsx from data/lenders.csv
python3 tests/test_build.py    # rebuilds, recalcs in LibreOffice, checks vs ground truth
```

Requires Python 3 + openpyxl; the test additionally needs LibreOffice
(`soffice`) for the headless recalculation check.

## Files

- `lendermatch.xlsx` — the workbook (Guide, Lenders, Scenario, Comparison, Dashboard, Checks)
- `generate.py` — builds the workbook from `data/lenders.csv`
- `data/lenders.csv` — the 14-lender source table
- `tests/test_build.py` — rebuild + formula scan + LibreOffice recalc vs Python ground truth

No macros. Plain A1 ranges throughout (LibreOffice-safe); rates stored as
0–1 fractions; ineligible rows show blanks for payment/interest/fees/DSCR
and score 0, with sentinel helper columns (hidden) feeding the Dashboard
aggregates so no `""` ever enters arithmetic.
