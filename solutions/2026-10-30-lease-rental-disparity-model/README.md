# LeaseDisparity — Contracted vs Market Rent: Disparity Valuation

A working demo built around a real freelance brief: an Excel workbook that
compares contracted rent with current market rent for up to 12 leases and
values the disparity over each remaining term at a single yield rate.

## The problem

A landlord (or buyer doing due diligence) holds leases signed at different
times at different rents. Some are under-rented — the market has moved above
the contract, so income should step up at renewal. Some are over-rented —
the tenant is paying above market, so renewal carries re-letting risk. The
question is not just "how much per month", but "what is that gap worth over
the time left on each lease, in today's money, at one consistent yield?"

## The approach

- **One yield, applied identically.** Annual yield 8.5% (Assumptions sheet,
  change it and the whole valuation reprices). Monthly rate = 8.5% / 12.
- **Monthly annuity in arrears.** For each lease:
  `PV = monthly disparity × (1 − (1+r)^−n) / r`, where the monthly
  disparity is market rent minus contracted rent and `n` is the remaining
  term in months.
- **Every figure is a live formula.** No hard-coded totals anywhere: the
  Leases sheet derives disparity, annual disparity, undiscounted disparity,
  PV and status from the seven value columns; the Dashboard, the 60-month
  run-off Schedule and the Checks sheet are all formulas over those cells.
- **Status with a tolerance.** Within $25/month of market = "At market";
  above = "Under-rented"; below = "Over-rented".
- **Run-off schedule.** Month by month for 60 months: the portfolio
  disparity still running (leases drop out as their terms expire),
  discounted back at the same yield, cumulating to the same PV total —
  proved by the Checks sheet.

## Verified figures (fictional demo portfolio, 12 Los Angeles-area leases)

- Contracted rent: **$76,100/month**; market rent: **$82,200/month**
- Monthly disparity: **$6,100/month** ($73,200/year)
- Undiscounted disparity over remaining terms: **$299,100**
- Present value of the disparity at 8.5%: **$251,193.47**
- Mix: 8 under-rented, 3 over-rented, 1 at market; average remaining
  term 31.2 months. The 60-month schedule's cumulative discounted total
  equals the lease-by-lease PV total to the cent in the ground-truth model.

Verification: `tests/test_build.py` rebuilds the workbook from
`data/leases.csv`, recalculates it headlessly in LibreOffice, and compares
every checked figure against an independent Python recomputation
(`main.py`, which never reads the workbook). All checks passed, including
all 9 Checks-sheet rows reading PASS and zero formula errors.

## How to run it

```bash
python3 generate.py          # (re)create data/leases.csv — deterministic demo data
python3 build_workbook.py    # build lease-disparity.xlsx
python3 tests/test_build.py  # rebuild + LibreOffice recalc + ground-truth comparison
```

Requires Python 3.10+ and `openpyxl`; the test additionally needs
LibreOffice (`soffice`) for the headless recalculation.

To use it with real leases: replace the values in `Leases` columns A–G
with your own leases (up to 12 rows; extend the formula ranges if you add
rows) and set your yield on the `Assumptions` sheet. The Power BI DAX
measures in `powerbi/MEASURES.md` rebuild the same dashboard as a `.pbix`.

## Files

- `lease-disparity.xlsx` — the workbook (Guide, Dashboard, Leases,
  Schedule, Assumptions, Checks; 3 charts on the Dashboard)
- `generate.py` — deterministic synthetic lease data (fictional; no real
  client data)
- `build_workbook.py` — the workbook builder (all formulas, A1 ranges only)
- `main.py` — independent ground-truth recomputation
- `data/leases.csv` — the 12-lease demo dataset
- `tests/test_build.py` — build + LibreOffice recalc + ground-truth tests
- `powerbi/MEASURES.md` — DAX measures for a Power BI version

## Reference

This is a demonstration implementation built from a public freelance
brief, not client work. Figures above are the demo portfolio's, computed
from the fictional data in this folder.
