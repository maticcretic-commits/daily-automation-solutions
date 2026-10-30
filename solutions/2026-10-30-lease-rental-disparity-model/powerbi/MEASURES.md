# Power BI measures — LeaseDisparity

DAX measures for rebuilding the LeaseDisparity dashboard as a `.pbix`
from the same `leases.csv` (import it as a table named `Leases`, with a
single-value parameter or disconnected table holding the annual yield).
These are documentation, not a verified deliverable: the verified
deliverable is the formula-driven Excel workbook in this folder, whose
figures were checked against an independent Python ground truth and a
LibreOffice headless recalculation.

```dax
Annual Yield = 0.085
Monthly Yield = [Annual Yield] / 12

Monthly Disparity =
    SUMX ( Leases, Leases[MarketRentMonthly] - Leases[ContractRentMonthly] )

PV of Disparity =
    SUMX (
        Leases,
        ( Leases[MarketRentMonthly] - Leases[ContractRentMonthly] )
            * ( 1 - POWER ( 1 + [Monthly Yield], -Leases[RemainingMonths] ) )
            / [Monthly Yield]
    )

Undiscounted Disparity =
    SUMX (
        Leases,
        ( Leases[MarketRentMonthly] - Leases[ContractRentMonthly] )
            * Leases[RemainingMonths]
    )

Under-Rented Leases =
    COUNTROWS (
        FILTER (
            Leases,
            Leases[MarketRentMonthly] - Leases[ContractRentMonthly] > 25
        )
    )

Over-Rented Leases =
    COUNTROWS (
        FILTER (
            Leases,
            Leases[MarketRentMonthly] - Leases[ContractRentMonthly] < -25
        )
    )
```

Visuals: clustered bar of contracted vs market rent by lease, column
chart of PV of disparity by lease, and a line of cumulative discounted
disparity over the 60-month run-off (build the month table in Power Query
with the same SUMPRODUCT-style expiry logic as the Excel Schedule sheet).
