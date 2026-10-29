# Power BI measures — RenewableMIS (DAX equivalents of the Excel KPIs)

Source tables: `Raw_Plan` (one row per site-month: PlannedMWh, PlannedRevenueLakh)
and `Raw_Actual` (ActualMWh, ActualRevenueLakh, CurtailmentMWh, OutageHours).
Relate both to `Sites` on SiteID and to a Date table on Month. Every Excel
formula in `renewable-mis.xlsx` maps to one measure below; no measure
hard-codes a site, state or month.

```dax
Planned MWh = SUM ( Raw_Plan[PlannedMWh] )

Actual MWh = SUM ( Raw_Actual[ActualMWh] )

Generation Achievement % =
DIVIDE ( [Actual MWh], [Planned MWh] )

Planned Revenue (Rs lakh) = SUM ( Raw_Plan[PlannedRevenueLakh] )

Actual Revenue (Rs lakh) = SUM ( Raw_Actual[ActualRevenueLakh] )

Revenue Achievement % =
DIVIDE ( [Actual Revenue (Rs lakh)], [Planned Revenue (Rs lakh)] )

Curtailment MWh = SUM ( Raw_Actual[CurtailmentMWh] )

Curtailment % of Planned =
DIVIDE ( [Curtailment MWh], [Planned MWh] )

Outage Hours = SUM ( Raw_Actual[OutageHours] )

Weighted Availability % =
1 - DIVIDE (
    [Outage Hours],
    DISTINCTCOUNT ( Sites[SiteID] ) * SUM ( 'Date'[DaysInMonth] ) * 24
)
```

## Pages (mirror the Excel sheets)

1. **Overview** — KPI cards for the 10 measures above; column chart Planned vs
   Actual MWh by Technology; pie of Actual MWh by Technology; line of
   Generation Achievement % by Month.
2. **Monthly** — table by Month: Planned / Actual / Variance / Achievement /
   Revenue achievement / Curtailment / Availability, with the AOP baseline
   as the comparison series.
3. **State & Technology** — matrix by State and by Technology with the same
   measures; slicers on State, Technology and Month.

## Refresh

Point Power Query at the same two raw exports the Excel version consumes
(one row per site-month). Append a new month's rows and every measure
recalculates; the AOP baseline is just the Planned table, so a revised
Annual Operating Plan is a data refresh, not a remodel. Scheduled refresh
in Power BI Service replaces the manual Excel copy/link assembly the
client does today.
