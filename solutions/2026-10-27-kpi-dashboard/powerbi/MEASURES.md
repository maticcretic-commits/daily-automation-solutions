# Power BI mapping for KPI Pulse (from the Excel demo)

The Excel workbook is the proof of the model: every figure below already
reconciles in Excel. In Power BI, load `data/raw_transactions.csv` and
`data/monthly_targets.csv`, relate Targets[Month] → Raw[Month] (many-to-one
on the Month text column — never on dates, which is where refresh breaks).

## DAX measures (paste into the model, one per line)

```dax
Revenue = SUM ( Raw[Revenue] )
Cost = SUM ( Raw[Cost] )
Orders = SUM ( Raw[Orders] )
Gross Profit = [Revenue] - [Cost]
Gross Margin % = DIVIDE ( [Gross Profit], [Revenue] )
Avg Order Value = DIVIDE ( [Revenue], [Orders] )
On-Time % = DIVIDE ( SUM ( Raw[OnTimeOrders] ), [Orders] )
CSAT = DIVIDE ( SUM ( Raw[CSATSum] ), [Orders] )
Revenue Target = SUM ( Targets[RevenueTarget] )
Target Attainment = DIVIDE ( [Revenue], [Revenue Target] )
Variance % = DIVIDE ( [Revenue] - [Revenue Target], [Revenue Target] )
```

## Pages

1. **KPI cards** — Revenue, Gross Margin %, On-Time %, CSAT, Target
   Attainment (same six cards as the Excel Dashboard sheet).
2. **Revenue vs target** — clustered column, Month on the axis, Revenue
   and Revenue Target as values (Excel chart 1).
3. **Margin trend** — line chart of Gross Margin % by Month (Excel chart 2).
4. **Channel mix** — donut of Revenue by Channel (Excel chart 3).

## Refresh

Append new daily rows to Raw with the Month column filled in; scheduled
refresh then picks them up with no model change. The Checks sheet logic
becomes a Power Query row-count assertion: `COUNTROWS(Raw)` must equal
the sum of monthly order rows after each refresh.
