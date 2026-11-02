# Power BI mapping for the FinancePulse dashboard

The brief lists Power BI as an accepted delivery format. The Excel
workbook is the reference implementation; this note maps each of its
KPIs to DAX measures over the same four tables (Income, Expenses,
Sales, Investments) plus an Assumptions/parameters table holding the
opening cash balance. Import the four `data/*.csv` files with Power
Query, add a Calendar table, and relate each fact table to Calendar on
its Date column.

## Measures

```dax
Total Income = SUM ( Income[Amount] )
Total Expenses = SUM ( Expenses[Amount] )
Net Cash Flow = [Total Income] - [Total Expenses]
Cash Balance =
    [Opening Cash] + CALCULATE ( [Net Cash Flow], ALL ( 'Calendar' ) )
Total Sales Revenue = SUMX ( Sales, Sales[Qty] * Sales[Unit price] )
Total Sales Cost = SUMX ( Sales, Sales[Qty] * Sales[Cost per unit] )
Gross Profit = [Total Sales Revenue] - [Total Sales Cost]
Savings Rate = DIVIDE ( [Net Cash Flow], [Total Income] )
Total Invested = SUMX ( Investments, Investments[Units] * Investments[Buy price] )
Portfolio Value = SUMX ( Investments, Investments[Units] * Investments[Current price] )
Total Gain = [Portfolio Value] - [Total Invested]
Gain % = DIVIDE ( [Total Gain], [Total Invested] )
```

`[Opening Cash]` is a single-value parameter (125,000 in the demo);
in Power BI, hold it in a one-row Parameters table or a What-if
parameter so the owner can change it without editing a measure.

## Pages (mirroring the Excel Dashboard)

1. **Overview** — KPI cards for the 12 measures above; clustered column
   chart Income vs Expenses by month; line chart Cumulative Cash
   (running total of Net Cash Flow plus Opening Cash); donut chart
   Expenses by Category.
2. **Monthly** — matrix of Income / Expenses / Net / Sales revenue /
   Sales cost / Gross profit / Cumulative cash / Savings rate by month.
3. **Investments** — table of holdings with Invested / Value / Gain /
   Gain % per holding, plus portfolio totals.

## Refresh

Replace the four CSVs (or point Power Query at the owner's folders)
and refresh: no measure changes are needed. Keep the same rule as the
Excel version — raw tables hold values only; every derived figure is a
measure, never a pasted value.
