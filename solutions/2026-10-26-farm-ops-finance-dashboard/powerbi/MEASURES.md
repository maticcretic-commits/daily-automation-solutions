# Power BI measures for FarmPulse

Import `Raw_Income` and `Raw_Expense` as two fact tables. No relationships
are needed between them (they share conformed dimensions: Quarter,
Enterprise, Paddock/Herd, Cost Centre — use them as slicers, not joins).

```dax
Total Income  = SUM ( Raw_Income[Amount] )
Total Expenses = SUM ( Raw_Expense[Amount] )
Net Profit    = [Total Income] - [Total Expenses]
Margin %      = DIVIDE ( [Net Profit], [Total Income] )
Expense Ratio = DIVIDE ( [Total Expenses], [Total Income] )
```

Visuals that reproduce the Excel workbook:

- **Profit by enterprise** — bar chart, axis = Enterprise, value =
  Net Profit computed per enterprise (income minus expenses for that
  enterprise; in DAX, use two measures filtered by enterprise or a
  calculated table — do not sum the two fact tables directly).
- **Income vs expense by quarter** — clustered column, axis = Quarter,
  values = Total Income and Total Expenses.
- **Expenses by category** — pie/donut, legend = Category, value =
  Total Expenses.
- **Drill-through** — from any summary visual to a transaction table
  visual on the Raw tables, filtered by the same Enterprise / Quarter /
  Paddock-Herd / Cost Centre slicers.
- **Traffic lights** — conditional formatting on Margin %: green > 20%,
  red < 0%.

Refresh: replace the two fact tables with the next quarter's exports;
all measures and visuals update unchanged.
