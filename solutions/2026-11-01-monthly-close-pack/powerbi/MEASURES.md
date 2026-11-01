# Power BI mapping for the Monthly Close Pack

The same monthly reporting pack as a `.pbix`. Source tables are the
two CSVs in `data/` (imported via Power Query, no transformations
beyond typing dates and adding a calendar table).

## Model (star schema)

- `Transactions` (fact): Date, Description, Category, Type, Amount.
- `Budget` (fact): Month, Category, Type, Budget.
- `Calendar` (dimension): one row per date, Jan 2026 – Oct 2026, with
  YearMonth, MonthName, Quarter. Both facts relate to `Calendar` on
  date/month.
- `Categories` (dimension, optional): Category → Type, for slicers.

## Measures (DAX)

```dax
Revenue = CALCULATE(SUM(Transactions[Amount]), Transactions[Type] = "Revenue")
Direct Costs = CALCULATE(SUM(Transactions[Amount]), Transactions[Type] = "Direct")
Operating Expenses = CALCULATE(SUM(Transactions[Amount]), Transactions[Type] = "Operating")
Gross Profit = [Revenue] - [Direct Costs]
Net Income = [Gross Profit] - [Operating Expenses] - 5000 * DISTINCTCOUNT(Calendar[YearMonth])
Net Margin % = DIVIDE([Net Income], [Revenue])
Budget Revenue = CALCULATE(SUM(Budget[Budget]), Budget[Type] = "Revenue")
Budget Expenses = CALCULATE(SUM(Budget[Budget]), Budget[Type] IN {"Direct", "Operating"})
Revenue vs Budget % = DIVIDE([Revenue], [Budget Revenue])
Expense Variance % = DIVIDE(([Direct Costs] + [Operating Expenses]) - [Budget Expenses], [Budget Expenses])
Cumulative Net Income = CALCULATE([Net Income], FILTER(ALL(Calendar), Calendar[Date] <= MAX(Calendar[Date])))
```

Balance-sheet and cash-flow pages in Power BI need the same
Assumptions parameters (opening balances, 20%/10% shares, depreciation,
purchases, drawings) as What-if parameters, with measures mirroring
the Excel walk: closing cash = opening cash + collections − payments −
purchases − drawings; the balance check measure must read 0 and is
shown on the page as a guard visual.

## Pages

1. **Dashboard** — KPI cards (the 8 Excel KPIs), revenue-vs-budget
   line chart, cumulative net income column chart, expense-mix donut.
2. **P&L** — matrix by month × P&L line, with the measures above.
3. **Budget vs Actual** — table by category with variance % and
   conditional formatting (red over 5% for expenses).
4. **Balance Sheet & Cash Flow** — the articulated walk, with the
   balance-check guard visual.
5. **Data** — the two raw tables, for audit.

Refresh: drop a new month's CSV rows into the source folder and
refresh; the calendar table must extend to cover the new month.
