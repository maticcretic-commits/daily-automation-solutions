# Power BI mapping for StatementPulse

The Excel workbook is the reference implementation. The same model in
Power BI uses two fact tables (the CSVs in `data/`) and the measures below.
All measures are plain DAX over the imported tables; no calculated columns
are needed beyond the Month key already present in `monthly.csv`.

## Tables

- `Statements` — `statements.csv` (Section, LineItem, Year, Amount).
  Filter `Year` to FY2024/FY2025/FY2026 for the statement views; the
  `Opening` rows feed working-capital movement measures only.
- `Monthly` — `monthly.csv` (Month, Revenue, COGS, Opex, Receivables,
  Inventory, Payables).

## Measures (DAX)

```dax
Revenue = CALCULATE(SUM(Statements[Amount]),
    Statements[LineItem] = "Revenue")

PAT =
VAR PBT = [Revenue]
    - CALCULATE(SUM(Statements[Amount]), Statements[LineItem] = "Cost of goods sold")
    - CALCULATE(SUM(Statements[Amount]), Statements[LineItem] = "Operating expenses")
    - CALCULATE(SUM(Statements[Amount]), Statements[LineItem] = "Depreciation")
    - CALCULATE(SUM(Statements[Amount]), Statements[LineItem] = "Interest expense")
RETURN PBT - CALCULATE(SUM(Statements[Amount]), Statements[LineItem] = "Tax expense")

Net Margin = DIVIDE([PAT], [Revenue])

Current Ratio =
VAR CA = CALCULATE(SUM(Statements[Amount]),
    Statements[LineItem] IN {"Cash and bank", "Trade receivables", "Inventories", "Other current assets"})
VAR CL = CALCULATE(SUM(Statements[Amount]),
    Statements[LineItem] IN {"Trade payables", "Short-term borrowings", "Other current liabilities"})
RETURN DIVIDE(CA, CL)

Receivables Days = DIVIDE(
    CALCULATE(SUM(Statements[Amount]), Statements[LineItem] = "Trade receivables"),
    [Revenue]) * 365

Monthly Revenue = SUM(Monthly[Revenue])
Monthly Margin % = DIVIDE(SUM(Monthly[Revenue]) - SUM(Monthly[COGS]) - SUM(Monthly[Opex]),
    SUM(Monthly[Revenue]))
Working Capital (month-end) = SUM(Monthly[Receivables]) + SUM(Monthly[Inventory])
    - SUM(Monthly[Payables])
```

## Pages

1. **Executive** — KPI cards (Revenue, PAT, Net Margin, Current Ratio,
   Receivables Days, CFO/PAT) + clustered column Revenue vs PAT by Year.
2. **Ratios** — table of the 13 workbook ratios by year, with conditional
   formatting on the FY2026 FLAG statuses from the Ratios sheet benchmarks.
3. **Trend** — line chart of Monthly Revenue + Monthly Margin % by Month,
   with the working-capital area chart beneath.

Refresh: replace the two CSVs and refresh — no model changes needed, since
every figure is a measure over the raw line items.
