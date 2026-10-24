# Power BI measures — RefreshForge

Point Power BI Desktop at the `history/*.csv` files (Get Data → Text/CSV)
plus `refresh_log.json`. One fact table per source; `dim_sku` comes from
`data/skus.csv`; a `dim_date` table covers 2026-10-01 .. 2026-12-31.

Relationships (single direction, many-to-one):
- sales[sku] → dim_sku[sku]; production[sku] → dim_sku[sku]
- inventory[sku] → dim_sku[sku]; issues[date] → dim_date[date]
- sales[date], production[date], inventory[date] → dim_date[date]

Scheduled refresh in Power BI Service replaces `refresh.py`'s local run:
the same validate → dedupe → append logic becomes the Power Query steps,
and the Refresh_Log sheet is the audit the Power Automate flow writes.

```dax
Sales Revenue      = SUMX ( sales, sales[qty] * sales[unit_price] )
MTD Sales Revenue  = TOTALMTD ( [Sales Revenue], dim_date[date] )

MTD Production Qty = TOTALMTD ( SUM ( production[qty_produced] ), dim_date[date] )

Average Yield      =
DIVIDE (
    SUM ( production[qty_produced] ),
    SUM ( production[qty_produced] ) + SUM ( production[qty_rejected] )
)

Open Issues        = COUNTROWS ( FILTER ( issues, issues[status] = "Open" ) )

Inventory On Hand  =
SUM ( dim_sku[opening_qty] )
    + SUMX ( FILTER ( inventory, inventory[movement] = "IN" ), inventory[qty] )
    - SUMX ( FILTER ( inventory, inventory[movement] = "OUT" ), inventory[qty] )

Refresh Health     = MAX ( refresh_log[timestamp] )   // card: last run
Rows Last Refresh  =
CALCULATE ( SUM ( refresh_log[rows_new] ),
    FILTER ( refresh_log, refresh_log[run_id] = MAX ( refresh_log[run_id] ) ) )
```

Report pages (mirror the Excel Dashboard):
1. Overview — KPI cards (the six measures above) + daily line chart
   (dim_date[date] × [Sales Revenue] + [MTD Production Qty]).
2. Quality — issues-by-category donut + severity slicer + open-issues table.
3. Production — qty by line bar chart + yield trend line.
4. Refresh audit — refresh_log table (run, source, seen/new/dup/quarantined,
   status) + Refresh Health card.
