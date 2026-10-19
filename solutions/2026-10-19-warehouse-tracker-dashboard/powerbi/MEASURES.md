# WarehousePulse — DAX measures

Import the four CSVs, create the relationships
(`fact_inventory_daily[sku]` → `dim_sku[sku]`,
`fact_inventory_daily[bin]` → `dim_bin[bin]`,
`fact_inventory_daily[date]` → `dim_date[date]` — all many-to-one,
single direction), then add these measures.

```dax
Closing Qty =
SUM ( fact_inventory_daily[closing_qty] )

Stock Value =
SUMX (
    fact_inventory_daily,
    fact_inventory_daily[closing_qty] * RELATED ( dim_sku[wac] )
)

Receipts Value =
SUM ( fact_inventory_daily[receipts_value] )

Issues Qty =
SUM ( fact_inventory_daily[issues_qty] )

Days of Cover =
DIVIDE ( [Closing Qty], DIVIDE ( SUM ( fact_inventory_daily[issues_qty] ), 18 ) )

Reorder SKUs =
COUNTROWS (
    FILTER (
        VALUES ( dim_sku[sku] ),
        [Closing Qty] <= MAX ( dim_sku[reorder_level] )
            && [Closing Qty] > 0
    )
)

Stockout SKUs =
COUNTROWS (
    FILTER ( VALUES ( dim_sku[sku] ), [Closing Qty] = 0 )
)

Overstock Value =
SUMX (
    FILTER (
        fact_inventory_daily,
        DIVIDE (
            fact_inventory_daily[closing_qty],
            DIVIDE ( fact_inventory_daily[issues_qty], 18 )
        ) > 60
    ),
    fact_inventory_daily[closing_qty] * RELATED ( dim_sku[wac] )
)

Bin Utilisation % =
DIVIDE ( SUM ( fact_inventory_daily[closing_qty] ), SUM ( dim_bin[capacity] ) )
```

Suggested page: KPI cards (Closing Qty, Stock Value, Reorder SKUs,
Stockout SKUs), bar visual of Stock Value by `dim_sku[category]`,
line visual of Receipts/Issues Qty by `dim_date[date]`, donut of SKUs by
status (add a calculated column on `dim_sku` mirroring the workbook's
OK / REORDER / STOCKOUT / OVERSTOCK rule), and a drill-through page
filtered by `dim_sku[category]` showing the top SKUs by Stock Value.
