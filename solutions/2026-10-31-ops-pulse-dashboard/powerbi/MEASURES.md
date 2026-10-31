# Power BI measures — OpsPulse

Rebuild the Excel dashboard as a `.pbix` from the same three CSV feeds
(`Movements`, `Raw_Shifts`, `Orders`) plus the `Products` dimension.
Load each CSV with Power Query, promote headers, set date columns to
Date type, and relate `Movements[ProductID]` and `Orders[ProductID]` to
`Products[ProductID]`. Add a calendar table with a Monday `WeekStart`
column and relate it to each fact table's `WeekStart` date.

Standards as parameters (or a one-row `Assumptions` table):
`StandardHours = 8`, `OnTimeTargetDays = 3`.

```dax
SKUs in Range =
COUNTROWS ( Products )

SKUs Needing Reorder =
COUNTROWS (
    FILTER (
        Products,
        [Stock On Hand] <= Products[ReorderLevel]
    )
)

Stock On Hand =
SUM ( Products[OpeningStock] )
    + CALCULATE (
        SUM ( Movements[Quantity] ),
        Movements[MovementType] = "Receipt"
    )
    - CALCULATE (
        SUM ( Movements[Quantity] ),
        Movements[MovementType] = "Sale"
    )
    + CALCULATE (
        SUM ( Movements[Quantity] ),
        Movements[MovementType] = "Adjustment"
    )
```

Note: `Stock On Hand` as written is a portfolio total. For a
per-product visual, use the same expression inside a measure evaluated
in the product context (the `SUMIFS` logic in the workbook's Inventory
sheet is the row-level equivalent), or add a calculated column on
`Products` with the same Receipt/Sale/Adjustment sums filtered to that
product.

```dax
Stock Value (Rs) =
SUMX (
    Products,
    [Stock On Hand] * Products[UnitCost]
)

Orders (8 Weeks) =
COUNTROWS ( Orders )

Order Value (Rs) =
SUMX ( Orders, Orders[Quantity] * Orders[UnitPrice] )

Fulfilment Days =
AVERAGEX (
    Orders,
    DATEDIFF ( Orders[OrderDate], Orders[ShipDate], DAY )
)

On-Time Fulfilment % =
DIVIDE (
    COUNTROWS (
        FILTER (
            Orders,
            DATEDIFF ( Orders[OrderDate], Orders[ShipDate], DAY )
                <= [OnTimeTargetDays]
        )
    ),
    COUNTROWS ( Orders )
)

Overtime Hours =
SUMX (
    Raw_Shifts,
    MAX ( 0, Raw_Shifts[WorkedHours] - [StandardHours] )
)

Staff Utilisation =
DIVIDE (
    SUM ( Raw_Shifts[WorkedHours] ),
    SUM ( Raw_Shifts[ScheduledHours] )
)

Reorder Share =
DIVIDE ( [SKUs Needing Reorder], [SKUs in Range] )
```

Pages mirroring the workbook: Dashboard (KPI cards + the three charts),
Inventory (stock, value, REORDER flags), Staff (overtime by employee/team,
utilisation), Weekly (orders, value, on-time %, fulfilment days,
receipts, sales, overtime by WeekStart). Add WeekStart, Team and Category
slicers; drill-through from Weekly to the underlying Orders/Movements/
Raw_Shifts rows. Refresh = drop the new weekly CSVs into the source
folder and refresh — no hard-coded paths if the folder is a Power Query
parameter, matching the brief's acceptance criteria.
