# DAX measures for the SalesOpsPulse Power BI model

Star schema (tables in `sales_ops.db`, same names as the SQL views):
fact_orders, fact_inventory_daily, fact_receivables, dim_product,
dim_customer, dim_date. Relationships: fact_orders[sku] → dim_product[sku],
fact_orders[customer] → dim_customer[customer], fact_orders[date] →
dim_date[date]; same keys for the other facts. One date table, one
as-of: 2026-11-10.

## Page 1 — Sales overview

```dax
Total Revenue =
SUMX ( fact_orders,
    fact_orders[qty] * fact_orders[unit_price]
        * ( 1 - fact_orders[discount_pct] / 100 ) )

Total Margin =
SUMX ( fact_orders,
    fact_orders[qty]
        * ( fact_orders[unit_price] * ( 1 - fact_orders[discount_pct] / 100 )
            - RELATED ( dim_product[cost_price] ) ) )

Margin % = DIVIDE ( [Total Margin], [Total Revenue] )

YTD Revenue =
TOTALYTD ( [Total Revenue], dim_date[date] )

November Pacing =
VAR NovRev =
    CALCULATE ( [Total Revenue],
        dim_date[month] = "Nov 2026" )
VAR NovTarget = 400000
VAR DaysElapsed = 10
VAR DaysInMonth = 30
RETURN
    DIVIDE ( NovRev, NovTarget ) * DaysInMonth / DaysElapsed

Revenue by Channel =
CALCULATE ( [Total Revenue],
    ALLEXCEPT ( fact_orders, dim_customer[channel] ) )
```

## Page 2 — Inventory / reorder

```dax
Closing Stock =
SUM ( dim_product[opening_stock] )
    + SUM ( fact_receipts[qty] )
    - SUM ( fact_orders[qty] )

Stock Value =
SUMX ( dim_product,
    [Closing Stock] * dim_product[cost_price] )

Stock Status =
VAR Closing = [Closing Stock]
VAR Reorder = SELECTEDVALUE ( dim_product[reorder_level] )
RETURN
    IF ( Closing <= 0, "OUT",
        IF ( Closing <= Reorder, "LOW", "OK" ) )

Reorder Alerts =
COUNTROWS (
    FILTER ( dim_product, [Stock Status] IN { "LOW", "OUT" } ) )

Days of Cover =
VAR DailySales =
    DIVIDE ( SUM ( fact_orders[qty] ), 224 )
RETURN
    DIVIDE ( [Closing Stock], DailySales )
```

## Page 3 — Delivery routes

```dax
On-Time Delivery % =
DIVIDE (
    COUNTROWS (
        FILTER ( fact_deliveries,
            fact_deliveries[actual_date] <= fact_deliveries[planned_date] ) ),
    COUNTROWS ( fact_deliveries ) )

Avg Delay (days) =
AVERAGEX ( fact_deliveries,
    DATEDIFF ( fact_deliveries[planned_date],
        fact_deliveries[actual_date], DAY ) )
```

## Page 4 — Calling plan

```dax
Call Conversion % =
DIVIDE (
    COUNTROWS (
        FILTER ( fact_calls, fact_calls[outcome] = "Order placed" ) ),
    COUNTROWS ( fact_calls ) )

Call Order Value =
SUMX (
    FILTER ( fact_calls, fact_calls[outcome] = "Order placed" ),
    fact_calls[order_value] )
```

## Page 5 — Receivables & ageing

```dax
Outstanding =
SUMX ( fact_receivables,
    fact_receivables[amount] - fact_receivables[paid_amount] )

DPD =
VAR Due = SELECTEDVALUE ( fact_receivables[due_date] )
RETURN
    IF ( [Outstanding] = 0, 0,
        MAX ( 0, DATEDIFF ( Due, DATE ( 2026, 11, 10 ), DAY ) ) )

Ageing Bucket =
VAR Dpd = [DPD]
RETURN
    IF ( [Outstanding] = 0, "Paid",
        IF ( Dpd = 0, "Current",
            IF ( Dpd <= 30, "0-30",
                IF ( Dpd <= 60, "31-60",
                    IF ( Dpd <= 90, "61-90", "90+" ) ) ) ) )

Overdue 90+ =
CALCULATE ( [Outstanding], [Ageing Bucket] = "90+" )
```

## Publish + refresh (Power BI Service)

1. Import `sales_ops.db` via Power Query (ODBC/SQLite) or export the
   views to CSV and load them as the fact/dimension tables.
2. Set the scheduled refresh to daily; the as-of date is 2026-11-10 —
   move it to the refresh date and the DPD / pacing measures follow.
3. Alert rule: email when `Reorder Alerts > 0` or `Overdue 90+`
   rises week on week.
