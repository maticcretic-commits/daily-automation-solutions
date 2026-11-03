# Power BI mapping — Payment Tracker

The workbook is the system of record; this file maps each workbook
figure to its Power BI equivalent so the same model can be rebuilt as a
`.pbix` without re-deriving the logic. No `.pbix` is shipped: the brief
asks for one `.xlsx` with no macros.

## Tables

- `Payments` — import the Payments sheet (rows 2..121, the filled rows;
  exclude the blank headroom rows on import, or filter `Customer` to
  non-blank). Keep `Pending`, `Status`, `Days Overdue`, `Month Key`
  as calculated columns only if the Excel formulas are not carried
  over; otherwise recompute with the measures below.
- `Customers` — one row per customer (the Customers sheet).
- `Monthly` — one row per invoice month.

## Measures (DAX)

```dax
Total Invoiced = SUM(Payments[Invoice Amount (Rs)])
Total Collected = SUM(Payments[Amount Paid (Rs)])
Total Pending = SUM(Payments[Pending (Rs)])
Collection Rate = DIVIDE([Total Collected], [Total Invoiced])
Overdue Invoices =
    CALCULATE(
        COUNTROWS(Payments),
        Payments[Pending (Rs)] > 0,
        Payments[Due Date] < [AsOfDate]
    )
Overdue Amount =
    CALCULATE(
        SUM(Payments[Pending (Rs)]),
        Payments[Pending (Rs)] > 0,
        Payments[Due Date] < [AsOfDate]
    )
Avg Days Overdue =
    AVERAGEX(
        FILTER(Payments, Payments[Days Overdue] > 0),
        Payments[Days Overdue]
    )
```

`AsOfDate` is a single-value parameter, default 2026-11-03, matching
Assumptions!B2 in the workbook. Status is a calculated column:

```dax
Status =
    IF(
        Payments[Pending (Rs)] <= 0, "Paid",
        IF(
            Payments[Due Date] < [AsOfDate],
            IF(Payments[Amount Paid (Rs)] > 0, "Partial - Overdue", "Unpaid - Overdue"),
            IF(Payments[Amount Paid (Rs)] > 0, "Partial", "Unpaid")
        )
    )
```

## Visuals

- KPI cards: Total Invoiced, Total Collected, Total Pending,
  Collection Rate, Overdue Invoices, Overdue Amount.
- Clustered column: Invoiced vs Collected by Month (Monthly table).
- Pie: Invoices by Status.
- Line: Cumulative Collected by Month.
- Table: Overdue list (Customer, Invoice No, Due Date, Pending,
  Days Overdue, Status), filtered to Pending > 0 and Due Date <
  AsOfDate, sorted by Due Date ascending, top 50.
