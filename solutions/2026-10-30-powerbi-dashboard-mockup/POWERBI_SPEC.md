# Power BI build spec — Demo Traders sales dashboard

Status: **SPEC ONLY — not tested in Power BI.** Power BI Desktop is
Windows-only and was not available on the machine that produced this package.
No Power BI file was produced here. Every DAX measure below is a written
specification to create in Power BI Desktop; none has been run or verified
in Power BI. The HTML mockup (`index.html`) is a visual preview only.

Data: `sales.csv` (600 rows) and `targets.csv` (48 rows). See
`DATA_DICTIONARY.md`. Synthetic data only.

## 1. Tables and relationships (star schema)

Import both CSVs with Power Query (Get Data → Text/CSV). Do not merge the
files; keep two tables.

| Table | Grain | Key columns |
|---|---|---|
| `Sales` (from sales.csv) | One row per order line | `order_id`, `date`, `region`, `category`, `product`, `salesperson`, `qty`, `unit_price`, `discount_pct`, `sales_amount` |
| `Targets` (from targets.csv) | One row per category × month | `month` (`YYYY-MM` text), `category`, `target_amount` |

Relationships to create in Model view:

1. `Sales[category]` → `Targets[category]` is NOT a valid single relationship
   on its own (many-to-many on category alone). Instead, add a calculated
   month column to `Sales` first (see step 2), then relate:
   - `Sales[category]` + `Sales[month]` → `Targets[category]` + `Targets[month]`
     as a composite relationship, OR (simpler, recommended for a first
     build) create a small `Category` dimension table (distinct category
     values) and a `Calendar` table (distinct months), and relate:
     - `Category[category]` 1-to-many `Sales[category]`
     - `Category[category]` 1-to-many `Targets[category]`
     - `Calendar[month]` 1-to-many `Sales[month]`
     - `Calendar[month]` 1-to-many `Targets[month]`
2. In Power Query, add a `month` column to `Sales`:
   `month = Date.ToText([date], "yyyy-MM")` (Power Query / M — spec only,
   untested here; the CSVs already carry the `date`, so this is a
   transformation step in the build, not a claim that Power Query was used
   in this package).

No SQL was used in this package, and no SQL proficiency is claimed.
The import above is a plain CSV import.

## 2. DAX measures to create (SPEC — untested in Power BI)

Create these as measures on the `Sales` / `Targets` tables in Power BI
Desktop. Definitions are written out here so the build is reproducible;
they have NOT been executed or validated in Power BI.

```dax
Total Sales = SUM(Sales[sales_amount])
```

```dax
Total Qty = SUM(Sales[qty])
```

```dax
Avg Discount % = AVERAGE(Sales[discount_pct]) * 100
```

```dax
Total Target = SUM(Targets[target_amount])
```

```dax
Sales vs Target % =
DIVIDE ( [Total Sales], [Total Target] ) * 100
```

```dax
Top Category Sales =
MAXX (
    VALUES ( Sales[category] ),
    CALCULATE ( [Total Sales] )
)
```

`Top Category` (the name, for a card title) — SPEC:

```dax
Top Category =
MAXX (
    TOPN ( 1, VALUES ( Sales[category] ), [Total Sales] ),
    Sales[category]
)
```

Expected values on this synthetic data (computed independently in Python,
see CHECKS.txt — use them to sanity-check the build, not as proof the
measures work): Total Sales Rs 7,189,354.50; Total Qty 4,030;
Avg Discount % 5.73%; Sales vs Target % 124.84%; Top Category Furniture.

## 3. Report pages and visuals

### Page 1 — Sales overview

- KPI cards: Total Sales, Total Qty, Avg Discount %, Sales vs Target %.
- Clustered column chart: Total Sales by month (x-axis: month; values:
  Total Sales). Sort months chronologically.
- Card: Top Category + Top Category Sales.
- Slicers: Region, Category, Salesperson.

### Page 2 — Category / region breakdown

- Bar chart: Total Sales by Category.
- Bar chart: Total Sales by Region.
- Table: Salesperson, Total Sales, Total Qty (sorted by Total Sales
  descending).
- Slicers: Region, Category, Month.

### Page 3 — Target tracking

- Clustered column chart: Total Sales vs Total Target by month.
- Table: Category, Total Sales, Total Target, Sales vs Target %.
- KPI card: overall Sales vs Target %.
- Slicers: Category, Month.

Formatting guidance: currency in Rs with 2 decimals on cards; % with
1–2 decimals; chronological month order (use the `Calendar` table's
month sort, not alphabetical).

## 4. Step-by-step build instructions (Power BI Desktop)

1. Open Power BI Desktop (Windows). Get Data → Text/CSV → import
   `sales.csv`, then `targets.csv`.
2. In Power Query, add the `month` column to `Sales` as in §1, step 2.
   Close & Apply.
3. In Model view, create the `Category` and `Calendar` dimension tables
   (or the composite relationship) as in §1.
4. Create the six measures in §2 exactly as written. Do not change the
   definitions without re-checking the expected values in CHECKS.txt.
5. Build the three report pages in §3, adding the slicers listed per page.
6. Check the page-1 KPI cards against CHECKS.txt. If a card disagrees,
   the model relationships are the first place to look.
7. Save your work in Power BI Desktop. (The saved file is the Power BI
   file format produced by that application; no such file is included
   in this package.)

## 5. Limitations

- This spec was written without access to Power BI Desktop; the DAX
  measures are untested and may need small corrections (e.g. filter
  context or relationship behaviour) when built for real.
- SQL and Power Query were not used to produce this package and are not
  claimed as skills evidenced here. The Power Query step in §1 is part
  of the build instructions only.
- All data is synthetic. Nothing here represents a real client, real
  sales, or a delivered dashboard.
