#!/usr/bin/env python3
"""WarehousePulse — export the Power BI star schema.

Reads data/*.csv and writes powerbi/:
  dim_sku.csv, dim_bin.csv, dim_date.csv, fact_inventory_daily.csv
plus powerbi/MEASURES.md with the DAX measures.
Grain of the fact: one row per SKU per day (2026-10-01 .. 2026-10-18).
"""
import csv
import os
from datetime import date, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
PBI = os.path.join(BASE, "powerbi")
os.makedirs(PBI, exist_ok=True)

PERIOD_START = date(2026, 10, 1)
PERIOD_DAYS = 18
DIM_END = date(2026, 12, 31)


def read(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def write(name, header, rows):
    with open(os.path.join(PBI, name), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def main():
    skus = read("skus.csv")
    opening = {r["sku"]: int(r["opening_qty"]) for r in read("opening.csv")}
    receipts = read("receipts.csv")
    issues = read("issues.csv")
    adjustments = read("adjustments.csv")
    bins = read("bins.csv")

    # WAC per SKU (independent ground truth, mirrors the workbook formula)
    wac = {}
    for s in skus:
        sku = s["sku"]
        oq = opening[sku]
        rq = sum(int(r["qty"]) for r in receipts if r["sku"] == sku)
        rv = sum(int(r["qty"]) * float(r["unit_cost"])
                 for r in receipts if r["sku"] == sku)
        std = float(s["std_cost"])
        wac[sku] = std if (oq + rq) == 0 else (oq * std + rv) / (oq + rq)

    write("dim_sku.csv",
          ["sku", "name", "category", "uom", "reorder_level", "std_cost",
           "wac", "bin", "opening_qty"],
          [[s["sku"], s["name"], s["category"], s["uom"], s["reorder_level"],
            s["std_cost"], f"{wac[s['sku']]:.4f}", s["bin"],
            opening[s["sku"]]] for s in skus])
    write("dim_bin.csv", ["bin", "capacity"],
          [[b["bin"], b["capacity"]] for b in bins])

    d = PERIOD_START
    date_rows = []
    while d <= DIM_END:
        date_rows.append([d.isoformat(), d.day, d.month,
                          d.strftime("%b"), d.strftime("%a"),
                          "Y" if d.weekday() >= 5 else "N"])
        d += timedelta(days=1)
    write("dim_date.csv",
          ["date", "day", "month", "month_name", "weekday", "is_weekend"],
          date_rows)

    # daily movements per SKU
    by_day = {}
    for r in receipts:
        by_day.setdefault((r["date"], r["sku"]), [0, 0.0, 0, 0])
        by_day[(r["date"], r["sku"])][0] += int(r["qty"])
        by_day[(r["date"], r["sku"])][1] += int(r["qty"]) * float(r["unit_cost"])
    for r in issues:
        by_day.setdefault((r["date"], r["sku"]), [0, 0.0, 0, 0])
        by_day[(r["date"], r["sku"])][2] += int(r["qty"])
    for r in adjustments:
        by_day.setdefault((r["date"], r["sku"]), [0, 0.0, 0, 0])
        by_day[(r["date"], r["sku"])][3] += int(r["qty"])

    fact, running = [], {s["sku"]: opening[s["sku"]] for s in skus}
    for i in range(PERIOD_DAYS):
        day = (PERIOD_START + timedelta(days=i)).isoformat()
        for s in skus:
            sku = s["sku"]
            rq, rv, iq, aq = by_day.get((day, sku), [0, 0.0, 0, 0])
            running[sku] += rq - iq + aq
            fact.append([day, sku, s["bin"], opening[sku], rq,
                         f"{rv:.2f}", iq, aq, running[sku]])
    write("fact_inventory_daily.csv",
          ["date", "sku", "bin", "opening_qty", "receipts_qty",
           "receipts_value", "issues_qty", "adj_qty", "closing_qty"], fact)

    with open(os.path.join(PBI, "MEASURES.md"), "w") as f:
        f.write(MEASURES)
    print(f"wrote powerbi/: {len(fact)} fact rows, "
          f"{len(date_rows)} date rows")


MEASURES = """# WarehousePulse — DAX measures

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
"""

if __name__ == "__main__":
    main()
