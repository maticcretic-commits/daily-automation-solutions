#!/usr/bin/env python3
"""Independent ground truth: recompute every workbook figure from data/*.csv.

Used by tests/test_build.py; never reads the workbook, so a formula bug in
build_workbook.py cannot hide behind its own numbers. Standards: standard
shift 8 hours, overtime = max(0, worked - 8), on-time = fulfilment in
3 days or fewer, reorder flag = stock on hand at or below the reorder level.
"""
import csv
import datetime
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
STANDARD_HOURS = 8.0
ON_TIME_DAYS = 3
WEEKS = [datetime.date(2026, 9, 7) + datetime.timedelta(weeks=i) for i in range(8)]


def read(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def compute():
    products = read("products.csv")
    movements = read("movements.csv")
    shifts = read("shifts.csv")
    orders = read("orders.csv")

    inventory = []
    for p in products:
        pid = p["ProductID"]
        opening = int(p["OpeningStock"])
        receipts = sum(int(m["Quantity"]) for m in movements
                       if m["ProductID"] == pid and m["MovementType"] == "Receipt")
        sales = sum(int(m["Quantity"]) for m in movements
                    if m["ProductID"] == pid and m["MovementType"] == "Sale")
        adj = sum(int(m["Quantity"]) for m in movements
                  if m["ProductID"] == pid and m["MovementType"] == "Adjustment")
        stock = opening + receipts - sales + adj
        inventory.append({
            "id": pid, "opening": opening, "receipts": receipts, "sales": sales,
            "adjustments": adj, "stock": stock,
            "value": stock * float(p["UnitCost"]),
            "reorder": int(p["ReorderLevel"]),
            "status": "REORDER" if stock <= int(p["ReorderLevel"]) else "OK",
        })

    staff = []
    for name in dict.fromkeys(s["Employee"] for s in shifts):
        rows = [s for s in shifts if s["Employee"] == name]
        worked = sum(float(s["WorkedHours"]) for s in rows)
        sched = sum(float(s["ScheduledHours"]) for s in rows)
        ot = sum(max(0.0, float(s["WorkedHours"]) - STANDARD_HOURS) for s in rows)
        staff.append({
            "name": name, "team": rows[0]["Team"], "shifts": len(rows),
            "scheduled": sched, "worked": worked, "overtime": ot,
            "utilisation": worked / sched if sched else 0.0,
        })

    order_rows = []
    for o in orders:
        days = (datetime.date.fromisoformat(o["ShipDate"])
                - datetime.date.fromisoformat(o["OrderDate"])).days
        order_rows.append({
            "id": o["OrderID"], "days": days,
            "value": int(o["Quantity"]) * float(o["UnitPrice"]),
            "on_time": days <= ON_TIME_DAYS, "week": o["WeekStart"],
        })

    weekly = []
    for wk in WEEKS:
        w = wk.isoformat()
        wk_orders = [o for o in order_rows if o["week"] == w]
        wk_mov = [m for m in movements if m["WeekStart"] == w]
        wk_sh = [s for s in shifts if s["WeekStart"] == w]
        weekly.append({
            "week": w, "orders": len(wk_orders),
            "value": sum(o["value"] for o in wk_orders),
            "on_time": sum(1 for o in wk_orders if o["on_time"]),
            "on_time_pct": (sum(1 for o in wk_orders if o["on_time"]) / len(wk_orders)
                            if wk_orders else 0.0),
            "avg_days": (sum(o["days"] for o in wk_orders) / len(wk_orders)
                         if wk_orders else 0.0),
            "receipts": sum(int(m["Quantity"]) for m in wk_mov
                            if m["MovementType"] == "Receipt"),
            "sales": sum(int(m["Quantity"]) for m in wk_mov
                         if m["MovementType"] == "Sale"),
            "overtime": sum(max(0.0, float(s["WorkedHours"]) - STANDARD_HOURS)
                            for s in wk_sh),
        })

    net_movement = sum(
        (-int(m["Quantity"]) if m["MovementType"] == "Sale" else int(m["Quantity"]))
        for m in movements)
    return {
        "inventory": inventory,
        "staff": staff,
        "orders": order_rows,
        "weekly": weekly,
        "count": len(products),
        "reorder_count": sum(1 for i in inventory if i["status"] == "REORDER"),
        "stock_value": sum(i["value"] for i in inventory),
        "stock_units": sum(i["stock"] for i in inventory),
        "order_count": len(order_rows),
        "order_value": sum(o["value"] for o in order_rows),
        "on_time_count": sum(1 for o in order_rows if o["on_time"]),
        "on_time_pct": (sum(1 for o in order_rows if o["on_time"]) / len(order_rows)
                        if order_rows else 0.0),
        "avg_days": sum(o["days"] for o in order_rows) / len(order_rows),
        "overtime": sum(s["overtime"] for s in staff),
        "utilisation": (sum(s["worked"] for s in staff)
                        / sum(s["scheduled"] for s in staff)),
        "net_movement": net_movement,
        "negative_stock": sum(1 for i in inventory if i["stock"] < 0),
    }


if __name__ == "__main__":
    t = compute()
    print(f"skus={t['count']} reorder={t['reorder_count']} "
          f"stock_value=Rs {t['stock_value']:,.0f} stock_units={t['stock_units']} "
          f"orders={t['order_count']} order_value=Rs {t['order_value']:,.0f} "
          f"on_time={t['on_time_pct']:.1%} avg_days={t['avg_days']:.2f} "
          f"overtime={t['overtime']:.1f}h utilisation={t['utilisation']:.1%}")
    for i in t["inventory"]:
        print(f"  {i['id']} stock={i['stock']} value=Rs {i['value']:,.0f} {i['status']}")
    for w in t["weekly"]:
        print(f"  {w['week']} orders={w['orders']} value=Rs {w['value']:,.0f} "
              f"on_time={w['on_time_pct']:.0%} avg_days={w['avg_days']:.2f} "
              f"overtime={w['overtime']:.1f}h")
