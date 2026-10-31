#!/usr/bin/env python3
"""Generate deterministic synthetic operations data for the OpsPulse demo.

Fictional company: Mehta Trading Co., Mumbai — a wholesaler whose brief
(Power BI Operations Dashboard Development, Freelancer ID 40794911) asks
for a refreshable dashboard from weekly CSV exports covering product
movements, shift schedules and order fulfilment. All names, quantities
and prices are invented for the demo; no real client data. Seed 20261031.
Period: 8 weeks, Mon 2026-09-07 to Sat 2026-10-31.
"""
import csv
import datetime
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
RNG = random.Random(20261031)

START = datetime.date(2026, 9, 7)
END = datetime.date(2026, 10, 31)
ORDER_LAST = datetime.date(2026, 10, 27)

# ProductID, Product, Category, Unit, OpeningStock, ReorderLevel, UnitCost
PRODUCTS = [
    ("P01", "Basmati Rice 5kg", "FMCG", "bag", 180, 50, 420),
    ("P02", "Mustard Oil 1L", "FMCG", "bottle", 150, 45, 165),
    ("P03", "Wheat Flour 10kg", "FMCG", "bag", 200, 55, 380),
    ("P04", "Sugar 5kg", "FMCG", "bag", 160, 40, 230),
    ("P05", "Tea Powder 500g", "FMCG", "pack", 120, 35, 310),
    ("P06", "Toor Dal 2kg", "FMCG", "pack", 140, 40, 285),
    ("P07", "LED Bulb 9W", "Electronics", "unit", 220, 60, 95),
    ("P08", "Extension Board 5m", "Electronics", "unit", 90, 30, 340),
    ("P09", "USB Cable 1m", "Electronics", "unit", 260, 70, 85),
    ("P10", "Table Fan 400mm", "Electronics", "unit", 70, 25, 1450),
    ("P11", "Mixer Grinder 750W", "Electronics", "unit", 55, 20, 2850),
    ("P12", "Iron Box Steam", "Electronics", "unit", 65, 25, 1180),
    ("P13", "Cotton Shirt M", "Apparel", "piece", 110, 35, 520),
    ("P14", "Denim Jeans 32", "Apparel", "piece", 95, 30, 890),
    ("P15", "Kids T-Shirt Set", "Apparel", "set", 130, 40, 460),
    ("P16", "Saree Cotton Print", "Apparel", "piece", 75, 25, 740),
    ("P17", "Bedsheet King", "Apparel", "piece", 85, 30, 980),
    ("P18", "Towel Set x3", "Apparel", "set", 105, 35, 390),
    ("P19", "Steel Dinner Set", "Home", "set", 60, 20, 1650),
    ("P20", "Pressure Cooker 5L", "Home", "unit", 80, 25, 1120),
    ("P21", "Storage Box 40L", "Home", "unit", 150, 45, 320),
    ("P22", "Floor Cleaner 1L", "Home", "bottle", 190, 50, 145),
    ("P23", "Dish Wash Gel 500ml", "Home", "bottle", 170, 50, 120),
    ("P24", "Broom + Dustpan Set", "Home", "set", 115, 35, 210),
]

# Employees: name, team. 12 shifts each = 120 rows.
EMPLOYEES = [
    ("Asha Verma", "Warehouse"), ("Rahul Singh", "Warehouse"),
    ("Imran Khan", "Warehouse"), ("Priya Nair", "Packing"),
    ("Sanjay Yadav", "Packing"), ("Meera Joshi", "Packing"),
    ("Vikram Rao", "Delivery"), ("Anita Desai", "Delivery"),
    ("Kiran Patel", "Delivery"), ("Neha Gupta", "Admin"),
]

LOW_STOCK = {"P02", "P05", "P09", "P11", "P16", "P20"}


def monday(d):
    return d - datetime.timedelta(days=d.weekday())


def rand_date(lo, hi):
    return lo + datetime.timedelta(days=RNG.randint(0, (hi - lo).days))


def main():
    os.makedirs(DATA, exist_ok=True)

    with open(os.path.join(DATA, "products.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ProductID", "Product", "Category", "Unit", "OpeningStock",
                    "ReorderLevel", "UnitCost"])
        w.writerows(PRODUCTS)
    cost = {p[0]: p[6] for p in PRODUCTS}
    selling = {p[0]: round(p[6] * 1.35) for p in PRODUCTS}

    # ---- Movements: 7 per product = 168 rows ----
    movements = []
    for pid, _name, _cat, _unit, opening, reorder, _c in PRODUCTS:
        balance = opening
        rows = []
        for _ in range(6):
            d = rand_date(START, END)
            kind = RNG.choice(["Receipt", "Receipt", "Sale", "Sale", "Sale", "Adjustment"])
            if kind == "Receipt":
                qty = RNG.randint(20, 80)
                balance += qty
            elif kind == "Sale":
                qty = min(balance - 5, RNG.randint(10, 55))
                qty = max(qty, 5)
                balance -= qty
            else:
                qty = RNG.randint(-5, 5)
                balance += qty
            rows.append([d.isoformat(), pid, kind, monday(d).isoformat(), qty])
        # 7th movement lands the closing stock on target: below reorder for
        # the LOW_STOCK set, comfortably above it for the rest.
        target = max(0, reorder - 8) if pid in LOW_STOCK else reorder + 25
        d = rand_date(datetime.date(2026, 10, 19), END)
        diff = target - balance
        if diff >= 0:
            rows.append([d.isoformat(), pid, "Receipt", monday(d).isoformat(), diff])
            balance += diff
        else:
            qty = min(balance, -diff)
            rows.append([d.isoformat(), pid, "Sale", monday(d).isoformat(), qty])
            balance -= qty
        assert balance == target, (pid, balance, target)
        movements.extend(rows)
    movements.sort(key=lambda r: (r[0], r[1]))
    assert len(movements) == 168
    with open(os.path.join(DATA, "movements.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["MovementDate", "ProductID", "MovementType", "WeekStart", "Quantity"])
        w.writerows(movements)

    # ---- Shifts: 12 per employee = 120 rows ----
    shifts = []
    for name, team in EMPLOYEES:
        for _ in range(12):
            d = rand_date(START, END)
            worked = RNG.choice([6.5, 7.0, 7.5, 8.0, 8.0, 8.5, 9.0, 9.5, 10.0, 10.5, 11.0])
            shifts.append([d.isoformat(), name, team, monday(d).isoformat(), 8, worked])
    shifts.sort(key=lambda r: (r[0], r[1]))
    assert len(shifts) == 120
    with open(os.path.join(DATA, "shifts.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ShiftDate", "Employee", "Team", "WeekStart",
                    "ScheduledHours", "WorkedHours"])
        w.writerows(shifts)

    # ---- Orders: 96 rows, all shipped by 2026-10-31 ----
    orders = []
    for i in range(1, 97):
        pid = RNG.choice(PRODUCTS)[0]
        d = rand_date(START, ORDER_LAST)
        lead = RNG.randint(1, 5)
        ship = min(d + datetime.timedelta(days=lead), END)
        qty = RNG.randint(2, 18)
        orders.append([f"O{i:03d}", d.isoformat(), pid, qty, selling[pid],
                       ship.isoformat(), monday(d).isoformat()])
    orders.sort(key=lambda r: r[1])
    assert len(orders) == 96
    with open(os.path.join(DATA, "orders.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["OrderID", "OrderDate", "ProductID", "Quantity", "UnitPrice",
                    "ShipDate", "WeekStart"])
        w.writerows(orders)

    print(f"products={len(PRODUCTS)} movements={len(movements)} "
          f"shifts={len(shifts)} orders={len(orders)}")


if __name__ == "__main__":
    main()
