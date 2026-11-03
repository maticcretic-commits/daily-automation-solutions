#!/usr/bin/env python3
"""Independent Python ground truth for the PaymentTracker dashboard.

Recomputes every checked figure straight from data/payments.csv plus
the as-of date (duplicated here deliberately — the test also asserts
the workbook's Assumptions sheet carries the same date).
"""
import csv
import datetime
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
AS_OF = datetime.date(2026, 11, 3)


def status_of(amount, paid, due):
    pending = amount - paid
    if pending <= 0:
        return "Paid"
    if due < AS_OF:
        return "Partial - Overdue" if paid > 0 else "Unpaid - Overdue"
    return "Partial" if paid > 0 else "Unpaid"


def compute():
    rows = list(csv.DictReader(open(os.path.join(DATA, "payments.csv"))))
    for r in rows:
        r["amount"] = int(r["amount"])
        r["paid"] = int(r["paid"])
        r["pending"] = r["amount"] - r["paid"]
        r["due"] = datetime.date.fromisoformat(r["due_date"])
        r["inv"] = datetime.date.fromisoformat(r["invoice_date"])
        r["status"] = status_of(r["amount"], r["paid"], r["due"])
        r["days_overdue"] = ((AS_OF - r["due"]).days
                             if r["pending"] > 0 and r["due"] < AS_OF else 0)
        r["month_key"] = r["inv"].year * 100 + r["inv"].month

    tot_inv = sum(r["amount"] for r in rows)
    tot_paid = sum(r["paid"] for r in rows)
    tot_pending = sum(r["pending"] for r in rows)
    overdue = [r for r in rows if r["pending"] > 0 and r["due"] < AS_OF]
    overdue_amt = sum(r["pending"] for r in overdue)
    avg_overdue = (sum(r["days_overdue"] for r in overdue) / len(overdue)
                   if overdue else 0)
    status_counts = defaultdict(int)
    for r in rows:
        status_counts[r["status"]] += 1

    monthly = {}
    for m in range(4, 11):
        key = 202600 + m
        sel = [r for r in rows if r["month_key"] == key]
        inv = sum(r["amount"] for r in sel)
        paid = sum(r["paid"] for r in sel)
        monthly[key] = {"invoiced": inv, "collected": paid,
                        "pending": inv - paid, "count": len(sel)}

    customers = {}
    for r in rows:
        c = customers.setdefault(r["customer"], {"count": 0, "invoiced": 0,
                                                 "paid": 0, "pending": 0,
                                                 "overdue_amt": 0,
                                                 "overdue_count": 0})
        c["count"] += 1
        c["invoiced"] += r["amount"]
        c["paid"] += r["paid"]
        c["pending"] += r["pending"]
        if r in overdue:
            c["overdue_amt"] += r["pending"]
            c["overdue_count"] += 1

    return {
        "count": len(rows), "customers": len(customers),
        "tot_invoiced": tot_inv, "tot_collected": tot_paid,
        "tot_pending": tot_pending,
        "collection_rate": tot_paid / tot_inv if tot_inv else 0,
        "paid_count": status_counts["Paid"],
        "overdue_count": len(overdue), "overdue_amt": overdue_amt,
        "avg_overdue": avg_overdue,
        "partial_count": (status_counts["Partial"]
                          + status_counts["Partial - Overdue"]),
        "status_counts": dict(status_counts),
        "monthly": monthly, "customers_detail": customers,
        "overdue_list": overdue,
    }


if __name__ == "__main__":
    t = compute()
    print(f"{t['count']} invoices, {t['customers']} customers; "
          f"invoiced Rs {t['tot_invoiced']:,}, collected "
          f"Rs {t['tot_collected']:,}, pending Rs {t['tot_pending']:,}, "
          f"overdue {t['overdue_count']} invoices / Rs {t['overdue_amt']:,}")
