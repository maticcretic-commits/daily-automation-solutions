"""Deterministic sample data for the Excel Finance Tracker demo.

Writes three CSVs next to this script:
  sample_transactions.csv  Date,Type,Category,Description,Amount
  sample_invoices.csv      InvoiceNo,Client,IssueDate,DueDate,Amount,Status
  sample_budgets.csv       Category,MonthlyBudget
"""
import csv
import os
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))

# ---- monthly expense schedules (Jan..Sep 2026) ----
MARKETING = [8000, 9000, 7500, 11000, 9500, 8500, 12000, 10000, 9000]
TRAVEL = [4000, 3500, 6000, 3000, 4500, 5000, 3500, 4000, 5500]
SUPPLIES = [1500, 1800, 1200, 2000, 1600, 1400, 2100, 1700, 1900]
INSURANCE = {1: 6000, 4: 6000, 7: 6000}          # quarterly premium
PROF_FEES = {2: 4500, 5: 4500, 8: 4500}          # CA retainer, quarterly

BUDGETS = {
    "Rent": 12000,
    "Salaries": 55000,
    "Marketing": 9000,
    "Travel": 4500,
    "Utilities": 2500,
    "Software": 3200,
    "Office Supplies": 1700,
    "Insurance": 2000,
    "Professional Fees": 1500,
}


def transactions():
    rows = []
    for m in range(1, 10):
        d = date(2026, m, 5)
        rows.append((d, "Income", "Client Retainer", "Monthly retainer",
                     42000 if m <= 3 else 45000 if m <= 6 else 48000))
        rows.append((date(2026, m, 20), "Income", "Product Sales",
                     "Online product sales", 17500 + 500 * m))
        rows.append((date(2026, m, 1), "Expense", "Rent", "Office rent", 12000))
        rows.append((date(2026, m, 3), "Expense", "Salaries", "Team payroll", 55000))
        rows.append((date(2026, m, 10), "Expense", "Marketing", "Ads & campaigns",
                     MARKETING[m - 1]))
        rows.append((date(2026, m, 15), "Expense", "Travel", "Field travel",
                     TRAVEL[m - 1]))
        rows.append((date(2026, m, 12), "Expense", "Utilities", "Power & internet",
                     2500))
        rows.append((date(2026, m, 8), "Expense", "Software", "SaaS subscriptions",
                     3200))
        rows.append((date(2026, m, 18), "Expense", "Office Supplies", "Stationery",
                     SUPPLIES[m - 1]))
        if m in INSURANCE:
            rows.append((date(2026, m, 2), "Expense", "Insurance",
                         "Quarterly premium", INSURANCE[m]))
        if m in PROF_FEES:
            rows.append((date(2026, m, 25), "Expense", "Professional Fees",
                         "CA retainer", PROF_FEES[m]))
    return rows


# InvoiceNo, Client, IssueDate, Amount, Status (DueDate = Issue + 30d)
INVOICES = [
    ("INV-26001", "Acme Corp", date(2026, 6, 2), 45000, "Paid"),
    ("INV-26002", "Beta Ltd", date(2026, 6, 9), 28000, "Paid"),
    ("INV-26003", "Gamma Inc", date(2026, 6, 16), 52000, "Paid"),
    ("INV-26004", "Delta Co", date(2026, 6, 23), 31000, "Paid"),
    ("INV-26005", "Epsilon LLC", date(2026, 6, 30), 47000, "Paid"),
    ("INV-26006", "Acme Corp", date(2026, 7, 7), 39000, "Paid"),
    ("INV-26007", "Beta Ltd", date(2026, 7, 14), 55000, "Paid"),
    ("INV-26008", "Gamma Inc", date(2026, 7, 21), 26000, "Paid"),
    ("INV-26009", "Delta Co", date(2026, 7, 28), 61000, "Paid"),
    ("INV-26010", "Epsilon LLC", date(2026, 8, 4), 33000, "Paid"),
    ("INV-26011", "Acme Corp", date(2026, 8, 11), 48000, "Paid"),
    ("INV-26012", "Beta Ltd", date(2026, 8, 18), 29000, "Paid"),
    ("INV-26013", "Gamma Inc", date(2026, 8, 25), 57000, "Paid"),
    ("INV-26014", "Delta Co", date(2026, 8, 16), 34000, "Unpaid"),  # due 2026-09-15 OVERDUE
    ("INV-26015", "Epsilon LLC", date(2026, 8, 21), 42000, "Unpaid"),  # due 2026-09-20 OVERDUE
    ("INV-26016", "Acme Corp", date(2026, 8, 29), 51000, "Unpaid"),  # due 2026-09-28 OVERDUE
    ("INV-26017", "Beta Ltd", date(2026, 9, 20), 36000, "Unpaid"),  # due 2026-10-20 outstanding
    ("INV-26018", "Gamma Inc", date(2026, 9, 25), 44000, "Unpaid"),  # due 2026-10-25 outstanding
    ("INV-26019", "Delta Co", date(2026, 10, 5), 38000, "Unpaid"),  # due 2026-11-04 outstanding
    ("INV-26020", "Epsilon LLC", date(2026, 9, 1), 46000, "Paid"),
]


def main():
    tx = transactions()
    with open(os.path.join(HERE, "sample_transactions.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Type", "Category", "Description", "Amount"])
        for r in tx:
            w.writerow([r[0].isoformat(), r[1], r[2], r[3], r[4]])
    with open(os.path.join(HERE, "sample_invoices.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["InvoiceNo", "Client", "IssueDate", "DueDate", "Amount", "Status"])
        for no, client, issued, amt, status in INVOICES:
            from datetime import timedelta
            w.writerow([no, client, issued.isoformat(),
                        (issued + timedelta(days=30)).isoformat(), amt, status])
    with open(os.path.join(HERE, "sample_budgets.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Category", "MonthlyBudget"])
        for cat, amt in BUDGETS.items():
            w.writerow([cat, amt])
    print(f"wrote {len(tx)} transactions, {len(INVOICES)} invoices, "
          f"{len(BUDGETS)} budgets")


if __name__ == "__main__":
    main()
