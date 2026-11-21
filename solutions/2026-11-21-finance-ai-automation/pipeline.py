#!/usr/bin/env python3
"""FinanceFlow AI — extraction + classification pipeline (stdlib only).

Reads data/bank_statement.csv (messy bank narrations) and
data/chart_of_accounts.csv, and produces auditable outputs in out/:

  classified_transactions.csv — every row with extracted counterparty,
      category, confidence, derived net + month, and anomaly flags
  monthly_summary.csv         — income / expenses / net per month
  category_summary.csv        — totals per chart-of-accounts category
  anomalies.csv               — every flagged row (duplicate, balance
                                mismatch, unclassified, large amount)

Classification is a deterministic ordered rule engine over the narration
keywords (the chart of accounts is the rule source). The engine is
deliberately swappable: in production the same interface fronts an LLM
call (e.g. Claude) for narrations the rules cannot classify — here the
rules cover everything except the genuinely cryptic lines, which are
flagged for human review instead of guessed.

Run:  python3 pipeline.py            (reads data/, writes out/)
"""
import csv
import json
import os
import re
from collections import defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "out")
OPENING_BALANCE = 85000.00
LARGE_AMOUNT = 60000.00

# Rule order matters: specific categories first, the generic PAYMENT
# fallback for Sales last. (category, keywords, confidence)
RULES = [
    ("Purchases (COGS)", ["PURCHASE", "STOCK", "SUPPLIER"], 1.0),
    ("Rent", ["RENT"], 1.0),
    ("Salaries", ["SALARY"], 1.0),
    ("Utilities", ["ELECTRICITY", "NBPDCL", "BSNL", "BROADBAND", "WATER"], 1.0),
    ("Transport & Fuel", ["PETROL", "DIESEL", "OLA", "DELIVERY", "TRANSPORT"], 1.0),
    ("GST Paid", ["GST"], 1.0),
    ("Bank Charges", ["CHARGES", "SMS", "ATM"], 1.0),
    ("Interest Income", ["INT CREDIT", "INTEREST"], 1.0),
    ("Owner Drawings", ["SELF", "TRANSFER"], 1.0),
    ("Sales Income", ["SALES"], 1.0),
    ("Sales Income", ["PAYMENT"], 0.7),
]

KIND = {}  # filled from chart_of_accounts.csv


def load_coa():
    with open(os.path.join(DATA, "chart_of_accounts.csv"), newline="") as f:
        for row in csv.DictReader(f):
            KIND[row["category"]] = row["kind"]


def classify(narration):
    """Return (category, confidence). Unmatched -> Unclassified, 0.0."""
    upper = narration.upper()
    for category, keywords, conf in RULES:
        if any(k in upper for k in keywords):
            return category, conf
    return "Unclassified", 0.0


def extract_counterparty(narration):
    """Best-effort counterparty from a messy narration.

    Strip the UPI/NEFT/IMPS payment-mode prefix first, then:
    "SALARY - X" -> X; otherwise the first non-numeric '/'-segment.
    """
    s = narration
    m = re.match(r"^(?:UPI|NEFT|IMPS)/(.*)$", s, re.I)
    if m:
        s = m.group(1)
    m = re.match(r"^SALARY\s*-\s*(.+)$", s, re.I)
    if m:
        return m.group(1).strip().title()
    parts = [p for p in s.split("/") if p.strip() and not p.strip().isdigit()]
    if parts:
        return parts[0].strip().title()
    return s.strip().title()[:40]


def run():
    load_coa()
    os.makedirs(OUT, exist_ok=True)

    with open(os.path.join(DATA, "bank_statement.csv"), newline="") as f:
        raw = list(csv.DictReader(f))

    classified = []
    anomalies = []
    seen = set()
    running = OPENING_BALANCE
    monthly = defaultdict(lambda: {"income": 0.0, "expenses": 0.0, "count": 0})
    by_cat = defaultdict(lambda: {"total": 0.0, "count": 0})

    for i, row in enumerate(raw, start=2):  # row 1 is the CSV header
        d = row["date"]
        narr = row["narration"]
        debit = float(row["debit"])
        credit = float(row["credit"])
        stated = float(row["balance"])
        running = round(running - debit + credit, 2)
        month = d[:7]

        category, conf = classify(narr)
        kind = KIND.get(category, "expense" if debit else "income")
        if category == "Unclassified":
            kind = "unclassified"
        net = round(credit - debit, 2)
        flags = []

        key = (d, narr, debit, credit)
        if key in seen:
            flags.append("DUPLICATE")
            anomalies.append((i, d, narr, "DUPLICATE",
                              "identical date/amount/narration seen earlier"))
        seen.add(key)

        if abs(stated - running) > 0.005:
            flags.append("BALANCE_MISMATCH")
            anomalies.append((i, d, narr, "BALANCE_MISMATCH",
                              "stated %.2f vs recomputed %.2f (diff %.2f)"
                              % (stated, running, stated - running)))

        if category == "Unclassified":
            flags.append("UNCLASSIFIED")
            anomalies.append((i, d, narr, "UNCLASSIFIED",
                              "no chart-of-accounts keyword matched; "
                              "needs human review, not guessed"))

        if max(debit, credit) > LARGE_AMOUNT:
            flags.append("LARGE_AMOUNT")
            anomalies.append((i, d, narr, "LARGE_AMOUNT",
                              "single leg %.2f exceeds %.2f review threshold"
                              % (max(debit, credit), LARGE_AMOUNT)))

        classified.append({
            "row": i, "date": d, "narration": narr,
            "counterparty": extract_counterparty(narr),
            "category": category, "kind": kind, "confidence": conf,
            "debit": debit, "credit": credit, "net": net, "month": month,
            "flags": ";".join(flags),
        })
        monthly[month]["income"] += credit
        monthly[month]["expenses"] += debit
        monthly[month]["count"] += 1
        by_cat[category]["total"] += net
        by_cat[category]["count"] += 1

    # --- outputs -------------------------------------------------------------
    with open(os.path.join(OUT, "classified_transactions.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["row", "date", "narration",
                                          "counterparty", "category", "kind",
                                          "confidence", "debit", "credit",
                                          "net", "month", "flags"])
        w.writeheader()
        w.writerows(classified)

    months = sorted(monthly)
    with open(os.path.join(OUT, "monthly_summary.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "income", "expenses", "net", "txn_count"])
        for m in months:
            v = monthly[m]
            w.writerow([m, "%.2f" % v["income"], "%.2f" % v["expenses"],
                        "%.2f" % (v["income"] - v["expenses"]), v["count"]])

    total_net = sum(v["total"] for v in by_cat.values())
    total_abs = sum(abs(v["total"]) for v in by_cat.values())
    with open(os.path.join(OUT, "category_summary.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["category", "kind", "total", "count", "share_pct"])
        for cat in sorted(by_cat, key=lambda c: -abs(by_cat[c]["total"])):
            v = by_cat[cat]
            share = (abs(v["total"]) / total_abs * 100) if total_abs else 0.0
            w.writerow([cat, KIND.get(cat, "unclassified"),
                        "%.2f" % v["total"], v["count"], "%.2f" % share])

    with open(os.path.join(OUT, "anomalies.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["row", "date", "narration", "flag", "detail"])
        w.writerows(anomalies)

    summary = {
        "transactions": len(classified),
        "months": months,
        "opening_balance": OPENING_BALANCE,
        "closing_balance": running,
        "total_income": round(sum(v["income"] for v in monthly.values()), 2),
        "total_expenses": round(sum(v["expenses"] for v in monthly.values()), 2),
        "net": round(sum(v["income"] - v["expenses"] for v in monthly.values()), 2),
        "unclassified": sum(1 for c in classified if c["category"] == "Unclassified"),
        "anomalies": len(anomalies),
        "anomaly_flags": sorted({a[3] for a in anomalies}),
    }
    with open(os.path.join(OUT, "summary.json"), "w") as f:
        json.dump(summary, f, indent=2)

    print(json.dumps(summary, indent=2))
    return summary


if __name__ == "__main__":
    run()
