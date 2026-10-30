#!/usr/bin/env python3
"""Tests for the Demo Traders Power BI demo proof.

Asserts:
 1. CSVs load with expected columns and row counts.
 2. No nulls/empties in key columns; sales_amount positive; dates in 2025.
 3. Totals in CHECKS.txt match a fresh independent recomputation.
 4. Mockup index.html exists, carries the disclaimer header, and contains
    the headline numbers.
 5. No Power BI (.pbix) file exists, and no text claims a .pbix/Power BI
    build was produced (instructional/future-tense mentions and the
    mockup disclaimer are allowed; existence claims are not).

Run: python3 test_demo.py
Exit code 0 = all pass.
"""

import csv
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent

FAILURES = []


def check(name, cond, detail=""):
    print(("PASS  " if cond else "FAIL  ") + name + (f" — {detail}" if detail else ""))
    if not cond:
        FAILURES.append(name)


def read_checks():
    d = {}
    for line in (BASE / "CHECKS.txt").read_text(encoding="utf-8").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            d[k.strip()] = v.strip()
    return d


def main():
    # 1. CSVs load.
    with open(BASE / "sales.csv", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fields = reader.fieldnames
        sales = list(reader)
    check(
        "sales.csv columns",
        fields == ["order_id", "date", "region", "category", "product",
                    "salesperson", "qty", "unit_price", "discount_pct",
                    "sales_amount"],
        str(fields),
    )
    check("sales.csv row count is 600", len(sales) == 600, f"got {len(sales)}")
    with open(BASE / "targets.csv", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        tfields = reader.fieldnames
        targets = list(reader)
    check(
        "targets.csv columns",
        tfields == ["month", "category", "target_amount"],
        str(tfields),
    )
    check("targets.csv row count is 48", len(targets) == 48, f"got {len(targets)}")

    # 2. No nulls in key columns; values sane.
    key_cols = ["order_id", "date", "region", "category", "product",
                "salesperson", "qty", "unit_price", "discount_pct", "sales_amount"]
    nulls = sum(1 for r in sales for c in key_cols if r[c] is None or r[c] == "")
    check("no nulls/empties in sales key columns", nulls == 0, f"null cells={nulls}")
    check(
        "all sales_amount positive",
        all(float(r["sales_amount"]) > 0 for r in sales),
    )
    check(
        "all dates in 2025",
        all(r["date"].startswith("2025-") for r in sales),
    )
    check(
        "sales_amount equals qty*price*(1-discount)",
        all(
            abs(float(r["sales_amount"]) - round(
                int(r["qty"]) * float(r["unit_price"]) * (1 - float(r["discount_pct"])), 2
            )) < 0.011
            for r in sales
        ),
    )
    check(
        "no nulls in targets",
        all(r[c] not in (None, "") for r in targets for c in tfields),
    )

    # 3. CHECKS.txt matches fresh recomputation.
    checks = read_checks()
    total_sales = round(sum(float(r["sales_amount"]) for r in sales), 2)
    total_qty = sum(int(r["qty"]) for r in sales)
    avg_disc = round(sum(float(r["discount_pct"]) for r in sales) / len(sales) * 100, 2)
    total_target = round(sum(float(r["target_amount"]) for r in targets), 2)
    vs_target = round(total_sales / total_target * 100, 2)
    from collections import defaultdict
    by_cat = defaultdict(float)
    for r in sales:
        by_cat[r["category"]] += float(r["sales_amount"])
    top_cat = max(by_cat, key=by_cat.get)
    expected = {
        "rows_sales": "600",
        "rows_targets": "48",
        "total_sales": f"{total_sales:.2f}",
        "total_qty": str(total_qty),
        "avg_discount_pct": f"{avg_disc:.2f}",
        "total_target": f"{total_target:.2f}",
        "sales_vs_target_pct": f"{vs_target:.2f}",
        "top_category": top_cat,
        "top_category_sales": f"{by_cat[top_cat]:.2f}",
    }
    for k, v in expected.items():
        check(f"CHECKS.txt {k} matches recomputation", checks.get(k) == v,
              f"checks={checks.get(k)} recomputed={v}")

    # 4. Mockup exists, disclaimed, contains headline numbers.
    html_path = BASE / "index.html"
    check("index.html exists", html_path.exists())
    html = html_path.read_text(encoding="utf-8") if html_path.exists() else ""
    check(
        "mockup header disclaimer present",
        "MOCKUP" in html and "VISUAL PREVIEW ONLY, NOT A POWER BI FILE" in html,
    )
    for needle in [
        f"Rs {total_sales:,.2f}",
        f"{total_qty:,}",
        f"{avg_disc:.2f}%",
        f"{vs_target:.2f}%",
        top_cat,
    ]:
        check(f"mockup contains headline number {needle!r}", needle in html)

    # 5. No .pbix file; no existence claim for a Power BI build.
    pbix_files = list(BASE.rglob("*.pbix"))
    check("no .pbix file in folder", not pbix_files, str(pbix_files))
    claim_patterns = [
        r"pbix file was created",
        r"pbix was produced",
        r"pbix file is included",
        r"produced a \.pbix",
        r"built in power bi desktop",
        r"created in power bi",
        r"dashboard was built in power bi",
        r"the power bi file is (attached|included|available)",
        r"power bi file was (built|created|produced)",
    ]
    text_files = [p for p in BASE.rglob("*")
                  if p.is_file() and p.suffix.lower() in {".md", ".html", ".txt", ".py", ".csv"}
                  and p.name not in {"test_demo.py", "VERIFY.txt"}]
    hits = []
    for p in text_files:
        text = p.read_text(encoding="utf-8", errors="ignore")
        for line in text.splitlines():
            low = line.lower()
            # Negated/disclaimer lines ("not built in Power BI", "no .pbix
            # file was produced", "no such file is included") are honest
            # denials, not existence claims — skip them.
            if re.search(r"\b(not|no|never|n't)\b", low):
                continue
            for pat in claim_patterns:
                if re.search(pat, line, re.IGNORECASE):
                    hits.append(f"{p.name}: {pat}: {line.strip()[:100]}")
    check("no .pbix / Power BI build existence claim", not hits, "; ".join(hits))
    # .pbix string may appear only as disclaimer/instruction/spec wording:
    # assert it never appears in the CSV data files.
    csv_hits = [p.name for p in (BASE / "sales.csv", BASE / "targets.csv")
                if ".pbix" in p.read_text(encoding="utf-8", errors="ignore").lower()]
    check("no .pbix mention inside CSV data", not csv_hits, str(csv_hits))

    print()
    if FAILURES:
        print(f"{len(FAILURES)} FAILURE(S): {FAILURES}")
        return 1
    print("ALL TESTS PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
