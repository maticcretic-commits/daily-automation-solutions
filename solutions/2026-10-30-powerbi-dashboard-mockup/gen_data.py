#!/usr/bin/env python3
"""Generate synthetic datasets for the Demo Traders Power BI demo proof.

Deterministic (seed 42). Fictional business, fictional people, synthetic
data only. Also computes headline numbers, writes CHECKS.txt, and builds
the HTML mockup (index.html) from the same CSVs so the mockup numbers
match the data exactly.

Run:  python3 gen_data.py
Outputs: sales.csv, targets.csv, CHECKS.txt, index.html
"""

import csv
import random
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

SEED = 42
BASE = Path(__file__).resolve().parent
N_ROWS = 600

CATEGORIES = {
    "Electronics": [("Laptop Stand", 2400), ("Wireless Mouse", 650), ("USB Hub", 900)],
    "Furniture": [("Office Chair", 5500), ("Desk Lamp", 1200), ("Filing Cabinet", 3800)],
    "Office Supplies": [
        ("Printer Paper (ream)", 320),
        ("Ball Pen (pack)", 180),
        ("Notebook (A4)", 240),
    ],
    "Appliances": [("Electric Kettle", 1400), ("Table Fan", 2100), ("Room Heater", 2800)],
}
# Monthly target (INR) per category before seeded variation.
BASE_TARGET = {
    "Electronics": 180000,
    "Furniture": 150000,
    "Office Supplies": 45000,
    "Appliances": 105000,
}
REGIONS = ["North", "South", "East", "West"]
REGION_WEIGHTS = [0.30, 0.28, 0.22, 0.20]
CATEGORY_WEIGHTS = [0.32, 0.27, 0.21, 0.20]
SALESPEOPLE = [
    "Ananya Iyer",
    "Vikram Rao",
    "Meera Nair",
    "Arjun Mehta",
    "Kavya Reddy",
    "Rohan Desai",
]
DISCOUNTS = [0.00, 0.00, 0.05, 0.05, 0.10, 0.15]


def fmt_inr(x: float) -> str:
    return f"Rs {x:,.2f}"


def fmt_num(x: float) -> str:
    return f"{x:,.2f}"


def bar_chart_svg(title, labels, values, width=640, height=260, color="#2f5597"):
    """Simple inline SVG vertical bar chart. No external dependency."""
    pad_l, pad_b, pad_t = 56, 34, 26
    plot_w = width - pad_l - 24
    plot_h = height - pad_t - pad_b
    vmax = max(values) if values else 1
    n = len(values)
    slot = plot_w / n
    bar_w = min(56, slot * 0.62)
    parts = [
        f'<svg viewBox="0 0 {width} {height}" width="100%" '
        f'role="img" aria-label="{title}" xmlns="http://www.w3.org/2000/svg">',
        f'<text x="{pad_l}" y="18" font-size="14" font-weight="bold" '
        f'fill="#1f3864">{title}</text>',
    ]
    for i, (lab, val) in enumerate(zip(labels, values)):
        h = (val / vmax) * plot_h if vmax else 0
        x = pad_l + i * slot + (slot - bar_w) / 2
        y = pad_t + (plot_h - h)
        parts.append(
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{h:.1f}" '
            f'fill="{color}"><title>{lab}: {fmt_num(val)}</title></rect>'
        )
        parts.append(
            f'<text x="{x + bar_w / 2:.1f}" y="{y - 5:.1f}" font-size="11" '
            f'text-anchor="middle" fill="#333">{fmt_num(val)}</text>'
        )
        parts.append(
            f'<text x="{x + bar_w / 2:.1f}" y="{height - 12}" font-size="11" '
            f'text-anchor="middle" fill="#333">{lab}</text>'
        )
    parts.append("</svg>")
    return "\n".join(parts)


def main():
    rng = random.Random(SEED)
    start = date(2025, 1, 1)

    rows = []
    for i in range(1, N_ROWS + 1):
        day = start + timedelta(days=rng.randrange(0, 365))
        region = rng.choices(REGIONS, weights=REGION_WEIGHTS, k=1)[0]
        category = rng.choices(
            list(CATEGORIES), weights=CATEGORY_WEIGHTS, k=1
        )[0]
        product, unit_price = rng.choice(CATEGORIES[category])
        qty = rng.randint(1, 12)
        discount = rng.choice(DISCOUNTS)
        salesperson = rng.choice(SALESPEOPLE)
        amount = round(qty * unit_price * (1 - discount), 2)
        rows.append(
            {
                "order_id": f"DT-{i:04d}",
                "date": day.isoformat(),
                "region": region,
                "category": category,
                "product": product,
                "salesperson": salesperson,
                "qty": qty,
                "unit_price": unit_price,
                "discount_pct": discount,
                "sales_amount": amount,
            }
        )

    sales_path = BASE / "sales.csv"
    with open(sales_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(
            f,
            fieldnames=[
                "order_id",
                "date",
                "region",
                "category",
                "product",
                "salesperson",
                "qty",
                "unit_price",
                "discount_pct",
                "sales_amount",
            ],
        )
        w.writeheader()
        for r in rows:
            w.writerow(r)

    # Targets: one row per category x month, seeded variation +/-8%.
    target_rows = []
    for month in range(1, 13):
        for category, base in BASE_TARGET.items():
            target = round(base * (1 + rng.uniform(-0.08, 0.08)), 2)
            target_rows.append(
                {"month": f"2025-{month:02d}", "category": category, "target_amount": target}
            )
    targets_path = BASE / "targets.csv"
    with open(targets_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["month", "category", "target_amount"])
        w.writeheader()
        for r in target_rows:
            w.writerow(r)

    # Headline numbers, computed from the CSVs just written.
    total_sales = round(sum(r["sales_amount"] for r in rows), 2)
    total_qty = sum(r["qty"] for r in rows)
    avg_discount = round(
        sum(r["discount_pct"] for r in rows) / len(rows) * 100, 2
    )
    total_target = round(sum(r["target_amount"] for r in target_rows), 2)
    vs_target = round(total_sales / total_target * 100, 2)
    by_category = defaultdict(float)
    for r in rows:
        by_category[r["category"]] += r["sales_amount"]
    top_category = max(by_category, key=by_category.get)
    top_sales = round(by_category[top_category], 2)

    checks = (
        f"rows_sales={len(rows)}\n"
        f"rows_targets={len(target_rows)}\n"
        f"total_sales={total_sales:.2f}\n"
        f"total_qty={total_qty}\n"
        f"avg_discount_pct={avg_discount:.2f}\n"
        f"total_target={total_target:.2f}\n"
        f"sales_vs_target_pct={vs_target:.2f}\n"
        f"top_category={top_category}\n"
        f"top_category_sales={top_sales:.2f}\n"
    )
    (BASE / "CHECKS.txt").write_text(checks, encoding="utf-8")

    # Aggregations for the mockup.
    months = [f"2025-{m:02d}" for m in range(1, 13)]
    month_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    by_month = defaultdict(float)
    by_month_target = defaultdict(float)
    for r in rows:
        by_month[r["date"][:7]] += r["sales_amount"]
    for r in target_rows:
        by_month_target[r["month"]] += r["target_amount"]
    month_sales = [round(by_month.get(m, 0.0), 2) for m in months]
    month_targets = [round(by_month_target.get(m, 0.0), 2) for m in months]
    month_pct = [
        round(s / t * 100, 1) if t else 0.0 for s, t in zip(month_sales, month_targets)
    ]
    cat_order = sorted(by_category, key=by_category.get, reverse=True)
    cat_sales = [round(by_category[c], 2) for c in cat_order]
    by_region = defaultdict(float)
    for r in rows:
        by_region[r["region"]] += r["sales_amount"]
    region_order = ["North", "South", "East", "West"]
    region_sales = [round(by_region.get(rg, 0.0), 2) for rg in region_order]
    by_person = defaultdict(float)
    for r in rows:
        by_person[r["salesperson"]] += r["sales_amount"]
    person_rows = "".join(
        f"<tr><td>{p}</td><td class='num'>{fmt_inr(v)}</td></tr>"
        for p, v in sorted(by_person.items(), key=lambda kv: kv[1], reverse=True)
    )
    target_rows_html = "".join(
        f"<tr><td>{c}</td><td class='num'>{fmt_inr(round(by_category[c], 2))}</td>"
        f"<td class='num'>{fmt_inr(round(sum(r['target_amount'] for r in target_rows if r['category'] == c), 2))}</td>"
        f"<td class='num'>{round(by_category[c] / sum(r['target_amount'] for r in target_rows if r['category'] == c) * 100, 1)}%</td></tr>"
        for c in cat_order
    )
    monthly_table = "".join(
        f"<tr><td>{lab} 2025</td><td class='num'>{fmt_inr(s)}</td>"
        f"<td class='num'>{fmt_inr(t)}</td><td class='num'>{p}%</td></tr>"
        for lab, s, t, p in zip(month_labels, month_sales, month_targets, month_pct)
    )

    chart_month = bar_chart_svg("Sales by month (Rs)", month_labels, month_sales)
    chart_cat = bar_chart_svg("Sales by category (Rs)", cat_order, cat_sales, color="#548235")
    chart_region = bar_chart_svg("Sales by region (Rs)", region_order, region_sales, color="#bf8f00")
    chart_target = bar_chart_svg(
        "Target achievement by month (%)", month_labels, month_pct, color="#7030a0"
    )

    html = f"""<!DOCTYPE html>
<html lang="en-GB">
<head>
<meta charset="utf-8">
<title>Demo Traders — Dashboard Mockup (visual preview only)</title>
<style>
body {{ font-family: Arial, Helvetica, sans-serif; margin: 24px; color: #222; background: #f2f2f2; }}
.banner {{ background: #c00000; color: #fff; padding: 12px 16px; font-weight: bold; }}
.card {{ background: #fff; border: 1px solid #ccc; padding: 16px; margin: 12px 0; }}
.kpis {{ display: flex; gap: 12px; flex-wrap: wrap; }}
.kpi {{ background: #fff; border: 1px solid #ccc; padding: 12px 18px; min-width: 170px; }}
.kpi .v {{ font-size: 22px; font-weight: bold; color: #1f3864; }}
.kpi .l {{ font-size: 12px; color: #555; }}
.tabs button {{ padding: 8px 16px; margin-right: 6px; cursor: pointer; }}
table {{ border-collapse: collapse; width: 100%; }}
th, td {{ border: 1px solid #ccc; padding: 6px 10px; text-align: left; }}
td.num, th.num {{ text-align: right; }}
.note {{ font-size: 13px; color: #555; }}
</style>
</head>
<body>
<div class="banner">MOCKUP &mdash; VISUAL PREVIEW ONLY, NOT A POWER BI FILE. No .pbix file was produced. Synthetic data only.</div>
<h1>Demo Traders &mdash; Sales Dashboard (mockup)</h1>
<p class="note">Fictional business &ldquo;Demo Traders&rdquo;. All data synthetic, generated with seed 42 (see DATA_DICTIONARY.md). This page is a static visual preview of the dashboard described in POWERBI_SPEC.md. It is not Power BI and was not built in Power BI Desktop.</p>

<div class="kpis">
<div class="kpi"><div class="v">{fmt_inr(total_sales)}</div><div class="l">Total Sales (600 orders, 2025)</div></div>
<div class="kpi"><div class="v">{total_qty:,}</div><div class="l">Total Qty</div></div>
<div class="kpi"><div class="v">{avg_discount:.2f}%</div><div class="l">Avg Discount %</div></div>
<div class="kpi"><div class="v">{vs_target:.2f}%</div><div class="l">Sales vs Target %</div></div>
<div class="kpi"><div class="v">{top_category}</div><div class="l">Top Category ({fmt_inr(top_sales)})</div></div>
</div>

<div class="tabs">
<button onclick="showPage(1)">Page 1: Sales overview</button>
<button onclick="showPage(2)">Page 2: Category / region breakdown</button>
<button onclick="showPage(3)">Page 3: Target tracking</button>
</div>

<div class="card" id="page1">
<h2>Page 1 &mdash; Sales overview</h2>
{chart_month}
<p class="note">Headline numbers: Total Sales {fmt_inr(total_sales)}; Total Qty {total_qty:,}; Avg Discount {avg_discount:.2f}%; Sales vs Target {vs_target:.2f}%; Top Category {top_category} ({fmt_inr(top_sales)}).</p>
</div>

<div class="card" id="page2" style="display:none">
<h2>Page 2 &mdash; Category / region breakdown</h2>
{chart_cat}
{chart_region}
<h3>Sales by salesperson</h3>
<table><tr><th>Salesperson (fictional)</th><th class="num">Sales (Rs)</th></tr>{person_rows}</table>
</div>

<div class="card" id="page3" style="display:none">
<h2>Page 3 &mdash; Target tracking</h2>
{chart_target}
<h3>Category vs target (full year 2025)</h3>
<table><tr><th>Category</th><th class="num">Sales (Rs)</th><th class="num">Target (Rs)</th><th class="num">Achievement</th></tr>{target_rows_html}</table>
<h3>Monthly sales vs target</h3>
<table><tr><th>Month</th><th class="num">Sales (Rs)</th><th class="num">Target (Rs)</th><th class="num">Achievement</th></tr>{monthly_table}</table>
<p class="note">Total target {fmt_inr(total_target)}; overall achievement {vs_target:.2f}%.</p>
</div>

<script>
function showPage(n) {{
  document.getElementById('page1').style.display = n === 1 ? 'block' : 'none';
  document.getElementById('page2').style.display = n === 2 ? 'block' : 'none';
  document.getElementById('page3').style.display = n === 3 ? 'block' : 'none';
}}
</script>
</body>
</html>
"""
    (BASE / "index.html").write_text(html, encoding="utf-8")
    print(checks)
    print("Wrote sales.csv, targets.csv, CHECKS.txt, index.html")


if __name__ == "__main__":
    main()
