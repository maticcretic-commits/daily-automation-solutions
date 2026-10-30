# Data dictionary — Demo Traders synthetic data

Fictional business: **Demo Traders** (a made-up wholesale trader). All data is
synthetic, generated deterministically by `gen_data.py` with seed 42. No real
company, customer, or salesperson is represented. Person names are fictional;
any resemblance to a real person is coincidental.

Regenerate exactly: `python3 gen_data.py` (Python 3 standard library only).
Re-running produces byte-identical CSVs.

## sales.csv — 600 rows, one row per order line

| Column | Type | Meaning | How generated |
|---|---|---|---|
| `order_id` | text | Unique order-line id, `DT-0001` … `DT-0600` | Sequential |
| `date` | date (ISO) | Order date in 2025 | Uniform random day in 2025 (seeded) |
| `region` | text | Sales region: North, South, East, West | Weighted random (30/28/22/20%) |
| `category` | text | Product category: Electronics, Furniture, Office Supplies, Appliances | Weighted random (32/27/21/20%) |
| `product` | text | Product name within the category | Random choice from the category's 3 products |
| `salesperson` | text | Fictional salesperson name (6 names) | Uniform random |
| `qty` | integer | Units sold, 1–12 | Uniform random integer |
| `unit_price` | number (Rs) | Fixed list price for the product | Fixed price list in `gen_data.py` (e.g. Office Chair Rs 5,500; Ball Pen pack Rs 180) |
| `discount_pct` | decimal | Discount fraction: 0.00, 0.05, 0.10 or 0.15 | Random choice, 0.00 weighted double |
| `sales_amount` | number (Rs, 2 dp) | Net line amount | `qty * unit_price * (1 - discount_pct)`, rounded to 2 dp |

No nulls. `sales_amount` is always positive. Dates cover all 12 months of 2025.

## targets.csv — 48 rows, one row per category × month

| Column | Type | Meaning | How generated |
|---|---|---|---|
| `month` | text (`YYYY-MM`) | Target month in 2025 | All 12 months × 4 categories |
| `category` | text | Product category (same 4 as sales.csv) | All 4 categories |
| `target_amount` | number (Rs, 2 dp) | Monthly sales target for the category | Base monthly target per category (Electronics Rs 180,000; Furniture Rs 150,000; Office Supplies Rs 45,000; Appliances Rs 105,000) × uniform variation of −8% to +8% (seeded) |

## Headline numbers (from CHECKS.txt, recomputed by `gen_data.py`)

- Total Sales: Rs 7,189,354.50
- Total Qty: 4,030
- Avg Discount %: 5.73% (simple mean of `discount_pct` across rows × 100)
- Total Target: Rs 5,758,730.18
- Sales vs Target %: 124.84%
- Top Category: Furniture (Rs 3,582,915.00)

Note: targets were set before seeing the generated sales, so achievement above
100% is a property of this synthetic draw, not a claim about any real business.
