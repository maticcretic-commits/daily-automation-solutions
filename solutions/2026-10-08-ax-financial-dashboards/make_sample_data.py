"""Generate a deterministic Dynamics AX 2012-style GL trial-balance export.

Output: sample_ax_export.csv
Columns: FiscalYear, Period, MainAccount, AccountName, AccountCategory, Debit, Credit

Double-entry construction (this is what makes it behave like a real export):
  * P&L and balance-sheet accounts (except Cash) get realistic synthetic values.
  * Cash & Bank is the RESIDUAL: cash(p) = -(sum of all other trial-balance
    nets for p), exactly how cash absorbs every transaction in a real ledger.
    This guarantees total debits == total credits every period with no fudge.
  * Retained Earnings carries the OPENING balance only (50,000 credit); the
    current period's profit is NOT closed into it yet -- the dashboard rolls it
    in when drawing the balance sheet, which is standard pre-close practice.
"""
import csv
import random

SEED = 42
FISCAL_YEAR = 2026
OUT = "sample_ax_export.csv"
OPENING_RETAINED_EARNINGS = -50000.0  # trial-balance net (credit normal)

ACCOUNTS = [
    # (main account, name, category)
    ("401000", "Product Sales", "Revenue"),
    ("402000", "Service Revenue", "Revenue"),
    ("403000", "Other Income", "Revenue"),
    ("501000", "Cost of Goods Sold", "Expense"),
    ("502000", "Salaries & Benefits", "Expense"),
    ("503000", "Rent & Facilities", "Expense"),
    ("504000", "Marketing & Advertising", "Expense"),
    ("505000", "Travel & Entertainment", "Expense"),
    ("506000", "IT & Software", "Expense"),
    ("507000", "Utilities", "Expense"),
    ("508000", "Depreciation", "Expense"),
    ("509000", "Interest Expense", "Expense"),
    ("510000", "Tax Provision", "Expense"),
    ("101000", "Cash & Bank", "Asset"),          # residual, computed last
    ("102000", "Accounts Receivable", "Asset"),
    ("103000", "Inventory", "Asset"),
    ("104000", "Prepaid Expenses", "Asset"),
    ("105000", "Fixed Assets (Gross)", "Asset"),
    ("106000", "Accumulated Depreciation", "Asset-Contra"),
    ("201000", "Accounts Payable", "Liability"),
    ("202000", "Accrued Expenses", "Liability"),
    ("203000", "Short-Term Debt", "Liability"),
    ("204000", "Long-Term Debt", "Liability"),
    ("205000", "Taxes Payable", "Liability"),
    ("301000", "Share Capital", "Equity"),
    ("302000", "Retained Earnings", "Equity"),
]

N_PERIODS = 12


def j(rng, base, pct):
    """Deterministic jitter: base * (1 +/- pct)."""
    return base * (1 + rng.uniform(-pct, pct))


def build_rows():
    rng = random.Random(SEED)
    rows = []
    for period in range(1, N_PERIODS + 1):
        rev = j(rng, 180000 * (1.02 ** (period - 1)), 0.10)
        cogs = j(rng, 0.45 * rev, 0.08)
        # trial-balance net amounts (Debit - Credit) per account for this period
        nets = {
            "401000": -j(rng, 0.78 * rev, 0.03),
            "402000": -j(rng, 0.19 * rev, 0.05),
            "403000": -j(rng, 0.03 * rev, 0.10),
            "501000": cogs,
            "502000": j(rng, 48000 * (1.01 ** (period - 1)), 0.05),
            "503000": 12000.0,
            "504000": j(rng, 9000, 0.20),
            "505000": j(rng, 6000, 0.25),
            "506000": j(rng, 4000, 0.10),
            "507000": j(rng, 2500, 0.10),
            "508000": 7000.0,
            "509000": 3000.0,
            "510000": 0.0,  # placeholder, computed below
            "102000": j(rng, 0.15 * rev, 0.12),
            "103000": j(rng, 90000 + 500 * period, 0.06),
            "104000": max(0.0, 8000 - 400 * period),
            "105000": 500000.0,
            "106000": -(35000 + 7000 * period),  # contra asset, credit normal
            "201000": -j(rng, 30000 + 1000 * period, 0.10),
            "202000": -j(rng, 12000, 0.10),
            "203000": -50000.0,
            "204000": -(200000 - 2000 * period),
            "205000": -j(rng, 18000, 0.15),
            "301000": -300000.0,  # equity, credit normal
            "302000": OPENING_RETAINED_EARNINGS,
        }
        # tax provision = 22% of pre-tax profit for the period
        revenue_total = -(nets["401000"] + nets["402000"] + nets["403000"])
        expense_total = sum(v for k, v in nets.items()
                            if k.startswith("5") and k != "510000")
        nets["510000"] = round(0.22 * max(0.0, revenue_total - expense_total), 2)
        # Cash & Bank = residual: absorbs every other movement, so the
        # trial balance proves out exactly (double-entry, no fudge account).
        others = round(sum(nets.values()), 2)
        nets["101000"] = -others
        for main, name, cat in ACCOUNTS:
            net = round(nets[main], 2)
            debit = net if net >= 0 else 0.0
            credit = -net if net < 0 else 0.0
            rows.append({
                "FiscalYear": FISCAL_YEAR,
                "Period": period,
                "MainAccount": main,
                "AccountName": name,
                "AccountCategory": cat,
                "Debit": round(debit, 2),
                "Credit": round(credit, 2),
            })
    return rows


def main():
    rows = build_rows()
    with open(OUT, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["FiscalYear", "Period", "MainAccount",
                                          "AccountName", "AccountCategory",
                                          "Debit", "Credit"])
        w.writeheader()
        w.writerows(rows)
    print(f"wrote {OUT}: {len(rows)} rows")


if __name__ == "__main__":
    main()
