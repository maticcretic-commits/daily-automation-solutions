"""DPRPulse ground truth — deterministic financial model for a fictional
turmeric-processing Detailed Project Report (Sharma Agro Foods, Patna, Bihar).

All money figures in INR lakh. The case is fictional; every number is produced
by the formulas below, which the test suite re-checks against the workbook
after a LibreOffice headless recalculation.

Accounting identities enforced:
  * cash accruals = PAT + depreciation (+ amortisation)
  * net cash to equity = cash accruals - principal repaid
      (interest is already inside PAT, so it is NOT subtracted again)
  * balance-sheet identity: net block + prelim + cash = equity + reserves + loan
"""
from datetime import date

AS_OF = date(2026, 10, 21)

# ---------------------------------------------------------------- project cost
CAPEX = [
    ("Land & site development", 30.0),
    ("Building & civil works", 55.0),
    ("Plant & machinery", 110.0),
    ("Contingency", 10.0),
    ("Electricals & utilities", 18.0),
    ("Preliminary & pre-operative", 12.0),
    ("Margin for working capital", 5.0),
]
PROJECT_COST = sum(v for _, v in CAPEX)          # 240.0

PM_BASE = 110.0 + 10.0                            # P&M incl. contingency -> 120
BLOCK_GROSS = 30.0 + 55.0 + PM_BASE + 18.0        # 223.0
DEPR_BUILDING = 55.0 / 20                         # 2.75
DEPR_PM = PM_BASE / 10                            # 12.0
DEPR_ELEC = 18.0 / 10                             # 1.8
DEPR_BASE = DEPR_BUILDING + DEPR_PM + DEPR_ELEC   # 16.55 /yr
PRELIM_AMORT = 12.0 / 5                           # 2.4 /yr, years 1..5 only

# ------------------------------------------------------------------- finance
EQUITY = 60.0
LOAN = 180.0
RATE = 0.10
MORATORIUM = 1            # year 1: interest serviced, no principal
REPAY_YEARS = 6           # annual instalments in years 2..7

# ---------------------------------------------------------------- operations
CAP_KG = 1_200_000
PRICE0 = 150.0            # Rs/kg in year 1
PRICE_G = 0.03
UTIL = [0.55, 0.65, 0.75, 0.85, 0.90, 0.90, 0.90]
RAW_PCT = 0.76            # of revenue
PACK_PCT = 0.06           # of revenue
ADMIN_PCT = 0.05          # of revenue
POWER0, POWER_G = 22.0, 0.04
SAL0, SAL_G = 34.0, 0.05
REPAIRS = 2.2
INS0, INS_G = 1.5, 0.03
TAX_RATE = 0.25
DISC = 0.12

SCENARIOS = [
    ("S0 Base", 1.00, 0.76),
    ("S1 Revenue -5%", 0.95, 0.76),
    ("S2 Raw material +3pp", 1.00, 0.79),
    ("S3 Combined stress", 0.95, 0.79),
]
N_YEARS = 7


def loan_schedule():
    """Annual schedule. Lists indexed 0..7 (Y0..Y7)."""
    opening, interest, payment, principal, closing = [], [], [], [], []
    bal = LOAN
    factor = RATE * (1 + RATE) ** REPAY_YEARS / ((1 + RATE) ** REPAY_YEARS - 1)
    annuity = LOAN * factor
    for t in range(0, N_YEARS + 1):
        opening.append(bal)
        if t == 0:
            interest.append(0.0); payment.append(0.0); principal.append(0.0)
        elif t <= MORATORIUM:
            interest.append(bal * RATE); payment.append(bal * RATE)
            principal.append(0.0)
        else:
            i = bal * RATE
            interest.append(i); payment.append(annuity); principal.append(annuity - i)
        bal = opening[-1] - principal[-1]
        closing.append(bal)
    return {"opening": opening, "interest": interest, "payment": payment,
            "principal": principal, "closing": closing, "annuity": annuity}


def operating_year(t, rev_factor=1.0, raw_pct=RAW_PCT):
    """t is 1-based year."""
    prod_kg = CAP_KG * UTIL[t - 1]
    price = PRICE0 * (1 + PRICE_G) ** (t - 1)
    revenue = prod_kg * price / 1e5 * rev_factor
    raw = revenue * raw_pct
    packing = revenue * PACK_PCT
    power = POWER0 * (1 + POWER_G) ** (t - 1)
    salary = SAL0 * (1 + SAL_G) ** (t - 1)
    repairs = REPAIRS
    admin = revenue * ADMIN_PCT
    insurance = INS0 * (1 + INS_G) ** (t - 1)
    total_cost = raw + packing + power + salary + repairs + admin + insurance
    return {"prod_kg": prod_kg, "price": price, "revenue": revenue,
            "raw": raw, "packing": packing, "power": power, "salary": salary,
            "repairs": repairs, "admin": admin, "insurance": insurance,
            "total_cost": total_cost, "ebitda": revenue - total_cost}


def depr(t):
    return DEPR_BASE + (PRELIM_AMORT if t <= 5 else 0.0)


def pl_series(rev_factor=1.0, raw_pct=RAW_PCT):
    loan = loan_schedule()
    out = {k: [] for k in ("revenue", "total_cost", "ebitda", "depr", "interest",
                           "pbt", "tax", "pat", "cash_accruals")}
    for t in range(1, N_YEARS + 1):
        op = operating_year(t, rev_factor, raw_pct)
        d = depr(t)
        interest = loan["interest"][t]
        pbt = op["ebitda"] - d - interest
        tax = TAX_RATE * pbt if pbt > 0 else 0.0
        pat = pbt - tax
        out["revenue"].append(op["revenue"])
        out["total_cost"].append(op["total_cost"])
        out["ebitda"].append(op["ebitda"])
        out["depr"].append(d)
        out["interest"].append(interest)
        out["pbt"].append(pbt)
        out["tax"].append(tax)
        out["pat"].append(pat)
        out["cash_accruals"].append(pat + d)
    return out


def cashflow_series(rev_factor=1.0, raw_pct=RAW_PCT):
    pl = pl_series(rev_factor, raw_pct)
    loan = loan_schedule()
    principal = loan["principal"][1:]
    debt_service = [loan["interest"][t] + loan["principal"][t]
                    for t in range(1, N_YEARS + 1)]
    net_cash = [ca - p for ca, p in zip(pl["cash_accruals"], principal)]
    dscr = [ca / ds for ca, ds in zip(pl["cash_accruals"], debt_service)]
    equity_cf = [-EQUITY] + net_cash
    cum_eq, run = [], -EQUITY
    for nc in net_cash:
        run += nc
        cum_eq.append(run)
    payback = next((t for t, c in enumerate(cum_eq, start=1) if c >= 0), None)
    return {"debt_service": debt_service, "principal": principal,
            "net_cash": net_cash, "cum_eq": cum_eq, "dscr": dscr,
            "equity_cf": equity_cf, "payback": payback,
            "avg_dscr": sum(dscr) / len(dscr),
            "min_dscr_repay": min(dscr[1:]),
            "npv": npv_series(equity_cf), "irr": irr_bisect(equity_cf)}


def npv_series(cf):
    return sum(c / (1 + DISC) ** t for t, c in enumerate(cf))


def irr_bisect(cf, lo=-0.9, hi=5.0, iters=200):
    def npv_at(r):
        return sum(c / (1 + r) ** t for t, c in enumerate(cf))
    for _ in range(iters):
        mid = (lo + hi) / 2
        if npv_at(mid) > 0:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def balance_sheet():
    pl = pl_series()
    cf = cashflow_series()
    loan = loan_schedule()
    years = list(range(0, N_YEARS + 1))
    net_block = [BLOCK_GROSS - DEPR_BASE * t for t in years]
    prelim = [12.0 - PRELIM_AMORT * min(t, 5) for t in years]
    cash = [5.0] + [5.0 + c for c in cf["cum_eq"]]
    # NOTE: cum_eq starts at -EQUITY; cash Y0 must be exactly the WC margin 5.0,
    # so rebuild cash from net_cash instead:
    cash = [5.0]
    run = 5.0
    for nc in cf["net_cash"]:
        run += nc
        cash.append(run)
    assets = [nb + pr + ca for nb, pr, ca in zip(net_block, prelim, cash)]
    reserves, run = [0.0], 0.0
    for pat in pl["pat"]:
        run += pat
        reserves.append(run)
    liab = [EQUITY + r + l for r, l in zip(reserves, loan["closing"])]
    return {"net_block": net_block, "prelim": prelim, "cash": cash,
            "assets": assets, "reserves": reserves, "loan_os": loan["closing"],
            "liab": liab,
            "identity_gap": [a - l for a, l in zip(assets, liab)]}


def break_even_pct():
    op1 = operating_year(1)
    fixed = (op1["power"] + op1["salary"] + op1["repairs"] + op1["insurance"]
             + depr(1) + loan_schedule()["interest"][1])
    contrib = 1 - RAW_PCT - PACK_PCT - ADMIN_PCT
    full_cap_rev = CAP_KG * PRICE0 / 1e5
    return fixed / contrib / full_cap_rev


def scenario(name, rev_factor, raw_pct):
    cf = cashflow_series(rev_factor, raw_pct)
    return {"name": name, "rev_factor": rev_factor, "raw_pct": raw_pct,
            "min_dscr_repay": cf["min_dscr_repay"], "avg_dscr": cf["avg_dscr"],
            "npv": cf["npv"], "irr": cf["irr"], "net_cash": cf["net_cash"],
            "dscr": cf["dscr"]}


def ground_truth():
    loan = loan_schedule()
    pl = pl_series()
    cf = cashflow_series()
    bs = balance_sheet()
    return {
        "project_cost": PROJECT_COST,
        "equity": EQUITY, "loan": LOAN,
        "promoter_margin": EQUITY / PROJECT_COST,
        "debt_equity": LOAN / EQUITY,
        "annuity": loan["annuity"],
        "loan_closing_y7": loan["closing"][7],
        "loan_interest_y1": loan["interest"][1],
        "principal_y2": loan["principal"][2],
        "revenue": pl["revenue"], "ebitda": pl["ebitda"], "pat": pl["pat"],
        "cash_accruals": pl["cash_accruals"],
        "dscr": cf["dscr"], "avg_dscr": cf["avg_dscr"],
        "min_dscr_repay": cf["min_dscr_repay"],
        "npv": cf["npv"], "irr": cf["irr"], "payback": cf["payback"],
        "net_cash": cf["net_cash"], "cum_eq": cf["cum_eq"],
        "identity_gap": bs["identity_gap"],
        "cash_min": min(bs["cash"]),
        "break_even_pct": break_even_pct(),
        "scenarios": [scenario(n, f, r) for n, f, r in SCENARIOS],
    }


if __name__ == "__main__":
    gt = ground_truth()
    print(f'project_cost={gt["project_cost"]} margin={gt["promoter_margin"]:.0%} '
          f'D:E={gt["debt_equity"]:.2f} annuity={gt["annuity"]:.2f}')
    print(f'minDSCR(repay)={gt["min_dscr_repay"]:.2f} avgDSCR={gt["avg_dscr"]:.2f} '
          f'NPV={gt["npv"]:.1f} IRR={gt["irr"]:.1%} payback=Y{gt["payback"]} '
          f'BEP={gt["break_even_pct"]:.1%}')
    print(f'identity_gap max={max(abs(g) for g in gt["identity_gap"]):.2e} '
          f'cash_min={gt["cash_min"]:.1f}')
    for s in gt["scenarios"]:
        print(f'{s["name"]}: minDSCR={s["min_dscr_repay"]:.2f} '
              f'avgDSCR={s["avg_dscr"]:.2f} NPV={s["npv"]:.1f} '
              f'IRR={s["irr"]:.1%}')
