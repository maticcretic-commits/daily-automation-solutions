"""CollectionPulse data generator — fictional micro-loan collection portfolio.

Produces deterministic (seeded) CSVs:
  data/schedule.csv — one row per scheduled EMI due (200 accounts x 6 months)
  data/payments.csv — one row per actual receipt (full, late, partial, or missing)

All names/places are fictional. 'As of' date for DPD maths: 2026-10-20.
"""
import csv
import random
from datetime import date, timedelta

SEED = 20261020
AS_OF = date(2026, 10, 20)
FIRST_NAMES = ["Ramesh", "Suresh", "Amit", "Vikash", "Prakash", "Manoj", "Rajesh",
               "Sunita", "Pooja", "Anita", "Rina", "Kiran", "Deepak", "Sanjay",
               "Mohan", "Geeta", "Ravi", "Naresh", "Umesh", "Bablu"]
LAST_NAMES = ["Prasad", "Kumar", "Singh", "Yadav", "Sharma", "Verma", "Gupta",
              "Mandal", "Sah", "Paswan", "Chaudhary", "Rai", "Tiwari", "Mishra"]
BRANCHES = ["Patna", "Gaya", "Muzaffarpur", "Bhagalpur", "Darbhanga"]
COLLECTORS = ["A. Verma", "S. Gupta", "R. Singh", "P. Yadav", "K. Sharma", "M. Mandal"]
LOAN_TYPES = ["Agri", "Trade", "Vehicle", "Personal", "Housing"]
DUE_MONTHS = [date(2026, m, 5) for m in (4, 5, 6, 7, 8, 9)]


def payment_plan(rng, archetype, due_date, emi):
    """Return list of (pay_date, amount) for one due."""
    if archetype == "good":
        delay = rng.randint(-3, 3)
        return [(due_date + timedelta(days=delay), emi)]
    if archetype == "slow":
        delay = rng.randint(10, 40)
        return [(due_date + timedelta(days=delay), emi)]
    if archetype == "erratic":
        r = rng.random()
        if r < 0.35:
            delay = rng.randint(0, 10)
            return [(due_date + timedelta(days=delay), emi)]
        if r < 0.70:
            first = round(emi * rng.uniform(0.4, 0.7))
            return [(due_date + timedelta(days=rng.randint(5, 20)), first),
                    (due_date + timedelta(days=rng.randint(25, 55)), emi - first)]
        delay = rng.randint(35, 75)
        return [(due_date + timedelta(days=delay), emi)]
    return []  # defaulter: no payments at all for dues past its cutoff


def main():
    rng = random.Random(SEED)
    sched_rows, pay_rows = [], []
    pay_seq = 0
    for a in range(200):
        acct = f"A{a+1:04d}"
        borrower = f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}"
        branch = rng.choice(BRANCHES)
        collector = rng.choice(COLLECTORS)
        loan_type = rng.choice(LOAN_TYPES)
        amount = rng.randrange(20, 201) * 1000
        emi = round(amount / 6)
        arch = rng.choices(["good", "slow", "erratic", "defaulter"],
                           weights=[55, 25, 12, 8])[0]
        cutoff = rng.randint(2, 4) if arch == "defaulter" else 6
        for i, due_date in enumerate(DUE_MONTHS):
            due_id = f"D{a*6+i+1:04d}"
            sched_rows.append([due_id, acct, borrower, branch, collector,
                               loan_type, due_date.isoformat(), emi])
            if i >= cutoff:
                continue
            for pay_date, amt in payment_plan(rng, arch, due_date, emi):
                if pay_date > AS_OF:  # not yet happened
                    continue
                pay_seq += 1
                pay_rows.append([f"P{pay_seq:05d}", due_id, acct,
                                 pay_date.isoformat(), amt,
                                 rng.choice(["UPI", "Cash", "NEFT", "Cash", "UPI"])])
    with open("data/schedule.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["DueID", "Account", "Borrower", "Branch", "Collector",
                    "LoanType", "DueDate", "DueAmount"])
        w.writerows(sched_rows)
    with open("data/payments.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["PaymentID", "DueID", "Account", "PayDate", "Amount", "Mode"])
        w.writerows(pay_rows)
    print(f"dues={len(sched_rows)} payments={len(pay_rows)}")


if __name__ == "__main__":
    main()
