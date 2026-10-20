"""Assemble the CollectionPulse workbook."""
import csv
from datetime import date
from build_workbook import build
from build_part2 import build_followup, build_dashboard, build_ai_report, build_checks

AS_OF = date(2026, 10, 20)


def ground_truth():
    sched = list(csv.DictReader(open("data/schedule.csv")))
    pays = list(csv.DictReader(open("data/payments.csv")))
    paid, lastpay = {}, {}
    for p in pays:
        paid[p["DueID"]] = paid.get(p["DueID"], 0) + int(p["Amount"])
        if p["PayDate"] > lastpay.get(p["DueID"], ""):
            lastpay[p["DueID"]] = p["PayDate"]
    gt = []
    for s in sched:
        due = int(s["DueAmount"])
        pd = paid.get(s["DueID"], 0)
        out = due - pd
        status = "PAID" if out <= 0.005 else ("PARTIAL" if pd > 0 else "OVERDUE")
        if status == "PAID":
            dpd = max(0, (date.fromisoformat(lastpay[s["DueID"]])
                           - date.fromisoformat(s["DueDate"])).days)
        else:
            dpd = (AS_OF - date.fromisoformat(s["DueDate"])).days
        bucket = ("COLLECTED" if out <= 0.005 else "1-30" if dpd <= 30
                  else "31-60" if dpd <= 60 else "61-90" if dpd <= 90 else "90+")
        gt.append({"due": s["DueID"], "due_amt": due, "paid": pd, "out": out,
                   "status": status, "dpd": dpd, "bucket": bucket,
                   "collector": s["Collector"]})
    return gt


def main():
    gt = ground_truth()
    overdue = sum(1 for g in gt if g["out"] > 0.005)
    wb, sched, pays, ns, led_last, collectors = build()
    fu_last = build_followup(wb, led_last, overdue)
    build_dashboard(wb, led_last, collectors)
    build_ai_report(wb, led_last, fu_last)
    n_checks = build_checks(wb, led_last, fu_last)
    wb.save("payment-collections-tracker.xlsx")
    print(f"saved: dues={ns} overdue={overdue} followup_rows={fu_last-1} checks={n_checks}")


if __name__ == "__main__":
    main()
