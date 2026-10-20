# CollectionPulse — Payment Collections Tracker

A formula-driven Excel workbook for tracking loan/EMI collections: per-due
receivables ledger, DPD aging, collector scorecards, an auto-built follow-up
action list, KPI dashboard with native charts, and a rule-based portfolio
narrative. Built as a demo around a typical client ask: *"a client-payment
Excel tracker with a summary dashboard (monthly totals, % collected)"*.

All data is **fictional** (deterministic seeded generator) — replace the two
raw sheets with real exports to go live.

## What's inside

| Sheet | Contents |
|---|---|
| `Guide` | As-of date (2026-10-20, anchors all DPD maths), sheet map, refresh steps |
| `Raw_Schedule` | 1,200 scheduled dues: DueID, account, borrower, branch, collector, loan type, due date, amount |
| `Raw_Payments` | ~1,078 receipts: PaymentID, DueID, account, pay date, amount, mode |
| `Receivables_Ledger` | Per-due ledger — paid, outstanding, status (PAID/PARTIAL/OVERDUE), last pay date, days overdue, DPD bucket, severity. **All formulas** |
| `Aging` | Outstanding by bucket: COLLECTED / 1–30 / 31–60 / 61–90 / 90+ DPD |
| `Collector_Scorecard` | Per-collector dues, efficiency %, overdue load, avg DPD, rank |
| `FollowUp` | Auto-built action list: every overdue due with severity (MILD/WATCH/ESCALATE/LOST) and next-action date |
| `Dashboard` | 8 KPIs + 3 native charts (monthly efficiency, aging, collector leaderboard) |
| `AI_Report` | Rule-based narrative summary + 8 status flags |
| `Checks` | 8 integrity checks — all must read TRUE |

No values are pasted into computed columns: the ledger pulls schedule fields
via `INDEX/MATCH`, payments aggregate via `SUMIFS`, DPD uses the Guide as-of
date, and the FollowUp list filters the ledger with `SMALL/INDEX`. Paste new
dues/receipts, extend the ledger formulas down, and everything recalculates.

## Reproduce

```bash
python3 generate.py        # data/schedule.csv + data/payments.csv (seed 20261020)
python3 main.py            # payment-collections-tracker.xlsx
python3 -m pytest tests/ -q
```

Tests recalculate the shipped workbook headlessly with LibreOffice and compare
every KPI, aging bucket, scorecard row, all 1,200 ledger rows (paid,
outstanding, status, DPD, bucket), last-pay-date spot checks, and the full
follow-up list against an independent Python ground truth computed straight
from the CSVs. **13/13 pass.**

Demo state (as of 2026-10-20): ₹2,16,16,998 scheduled · ₹1,87,51,276 collected
(86.7%) · 168 overdue dues · ₹28,65,722 outstanding · avg DPD 116 days.
