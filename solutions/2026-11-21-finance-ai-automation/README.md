# FinanceFlow AI — Finance Automation & Reporting

A working demo of the **Finance AI Automation** pattern: messy bank-statement
narrations go in, a classified, anomaly-flagged financial report comes out.
Built 2026-11-21 around a live Freelancer lead — *Finance AI Automation
Implementation Partner* (Claude AI finance automation: data extraction,
analysis, financial reporting; Excel in the client's stack).

Fictional data: **Kesarwani Traders, Patna** (wholesale FMCG), Apr–Oct 2026,
268 bank-statement rows, opening balance ₹85,000. Deterministic
(`seed 20261121`) — re-running the pipeline reproduces identical outputs.

## What it does

```
data/bank_statement.csv ──┐
data/chart_of_accounts.csv ┤
                           ▼
              pipeline.py  (stdlib only)
                · extract counterparty from messy narrations
                · classify each row against the chart of accounts
                  (ordered keyword rules; confidence 1.0 / 0.7 / 0.0)
                · flag anomalies: duplicates, stated-vs-recomputed
                  balance mismatches, unclassified lines, single
                  legs over ₹60,000
                           ▼
              out/*.csv + summary.json  (auditable, human-readable)
                           ▼
              generate.py  (openpyxl, no macros)
                           ▼
              financeflow-ai-report.xlsx
```

The workbook is **genuinely formula-driven** — every figure is a live
Excel formula, so re-classifying a row on `Raw` re-computes the whole
report:

| Sheet | Contents |
|---|---|
| Guide | What it is, how to refresh, confidence scale |
| Raw | 268 classified transactions; Net (`=Credit−Debit`) and Month (`=TEXT(date,"yyyy-mm")`) are formulas |
| Monthly | Income / expenses / net / txn count per month — all `SUMIFS`/`COUNTIFS` over Raw, plus TOTAL |
| Categories | Totals per chart-of-accounts category (`SUMIFS`), share % via `SUMPRODUCT(ABS(...))` |
| Dashboard | Focus-month dropdown driving 8 KPIs + 3 native charts (monthly net, category share, income vs expenses) |
| Checks | 12 integrity checks against the independent Python ground truth |

## Results (Apr–Oct 2026)

- **Income ₹25,41,314 · Expenses ₹22,85,506 · Net +₹2,55,809**
- Closing balance ₹3,40,809 (opening ₹85,000 + net — ties out exactly)
- Sales ₹25,26,113 (110 txns) · Purchases ₹19,70,583 (66 txns) — 22% gross margin
- **6 anomalies caught**: 1 exact duplicate, 1 stated-balance error
  (+₹500 on 2026-09-15), 3 unclassifiable narrations (flagged for human
  review, never guessed), 1 outsized purchase (₹78,400)

## The classifier

An ordered rule engine over the chart of accounts (`data/chart_of_accounts.csv`).
Specific categories match first; the generic `PAYMENT` fallback for Sales
scores 0.7; anything unmatched is `Unclassified` (0.0) and queued for a
human. In production the same interface fronts an LLM call (e.g. Claude)
for narrations no rule can classify — the rules handle the routine 99%,
the model handles the cryptic 1%.

## Run it

```bash
python3 make_data.py    # regenerate the fictional statement (seeded)
python3 pipeline.py     # classify + flag -> out/
python3 generate.py     # build financeflow-ai-report.xlsx
python3 -m unittest discover -s tests -v
```

## Tests — 22/22 pass

- Classification rules (order, confidence, never-guess unclassified)
- Counterparty extraction from UPI/NEFT/IMPS narrations
- Anomaly detection (exact counts + the +₹500 balance diff)
- Monthly/category CSVs reconcile with the classified rows
- Workbook formulas present (no hard-coded money in Monthly)
- **LibreOffice headless recalc** of the generated workbook vs the
  independent Python ground truth: monthly totals, category totals,
  dashboard KPIs and all 12 Checks PASS, zero formula errors

## Files

```
make_data.py      fictional statement generator (seed 20261121)
pipeline.py       extraction + classification + anomaly flags
generate.py       formula-driven workbook builder
tests/test_build.py
data/             bank_statement.csv, chart_of_accounts.csv
out/              classified_transactions.csv, monthly_summary.csv,
                  category_summary.csv, anomalies.csv, summary.json
financeflow-ai-report.xlsx
```

Built as a daily demo for the DashboardForge wedge — no client data,
no secrets, no live credentials anywhere in this folder.
