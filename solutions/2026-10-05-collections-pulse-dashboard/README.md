# CollectionsPulse — NBFC Collections Dashboard + Follow-up Queue Factory

**Built:** 2026-10-05 · **Stack:** Python 3 (standard library only — no dependencies)
**Demo for:** finance/NBFC-adjacent reporting gigs — *"Hospitality Reporting Automation
Build"* (daily automated reporting from raw operational data), *"Full Financial
Statement Preparation"* (Excel + Power BI financial reporting) style work: turn a raw
loan ledger into a client-ready collections command center.

## The problem

Lending teams (NBFCs, microfinance, co-operative banks) track EMIs in raw ledger
exports — loan id, customer, branch, agent, due date, EMI, paid — but every morning the
collections manager still asks the same questions by hand: *How much is outstanding in
each DPD bucket? Which branch is collecting worst? Who do we call TODAY?* Hand-built
spreadsheets are slow, error-prone, and never rank the follow-up queue.

## The solution

`collections_pulse.py` reads any loan-ledger CSV with the columns
`LoanID, Customer, Phone, Branch, Agent, DueDate (YYYY-MM-DD), EMI, Paid` and
generates a real `.xlsx` workbook (genuine Office Open XML, written with the stdlib
`zipfile` module — no `openpyxl`, no pandas, no Excel required to build it) with:

| Sheet | Contents |
|---|---|
| **Raw Ledger** | Every loan + derived **Due / DPD / Bucket** columns · styled headers · AutoFilter · frozen top row · ₹ currency and date formats |
| **DPD Summary** | Pivot-style arrears table (Current / 1–30 / 31–60 / 61–90 / 90+) from **live Excel `SUMIF`/`COUNTIF`/`AVERAGEIF` formulas** over the raw rows · TOTAL row · red conditional formatting on the 90+ bucket |
| **Branch Summary** | Per-branch accounts, EMI due, collected, outstanding and **collection-efficiency %** (SUMIF formulas, division-by-zero guarded) · green/red conditional formatting on best/worst efficiency |
| **Dashboard** | Merged title · 4 KPI cells wired by formula to the summary totals (Total Outstanding, Accounts in Arrears, Collection Efficiency, **PAR>30**) · **3 native Excel charts** — column (outstanding by DPD bucket), pie (outstanding share), line (branch efficiency) |
| **Action Queue** | Today's priority-ranked follow-up list — the N accounts with the highest priority score `Due × (1 + DPD/30)`, with a formula priority column and a ready-to-send **bilingual (Hinglish/English) reminder template** per account |

It also writes `followup_reminders.csv` — the collections team's call list for the day:
LoanID, Customer, Phone, Branch, Agent, Due, DPD and a pre-filled WhatsApp/SMS
reminder message ("Namaste …, aapki EMI ₹… … din se pending hai…").

Because the summaries use live Excel formulas (not hard-coded numbers), the workbook
**recalculates automatically** if the team edits the raw ledger — and no VBA/macros
are needed, so there are no macro-security warnings on open.

## How to run

```bash
# 1. Generate the deterministic sample ledger (180 loans, 4 branches)
python collections_pulse.py --generate-sample sample_ledger_2026.csv --ref-date 2026-10-05

# 2. Build the dashboard workbook + today's follow-up list from any ledger CSV
python collections_pulse.py --input sample_ledger_2026.csv --output CollectionsPulse.xlsx \
    --reminders followup_reminders.csv --ref-date 2026-10-05 --top 50
# loans: 180 | in arrears: 95 | total outstanding: ₹582000
# dashboard written: CollectionsPulse.xlsx
# reminders written: followup_reminders.csv (50)

# 3. Open CollectionsPulse.xlsx in Excel / LibreOffice / Google Sheets
```

`--ref-date` fixes "today" for DPD math (defaults to the real today);
`--top N` controls how many accounts land in the Action Queue / reminders CSV.
Bad dates, non-numeric amounts, negative values, or Paid > EMI fail fast with a
clear error naming the row.

## Tests

```bash
python test_collections_pulse.py   # 30 tests, all passing
```

Coverage: DPD bucket boundaries (0/30/31/60/61/90/91), priority-score ordering,
reminder text contents, input validation (missing columns, bad date/number with row
numbers, negatives, Paid > EMI, empty file), branch ordering, workbook structure
(18 OOXML parts, 5 sheet names, 3 charts, drawing on the Dashboard), formula ranges
cover every raw row, conditional-formatting targets, queue ranking + priority formula,
reminder CSV contents, deterministic rebuilds (byte-identical), deterministic sample
generator.

**Independent formula verification:** every summary cell's XML was re-evaluated with
an independent parser against the raw sheet's values — all 28 checks (bucket
accounts/outstanding/avg-DPD, branch efficiency/outstanding, and all 4 dashboard KPIs)
matched the Python computation exactly (e.g. total outstanding ₹582,000, collection
efficiency 56.2%, PAR>30 45.9%).

## Files

- `collections_pulse.py` — the factory: CSV validation, DPD/priority engine, OOXML workbook builder, charts, reminder CSV writer, CLI
- `test_collections_pulse.py` — 30-test suite (stdlib `unittest`)
- `sample_ledger_2026.csv` — deterministic 180-loan sample ledger (seed `20261005`)
- `README.md` — this file

## Customizing for a client

- Change brand colors: edit the `fills`/`fonts` in `styles_xml()` (header blue is `2E75B6`).
- Change buckets or the priority formula: edit `BUCKETS` / `bucket_of()` / `priority_score()` — the summary sheets pick up new buckets automatically via the shared range pattern.
- English-only reminders: edit `reminder_message()`; the queue sheet and CSV both use it.
