# RefreshForge — automated report-refresh pipeline

A working demo built for a real freelance brief: **"Automate Microsoft Data
Workflows"** — a client wants Power Automate + Excel + Power BI so their
dashboard auto-refreshes on a schedule with zero manual intervention.

This repo is the local, testable mirror of that scheduled flow. Four
departmental CSV exports land in a watch folder; `refresh.py` validates,
dedupes and appends only new rows to a cumulative store, writes an audit
log, and `build_workbook.py` rebuilds a fully formula-driven Excel workbook —
the same shape as the client's Power Automate flow, verifiable end to end.

All data is fictional and deterministic (seed `20261024`): "Sharma Foods
Processing, Patna", 01–23 Oct 2026.

## The pipeline

```
drop/*.csv ──validate──▶ quarantine (counted, never appended)
      │
      ├──dedupe on natural key (order_id / batch_id / txn_id / issue_id)
      │
      ▼
history/*.csv ◀── append ONLY genuinely new rows (idempotent: re-running the
      │           same files appends ZERO rows)
      │
      ▼
refresh_log.json ── audit: run id, timestamp, per-source seen/new/dup/
                    quarantined, file checksums, duration, status
      │
      ▼
refresh-report.xlsx ── rebuilt from history/; every computed cell is a live
                       formula, so KPIs/charts/checks "auto-refresh"
```

## Run it

```bash
python3 generate.py        # (re)create the synthetic drop/ files
python3 refresh.py         # validate + dedupe + append + audit log
python3 build_workbook.py  # rebuild refresh-report.xlsx
python3 -m pytest tests/ -q
```

Try the refresh loop yourself:

```bash
python3 refresh.py            # second run: 0 new rows (idempotent)
# add a new daily export row to drop/sales.csv, then:
python3 refresh.py            # appends exactly the delta
python3 build_workbook.py     # dashboard recalculates
```

## The workbook (9 sheets)

- **Guide** — pipeline docs + the exact Power Automate flow mapping
  (Recurrence trigger → list files → parse → validate → dedupe → append →
  refresh Power BI dataset).
- **Refresh_Log** — every pipeline run, one row per source (the audit).
- **Raw_Sales / Raw_Production / Raw_Inventory / Raw_Issues** — cumulative
  history; derived columns (revenue, yield, signed qty) are live formulas.
- **Daily** — per-day SUMIFS/COUNTIFS summary; conditional formatting flags
  issue spikes (>2/day) and yield dips (<95%) — 14-Oct's breakdown day lights
  up red.
- **Dashboard** — 9 formula-linked KPI cards (MTD revenue, production, avg
  yield, open issues, inventory on hand, history rows, last refresh run /
  status / rows added) + 3 native charts (daily line, issues pie, line bar).
- **Checks** — 11 integrity checks (no duplicate keys, formula columns
  consistent, daily ties to raw, inventory KPI = opening + movements, every
  refresh run OK). All must read TRUE.

## Verification

17/17 tests pass, including a headless LibreOffice recalculation of the
rebuilt workbook against an independent Python ground truth: all 6 numeric
KPIs match exactly, zero formula errors, all 11 checks TRUE. Plain A1-style
ranges throughout (no Table objects, no structured references), so the
workbook recalculates identically in Excel, LibreOffice and Google Sheets.

`powerbi/MEASURES.md` holds the DAX measures and the report-page mapping for
the Power BI Service version of this dashboard.
