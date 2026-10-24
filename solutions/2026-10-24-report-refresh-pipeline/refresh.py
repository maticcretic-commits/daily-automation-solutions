#!/usr/bin/env python3
"""RefreshForge — the refresh pipeline.

This is the local, testable mirror of the scheduled flow the client asked
for ("Power Automate + Excel + Power BI dashboard auto-refresh"). Each run:

  1. reads every CSV in drop/ (the "watch folder"),
  2. validates each row against the source schema (bad rows are quarantined,
     counted and reported — never appended),
  3. dedupes on the source's natural key (within the file and against
     history), so re-running the same files appends ZERO rows,
  4. appends only genuinely new rows to history/*.csv (the cumulative store),
  5. records the run in refresh_log.json (audit trail: per-source
     seen/new/dup/quarantined, file checksums, duration, status).

Then build_workbook.py rebuilds refresh-report.xlsx from history/ — every
computed cell is a live formula, so the dashboard "auto-refreshes" the same
way the client's Power Automate flow would rebuild it.

Usage: python3 refresh.py [--dir DIR]   (DIR defaults to this folder)
"""
import argparse
import csv
import hashlib
import json
import os
import sys
import time
from datetime import date

SOURCES = {
    "sales": {
        "file": "sales.csv",
        "key": "order_id",
        "required": ["order_id", "date", "customer", "sku", "qty",
                     "unit_price"],
        "numeric": {"qty": (1, None), "unit_price": (0.01, None)},
    },
    "production": {
        "file": "production.csv",
        "key": "batch_id",
        "required": ["batch_id", "date", "line", "sku", "qty_produced",
                     "qty_rejected"],
        "numeric": {"qty_produced": (0, None), "qty_rejected": (0, None)},
    },
    "inventory": {
        "file": "inventory.csv",
        "key": "txn_id",
        "required": ["txn_id", "date", "sku", "movement", "qty"],
        "numeric": {"qty": (1, None)},
        "choices": {"movement": {"IN", "OUT"}},
    },
    "issues": {
        "file": "issues.csv",
        "key": "issue_id",
        "required": ["issue_id", "date", "category", "severity", "status",
                     "description"],
        "choices": {"severity": {"P1", "P2", "P3"},
                    "status": {"Open", "Closed"}},
    },
}


def _valid_date(v):
    try:
        y, m, d = (int(x) for x in v.split("-"))
        date(y, m, d)
        return True
    except Exception:
        return False


def validate_row(cfg, row, header):
    """Return None if valid, else a short reason string."""
    for col in cfg["required"]:
        if col not in header or row.get(col, "").strip() == "":
            return f"missing/empty {col}"
    if not _valid_date(row["date"].strip()):
        return f"bad date {row['date']!r}"
    for col, (lo, hi) in cfg.get("numeric", {}).items():
        try:
            v = float(row[col])
        except ValueError:
            return f"non-numeric {col}={row[col]!r}"
        if v < lo or (hi is not None and v > hi):
            return f"{col}={v} out of range"
    for col, allowed in cfg.get("choices", {}).items():
        if row[col].strip() not in allowed:
            return f"{col}={row[col]!r} not in {sorted(allowed)}"
    return None


def read_history(hist_path, key):
    keys, rows = set(), []
    if os.path.exists(hist_path):
        with open(hist_path, newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                keys.add(r[key])
                rows.append(r)
    return keys, rows


def run_refresh(base_dir):
    base_dir = os.path.abspath(base_dir)
    drop = os.path.join(base_dir, "drop")
    hist_dir = os.path.join(base_dir, "history")
    os.makedirs(hist_dir, exist_ok=True)
    log_path = os.path.join(base_dir, "refresh_log.json")

    if os.path.exists(log_path):
        with open(log_path) as f:
            log = json.load(f)
    else:
        log = {"runs": []}
    run_id = f"R{len(log['runs']) + 1:04d}"
    started = time.time()
    stamp = time.strftime("%Y-%m-%dT%H:%M:%S")

    entry = {"run_id": run_id, "timestamp": stamp, "sources": {},
             "status": "OK"}
    for name, cfg in SOURCES.items():
        src = os.path.join(drop, cfg["file"])
        hist_path = os.path.join(hist_dir, cfg["file"])
        stat = {"file": cfg["file"], "seen": 0, "new": 0,
                "duplicates": 0, "quarantined": 0, "errors": []}
        if not os.path.exists(src):
            stat["status"] = "missing file"
            entry["sources"][name] = stat
            continue
        with open(src, "rb") as f:
            stat["sha1"] = hashlib.sha1(f.read()).hexdigest()
        with open(src, newline="") as f:
            rdr = csv.DictReader(f)
            header = rdr.fieldnames or []
            rows = list(rdr)
        stat["seen"] = len(rows)
        existing_keys, _ = read_history(hist_path, cfg["key"])
        seen_keys = set()
        new_rows = []
        for i, row in enumerate(rows, start=2):
            reason = validate_row(cfg, row, header)
            if reason:
                stat["quarantined"] += 1
                if len(stat["errors"]) < 3:
                    stat["errors"].append(f"row {i}: {reason}")
                continue
            key = row[cfg["key"]].strip()
            if key in seen_keys or key in existing_keys:
                stat["duplicates"] += 1
                continue
            seen_keys.add(key)
            new_rows.append(row)
        if new_rows:
            write_header = not os.path.exists(hist_path)
            with open(hist_path, "a", newline="") as f:
                w = csv.DictWriter(f, fieldnames=header)
                if write_header:
                    w.writeheader()
                w.writerows(new_rows)
        stat["new"] = len(new_rows)
        if stat["quarantined"]:
            entry["status"] = "OK_WITH_QUARANTINE"
        entry["sources"][name] = stat

    entry["duration_s"] = round(time.time() - started, 2)
    log["runs"].append(entry)
    with open(log_path, "w") as f:
        json.dump(log, f, indent=2)
    return entry


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=os.path.dirname(os.path.abspath(__file__)))
    args = ap.parse_args()
    entry = run_refresh(args.dir)
    total_new = sum(s["new"] for s in entry["sources"].values())
    total_q = sum(s["quarantined"] for s in entry["sources"].values())
    print(f"{entry['run_id']} {entry['timestamp']} status={entry['status']} "
          f"new_rows={total_new} quarantined={total_q} "
          f"({entry['duration_s']}s)")
    for name, s in entry["sources"].items():
        print(f"  {name:11s} seen={s['seen']:4d} new={s['new']:4d} "
              f"dup={s['duplicates']:3d} quar={s['quarantined']:3d}")
        for e in s["errors"]:
            print(f"    ! {e}")


if __name__ == "__main__":
    main()
