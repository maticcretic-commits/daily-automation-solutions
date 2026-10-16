#!/usr/bin/env python3
"""build_db.py — Load the QuickBooks-style CSVs into a relational SQLite model.

Schema (the "relational DB" in the QBO -> DB -> BI pipeline):
    accounts(acct_no PK, acct_name, acct_type)
    journal_lines(id PK, date, journal_no, line_no, acct_no FK, debit, credit, memo, party)

Views:
    v_trial_balance  — per-account debit/credit/net totals (net = debit - credit)
    v_monthly_pl     — per-month revenue, cogs, expenses, gross & net profit
    v_balance_sheet  — asset/liability/equity totals as of a period end
"""
import csv
import sqlite3
from pathlib import Path

HERE = Path(__file__).parent
DB = HERE / "qb2bi.sqlite"

SCHEMA = """
DROP TABLE IF EXISTS journal_lines;
DROP TABLE IF EXISTS accounts;
CREATE TABLE accounts (
    acct_no   TEXT PRIMARY KEY,
    acct_name TEXT NOT NULL,
    acct_type TEXT NOT NULL CHECK (acct_type IN ('Asset','Liability','Equity','Income','Expense'))
);
CREATE TABLE journal_lines (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    date       TEXT NOT NULL,            -- ISO yyyy-mm-dd
    journal_no TEXT NOT NULL,
    line_no    INTEGER NOT NULL,
    acct_no    TEXT NOT NULL REFERENCES accounts(acct_no),
    debit      REAL NOT NULL DEFAULT 0,
    credit     REAL NOT NULL DEFAULT 0,
    memo       TEXT,
    party      TEXT
);
CREATE INDEX ix_jl_date ON journal_lines(date);
CREATE INDEX ix_jl_acct ON journal_lines(acct_no);
CREATE VIEW v_trial_balance AS
SELECT a.acct_no, a.acct_name, a.acct_type,
       ROUND(SUM(j.debit), 2)  AS tot_debit,
       ROUND(SUM(j.credit), 2) AS tot_credit,
       ROUND(SUM(j.debit - j.credit), 2) AS net
FROM accounts a LEFT JOIN journal_lines j ON j.acct_no = a.acct_no
GROUP BY a.acct_no, a.acct_name, a.acct_type
ORDER BY a.acct_no;
"""


def load():
    if DB.exists():
        DB.unlink()
    con = sqlite3.connect(DB)
    con.executescript(SCHEMA)
    with open(HERE / "data" / "chart_of_accounts.csv", newline="") as f:
        rows = list(csv.DictReader(f))
        con.executemany(
            "INSERT INTO accounts VALUES (?,?,?)",
            [(r["acct_no"], r["acct_name"], r["acct_type"]) for r in rows])
    with open(HERE / "data" / "journals.csv", newline="") as f:
        rows = list(csv.DictReader(f))
        con.executemany(
            """INSERT INTO journal_lines
               (date, journal_no, line_no, acct_no, debit, credit, memo, party)
               VALUES (?,?,?,?,?,?,?,?)""",
            [(r["date"], r["journal_no"], int(r["line_no"]), r["acct_no"],
              float(r["debit"]), float(r["credit"]), r["memo"], r["party"])
             for r in rows])
    con.commit()
    n_acct = con.execute("SELECT COUNT(*) FROM accounts").fetchone()[0]
    n_lines = con.execute("SELECT COUNT(*) FROM journal_lines").fetchone()[0]
    dr = con.execute("SELECT ROUND(SUM(debit),2) FROM journal_lines").fetchone()[0]
    cr = con.execute("SELECT ROUND(SUM(credit),2) FROM journal_lines").fetchone()[0]
    assert abs(dr - cr) < 0.01, f"ledger out of balance: {dr} vs {cr}"
    print(f"db OK: {n_acct} accounts, {n_lines} lines, dr={dr:,.2f} cr={cr:,.2f}")
    con.close()


if __name__ == "__main__":
    load()
