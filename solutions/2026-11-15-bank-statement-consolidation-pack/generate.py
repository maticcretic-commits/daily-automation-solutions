#!/usr/bin/env python3
"""Build bank-statement-consolidation-pack.xlsx from data/statements/*.csv.

Demo for the Freelancer lead "Bank Statement Excel Consolidation":
merge a stack of bank statements into ONE clean unified worksheet,
every transaction exactly as in the source, balances to the cent.

Sheets: Guide, Sources, Transactions, Monthly, Dashboard, Checks.
All money cells are live formulas (SUMIFS/COUNTIFS/INDEX-MATCH); the
only values are the source CSV figures themselves. Running balances use
a per-account sequence number so the merged, date-sorted ledger still
reconciles to each source statement's closing balance.
Also writes data/consolidated.csv — the merged ledger as a flat export.
"""
import csv
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent
DATA = BASE / "data" / "statements"
XLSX = BASE / "bank-statement-consolidation-pack.xlsx"

ACCOUNTS = [
    ("Canara Bank", "Current", "****4821", "canara.csv"),
    ("State Bank of India", "Savings", "****7734", "sbi.csv"),
    ("HDFC Bank", "Current", "****9156", "hdfc.csv"),
]
MONTHS = [f"2026-{m:02d}" for m in range(1, 11)]
MONEY = '#,##0.00'
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
SUB_FONT = Font(size=10, italic=True, color="595959")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")


def categorise(desc):
    if desc.startswith("UPI/NEFT from") or desc.startswith("Cash deposit"):
        return "Sales Receipt"
    if desc.startswith("NEFT to"):
        return "Supplier Payment"
    if desc.startswith("Godown rent"):
        return "Rent"
    if desc.startswith("Salary"):
        return "Salary"
    if desc.startswith("GST payment"):
        return "GST & Tax"
    if desc.startswith(("Electricity bill", "Mobile/internet", "Fuel",
                        "Packing material")):
        return "Utilities & Operations"
    if desc.startswith(("SMS charges", "Cheque book", "NEFT charges",
                        "Cash handling")):
        return "Bank Charges"
    if desc.startswith("Interest credited"):
        return "Interest Income"
    if desc.startswith("Own account transfer"):
        return "Transfer"
    if desc.startswith("Refund from supplier"):
        return "Supplier Refund"
    if desc.startswith("Scrap sale"):
        return "Other Income"
    return "Uncategorised"


def load_statements():
    """Return (accounts, rows): accounts = per-bank dicts, rows = merged
    transactions sorted by (date, bank, description) with per-account Seq."""
    accounts = []
    rows = []
    for (bank, typ, acct, fname) in ACCOUNTS:
        with open(DATA / fname, newline="") as f:
            r = list(csv.DictReader(f))
        credits = sum(float(x["Credit"] or 0) for x in r)
        debits = sum(float(x["Debit"] or 0) for x in r)
        opening = round(float(r[0]["Balance"]) - float(r[0]["Credit"] or 0)
                        + float(r[0]["Debit"] or 0), 2)
        accounts.append({"bank": bank, "type": typ, "acct": acct,
                         "file": fname, "n": len(r), "opening": opening,
                         "credits": credits, "debits": debits,
                         "closing": float(r[-1]["Balance"])})
        for x in r:
            rows.append({"date": x["Date"], "bank": bank, "acct": acct,
                         "desc": x["Description"],
                         "debit": float(x["Debit"] or 0),
                         "credit": float(x["Credit"] or 0),
                         "src_bal": float(x["Balance"])})
    rows.sort(key=lambda t: (t["date"], t["bank"], t["desc"]))
    seq = {a["bank"]: 0 for a in accounts}
    for t in rows:
        seq[t["bank"]] += 1
        t["seq"] = seq[t["bank"]]
        t["cat"] = categorise(t["desc"])
        t["month"] = t["date"][:7]
    return accounts, rows


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    ws.row_dimensions[row].height = 30


def build():
    accounts, rows = load_statements()
    n = len(rows)
    first, last = 5, 4 + n
    wb = openpyxl.Workbook()

    # ---------- Guide ----------
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "Bank Statement Consolidation Pack"
    ws["A1"].font = Font(bold=True, size=16, color="1F4E79")
    ws["A2"] = ("Demo for the Freelancer project “Bank Statement Excel "
                "Consolidation”: merge a stack of bank statements into one "
                "clean unified worksheet — every transaction exactly as in "
                "the source, balances to the cent.")
    ws["A2"].font = SUB_FONT
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:F2")
    ws.row_dimensions[2].height = 30
    guide = [
        ("Business (fictional)", "Sharma Traders, Patna — wholesale FMCG "
         "trader, Jan–Oct 2026. Three accounts: Canara Bank current "
         "****4821, SBI savings ****7734, HDFC Bank current ****9156."),
        ("Sources", "One sheet per source statement is NOT used — the "
         "register lists each file, its rows, opening balance and totals; "
         "the raw statements live in data/statements/*.csv."),
        ("Transactions", "The unified ledger: all 158 transactions, "
         "date-sorted, with category, per-account sequence and a live "
         "running balance per account. Use the filters to slice by bank, "
         "category or month."),
        ("Monthly", "Opening / credits / debits / closing per account per "
         "month, all SUMIFS — the October closing ties to Sources."),
        ("Dashboard", "Pick a focus account and month (B3/B4) — the KPIs, "
         "category table and charts follow. Defaults show everything."),
        ("Checks", "10 live checks: row counts, credit/debit totals, final "
         "running balance vs each source closing, transfers netting to "
         "zero, duplicates, uncategorised rows, monthly chain. All must "
         "read PASS."),
        ("Transfers", "Three own-account transfers appear in two statements "
         "each; they are categorised “Transfer” and excluded from income "
         "and expense totals in the category table."),
        ("Refresh", "Re-run make_data.py (regenerates the fictional "
         "statements) then generate.py (rebuilds this workbook)."),
        ("Note", "Fictional data for demonstration only — no real account "
         "numbers, no real transactions."),
    ]
    r = 4
    for (k, v) in guide:
        ws.cell(row=r, column=1, value=k).font = Font(bold=True, size=11)
        c = ws.cell(row=r, column=2, value=v)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 30
        r += 1
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 110

    # ---------- Sources ----------
    src = wb.create_sheet("Sources")
    src["A1"] = "Source statements — register"
    src["A1"].font = TITLE_FONT
    src["A2"] = ("Each row is one source file. Credits, debits and closing "
                 "are live SUMIFS/COUNTIFS over Transactions; “Match” "
                 "compares the computed closing with the source file's own "
                 "closing balance.")
    src["A2"].font = SUB_FONT
    src.merge_cells("A2:K2")
    headers = ["Bank", "Type", "Account", "Source file", "Rows", "Opening",
               "Credits", "Debits", "Closing (calc)", "Source closing",
               "Match"]
    for j, h in enumerate(headers, start=1):
        src.cell(row=4, column=j, value=h)
    style_header(src, 4, len(headers))
    for i, a in enumerate(accounts):
        row = 5 + i
        src.cell(row=row, column=1, value=a["bank"])
        src.cell(row=row, column=2, value=a["type"])
        src.cell(row=row, column=3, value=a["acct"])
        src.cell(row=row, column=4, value=a["file"])
        src.cell(row=row, column=5,
                value=f"=COUNTIFS(Transactions!$B$5:$B${last},$A{row})")
        src.cell(row=row, column=6, value=a["opening"]).number_format = MONEY
        src.cell(row=row, column=7,
                value=f"=SUMIFS(Transactions!$G$5:$G${last},"
                      f"Transactions!$B$5:$B${last},$A{row})").number_format = MONEY
        src.cell(row=row, column=8,
                value=f"=SUMIFS(Transactions!$F$5:$F${last},"
                      f"Transactions!$B$5:$B${last},$A{row})").number_format = MONEY
        src.cell(row=row, column=9,
                value=f"=F{row}+G{row}-H{row}").number_format = MONEY
        src.cell(row=row, column=10, value=a["closing"]).number_format = MONEY
        src.cell(row=row, column=11,
                value=f"=IF(ABS(I{row}-J{row})<0.01,\"OK\",\"MISMATCH\")")
    tot = 8
    src.cell(row=tot, column=1, value="TOTAL").font = Font(bold=True)
    for col in "EFGHIJ":
        src.cell(row=tot, column=ord(col) - 64,
                value=f"=SUM({col}5:{col}7)").font = Font(bold=True)
        src.cell(row=tot, column=ord(col) - 64).number_format = MONEY
    widths = [22, 10, 12, 14, 7, 14, 16, 16, 16, 15, 11]
    for j, w in enumerate(widths, start=1):
        src.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    # ---------- Transactions ----------
    tr = wb.create_sheet("Transactions")
    tr["A1"] = "Unified ledger — all statements, one tab"
    tr["A1"].font = TITLE_FONT
    tr["A2"] = ("Every transaction exactly as in its source statement. "
                "Running Balance is live: opening balance + credits − "
                "debits for that account up to this row (Seq). “Final row” "
                "marks each account's last transaction; “Dup flag” marks "
                "any exact repeat of an earlier row.")
    tr["A2"].font = SUB_FONT
    tr.merge_cells("A2:L2")
    theaders = ["Date", "Bank", "Account", "Description", "Category",
                "Debit", "Credit", "Seq", "Running Balance", "Month",
                "Final row", "Dup flag"]
    for j, h in enumerate(theaders, start=1):
        tr.cell(row=4, column=j, value=h)
    style_header(tr, 4, len(theaders))
    for i, t in enumerate(rows):
        row = 5 + i
        tr.cell(row=row, column=1, value=t["date"]).number_format = "yyyy-mm-dd"
        tr.cell(row=row, column=2, value=t["bank"])
        tr.cell(row=row, column=3, value=t["acct"])
        tr.cell(row=row, column=4, value=t["desc"])
        tr.cell(row=row, column=5, value=t["cat"])
        tr.cell(row=row, column=6,
               value=t["debit"] if t["debit"] else None).number_format = MONEY
        tr.cell(row=row, column=7,
               value=t["credit"] if t["credit"] else None).number_format = MONEY
        tr.cell(row=row, column=8, value=t["seq"])
        tr.cell(row=row, column=9,
               value=f"=INDEX(Sources!$F$5:$F$7,MATCH($B{row},"
                     f"Sources!$A$5:$A$7,0))"
                     f"+SUMIFS($G$5:$G${last},$B$5:$B${last},$B{row},"
                     f"$H$5:$H${last},\"<=\"&$H{row})"
                     f"-SUMIFS($F$5:$F${last},$B$5:$B${last},$B{row},"
                     f"$H$5:$H${last},\"<=\"&$H{row})").number_format = MONEY
        tr.cell(row=row, column=10,
               value=f"=TEXT($A{row},\"yyyy-mm\")")
        tr.cell(row=row, column=11,
               value=f"=IF($H{row}=COUNTIFS($B$5:$B${last},$B{row}),1,0)")
        tr.cell(row=row, column=12,
               value=f"=IF(SUMPRODUCT(($A$5:$A{row}=$A{row})"
                     f"*($B$5:$B{row}=$B{row})"
                     f"*($D$5:$D{row}=$D{row})"
                     f"*($F$5:$F{row}=IF($F{row}=\"\",0,$F{row}))"
                     f"*($G$5:$G{row}=IF($G{row}=\"\",0,$G{row})))>1,1,0)")
    tr.auto_filter.ref = f"A4:L{last}"
    tr.freeze_panes = "A5"
    trow = last + 1
    tr.cell(row=trow, column=1, value="TOTAL").font = Font(bold=True)
    tr.cell(row=trow, column=6,
           value=f"=SUM(F5:F{last})").font = Font(bold=True)
    tr.cell(row=trow, column=6).number_format = MONEY
    tr.cell(row=trow, column=7,
           value=f"=SUM(G5:G{last})").font = Font(bold=True)
    tr.cell(row=trow, column=7).number_format = MONEY
    twidths = [12, 22, 11, 42, 20, 14, 14, 6, 16, 9, 9, 9]
    for j, w in enumerate(twidths, start=1):
        tr.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    # ---------- Monthly ----------
    mo = wb.create_sheet("Monthly")
    mo["A1"] = "Monthly summary — per account"
    mo["A1"].font = TITLE_FONT
    mo["A2"] = ("Opening is the source opening plus all earlier months' "
                "net flow, so each month's closing feeds the next. All "
                "cells are live SUMIFS over Transactions.")
    mo["A2"].font = SUB_FONT
    mo.merge_cells("A2:G2")
    mheaders = ["Bank", "Month", "Opening", "Credits", "Debits", "Closing",
                "Txns"]
    for j, h in enumerate(mheaders, start=1):
        mo.cell(row=4, column=j, value=h)
    style_header(mo, 4, len(mheaders))
    mfirst = 5
    mrow = mfirst
    for a in accounts:
        for m in MONTHS:
            mo.cell(row=mrow, column=1, value=a["bank"])
            mo.cell(row=mrow, column=2, value=m)
            mo.cell(row=mrow, column=3,
                   value=f"=INDEX(Sources!$F$5:$F$7,MATCH($A{mrow},"
                         f"Sources!$A$5:$A$7,0))"
                         f"+SUMIFS(Transactions!$G$5:$G${last},"
                         f"Transactions!$B$5:$B${last},$A{mrow},"
                         f"Transactions!$J$5:$J${last},\"<\"&$B{mrow})"
                         f"-SUMIFS(Transactions!$F$5:$F${last},"
                         f"Transactions!$B$5:$B${last},$A{mrow},"
                         f"Transactions!$J$5:$J${last},\"<\"&$B{mrow})"
                   ).number_format = MONEY
            mo.cell(row=mrow, column=4,
                   value=f"=SUMIFS(Transactions!$G$5:$G${last},"
                         f"Transactions!$B$5:$B${last},$A{mrow},"
                         f"Transactions!$J$5:$J${last},$B{mrow})"
                   ).number_format = MONEY
            mo.cell(row=mrow, column=5,
                   value=f"=SUMIFS(Transactions!$F$5:$F${last},"
                         f"Transactions!$B$5:$B${last},$A{mrow},"
                         f"Transactions!$J$5:$J${last},$B{mrow})"
                   ).number_format = MONEY
            mo.cell(row=mrow, column=6,
                   value=f"=C{mrow}+D{mrow}-E{mrow}").number_format = MONEY
            mo.cell(row=mrow, column=7,
                   value=f"=COUNTIFS(Transactions!$B$5:$B${last},$A{mrow},"
                         f"Transactions!$J$5:$J${last},$B{mrow})")
            mrow += 1
    mlast = mrow - 1
    mo.auto_filter.ref = f"A4:G{mlast}"
    mo.freeze_panes = "A5"
    for j, w in enumerate([22, 10, 16, 16, 16, 16, 7], start=1):
        mo.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    # ---------- Dashboard ----------
    db = wb.create_sheet("Dashboard")
    db["A1"] = "Dashboard — consolidated position"
    db["A1"].font = TITLE_FONT
    db["A3"] = "Focus account:"
    db["A3"].font = Font(bold=True)
    db["B3"] = "All"
    db["A4"] = "Focus month:"
    db["A4"].font = Font(bold=True)
    db["B4"] = "All"
    dv_acct = DataValidation(type="list",
                             formula1='"All,Canara Bank,State Bank of India,'
                             'HDFC Bank"',
                             allow_blank=False)
    dv_acct.add("B3")
    db.add_data_validation(dv_acct)
    dv_month = DataValidation(type="list",
                              formula1='"All,' + ",".join(MONTHS) + '"',
                              allow_blank=False)
    dv_month.add("B4")
    db.add_data_validation(dv_month)
    db["A6"] = "Key figures (follow the filters above)"
    db["A6"].font = Font(bold=True, size=12, color="1F4E79")
    bank_crit = f'IF($B$3="All","*",$B$3)'
    month_crit = f'IF($B$4="All","*",$B$4)'
    kpis = [
        ("Total credits", f"=SUMIFS(Transactions!$G$5:$G${last},"
         f"Transactions!$B$5:$B${last},{bank_crit},"
         f"Transactions!$J$5:$J${last},{month_crit})"),
        ("Total debits", f"=SUMIFS(Transactions!$F$5:$F${last},"
         f"Transactions!$B$5:$B${last},{bank_crit},"
         f"Transactions!$J$5:$J${last},{month_crit})"),
        ("Net flow", "=B7-B8"),
        ("Closing balance", "=IF($B$3=\"All\",Sources!$I$8,"
         "INDEX(Sources!$I$5:$I$7,MATCH($B$3,Sources!$A$5:$A$7,0)))"),
        ("Transactions", f"=COUNTIFS(Transactions!$B$5:$B${last},"
         f"{bank_crit},Transactions!$J$5:$J${last},{month_crit})"),
        ("Largest receipt", f"=SUMPRODUCT(MAX(((Transactions!$B$5:$B${last}"
         f"=$B$3)+($B$3=\"All\"))*((Transactions!$J$5:$J${last}=$B$4)"
         f"+($B$4=\"All\"))*Transactions!$G$5:$G${last}))"),
        ("Largest payment", f"=SUMPRODUCT(MAX(((Transactions!$B$5:$B${last}"
         f"=$B$3)+($B$3=\"All\"))*((Transactions!$J$5:$J${last}=$B$4)"
         f"+($B$4=\"All\"))*Transactions!$F$5:$F${last}))"),
    ]
    for i, (label, formula) in enumerate(kpis):
        row = 7 + i
        db.cell(row=row, column=1, value=label).font = Font(bold=True)
        c = db.cell(row=row, column=2, value=formula)
        if label != "Transactions":
            c.number_format = MONEY
    # monthly net-flow helper table
    db["A16"] = "Monthly net flow (all accounts)"
    db["A16"].font = Font(bold=True, size=12, color="1F4E79")
    for j, h in enumerate(["Month", "Credits", "Debits", "Net"], start=1):
        db.cell(row=17, column=j, value=h)
    style_header(db, 17, 4)
    for i, m in enumerate(MONTHS):
        row = 18 + i
        db.cell(row=row, column=1, value=m)
        db.cell(row=row, column=2,
               value=f"=SUMIFS(Transactions!$G$5:$G${last},"
                     f"Transactions!$J$5:$J${last},$A{row})"
               ).number_format = MONEY
        db.cell(row=row, column=3,
               value=f"=SUMIFS(Transactions!$F$5:$F${last},"
                     f"Transactions!$J$5:$J${last},$A{row})"
               ).number_format = MONEY
        db.cell(row=row, column=4,
               value=f"=B{row}-C{row}").number_format = MONEY
    # category table (follows filters; transfers excluded from the
    # income/expense reading by listing them separately at the bottom)
    cats = sorted({t["cat"] for t in rows})
    cat_head = 29
    db.cell(row=cat_head, column=1,
           value="By category (follows the filters above)").font = \
        Font(bold=True, size=12, color="1F4E79")
    for j, h in enumerate(["Category", "Credits", "Debits"], start=1):
        db.cell(row=cat_head + 1, column=j, value=h)
    style_header(db, cat_head + 1, 3)
    for i, cat in enumerate(cats):
        row = cat_head + 2 + i
        db.cell(row=row, column=1, value=cat)
        db.cell(row=row, column=2,
               value=f"=SUMIFS(Transactions!$G$5:$G${last},"
                     f"Transactions!$E$5:$E${last},$A{row},"
                     f"Transactions!$B$5:$B${last},{bank_crit},"
                     f"Transactions!$J$5:$J${last},{month_crit})"
               ).number_format = MONEY
        db.cell(row=row, column=3,
               value=f"=SUMIFS(Transactions!$F$5:$F${last},"
                     f"Transactions!$E$5:$E${last},$A{row},"
                     f"Transactions!$B$5:$B${last},{bank_crit},"
                     f"Transactions!$J$5:$J${last},{month_crit})"
               ).number_format = MONEY
    cat_last = cat_head + 1 + len(cats)
    # charts
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "Monthly net flow"
    ch1.y_axis.title = "Rs"
    ch1.x_axis.title = "Month"
    ch1.add_data(Reference(db, min_col=4, min_row=17, max_row=27), titles_from_data=True)
    ch1.set_categories(Reference(db, min_col=1, min_row=18, max_row=27))
    ch1.height, ch1.width = 7.5, 15
    db.add_chart(ch1, "D6")
    ch2 = BarChart()
    ch2.type = "bar"
    ch2.title = "Credits vs debits by account"
    ch2.add_data(Reference(src, min_col=7, min_row=4, max_row=7), titles_from_data=True)
    ch2.set_categories(Reference(src, min_col=1, min_row=5, max_row=7))
    ch2.height, ch2.width = 7.5, 15
    db.add_chart(ch2, "D16")
    ch3 = BarChart()
    ch3.type = "bar"
    ch3.title = "Debits by category"
    ch3.add_data(Reference(db, min_col=3, min_row=cat_head + 1,
                           max_row=cat_last), titles_from_data=True)
    ch3.set_categories(Reference(db, min_col=1, min_row=cat_head + 2,
                                 max_row=cat_last))
    ch3.height, ch3.width = 7.5, 15
    db.add_chart(ch3, "D29")
    for j, w in enumerate([26, 18, 18, 16], start=1):
        db.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    # ---------- Checks ----------
    ck = wb.create_sheet("Checks")
    ck["A1"] = "Reconciliation checks — all must read PASS"
    ck["A1"].font = TITLE_FONT
    ck["A2"] = ("Each Actual is a live formula over Transactions, Sources "
                "or Monthly. A FAIL means the consolidation does not tie "
                "to the source statements.")
    ck["A2"].font = SUB_FONT
    ck.merge_cells("A2:E2")
    for j, h in enumerate(["#", "Check", "Expected", "Actual", "Result"],
                          start=1):
        ck.cell(row=4, column=j, value=h)
    style_header(ck, 4, 5)
    total_credits = sum(a["credits"] for a in accounts)
    total_debits = sum(a["debits"] for a in accounts)
    checks = [
        ("Transaction rows = source rows", n,
         f"=COUNTA(Transactions!$A$5:$A${last})"),
        ("Total credits = source files", round(total_credits, 2),
         f"=SUM(Transactions!$G$5:$G${last})"),
        ("Total debits = source files", round(total_debits, 2),
         f"=SUM(Transactions!$F$5:$F${last})"),
        ("Canara final balance = source closing",
         accounts[0]["closing"],
         "=SUMPRODUCT((Transactions!$B$5:$B$" + str(last) + "=\"Canara Bank\")"
         f"*Transactions!$K$5:$K${last}*Transactions!$I$5:$I${last})"),
        ("SBI final balance = source closing",
         accounts[1]["closing"],
         "=SUMPRODUCT((Transactions!$B$5:$B$" + str(last)
         + "=\"State Bank of India\")"
         f"*Transactions!$K$5:$K${last}*Transactions!$I$5:$I${last})"),
        ("HDFC final balance = source closing",
         accounts[2]["closing"],
         "=SUMPRODUCT((Transactions!$B$5:$B$" + str(last) + "=\"HDFC Bank\")"
         f"*Transactions!$K$5:$K${last}*Transactions!$I$5:$I${last})"),
        ("Own-account transfers net to zero", 0,
         "=SUMIFS(Transactions!$G$5:$G$" + str(last) +
         ",Transactions!$E$5:$E$" + str(last) + ",\"Transfer\")"
         "-SUMIFS(Transactions!$F$5:$F$" + str(last) +
         ",Transactions!$E$5:$E$" + str(last) + ",\"Transfer\")"),
        ("Duplicate transactions", 0,
         f"=SUM(Transactions!$L$5:$L${last})"),
        ("Uncategorised transactions", 0,
         f"=COUNTIF(Transactions!$E$5:$E${last},\"Uncategorised\")"
         f"+COUNTIF(Transactions!$E$5:$E${last},\"\")"),
        ("October monthly closings = total closing", round(
            sum(a["closing"] for a in accounts), 2),
         f"=SUMIFS(Monthly!$F$5:$F${mlast},Monthly!$B$5:$B${mlast},"
         f"\"2026-10\")"),
    ]
    for i, (label, expected, actual) in enumerate(checks):
        row = 5 + i
        ck.cell(row=row, column=1, value=i + 1)
        ck.cell(row=row, column=2, value=label)
        c = ck.cell(row=row, column=3, value=expected)
        c.number_format = MONEY
        d = ck.cell(row=row, column=4, value=actual)
        d.number_format = MONEY
        ck.cell(row=row, column=5,
               value=f"=IF(ABS(D{row}-C{row})<0.01,\"PASS\",\"FAIL\")")
    ck.conditional_formatting.add(f"E5:E14",
                                  CellIsRule(operator="equal",
                                             formula=['"PASS"'],
                                             fill=GREEN_FILL))
    ck.conditional_formatting.add(f"E5:E14",
                                  CellIsRule(operator="equal",
                                             formula=['"FAIL"'],
                                             fill=RED_FILL))
    for j, w in enumerate([5, 44, 16, 16, 10], start=1):
        ck.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    wb.save(XLSX)

    # flat merged export
    with open(BASE / "data" / "consolidated.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Bank", "Account", "Description", "Category",
                    "Debit", "Credit", "Running Balance"])
        # recompute running balances in file order per account
        seq_bal = {a["bank"]: a["opening"] for a in accounts}
        for t in rows:
            seq_bal[t["bank"]] = round(
                seq_bal[t["bank"]] + t["credit"] - t["debit"], 2)
            w.writerow([t["date"], t["bank"], t["acct"], t["desc"],
                        t["cat"],
                        f"{t['debit']:.2f}" if t["debit"] else "",
                        f"{t['credit']:.2f}" if t["credit"] else "",
                        f"{seq_bal[t['bank']]:.2f}"])
    print(f"built {XLSX.name}: {n} transactions, "
          f"credits Rs {total_credits:,.2f}, debits Rs {total_debits:,.2f}, "
          f"closing Rs {sum(a['closing'] for a in accounts):,.2f}")


if __name__ == "__main__":
    build()
