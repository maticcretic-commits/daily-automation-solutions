"""Build reconpulse.xlsx — formula-driven bank reconciliation workbook.

Reads data/*.csv (generate.py). Every computed cell is a live Excel formula;
no computed values are pasted. Verified by tests/test_build.py via a
LibreOffice headless recalculation against main.py ground truth.

LO-COMPAT notes (from earlier builds):
  * no MAXIFS/MINIFS/XLOOKUP — MATCH/INDEX/COUNTIFS/SUMIFS only
  * range literals repeat the column letter: Sheet!$B$2:$B$129
  * no bare currency symbol inside TEXT() format strings
  * SMALL(IF(...)) extraction pattern is LO-safe (proven 2026-10-20)
"""
import csv
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data")
OUT = os.path.join(ROOT, "reconpulse.xlsx")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=16, color="1F4E79")
SUB_FONT = Font(bold=True, size=12, color="1F4E79")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
MONEY = '"₹"#,##0.00'
PCT1 = "0.0%"
DATEF = "yyyy-mm-dd"
THIN = Border(*[Side(style="thin", color="B0B0B0")] * 4)


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def style_block(ws, r1, r2, c1, c2, money_cols=(), date_cols=(), pct_cols=()):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN
            if c in money_cols:
                cell.number_format = MONEY
            if c in date_cols:
                cell.number_format = DATEF
            if c in pct_cols:
                cell.number_format = PCT1


def build():
    bank = read_csv("bank_statement.csv")
    book = read_csv("book_ledger.csv")
    NB, NK = len(bank), len(book)
    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "ReconPulse — Bank Reconciliation Workbench"
    ws["A1"].font = TITLE_FONT
    lines = [
        ("What this is", "A formula-driven bank reconciliation for Mehta Trading Co., Ludhiana "
         "(fictional demo data, 1–21 Oct 2026): every bank-statement line is auto-matched "
         "to the book ledger by amount + side, with a ±3-day clearing window."),
        ("Inputs (yellow)", "Summary!B2 opening balance, Summary!B3 as-of date. "
         "Paste a fresh statement into Raw_Bank and the ledger into Raw_Book; "
         "extend the formula ranges on Recon/Unmatched/Daily/Summary/Checks."),
        ("How matching works", "Recon!F builds a key amount|side (R=receipt, P=payment); "
         "Recon!G finds the key in the ledger; Recon!K marks MATCHED only when the "
         "booked date is within 3 days of the bank date."),
        ("Unmatched items", "Bank-only items (charges, interest, unidentified deposits) and "
         "book-only items (unpresented cheques, deposits in transit) are extracted with "
         "a suggested action and days outstanding."),
        ("Reconciliation", "Summary walks both sides to one adjusted balance; "
         "Checks!B1:B8 must all read TRUE and the recon difference must be zero."),
        ("Demo state", f"{NB} bank lines, {NK} book lines; 4 bank-only and 3 book-only "
         "items deliberately unmatched."),
    ]
    r = 3
    for title, body in lines:
        ws[f"A{r}"] = title
        ws[f"A{r}"].font = SUB_FONT
        ws[f"B{r}"] = body
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 36
        r += 1
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 120

    # ---------------- Raw_Bank ----------------
    ws = wb.create_sheet("Raw_Bank")
    header(ws, 1, ["Date", "Description", "Ref", "Debit", "Credit", "Balance"],
           [13, 44, 16, 15, 15, 17])
    for i, row in enumerate(bank, 2):
        ws.cell(row=i, column=1, value=datetime.fromisoformat(row["Date"]).date())
        ws.cell(row=i, column=2, value=row["Description"])
        ws.cell(row=i, column=3, value=row["Ref"])
        ws.cell(row=i, column=4, value=float(row["Debit"]))
        ws.cell(row=i, column=5, value=float(row["Credit"]))
        if i == 2:
            ws.cell(row=i, column=6, value="=Summary!$B$2+E2-D2")
        else:
            ws.cell(row=i, column=6, value=f"=F{i-1}+E{i}-D{i}")
    style_block(ws, 2, NB + 1, 1, 6, money_cols=(4, 5, 6), date_cols=(1,))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{NB+1}"

    # ---------------- Raw_Book ----------------
    ws = wb.create_sheet("Raw_Book")
    header(ws, 1, ["Date", "Narration", "Voucher", "Debit", "Credit", "Balance",
                   "Key", "Matched?", "UPos"], [13, 44, 14, 15, 15, 17, 16, 11, 8])
    for i, row in enumerate(book, 2):
        ws.cell(row=i, column=1, value=datetime.fromisoformat(row["Date"]).date())
        ws.cell(row=i, column=2, value=row["Narration"])
        ws.cell(row=i, column=3, value=row["Voucher"])
        ws.cell(row=i, column=4, value=float(row["Debit"]))
        ws.cell(row=i, column=5, value=float(row["Credit"]))
        if i == 2:
            ws.cell(row=i, column=6, value="=Summary!$B$2+D2-E2")
        else:
            ws.cell(row=i, column=6, value=f"=F{i-1}+D{i}-E{i}")
        ws.cell(row=i, column=7,
                value=f'=TEXT(MAX(D{i},E{i}),"0.00")&"|"&IF(D{i}>0,"R","P")')
        ws.cell(row=i, column=8,
                value=(f"=COUNTIFS(Recon!$F$2:$F${NB+1},G{i},"
                       f'Recon!$K$2:$K${NB+1},"MATCHED")>0'))
        ws.cell(row=i, column=9, value=f'=IF(H{i}=FALSE,ROW()-1,"")')
    style_block(ws, 2, NK + 1, 1, 9, money_cols=(4, 5, 6), date_cols=(1,))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{NK+1}"

    # ---------------- Recon ----------------
    ws = wb.create_sheet("Recon")
    header(ws, 1, ["Bank date", "Description", "Ref", "Bank debit", "Bank credit",
                   "Key", "Match row", "Book date", "Book narration", "Date diff",
                   "Status", "Amount", "Diff", "UPos"],
           [12, 40, 15, 14, 14, 15, 11, 12, 40, 10, 12, 14, 12, 8])
    for i in range(2, NB + 2):
        ws.cell(row=i, column=1, value=f"=Raw_Bank!A{i}")
        ws.cell(row=i, column=2, value=f"=Raw_Bank!B{i}")
        ws.cell(row=i, column=3, value=f"=Raw_Bank!C{i}")
        ws.cell(row=i, column=4, value=f"=Raw_Bank!D{i}")
        ws.cell(row=i, column=5, value=f"=Raw_Bank!E{i}")
        ws.cell(row=i, column=6,
                value=f'=TEXT(MAX(D{i},E{i}),"0.00")&"|"&IF(E{i}>0,"R","P")')
        ws.cell(row=i, column=7, value=f"=MATCH(F{i},Raw_Book!$G$2:$G${NK+1},0)")
        ws.cell(row=i, column=8,
                value=f'=IF(ISNA(G{i}),"",INDEX(Raw_Book!$A$2:$A${NK+1},G{i}))')
        ws.cell(row=i, column=9,
                value=f'=IF(ISNA(G{i}),"",INDEX(Raw_Book!$B$2:$B${NK+1},G{i}))')
        ws.cell(row=i, column=10, value=f'=IF(ISNA(G{i}),"",ABS(A{i}-H{i}))')
        ws.cell(row=i, column=11,
                value=f'=IF(OR(ISNA(G{i}),J{i}>3),"UNMATCHED","MATCHED")')
        ws.cell(row=i, column=12, value=f"=MAX(D{i},E{i})")
        ws.cell(row=i, column=13, value=f'=IF(K{i}="MATCHED",0,L{i})')
        ws.cell(row=i, column=14, value=f'=IF(K{i}="UNMATCHED",ROW()-1,"")')
    style_block(ws, 2, NB + 1, 1, 14, money_cols=(4, 5, 12, 13), date_cols=(1, 8))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:N{NB+1}"

    # ---------------- Unmatched ----------------
    # LO-safe extraction: a numeric helper column (H) holds SMALL() positions
    # over the per-row UPos helpers — never SMALL(IF(range=...)) which LO
    # reduces by implicit intersection (lesson 2026-10-22).
    ws = wb.create_sheet("Unmatched")
    ws["A1"] = "Bank-only items — in the statement, not in the books"
    ws["A1"].font = SUB_FONT
    header(ws, 2, ["Item", "Date", "Amount", "Side", "Reason",
                   "Suggested action", "Days outstanding", "pos"],
           [40, 12, 14, 10, 22, 52, 16, 8])
    for r in range(3, 9):
        k = r - 2
        ws.cell(row=r, column=8,
                value=f'=IFERROR(SMALL(Recon!$N$2:$N${NB+1},{k}),"")')
        ws.cell(row=r, column=1,
                value=f'=IF($H{r}="","",INDEX(Recon!$B$2:$B${NB+1},$H{r}))')
        ws.cell(row=r, column=2,
                value=f'=IF($H{r}="","",INDEX(Recon!$A$2:$A${NB+1},$H{r}))')
        ws.cell(row=r, column=3,
                value=f'=IF($H{r}="","",INDEX(Recon!$L$2:$L${NB+1},$H{r}))')
        ws.cell(row=r, column=4, value=(
            f'=IF($H{r}="","",IF(INDEX(Recon!$E$2:$E${NB+1},$H{r})>0,'
            f'"Receipt","Payment"))'))
        ws.cell(row=r, column=5, value=(
            f'=IF(A{r}="","",IF(ISNUMBER(SEARCH("charg",A{r})),"Bank charges",'
            f'IF(ISNUMBER(SEARCH("Interest",A{r})),"Interest credited",'
            f'"Unidentified deposit"))'))
        ws.cell(row=r, column=6, value=(
            f'=IF(A{r}="","",IF(E{r}="Bank charges","Record the charges in the book ledger",'
            f'IF(E{r}="Interest credited","Record interest income in the books",'
            f'"Identify the depositor from the UTR; book the receipt or arrange a refund"))'))
        ws.cell(row=r, column=7, value=f'=IF(A{r}="","",Summary!$B$3-B{r})')
    ws["A10"] = "Book-only items — in the books, not yet at the bank"
    ws["A10"].font = SUB_FONT
    header(ws, 11, ["Item", "Date", "Amount", "Side", "Reason",
                    "Suggested action", "Days outstanding", "pos"],
           [40, 12, 14, 10, 22, 52, 16, 8])
    for r in range(12, 20):
        k = r - 11
        ws.cell(row=r, column=8,
                value=f'=IFERROR(SMALL(Raw_Book!$I$2:$I${NK+1},{k}),"")')
        ws.cell(row=r, column=1,
                value=f'=IF($H{r}="","",INDEX(Raw_Book!$B$2:$B${NK+1},$H{r}))')
        ws.cell(row=r, column=2,
                value=f'=IF($H{r}="","",INDEX(Raw_Book!$A$2:$A${NK+1},$H{r}))')
        ws.cell(row=r, column=3, value=(
            f'=IF($H{r}="","",MAX(INDEX(Raw_Book!$D$2:$D${NK+1},$H{r}),'
            f'INDEX(Raw_Book!$E$2:$E${NK+1},$H{r})))'))
        ws.cell(row=r, column=4, value=(
            f'=IF(A{r}="","",IF(ISNUMBER(SEARCH("transit",A{r})),"Receipt","Payment"))'))
        ws.cell(row=r, column=5, value=(
            f'=IF(A{r}="","",IF(ISNUMBER(SEARCH("unpresented",A{r})),'
            f'"Unpresented cheque","Deposit in transit"))'))
        ws.cell(row=r, column=6, value=(
            f'=IF(A{r}="","",IF(E{r}="Unpresented cheque",'
            f'"Follow up with the payee; reissue if stale (>3 months)",'
            f'"Confirm credit in the next statement; then it auto-matches"))'))
        ws.cell(row=r, column=7, value=f'=IF(A{r}="","",Summary!$B$3-B{r})')
    style_block(ws, 3, 8, 1, 8, money_cols=(3,), date_cols=(2,))
    style_block(ws, 12, 19, 1, 8, money_cols=(3,), date_cols=(2,))

    # ---------------- Daily ----------------
    ws = wb.create_sheet("Daily")
    header(ws, 1, ["Date", "Bank balance (EOD)", "Book balance (EOD)"],
           [13, 20, 20])
    for i, day in enumerate(range(1, 22), 2):
        ws.cell(row=i, column=1, value=datetime(2026, 10, day).date())
        ws.cell(row=i, column=2, value=(
            f"=Summary!$B$2+SUMIFS(Raw_Bank!$E$2:$E${NB+1},Raw_Bank!$A$2:$A${NB+1},"
            f'"<="&A{i})-SUMIFS(Raw_Bank!$D$2:$D${NB+1},Raw_Bank!$A$2:$A${NB+1},'
            f'"<="&A{i})'))
        ws.cell(row=i, column=3, value=(
            f"=Summary!$B$2+SUMIFS(Raw_Book!$D$2:$D${NK+1},Raw_Book!$A$2:$A${NK+1},"
            f'"<="&A{i})-SUMIFS(Raw_Book!$E$2:$E${NK+1},Raw_Book!$A$2:$A${NK+1},'
            f'"<="&A{i})'))
    style_block(ws, 2, 22, 1, 3, money_cols=(2, 3), date_cols=(1,))

    # ---------------- Summary ----------------
    ws = wb.create_sheet("Summary")
    ws["A1"] = "Bank reconciliation — Mehta Trading Co. (demo)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Opening balance 30-Sep-2026"
    ws["B2"] = 4500000.00
    ws["B2"].fill = INPUT_FILL
    ws["A3"] = "As-of date"
    ws["B3"] = datetime(2026, 10, 21).date()
    ws["B3"].fill = INPUT_FILL
    ws["B3"].number_format = DATEF
    ws["A4"] = "Balance per bank statement"
    ws["B4"] = f"=Raw_Bank!F{NB+1}"
    ws["A5"] = "Balance per books"
    ws["B5"] = f"=Raw_Book!F{NK+1}"
    ws["A6"] = "Walk — bank side"
    ws["A6"].font = SUB_FONT
    ws["A7"] = "Balance per bank statement"
    ws["B7"] = "=B4"
    ws["A8"] = "Add: bank charges not yet in books"
    ws["B8"] = '=SUMIFS(Unmatched!$C$3:$C$8,Unmatched!$E$3:$E$8,"Bank charges")'
    ws["A9"] = "Less: interest credited, not yet in books"
    ws["B9"] = ('=SUMIFS(Unmatched!$C$3:$C$8,Unmatched!$E$3:$E$8,'
                '"Interest credited")')
    ws["A10"] = "Less: unidentified deposits, not yet in books"
    ws["B10"] = ('=SUMIFS(Unmatched!$C$3:$C$8,Unmatched!$E$3:$E$8,'
                 '"Unidentified deposit")')
    ws["A11"] = "Adjusted bank balance"
    ws["B11"] = "=B7+B8-B9-B10"
    ws["B11"].font = Font(bold=True)
    ws["A12"] = "Walk — book side"
    ws["A12"].font = SUB_FONT
    ws["A13"] = "Balance per books"
    ws["B13"] = "=B5"
    ws["A14"] = "Add: cheques issued, not yet presented"
    ws["B14"] = ('=SUMIFS(Unmatched!$C$12:$C$19,Unmatched!$E$12:$E$19,'
                 '"Unpresented cheque")')
    ws["A15"] = "Less: deposits in transit"
    ws["B15"] = ('=SUMIFS(Unmatched!$C$12:$C$19,Unmatched!$E$12:$E$19,'
                 '"Deposit in transit")')
    ws["A16"] = "Adjusted book balance"
    ws["B16"] = "=B13+B14-B15"
    ws["B16"].font = Font(bold=True)
    ws["A17"] = "Reconciliation difference (must be 0)"
    ws["B17"] = "=B11-B16"
    ws["B17"].font = Font(bold=True)
    kpis = [
        ("Total receipts (bank)", f"=SUM(Raw_Bank!$E$2:$E${NB+1})"),
        ("Total payments (bank)", f"=SUM(Raw_Bank!$D$2:$D${NB+1})"),
        ("Lines matched", f'=COUNTIF(Recon!$K$2:$K${NB+1},"MATCHED")'),
        ("Unmatched — bank side", f'=COUNTIF(Recon!$K$2:$K${NB+1},"UNMATCHED")'),
        ("Unmatched — book side", f"=COUNTIF(Raw_Book!$H$2:$H${NK+1},FALSE)"),
        ("Match rate", "=B21/(B21+B22)"),
        ("Oldest outstanding item (days)", "=MAX(Unmatched!$G$3:$G$8,Unmatched!$G$12:$G$19)"),
    ]
    r = 19
    for label, formula in kpis:
        ws[f"A{r}"] = label
        ws[f"B{r}"] = formula
        r += 1
    for r in range(2, 26):
        ws[f"A{r}"].border = THIN
        ws[f"B{r}"].border = THIN
        if r in (2, 4, 5, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 19, 20):
            ws[f"B{r}"].number_format = MONEY
    ws["B24"].number_format = PCT1
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 20

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Reconciliation dashboard — as of 21 Oct 2026"
    ws["A1"].font = TITLE_FONT
    cards = [
        ("Match rate", "=Summary!B24", PCT1),
        ("Lines matched", "=Summary!B21", None),
        ("Unmatched items", "=Summary!B22+Summary!B23", None),
        ("Recon difference", "=Summary!B17", MONEY),
        ("Bank balance", "=Summary!B4", MONEY),
        ("Book balance", "=Summary!B5", MONEY),
        ("Adjusted balance", "=Summary!B11", MONEY),
        ("Oldest outstanding (days)", "=Summary!B25", None),
    ]
    for i, (label, formula, fmt) in enumerate(cards):
        r = 3 + (i % 4)
        c = 1 + (i // 4) * 3
        ws.cell(row=r, column=c, value=label).font = Font(bold=True, size=11)
        v = ws.cell(row=r, column=c + 1, value=formula)
        v.font = Font(bold=True, size=13, color="1F4E79")
        if fmt:
            v.number_format = fmt
        for cc in (c, c + 1):
            ws.cell(row=r, column=cc).border = THIN
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["D"].width = 24
    ws.column_dimensions["E"].width = 18
    # helper tables for charts
    ws["A9"] = "Status"
    ws["B9"] = "Lines"
    ws["A10"] = "Matched"
    ws["B10"] = "=Summary!B21"
    ws["A11"] = "Unmatched"
    ws["B11"] = "=Summary!B22"
    ws["A13"] = "Reason"
    ws["B13"] = "Amount"
    reasons = [("Bank charges", "=Summary!B8"),
               ("Interest credited", "=Summary!B9"),
               ("Unidentified deposit", "=Summary!B10"),
               ("Unpresented cheques", "=Summary!B14"),
               ("Deposits in transit", "=Summary!B15")]
    for i, (label, formula) in enumerate(reasons, 14):
        ws.cell(row=i, column=1, value=label)
        v = ws.cell(row=i, column=2, value=formula)
        v.number_format = MONEY

    pie = PieChart()
    pie.title = "Bank lines: matched vs unmatched"
    labels = Reference(ws, min_col=1, min_row=10, max_row=11)
    data = Reference(ws, min_col=2, min_row=9, max_row=11)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.height = 7
    ws.add_chart(pie, "G3")

    bar = BarChart()
    bar.type = "col"
    bar.title = "Unmatched value by reason"
    bar.y_axis.title = "₹"
    data = Reference(ws, min_col=2, min_row=13, max_row=18)
    cats = Reference(ws, min_col=1, min_row=14, max_row=18)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 7
    ws.add_chart(bar, "G18")

    line = LineChart()
    line.title = "Daily closing balance — bank vs books"
    line.y_axis.title = "₹"
    line.style = 2
    daily = wb["Daily"]
    cats = Reference(daily, min_col=1, min_row=2, max_row=22)
    for col, title in ((2, "Bank"), (3, "Books")):
        s = Reference(daily, min_col=col, min_row=1, max_row=22)
        line.add_data(s, titles_from_data=True)
    line.set_categories(cats)
    line.height = 8
    line.width = 22
    ws.add_chart(line, "A20")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    header(ws, 1, ["#", "Check", "Result"], [6, 64, 12])
    checks = [
        "Every bank line classified (MATCHED/UNMATCHED)",
        "Matched pairs agree on amount (total diff = 0)",
        "Reconciliation difference is zero",
        "124 book lines matched",
        "4 unmatched bank-side items",
        "3 unmatched book-side items",
        "Daily bank EOD ties to statement balance",
        "Daily book EOD ties to ledger balance",
    ]
    formulas = [
        f"=COUNTA(Recon!$K$2:$K${NB+1})={NB}",
        f'=SUMIFS(Recon!$M$2:$M${NB+1},Recon!$K$2:$K${NB+1},"MATCHED")=0',
        "=Summary!$B$17=0",
        f"=COUNTIF(Raw_Book!$H$2:$H${NK+1},TRUE)=124",
        "=Summary!$B$22=4",
        "=Summary!$B$23=3",
        "=Daily!$B$22=Summary!$B$4",
        "=Daily!$C$22=Summary!$B$5",
    ]
    for i, (label, formula) in enumerate(zip(checks, formulas), 2):
        ws.cell(row=i, column=1, value=i - 1)
        ws.cell(row=i, column=2, value=label)
        ws.cell(row=i, column=3, value=formula)
        for c in range(1, 4):
            ws.cell(row=i, column=c).border = THIN
    ws.column_dimensions["B"].width = 64

    for name in ("Guide", "Raw_Bank", "Raw_Book", "Recon", "Unmatched",
                 "Daily", "Summary", "Dashboard", "Checks"):
        wb[name].sheet_properties.tabColor = "1F4E79"

    wb.save(OUT)
    print(f"wrote {OUT} ({NB} bank rows, {NK} book rows)")


if __name__ == "__main__":
    build()
