"""Build payment-collections-tracker.xlsx — formula-driven collections dashboard.

Reads data/schedule.csv + data/payments.csv (from generate.py), writes the
workbook with live formulas only (no pasted values in computed columns).

LO-COMPAT NOTE: LibreOffice does not recognise MAXIFS via xlsx import
(#NAME?), so LastPayDate uses SUMPRODUCT(MAX((range=crit)*dates)) — verified.
"""
import csv
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

AS_OF = date(2026, 10, 20)
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=16, color="1F4E79")
MONEY_FMT = '#,##0'
PCT_FMT = '0.0%'
DATE_FMT = 'YYYY-MM-DD'
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BUCKETS = ["COLLECTED", "1-30", "31-60", "61-90", "90+"]


def style_header(ws, ncols, nrows=1):
    for r in range(1, nrows + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER


def style_range(ws, min_row, max_row, ncols):
    for r in range(min_row, max_row + 1):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER


def autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def load_csv(path):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def build():
    sched = load_csv("data/schedule.csv")
    pays = load_csv("data/payments.csv")
    ns, np_ = len(sched), len(pays)
    sched_last = ns + 1
    pay_last = np_ + 1
    led_last = ns + 1

    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "CollectionPulse — Payment Collections Tracker"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "As-of date (all DPD maths anchor here):"
    ws["B3"] = AS_OF
    ws["B3"].number_format = DATE_FMT
    ws["B3"].font = Font(bold=True)
    rows = [
        ("Sheet", "What it holds"),
        ("Raw_Schedule", "One row per scheduled EMI due (read-only source data)"),
        ("Raw_Payments", "One row per actual receipt (read-only source data)"),
        ("Receivables_Ledger", "Per-due ledger: paid, outstanding, status, DPD, bucket — all formulas"),
        ("Aging", "Outstanding by DPD bucket (COLLECTED / 1-30 / 31-60 / 61-90 / 90+)"),
        ("Collector_Scorecard", "Per-collector efficiency, overdue load, avg DPD, rank"),
        ("FollowUp", "Auto-built action list: every overdue due with severity + next-action date"),
        ("Dashboard", "8 KPIs + 3 native charts (monthly efficiency, aging, collectors)"),
        ("AI_Report", "Rule-based narrative summary + 8 status flags"),
        ("Checks", "8 integrity checks — all must read TRUE"),
    ]
    for i, (a, b) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=a)
        ws.cell(row=i, column=2, value=b)
    style_header(ws, 2)
    ws["A16"] = "How to refresh: paste new dues into Raw_Schedule and receipts into Raw_Payments,"
    ws["A17"] = "extend the Receivables_Ledger formulas down, and every sheet recalculates."
    ws["A19"] = "Demo data is fictional (seeded generator). Replace with real exports to go live."
    autofit(ws, [22, 80])

    # ---------------- Raw_Schedule ----------------
    ws = wb.create_sheet("Raw_Schedule")
    hdr = ["DueID", "Account", "Borrower", "Branch", "Collector",
           "LoanType", "DueDate", "DueAmount"]
    ws.append(hdr)
    for r in sched:
        ws.append([r["DueID"], r["Account"], r["Borrower"], r["Branch"],
                   r["Collector"], r["LoanType"],
                   date.fromisoformat(r["DueDate"]), int(r["DueAmount"])])
    style_header(ws, len(hdr))
    style_range(ws, 2, sched_last, len(hdr))
    for row in ws.iter_rows(min_row=2, max_row=sched_last, min_col=7, max_col=7):
        row[0].number_format = DATE_FMT
    for row in ws.iter_rows(min_row=2, max_row=sched_last, min_col=8, max_col=8):
        row[0].number_format = MONEY_FMT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{sched_last}"
    autofit(ws, [9, 10, 20, 13, 12, 10, 12, 11])

    # ---------------- Raw_Payments ----------------
    ws = wb.create_sheet("Raw_Payments")
    hdr = ["PaymentID", "DueID", "Account", "PayDate", "Amount", "Mode"]
    ws.append(hdr)
    for r in pays:
        ws.append([r["PaymentID"], r["DueID"], r["Account"],
                   date.fromisoformat(r["PayDate"]), int(r["Amount"]), r["Mode"]])
    style_header(ws, len(hdr))
    style_range(ws, 2, pay_last, len(hdr))
    for row in ws.iter_rows(min_row=2, max_row=pay_last, min_col=4, max_col=4):
        row[0].number_format = DATE_FMT
    for row in ws.iter_rows(min_row=2, max_row=pay_last, min_col=5, max_col=5):
        row[0].number_format = MONEY_FMT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{pay_last}"
    autofit(ws, [11, 9, 10, 12, 11, 8])

    # ---------------- Receivables_Ledger ----------------
    ws = wb.create_sheet("Receivables_Ledger")
    hdr = ["DueID", "Account", "Borrower", "Branch", "Collector", "LoanType",
           "DueDate", "DueAmount", "PaidTotal", "Outstanding", "Status",
           "LastPayDate", "DaysOverdue", "Bucket", "Severity", "RowKey"]
    ws.append(hdr)
    for i, s in enumerate(sched, start=2):
        ws.cell(row=i, column=1, value=s["DueID"])
        ws.cell(row=i, column=2, value=f"=INDEX(Raw_Schedule!B:B,MATCH(A{i},Raw_Schedule!A:A,0))")
        ws.cell(row=i, column=3, value=f"=INDEX(Raw_Schedule!C:C,MATCH(A{i},Raw_Schedule!A:A,0))")
        ws.cell(row=i, column=4, value=f"=INDEX(Raw_Schedule!D:D,MATCH(A{i},Raw_Schedule!A:A,0))")
        ws.cell(row=i, column=5, value=f"=INDEX(Raw_Schedule!E:E,MATCH(A{i},Raw_Schedule!A:A,0))")
        ws.cell(row=i, column=6, value=f"=INDEX(Raw_Schedule!F:F,MATCH(A{i},Raw_Schedule!A:A,0))")
        ws.cell(row=i, column=7, value=f"=INDEX(Raw_Schedule!G:G,MATCH(A{i},Raw_Schedule!A:A,0))")
        ws.cell(row=i, column=8, value=f"=INDEX(Raw_Schedule!H:H,MATCH(A{i},Raw_Schedule!A:A,0))")
        ws.cell(row=i, column=9, value=f"=SUMIFS(Raw_Payments!E:E,Raw_Payments!B:B,A{i})")
        ws.cell(row=i, column=10, value=f"=H{i}-I{i}")
        ws.cell(row=i, column=11, value=f'=IF(J{i}<=0.005,"PAID",IF(I{i}>0,"PARTIAL","OVERDUE"))')
        # LO has no MAXIFS via xlsx: SUMPRODUCT forces the array evaluation.
        ws.cell(row=i, column=12,
                value=f'=IF(COUNTIFS(Raw_Payments!B:B,A{i})=0,"",'
                      f'SUMPRODUCT(MAX((Raw_Payments!B$2:B${pay_last}=A{i})'
                      f'*Raw_Payments!D$2:D${pay_last})))')
        ws.cell(row=i, column=13,
                value=f'=IF(K{i}="PAID",MAX(0,L{i}-G{i}),Guide!$B$3-G{i})')
        ws.cell(row=i, column=14,
                value=f'=IF(J{i}<=0.005,"COLLECTED",IF(M{i}<=30,"1-30",'
                      f'IF(M{i}<=60,"31-60",IF(M{i}<=90,"61-90","90+"))))')
        ws.cell(row=i, column=15,
                value=f'=IF(N{i}="90+","LOST",IF(N{i}="61-90","ESCALATE",'
                      f'IF(N{i}="31-60","WATCH",IF(N{i}="1-30","MILD","OK"))))')
        ws.cell(row=i, column=16, value=f"=IF(J{i}>0.005,ROW()-1,\"\")")
    style_header(ws, len(hdr))
    style_range(ws, 2, led_last, len(hdr))
    for row in ws.iter_rows(min_row=2, max_row=led_last, min_col=7, max_col=7):
        row[0].number_format = DATE_FMT
    for row in ws.iter_rows(min_row=2, max_row=led_last, min_col=12, max_col=12):
        row[0].number_format = DATE_FMT
    for col in (8, 9, 10):
        for row in ws.iter_rows(min_row=2, max_row=led_last, min_col=col, max_col=col):
            row[0].number_format = MONEY_FMT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:P{led_last}"
    autofit(ws, [9, 10, 20, 13, 12, 10, 12, 11, 11, 12, 10, 12, 11, 10, 10, 8])

    # ---------------- Aging ----------------
    ws = wb.create_sheet("Aging")
    hdr = ["Bucket", "Dues", "Due (Rs)", "Collected (Rs)", "Outstanding (Rs)", "Collection %"]
    ws.append(hdr)
    for j, b in enumerate(BUCKETS, start=2):
        ws.cell(row=j, column=1, value=b)
        ws.cell(row=j, column=2, value=f"=COUNTIFS(Receivables_Ledger!N$2:N${led_last},A{j})")
        ws.cell(row=j, column=3, value=f"=SUMIFS(Receivables_Ledger!H$2:H${led_last},Receivables_Ledger!N$2:N${led_last},A{j})")
        ws.cell(row=j, column=4, value=f"=SUMIFS(Receivables_Ledger!I$2:I${led_last},Receivables_Ledger!N$2:N${led_last},A{j})")
        ws.cell(row=j, column=5, value=f"=SUMIFS(Receivables_Ledger!J$2:J${led_last},Receivables_Ledger!N$2:N${led_last},A{j})")
        ws.cell(row=j, column=6, value=f"=IF(C{j}=0,0,D{j}/C{j})")
    t = len(BUCKETS) + 2
    ws.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
    for c, col in zip(range(2, 6), "BCDE"):
        ws.cell(row=t, column=c, value=f"=SUM({col}2:{col}{t-1})").font = Font(bold=True)
    ws.cell(row=t, column=6, value=f"=IF(C{t}=0,0,D{t}/C{t})").font = Font(bold=True)
    style_header(ws, len(hdr))
    style_range(ws, 2, t, len(hdr))
    for col in (3, 4, 5):
        for row in ws.iter_rows(min_row=2, max_row=t, min_col=col, max_col=col):
            row[0].number_format = MONEY_FMT
    for row in ws.iter_rows(min_row=2, max_row=t, min_col=6, max_col=6):
        row[0].number_format = PCT_FMT
    autofit(ws, [12, 8, 13, 15, 16, 12])

    # ---------------- Collector_Scorecard ----------------
    ws = wb.create_sheet("Collector_Scorecard")
    hdr = ["Collector", "Dues", "Due (Rs)", "Collected (Rs)", "Efficiency %",
           "Overdue (Rs)", "Avg DPD", "Rank"]
    ws.append(hdr)
    collectors = sorted({s["Collector"] for s in sched})
    for j, name in enumerate(collectors, start=2):
        ws.cell(row=j, column=1, value=name)
        ws.cell(row=j, column=2, value=f"=COUNTIFS(Receivables_Ledger!E$2:E${led_last},A{j})")
        ws.cell(row=j, column=3, value=f"=SUMIFS(Receivables_Ledger!H$2:H${led_last},Receivables_Ledger!E$2:E${led_last},A{j})")
        ws.cell(row=j, column=4, value=f"=SUMIFS(Receivables_Ledger!I$2:I${led_last},Receivables_Ledger!E$2:E${led_last},A{j})")
        ws.cell(row=j, column=5, value=f"=IF(C{j}=0,0,D{j}/C{j})")
        ws.cell(row=j, column=6,
                value=f'=SUMIFS(Receivables_Ledger!J$2:J${led_last},Receivables_Ledger!E$2:E${led_last},A{j},'
                      f'Receivables_Ledger!J$2:J${led_last},">0.005")')
        ws.cell(row=j, column=7,
                value=f'=IF(COUNTIFS(Receivables_Ledger!E$2:E${led_last},A{j},'
                      f'Receivables_Ledger!J$2:J${led_last},">0.005")=0,0,'
                      f'AVERAGEIFS(Receivables_Ledger!M$2:M${led_last},Receivables_Ledger!E$2:E${led_last},A{j},'
                      f'Receivables_Ledger!J$2:J${led_last},">0.005"))')
        ws.cell(row=j, column=8, value=f"=RANK(E{j},$E$2:$E${len(collectors)+1})")
    style_header(ws, len(hdr))
    style_range(ws, 2, len(collectors) + 1, len(hdr))
    for col in (3, 4, 6):
        for row in ws.iter_rows(min_row=2, max_row=len(collectors) + 1, min_col=col, max_col=col):
            row[0].number_format = MONEY_FMT
    for row in ws.iter_rows(min_row=2, max_row=len(collectors) + 1, min_col=5, max_col=5):
        row[0].number_format = PCT_FMT
    autofit(ws, [13, 8, 13, 15, 12, 14, 10, 7])
    return wb, sched, pays, ns, led_last, collectors
