#!/usr/bin/env python3
"""Build payment-tracker.xlsx from data/payments.csv — every figure a live formula.

Sheets: Guide, Dashboard, Payments, Customers, Monthly, Overdue,
Assumptions, Checks.
Rules honoured (learned the hard way): plain A1 ranges only, range
literals repeat the column letter on both ends ($E$2:$E$321), no Table
objects, chart series titles come from header cells (titles_from_data),
dates are real date values, charts are added only after every source
sheet exists. No macros, no sheet protection — every formula is
unlocked and inspectable, as the brief asks. The Payments sheet carries
200 pre-filled formula rows of headroom, so new manual entries pasted
into columns A-F inside the headroom are picked up without editing any
formula; a blank customer yields blank formula results, never an error.
Status logic (one definition, used by the workbook, the ground truth
and the Guide): Paid when pending <= 0; otherwise Overdue/Partial/Unpaid
by due date vs the as-of date and whether anything was paid.
"""
import csv
import datetime
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
OUT = os.path.join(HERE, "payment-tracker.xlsx")
HEADROOM = 200
OVERDUE_ROWS = 50

NAVY = "1F3864"
LIGHT = "D9E2F3"
MONEY = "#,##0"
PCT = "0.0%"
DATEFMT = "yyyy-mm-dd"

MONTHS = [(202600 + m, label) for m, label in
          [(4, "Apr 2026"), (5, "May 2026"), (6, "Jun 2026"),
           (7, "Jul 2026"), (8, "Aug 2026"), (9, "Sep 2026"),
           (10, "Oct 2026")]]
STATUSES = ["Paid", "Partial", "Unpaid", "Partial - Overdue",
            "Unpaid - Overdue"]
AS_OF = datetime.date(2026, 11, 3)


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def dval(text):
    return datetime.date.fromisoformat(text)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def main():
    payments = read_csv("payments.csv")
    n = len(payments)
    last = 1 + n + HEADROOM  # last Payments row carrying formulas
    customers = sorted({r["customer"] for r in payments})
    ncust = len(customers)
    cust_last = 1 + ncust  # last customer data row

    E = f"Payments!$E$2:$E${last}"
    F = f"Payments!$F$2:$F${last}"
    G = f"Payments!$G$2:$G${last}"
    H = f"Payments!$H$2:$H${last}"
    I = f"Payments!$I$2:$I${last}"
    J = f"Payments!$J$2:$J${last}"
    K = f"Payments!$K$2:$K${last}"
    A = f"Payments!$A$2:$A${last}"
    D = f"Payments!$D$2:$D${last}"
    ASOF = "Assumptions!$B$2"

    wb = Workbook()

    # ---- Assumptions (created first so formulas can reference it) ----
    wa = wb.active
    wa.title = "Assumptions"
    wa["A1"] = "Setting"
    wa["B1"] = "Value"
    style_header(wa, 1, 2)
    wa["A2"] = "As-of date (statuses and overdue ageing are measured to this date)"
    wa["B2"] = AS_OF
    wa["B2"].number_format = DATEFMT
    wa["A3"] = "Payment terms (days)"
    wa["B3"] = 30
    wa["A4"] = "Headroom formula rows on Payments"
    wa["B4"] = HEADROOM
    wa["A5"] = "Currency"
    wa["B5"] = "INR (Rs)"
    wa["A7"] = ("Change B2 to re-age every status, overdue count and overdue "
                "amount in the workbook. Do not type over the formula "
                "columns G-L on Payments.")
    widths(wa, {"A": 62, "B": 16})

    # ---- Guide ----
    wg = wb.create_sheet("Guide")
    guide = [
        ("Payment Tracker — Sharma Traders, New Delhi", True),
        ("Built for the Automated Excel Payment Tracker brief: a raw "
         "customer-payment table turned into one clean workbook with "
         "formulas for pending amounts and paid / partial / unpaid "
         "status, colour cues, and a real-time summary. No macros — "
         "nothing that triggers a macro warning.", False),
        ("How to add a payment: paste or type the new row into columns "
         f"A-F of Payments, on any row up to row {last} — columns G-L "
         "already carry formulas there, so the pending amount, status, "
         "days overdue and month key fill in themselves and every "
         "summary updates. Do not paste values over columns G-L, and "
         "do not insert rows above row 2 (formulas reference fixed "
         "ranges).", False),
        ("Status rules: Paid = nothing pending. Partial = something "
         "paid, balance outstanding, due date not yet passed. Unpaid = "
         "nothing paid, due date not yet passed. Partial - Overdue / "
         "Unpaid - Overdue = balance outstanding and the due date has "
         "passed the as-of date on Assumptions.", False),
        ("Colour cues: green = Paid, amber = Partial, grey = Unpaid, "
         "red/orange = overdue statuses; pending amounts over Rs 0 are "
         "shaded pink and days overdue over 30 days are shaded red.",
         False),
        ("Sheets: Dashboard = KPIs + charts. Payments = the entry "
         "table. Customers = one row per customer. Monthly = invoiced / "
         "collected / pending by invoice month. Overdue = the 50 "
         "earliest-due outstanding invoices (live formulas, re-sorts "
         "itself as rows change). Checks = 10 integrity checks, all "
         "must read PASS.", False),
        ("Data in this demo is fictional (seeded, deterministic) — "
         "replace data/payments.csv contents with the real payment "
         "table and rebuild, or paste real rows straight into "
         "Payments.", False),
    ]
    for i, (text, bold) in enumerate(guide, start=1):
        c = wg.cell(row=i, column=1, value=text)
        c.font = Font(bold=bold, size=12 if i == 1 else 11)
        c.alignment = openpyxl_alignment_wrap()
    wg.column_dimensions["A"].width = 130

    # ---- Payments ----
    wp = wb.create_sheet("Payments")
    headers = ["Customer", "Invoice No", "Invoice Date", "Due Date",
               "Invoice Amount (Rs)", "Amount Paid (Rs)", "Pending (Rs)",
               "Status", "Days Overdue", "Month Key", "Overdue Row",
               "Age Bucket"]
    for col, h in enumerate(headers, start=1):
        wp.cell(row=1, column=col, value=h)
    style_header(wp, 1, len(headers))
    for r, row in enumerate(payments, start=2):
        wp.cell(row=r, column=1, value=row["customer"])
        wp.cell(row=r, column=2, value=row["invoice_no"])
        wp.cell(row=r, column=3, value=dval(row["invoice_date"])).number_format = DATEFMT
        wp.cell(row=r, column=4, value=dval(row["due_date"])).number_format = DATEFMT
        wp.cell(row=r, column=5, value=int(row["amount"])).number_format = MONEY
        wp.cell(row=r, column=6, value=int(row["paid"])).number_format = MONEY
    for r in range(2, last + 1):
        wp.cell(row=r, column=7,
                value=f'=IF($A{r}="","",$E{r}-$F{r})').number_format = MONEY
        wp.cell(row=r, column=8,
                value=(f'=IF($A{r}="","",IF($G{r}<=0,"Paid",'
                       f'IF($D{r}<{ASOF},IF($F{r}>0,"Partial - Overdue",'
                       f'"Unpaid - Overdue"),IF($F{r}>0,"Partial","Unpaid"))))'))
        wp.cell(row=r, column=9,
                value=(f'=IF($A{r}="","",IF(AND($G{r}>0,$D{r}<{ASOF}),'
                       f'{ASOF}-$D{r},0))'))
        wp.cell(row=r, column=10,
                value=f'=IF($A{r}="","",YEAR($C{r})*100+MONTH($C{r}))')
        wp.cell(row=r, column=11,
                value=(f'=IF(AND($A{r}<>"",$G{r}>0,$D{r}<{ASOF}),'
                       f'ROW()-1,"")'))
        wp.cell(row=r, column=12,
                value=(f'=IF($A{r}="","",IF($I{r}>90,"90+",'
                       f'IF($I{r}>60,"61-90",IF($I{r}>30,"31-60",'
                       f'IF($I{r}>0,"1-30","Current")))))'))
    wp.freeze_panes = "A2"
    wp.auto_filter.ref = f"A1:L{last}"
    widths(wp, {"A": 26, "B": 13, "C": 13, "D": 13, "E": 17, "F": 15,
                "G": 13, "H": 18, "I": 13, "J": 11, "K": 12, "L": 11})
    green = PatternFill("solid", fgColor="C6EFCE")
    amber = PatternFill("solid", fgColor="FFEB9C")
    grey = PatternFill("solid", fgColor="D9D9D9")
    red = PatternFill("solid", fgColor="FFC7CE")
    orange = PatternFill("solid", fgColor="FCE4D6")
    wp.conditional_formatting.add(
        f"H2:H{last}", FormulaRule(formula=['$H2="Paid"'], fill=green))
    wp.conditional_formatting.add(
        f"H2:H{last}", FormulaRule(formula=['$H2="Partial"'], fill=amber))
    wp.conditional_formatting.add(
        f"H2:H{last}", FormulaRule(formula=['$H2="Unpaid"'], fill=grey))
    wp.conditional_formatting.add(
        f"H2:H{last}", FormulaRule(formula=['$H2="Partial - Overdue"'], fill=orange))
    wp.conditional_formatting.add(
        f"H2:H{last}", FormulaRule(formula=['$H2="Unpaid - Overdue"'], fill=red))
    wp.conditional_formatting.add(
        f"G2:G{last}", CellIsRule(operator="greaterThan", formula=["0"], fill=red))
    wp.conditional_formatting.add(
        f"I2:I{last}", CellIsRule(operator="greaterThan", formula=["30"], fill=red))
    wp.conditional_formatting.add(
        f"I2:I{last}", CellIsRule(operator="between", formula=["1", "30"], fill=orange))

    # ---- Customers ----
    wc = wb.create_sheet("Customers")
    cheaders = ["Customer", "Invoices", "Invoiced (Rs)", "Paid (Rs)",
                "Pending (Rs)", "Overdue Amount (Rs)", "Overdue Invoices",
                "Collection %"]
    for col, h in enumerate(cheaders, start=1):
        wc.cell(row=1, column=col, value=h)
    style_header(wc, 1, len(cheaders))
    for r, cust in enumerate(customers, start=2):
        wc.cell(row=r, column=1, value=cust)
        wc.cell(row=r, column=2, value=f"=COUNTIF({A},$A{r})")
        wc.cell(row=r, column=3, value=f"=SUMIF({A},$A{r},{E})").number_format = MONEY
        wc.cell(row=r, column=4, value=f"=SUMIF({A},$A{r},{F})").number_format = MONEY
        wc.cell(row=r, column=5, value=f"=SUMIF({A},$A{r},{G})").number_format = MONEY
        wc.cell(row=r, column=6,
                value=(f"=SUMIFS({G},{A},$A{r},{G},\">0\",{D},\"<\"&{ASOF})")
                ).number_format = MONEY
        wc.cell(row=r, column=7,
                value=f"=COUNTIFS({A},$A{r},{G},\">0\",{D},\"<\"&{ASOF})")
        wc.cell(row=r, column=8,
                value=f"=IF($C{r}=0,0,$D{r}/$C{r})").number_format = PCT
    tr = cust_last + 1
    wc.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col in (2, 3, 4, 5, 6, 7):
        c = wc.cell(row=tr, column=col,
                    value=f"=SUM({chr(64 + col)}2:{chr(64 + col)}{cust_last})")
        c.font = Font(bold=True)
        if col in (3, 4, 5, 6):
            c.number_format = MONEY
    wc.cell(row=tr, column=8,
            value=f"=IF($C{tr}=0,0,$D{tr}/$C{tr})").number_format = PCT
    wc.cell(row=tr, column=8).font = Font(bold=True)
    wc.freeze_panes = "A2"
    widths(wc, {"A": 26, "B": 10, "C": 15, "D": 13, "E": 14,
                "F": 18, "G": 16, "H": 12})

    # ---- Monthly ----
    wm = wb.create_sheet("Monthly")
    mheaders = ["Month Key", "Month", "Invoiced (Rs)", "Collected (Rs)",
                "Pending (Rs)", "Collection %", "Cumulative Collected (Rs)",
                "Invoices"]
    for col, h in enumerate(mheaders, start=1):
        wm.cell(row=1, column=col, value=h)
    style_header(wm, 1, len(mheaders))
    for i, (key, label) in enumerate(MONTHS, start=2):
        wm.cell(row=i, column=1, value=key)
        wm.cell(row=i, column=2, value=label)
        wm.cell(row=i, column=3,
                value=f"=SUMIFS({E},{J},$A{i})").number_format = MONEY
        wm.cell(row=i, column=4,
                value=f"=SUMIFS({F},{J},$A{i})").number_format = MONEY
        wm.cell(row=i, column=5,
                value=f"=SUMIFS({G},{J},$A{i})").number_format = MONEY
        wm.cell(row=i, column=6,
                value=f"=IF($C{i}=0,0,$D{i}/$C{i})").number_format = PCT
        prev = f"$G{i - 1}" if i > 2 else "0"
        wm.cell(row=i, column=7,
                value=f"={prev}+$D{i}").number_format = MONEY
        wm.cell(row=i, column=8, value=f"=COUNTIF({J},$A{i})")
    mtot = 2 + len(MONTHS)
    wm.cell(row=mtot, column=1, value="TOTAL").font = Font(bold=True)
    wm.cell(row=mtot, column=2, value="Apr-Oct 2026").font = Font(bold=True)
    for col, letter in ((3, "C"), (4, "D"), (5, "E"), (8, "H")):
        c = wm.cell(row=mtot, column=col,
                    value=f"=SUM({letter}2:{letter}{mtot - 1})")
        c.font = Font(bold=True)
        if col in (3, 4, 5):
            c.number_format = MONEY
    wm.cell(row=mtot, column=6,
            value=f"=IF($C{mtot}=0,0,$D{mtot}/$C{mtot})").number_format = PCT
    wm.cell(row=mtot, column=6).font = Font(bold=True)
    wm.cell(row=mtot, column=7, value=f"=$G{mtot - 1}").number_format = MONEY
    wm.cell(row=mtot, column=7).font = Font(bold=True)
    widths(wm, {"A": 11, "B": 11, "C": 15, "D": 15, "E": 14, "F": 12,
                "G": 24, "H": 10})

    # ---- Overdue (live pull of the earliest-due outstanding invoices) ----
    wo = wb.create_sheet("Overdue")
    oheaders = ["Customer", "Invoice No", "Due Date", "Pending (Rs)",
                "Days Overdue", "Status", "Age Bucket"]
    for col, h in enumerate(oheaders, start=1):
        wo.cell(row=1, column=col, value=h)
    style_header(wo, 1, len(oheaders))
    src_cols = ["A", "B", "D", "G", "I", "H", "L"]
    for r in range(2, 2 + OVERDUE_ROWS):
        k = r - 1  # 1-based rank among overdue rows
        for col, srccol in enumerate(src_cols, start=1):
            src = f"Payments!${srccol}$2:${srccol}${last}"
            c = wo.cell(row=r, column=col,
                        value=f'=IFERROR(INDEX({src},SMALL({K},{k})),"")')
            if col == 3:
                c.number_format = DATEFMT
            if col == 4:
                c.number_format = MONEY
    wo.cell(row=2 + OVERDUE_ROWS + 1, column=1,
            value=("Live list: up to 50 earliest-due outstanding invoices, "
                   "in Payments row order (due-date order). Blank rows mean "
                   "fewer than 50 invoices are overdue."))
    wo.freeze_panes = "A2"
    widths(wo, {"A": 26, "B": 13, "C": 13, "D": 14, "E": 13, "F": 18, "G": 11})

    # ---- Dashboard (created after source sheets; charts added last) ----
    wd = wb.create_sheet("Dashboard")
    wd["A1"] = "Payment Tracker Dashboard — Sharma Traders, New Delhi"
    wd["A1"].font = Font(bold=True, size=13)
    wd["A2"] = "As of:"
    wd["B2"] = f"={ASOF}"
    wd["B2"].number_format = DATEFMT
    wd["A3"] = "KPI"
    wd["B3"] = "Value"
    style_header(wd, 3, 2)
    kpis = [
        ("Total invoiced (Rs)", f"=SUM({E})", MONEY),
        ("Total collected (Rs)", f"=SUM({F})", MONEY),
        ("Total pending (Rs)", f"=SUM({G})", MONEY),
        ("Collection rate", "=IF($B$4=0,0,$B$5/$B$4)", PCT),
        ("Invoices", f"=COUNTA({A})", "0"),
        ("Paid invoices", f'=COUNTIF({H},"Paid")', "0"),
        ("Overdue invoices",
         f'=COUNTIFS({G},">0",{D},"<"&{ASOF})', "0"),
        ("Overdue amount (Rs)",
         f'=SUMIFS({G},{G},">0",{D},"<"&{ASOF})', MONEY),
        ("Customers", f"=COUNTA(Customers!$A$2:$A${cust_last})", "0"),
        ("Avg days overdue (overdue only)",
         f'=IF($B$10=0,0,AVERAGEIF({I},">0"))', "0.0"),
        ("Partially paid invoices",
         f'=COUNTIF({H},"Partial")+COUNTIF({H},"Partial - Overdue")', "0"),
    ]
    for i, (label, formula, fmt) in enumerate(kpis, start=4):
        wd.cell(row=i, column=1, value=label)
        c = wd.cell(row=i, column=2, value=formula)
        c.number_format = fmt
    # Status mix table (rows 16-22): header row 16, data 17-21.
    wd["A16"] = "Status"
    wd["B16"] = "Invoices"
    style_header(wd, 16, 2)
    for i, status in enumerate(STATUSES, start=17):
        wd.cell(row=i, column=1, value=status)
        wd.cell(row=i, column=2, value=f'=COUNTIF({H},$A{i})')
    wd["A22"] = "TOTAL"
    wd["A22"].font = Font(bold=True)
    wd["B22"] = "=SUM($B$17:$B$21)"
    wd["B22"].font = Font(bold=True)
    widths(wd, {"A": 34, "B": 18, "C": 4, "D": 18, "E": 18})

    # ---- Checks ----
    wk = wb.create_sheet("Checks")
    for col, h in enumerate(["Check", "Difference", "Expected", "Result"],
                            start=1):
        wk.cell(row=1, column=col, value=h)
    style_header(wk, 1, 4)
    checks = [
        ("Dashboard invoiced = Monthly total invoiced",
         "=Dashboard!$B$4-Monthly!$C$9"),
        ("Dashboard collected = Monthly total collected",
         "=Dashboard!$B$5-Monthly!$D$9"),
        ("Dashboard pending = Monthly total pending",
         "=Dashboard!$B$6-Monthly!$E$9"),
        ("Pending = invoiced - collected",
         "=Dashboard!$B$6-(Dashboard!$B$4-Dashboard!$B$5)"),
        ("Customers invoiced = Dashboard invoiced",
         f"=SUM(Customers!$C$2:$C${cust_last})-Dashboard!$B$4"),
        ("Customers pending = Dashboard pending",
         f"=SUM(Customers!$E$2:$E${cust_last})-Dashboard!$B$6"),
        ("Payments pending sum = Dashboard pending",
         f"=SUM({G})-Dashboard!$B$6"),
        ("Customer invoice counts = Dashboard invoice count",
         f"=SUM(Customers!$B$2:$B${cust_last})-Dashboard!$B$8"),
        ("Status counts sum = invoice count",
         "=SUM(Dashboard!$B$17:$B$21)-Dashboard!$B$8"),
        ("No negative pending amounts",
         f'=COUNTIF({G},"<0")'),
    ]
    for i, (label, formula) in enumerate(checks, start=2):
        wk.cell(row=i, column=1, value=label)
        c = wk.cell(row=i, column=2, value=formula)
        c.number_format = MONEY
        wk.cell(row=i, column=3, value=0)
        wk.cell(row=i, column=4,
                value=f'=IF(ABS($B{i})<0.005,"PASS","FAIL")')
    widths(wk, {"A": 48, "B": 14, "C": 10, "D": 10})

    # ---- Charts (only now that every source sheet exists) ----
    bar = BarChart()
    bar.type = "col"
    bar.title = "Invoiced vs collected by month"
    bar.y_axis.title = "Rs"
    bar_data = Reference(wm, min_col=3, min_row=1, max_row=8, max_col=4)
    bar_cats = Reference(wm, min_col=2, min_row=2, max_row=8)
    bar.add_data(bar_data, titles_from_data=True)
    bar.set_categories(bar_cats)
    bar.height = 7.5
    bar.width = 14
    wd.add_chart(bar, "D3")

    pie = PieChart()
    pie.title = "Invoices by status"
    pie_data = Reference(wd, min_col=2, min_row=16, max_row=21)
    pie_cats = Reference(wd, min_col=1, min_row=17, max_row=21)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)
    pie.height = 7.5
    pie.width = 12
    wd.add_chart(pie, "D18")

    line = LineChart()
    line.title = "Cumulative collected"
    line.y_axis.title = "Rs"
    line_data = Reference(wm, min_col=7, min_row=1, max_row=8)
    line.add_data(line_data, titles_from_data=True)
    line.set_categories(bar_cats)
    line.height = 7.5
    line.width = 14
    wd.add_chart(line, "D33")

    # Sensible tab order and active sheet.
    wb._sheets = [wb["Guide"], wb["Dashboard"], wb["Payments"],
                  wb["Customers"], wb["Monthly"], wb["Overdue"],
                  wb["Assumptions"], wb["Checks"]]
    wb.active = wb["Dashboard"]
    wb.save(OUT)
    print(f"wrote {OUT} ({n} invoices, {ncust} customers, "
          f"Payments formulas to row {last})")


def openpyxl_alignment_wrap():
    from openpyxl.styles import Alignment
    return Alignment(wrap_text=True, vertical="top")


if __name__ == "__main__":
    main()
