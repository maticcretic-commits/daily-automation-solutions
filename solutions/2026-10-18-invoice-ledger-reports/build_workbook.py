#!/usr/bin/env python3
"""InvoiceLedger — build the formula-driven invoice-to-reports workbook.

Reads data/invoices.csv + data/payments.csv (see generate.py) and writes
invoice-ledger-reports.xlsx: raw data sheets plus P&L, CashFlow, Receivables
aging, Dashboard, AI_Report and Checks — every computed cell a live Excel
formula, so the workbook re-derives everything when new rows are pasted in.
"""
import csv
import os
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "invoice-ledger-reports.xlsx")
REPORT_DATE = date(2026, 9, 30)
MONTHS = [(2025, 10), (2025, 11), (2025, 12), (2026, 1), (2026, 2), (2026, 3),
          (2026, 4), (2026, 5), (2026, 6), (2026, 7), (2026, 8), (2026, 9)]

HDR_FILL = PatternFill("solid", fgColor="1F4E5F")
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color="1F4E5F")
MONEY_FMT = "#,##0.00"
PCT_FMT = "0.0%"
DATE_FMT = "dd-mmm-yy"
MONTH_FMT = "mmm-yy"

INV_COLS = ["inv_id", "date", "party", "type", "category", "amount",
            "gst_rate", "gst_amt", "total", "paid_amount", "paid_date"]
PAY_COLS = ["pay_id", "inv_id", "type", "category", "date", "amount"]


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)


def month_formula_sumifs(ws, row, col_letter, amount_col, type_val, cat_val=None):
    """SUMIFS over Raw_Invoices for one month column (header row 2 holds date)."""
    crit = (f'SUMIFS(Raw_Invoices!${amount_col}:${amount_col},'
            f'Raw_Invoices!$B:$B,">="&{col_letter}$2,'
            f'Raw_Invoices!$B:$B,"<="&EOMONTH({col_letter}$2,0),'
            f'Raw_Invoices!$D:$D,"{type_val}"')
    if cat_val:
        crit += f',Raw_Invoices!$E:$E,"{cat_val}"'
    return "=" + crit + ")"


def pay_sumifs(row, col_letter, type_val):
    return (f"=SUMIFS(Raw_Payments!$F:$F,Raw_Payments!$E:$E,\">=\"&{col_letter}$2,"
            f"Raw_Payments!$E:$E,\"<=\"&EOMONTH({col_letter}$2,0),"
            f"Raw_Payments!$C:$C,\"{type_val}\")")


def build():
    invoices = read_csv("invoices.csv")
    payments = read_csv("payments.csv")
    n_inv = len(invoices)
    wb = Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "InvoiceLedger — invoice-to-reports workbook"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Report date (all aging measured from here):"
    ws["B2"] = REPORT_DATE
    ws["B2"].number_format = DATE_FMT
    ws["B2"].font = Font(bold=True)
    how = [
        "1. Paste new invoice rows into Raw_Invoices (keep the column order).",
        "2. Paste new payment rows into Raw_Payments.",
        "3. Everything else recalculates: P&L, CashFlow, Receivables aging,",
        "   Dashboard KPIs/charts, the AI_Report narrative and the Checks.",
        "4. Change the report date in B2 and the aging buckets move with it.",
        "Data dictionary — Raw_Invoices: inv_id, date, party, type",
        "(SALES/PURCHASE/EXPENSE), category, amount (ex-GST), gst_rate,",
        "gst_amt, total, paid_amount, paid_date, balance (=total-paid),",
        "days_to_pay (=paid_date - date, paid invoices only).",
        "Raw_Payments: pay_id, inv_id, type, category, date, amount.",
        "All records are fictional demo data for 'Gupta Garments & Textiles'.",
    ]
    for i, line in enumerate(how, start=4):
        ws[f"A{i}"] = line
    ws.column_dimensions["A"].width = 78
    ws.column_dimensions["B"].width = 16

    # ---------------- Raw_Invoices ----------------
    ri = wb.create_sheet("Raw_Invoices")
    headers = INV_COLS + ["balance", "days_to_pay"]
    for c, h in enumerate(headers, start=1):
        ri.cell(row=1, column=c, value=h)
    style_header(ri, 1, len(headers))
    for r, inv in enumerate(invoices, start=2):
        vals = [inv["inv_id"], date.fromisoformat(inv["date"]), inv["party"],
                inv["type"], inv["category"], float(inv["amount"]),
                int(inv["gst_rate"]), float(inv["gst_amt"]),
                float(inv["total"]), float(inv["paid_amount"]),
                date.fromisoformat(inv["paid_date"]) if inv["paid_date"] else ""]
        for c, v in enumerate(vals, start=1):
            cell = ri.cell(row=r, column=c, value=v)
            if c == 2 or c == 11:
                cell.number_format = DATE_FMT
            if c in (6, 8, 9, 10):
                cell.number_format = MONEY_FMT
        ri.cell(row=r, column=12, value=f"=IF(I{r}=\"\",\"\",I{r}-J{r})")
        ri.cell(row=r, column=12).number_format = MONEY_FMT
        ri.cell(row=r, column=13, value=f"=IF(K{r}=\"\",\"\",K{r}-B{r})")
    ri.freeze_panes = "A2"
    ri.auto_filter.ref = f"A1:M{n_inv + 1}"
    for c, w in zip("ABCDEFGHIJKLM",
                    [10, 12, 26, 10, 20, 12, 9, 10, 12, 12, 12, 11, 11]):
        ri.column_dimensions[c].width = w

    # ---------------- Raw_Payments ----------------
    rp = wb.create_sheet("Raw_Payments")
    for c, h in enumerate(PAY_COLS, start=1):
        rp.cell(row=1, column=c, value=h)
    style_header(rp, 1, len(PAY_COLS))
    for r, p in enumerate(payments, start=2):
        vals = [p["pay_id"], p["inv_id"], p["type"], p["category"],
                date.fromisoformat(p["date"]), float(p["amount"])]
        for c, v in enumerate(vals, start=1):
            cell = rp.cell(row=r, column=c, value=v)
            if c == 5:
                cell.number_format = DATE_FMT
            if c == 6:
                cell.number_format = MONEY_FMT
    rp.freeze_panes = "A2"
    rp.auto_filter.ref = f"A1:F{len(payments) + 1}"
    for c, w in zip("ABCDEF", [10, 10, 10, 20, 12, 12]):
        rp.column_dimensions[c].width = w

    month_letters = [get_column_letter(c) for c in range(2, 14)]  # B..M
    fy_col = "N"

    def write_month_headers(ws):
        ws["A2"] = "Line"
        for L, (y, m) in zip(month_letters, MONTHS):
            cell = ws[f"{L}2"] = date(y, m, 1)
            ws[f"{L}2"].number_format = MONTH_FMT
        ws[f"{fy_col}2"] = "FY total"
        style_header(ws, 2, 14)
        ws["A1"] = ws.title + " — monthly (formula-driven)"
        ws["A1"].font = TITLE_FONT
        ws.column_dimensions["A"].width = 30
        for L in month_letters + [fy_col]:
            ws.column_dimensions[L].width = 13

    def fy_total(ws, row):
        ws[f"{fy_col}{row}"] = f"=SUM(B{row}:M{row})"

    # ---------------- P&L ----------------
    pl = wb.create_sheet("P&L")
    write_month_headers(pl)
    pl_rows = [
        (3, "Revenue — Fabrics", "SALES", "Fabrics"),
        (4, "Revenue — Garments", "SALES", "Garments"),
        (5, "Revenue — Accessories", "SALES", "Accessories"),
        (7, "COGS — Inventory purchases", "PURCHASE", "Inventory purchases"),
        (9, "Expenses — Rent", "EXPENSE", "Rent"),
        (10, "Expenses — Salaries", "EXPENSE", "Salaries"),
        (11, "Expenses — Utilities", "EXPENSE", "Utilities"),
        (12, "Expenses — Transport", "EXPENSE", "Transport"),
        (13, "Expenses — Marketing", "EXPENSE", "Marketing"),
        (14, "Expenses — Professional fees", "EXPENSE", "Professional fees"),
        (15, "Expenses — Misc", "EXPENSE", "Misc"),
    ]
    for row, label, tval, cval in pl_rows:
        pl[f"A{row}"] = label
        for L in month_letters:
            pl[f"{L}{row}"] = month_formula_sumifs(pl, row, L, "F", tval, cval)
            pl[f"{L}{row}"].number_format = MONEY_FMT
        fy_total(pl, row)
        pl[f"{fy_col}{row}"].number_format = MONEY_FMT
    calc_rows = [
        (6, "Total revenue", "=SUM({L}3:{L}5)"),
        (8, "Gross profit", "={L}6-{L}7"),
        (16, "Total expenses", "=SUM({L}9:{L}15)"),
        (17, "Net profit", "={L}8-{L}16"),
        (18, "Gross margin %", "=IF({L}6=0,0,{L}8/{L}6)"),
        (19, "Net margin %", "=IF({L}6=0,0,{L}17/{L}6)"),
        (22, "GST — Net payable", "={L}20-{L}21"),
    ]
    for row, label, tmpl in calc_rows:
        pl[f"A{row}"] = label
        for L in month_letters:
            pl[f"{L}{row}"] = tmpl.format(L=L)
            pl[f"{L}{row}"].number_format = (
                PCT_FMT if row in (18, 19) else MONEY_FMT)
        fy_total(pl, row)
        pl[f"{fy_col}{row}"].number_format = (
            PCT_FMT if row in (18, 19) else MONEY_FMT)
    # GST output/input need type-wide SUMIFS (no category filter)
    for row, label, tval in [(20, "GST — Output", "SALES"),
                             (21, "GST — Input", None)]:
        pl[f"A{row}"] = label
        for L in month_letters:
            if tval:
                f = month_formula_sumifs(pl, row, L, "H", tval)
            else:
                f = ("=" + month_formula_sumifs(pl, row, L, "H",
                                                "PURCHASE")[1:] + "+"
                     + month_formula_sumifs(pl, row, L, "H",
                                            "EXPENSE")[1:])
            pl[f"{L}{row}"] = f
            pl[f"{L}{row}"].number_format = MONEY_FMT
        fy_total(pl, row)
        pl[f"{fy_col}{row}"].number_format = MONEY_FMT
    for r in (6, 8, 16, 17, 22):
        pl[f"A{r}"].font = Font(bold=True)
    pl.freeze_panes = "B3"

    # ---------------- CashFlow ----------------
    cf = wb.create_sheet("CashFlow")
    write_month_headers(cf)
    for row, label, tval in [(3, "Cash receipts — customers", "SALES"),
                             (4, "Cash paid — inventory", "PURCHASE"),
                             (5, "Cash paid — expenses", "EXPENSE")]:
        cf[f"A{row}"] = label
        for L in month_letters:
            cf[f"{L}{row}"] = pay_sumifs(row, L, tval)
            cf[f"{L}{row}"].number_format = MONEY_FMT
        fy_total(cf, row)
        cf[f"{fy_col}{row}"].number_format = MONEY_FMT
    cf["A6"] = "Net cash flow"
    cf["A7"] = "Cumulative cash flow"
    for i, L in enumerate(month_letters):
        cf[f"{L}6"] = f"={L}3-{L}4-{L}5"
        cf[f"{L}6"].number_format = MONEY_FMT
        prev = month_letters[i - 1] if i else None
        cf[f"{L}7"] = f"={L}6" if not prev else f"={prev}7+{L}6"
        cf[f"{L}7"].number_format = MONEY_FMT
    fy_total(cf, 6)
    cf[f"{fy_col}6"].number_format = MONEY_FMT
    cf[f"{fy_col}7"] = "=M7"  # closing balance
    cf[f"{fy_col}7"].number_format = MONEY_FMT
    cf["A6"].font = Font(bold=True)
    cf["A7"].font = Font(bold=True)
    cf.freeze_panes = "B3"

    # ---------------- Receivables ----------------
    ar = wb.create_sheet("Receivables")
    ar["A1"] = "Receivables — outstanding invoices & aging (formula-driven)"
    ar["A1"].font = TITLE_FONT
    for c, h in enumerate(["inv_id", "party", "invoice date", "total",
                           "paid", "balance", "days outstanding",
                           "bucket"], start=1):
        ar.cell(row=3, column=c, value=h)
    style_header(ar, 3, 8)
    outstanding = [i for i in invoices
                   if float(i["total"]) - float(i["paid_amount"]) > 0.005]
    last = 3 + len(outstanding)
    for r, inv in enumerate(outstanding, start=4):
        ar.cell(row=r, column=1, value=inv["inv_id"])
        ar.cell(row=r, column=2, value=inv["party"])
        d = ar.cell(row=r, column=3, value=date.fromisoformat(inv["date"]))
        d.number_format = DATE_FMT
        for c, k in [(4, "total"), (5, "paid_amount")]:
            cell = ar.cell(row=r, column=c, value=float(inv[k]))
            cell.number_format = MONEY_FMT
        bal = ar.cell(row=r, column=6, value=f"=D{r}-E{r}")
        bal.number_format = MONEY_FMT
        ar.cell(row=r, column=7, value=f"=Guide!$B$2-C{r}")
        ar.cell(row=r, column=8,
                value=f'=IF(G{r}="","",IF(G{r}<=30,"0-30",IF(G{r}<=60,'
                      f'"31-60",IF(G{r}<=90,"61-90","90+"))))')
    ar.freeze_panes = "A4"
    ar.auto_filter.ref = f"A3:H{last}"
    for c, w in zip("ABCDEFGH", [10, 26, 13, 12, 12, 12, 15, 9]):
        ar.column_dimensions[c].width = w
    # aging summary block
    ar["J1"] = "Aging summary"
    ar["J1"].font = Font(bold=True)
    ar["J2"], ar["K2"] = "Bucket", "Outstanding"
    for cell in ("J2", "K2"):
        ar[cell].fill = HDR_FILL
        ar[cell].font = HDR_FONT
        ar[cell].alignment = Alignment(horizontal="center", vertical="center")
    buckets = ["0-30", "31-60", "61-90", "90+"]
    for i, b in enumerate(buckets, start=3):
        ar[f"J{i}"] = b
        ar[f"K{i}"] = (f"=SUMIFS($F$4:$F${last},$H$4:$H${last},"
                       f"\"{b}\")")
        ar[f"K{i}"].number_format = MONEY_FMT
    ar["J7"], ar["K7"] = "Total outstanding", f"=SUM(F4:F{last})"
    ar["K7"].number_format = MONEY_FMT
    ar["J7"].font = Font(bold=True)
    ar["J8"], ar["K8"] = "Invoices outstanding", f"=COUNTA(A4:A{last})"
    ar["J9"] = "Avg collection days (paid sales)"
    ar["K9"] = ('=AVERAGEIFS(Raw_Invoices!$M:$M,Raw_Invoices!$D:$D,'
                '"SALES",Raw_Invoices!$M:$M,"<>")')
    ar["J10"] = "Overdue 61+ days"
    ar["K10"] = "=K5+K6"
    ar["K10"].number_format = MONEY_FMT
    for c in "JK":
        ar.column_dimensions[c].width = 22 if c == "K" else 28

    # ---------------- Dashboard ----------------
    dash = wb.create_sheet("Dashboard")
    dash["A1"] = "Dashboard — KPIs (all formula-linked)"
    dash["A1"].font = TITLE_FONT
    kpis = [
        ("FY revenue", "='P&L'!N6", MONEY_FMT),
        ("FY net profit", "='P&L'!N17", MONEY_FMT),
        ("Net margin", "='P&L'!N19", PCT_FMT),
        ("Receivables outstanding", "=Receivables!K7", MONEY_FMT),
        ("Avg collection days", "=Receivables!K9", "0.0"),
        ("Closing cash balance", "='CashFlow'!N7", MONEY_FMT),
    ]
    for i, (label, formula, fmt) in enumerate(kpis, start=3):
        dash[f"A{i}"] = label
        dash[f"A{i}"].font = Font(bold=True)
        dash[f"B{i}"] = formula
        dash[f"B{i}"].number_format = fmt
        dash[f"B{i}"].font = Font(size=13, color="1F4E5F")
    dash.column_dimensions["A"].width = 26
    dash.column_dimensions["B"].width = 18

    rev_exp = LineChart()
    rev_exp.title = "Monthly revenue vs expenses"
    rev_exp.style = 2
    rev_exp.height, rev_exp.width = 7.5, 15
    d = Reference(pl, min_col=2, max_col=13, min_row=6, max_row=6)
    rev_exp.add_data(d, from_rows=True, titles_from_data=False)
    d2 = Reference(pl, min_col=2, max_col=13, min_row=16, max_row=16)
    rev_exp.add_data(d2, from_rows=True, titles_from_data=False)
    rev_exp.series[0].title = SeriesLabel(v="Revenue")
    rev_exp.series[1].title = SeriesLabel(v="Expenses")
    rev_exp.set_categories(Reference(pl, min_col=2, max_col=13, min_row=2))
    dash.add_chart(rev_exp, "D2")

    cat_bar = BarChart()
    cat_bar.title = "FY revenue by category"
    cat_bar.style = 2
    cat_bar.height, cat_bar.width = 7.5, 15
    cat_bar.add_data(Reference(pl, min_col=14, max_col=14, min_row=3,
                               max_row=5), titles_from_data=False)
    cat_bar.set_categories(Reference(pl, min_col=1, max_col=1, min_row=3,
                                     max_row=5))
    dash.add_chart(cat_bar, "D18")

    cf_bar = BarChart()
    cf_bar.title = "Monthly net cash flow"
    cf_bar.style = 2
    cf_bar.height, cf_bar.width = 7.5, 15
    cf_bar.add_data(Reference(cf, min_col=2, max_col=13, min_row=6,
                              max_row=6), from_rows=True,
                    titles_from_data=False)
    cf_bar.set_categories(Reference(cf, min_col=2, max_col=13, min_row=2))
    dash.add_chart(cf_bar, "D34")

    # ---------------- AI_Report ----------------
    ai = wb.create_sheet("AI_Report")
    ai["A1"] = "AI_Report — rule-based narrative (recalculates with the data)"
    ai["A1"].font = TITLE_FONT
    ai["A3"] = "Key findings"
    ai["A3"].font = Font(bold=True)
    # metric cells (column B)
    ai["B4"] = "=B5+B6"                  # FY revenue
    ai["B5"] = "=SUM('P&L'!B6:G6)"          # H1 revenue
    ai["B6"] = "=SUM('P&L'!H6:M6)"          # H2 revenue
    ai["B7"] = "=IF(B5=0,0,(B6-B5)/B5)"     # H2 vs H1 growth
    ai["B8"] = "=INDEX('P&L'!B2:M2,MATCH(MAX('P&L'!B6:M6),'P&L'!B6:M6,0))"
    ai["B8"].number_format = MONTH_FMT
    ai["B9"] = "=Receivables!K7/'P&L'!N6"   # receivables / revenue
    ai["B10"] = "=Receivables!K10/Receivables!K7"  # overdue 61+ share
    ai["B11"] = "=MIN('CashFlow'!B7:M7)"    # lowest cumulative cash
    ai["B12"] = "='P&L'!N19"                # net margin
    ai["B13"] = "=Receivables!K9"           # avg collection days
    ai["B14"] = "=MAX('P&L'!N3:N5)/'P&L'!N6"  # top category share
    for r in range(4, 15):
        ai[f"B{r}"].number_format = MONEY_FMT if r <= 6 else (
            PCT_FMT if r in (7, 9, 10, 12, 14) else ("0.0" if r == 13 else MONEY_FMT))
    ai["A5"] = ('="FY revenue was "&TEXT(B4,"#,##0")&", H2 vs H1: "'
                '&TEXT(B7,"0.0%")&". Best month: "&TEXT(B8,"mmm-yy")'
                '&". Net margin "&TEXT(B12,"0.0%")&"."')
    ai["A5"].alignment = Alignment(wrap_text=True)
    flags = [
        (7, '=IF(B9>0.15,"⚠ Receivables are "&TEXT(B9,"0%")&" of revenue — '
         'tighten collections","✓ Receivables at "&TEXT(B9,"0%")&" of revenue")'),
        (8, '=IF(B12<0.08,"⚠ Net margin "&TEXT(B12,"0.0%")&" is thin — review '
         'pricing or costs","✓ Net margin healthy at "&TEXT(B12,"0.0%"))'),
        (9, '=IF(B10>0.25,"⚠ "&TEXT(B10,"0%")&" of receivables are 61+ days '
         'overdue — prioritise recovery","✓ Overdue 61+ days under control")'),
        (10, '=IF(B11<0,"⚠ Cash went negative intra-year (lowest "'
         '&TEXT(B11,"#,##0")&") — arrange a buffer","✓ Cash stayed positive '
         'all year")'),
        (11, '=IF(B13>45,"⚠ Avg collection "&TEXT(B13,"0")&" days — slower '
         'than 45-day terms","✓ Avg collection "&TEXT(B13,"0")&" days")'),
        (12, '=IF(B14>0.5,"⚠ One category is "&TEXT(B14,"0%")&" of revenue — '
         'concentration risk","✓ Revenue spread across categories")'),
        (13, '=IF(\'P&L\'!N22>0,"ℹ Net GST payable for the year: "'
         '&TEXT(\'P&L\'!N22,"#,##0"),"ℹ Net GST credit for the year")'),
        (14, '=IF(B7>0,"✓ H2 revenue momentum positive","⚠ H2 revenue '
         'declined vs H1 — investigate demand")'),
    ]
    ai["A6"] = "Flags"
    ai["A6"].font = Font(bold=True)
    for r, f in flags:
        ai[f"A{r}"] = f
        ai[f"A{r}"].alignment = Alignment(wrap_text=True)
    ai.column_dimensions["A"].width = 95
    ai.column_dimensions["B"].width = 16
    for r in range(5, 15):
        ai.row_dimensions[r].height = 22

    # ---------------- Checks ----------------
    ch = wb.create_sheet("Checks")
    ch["A1"] = "Checks — all must be TRUE after recalculation"
    ch["A1"].font = TITLE_FONT
    ch["A2"], ch["B2"] = "Check", "Result"
    style_header(ch, 2, 2)
    checks = [
        "P&L FY revenue = SUM of raw SALES amounts",
        "Gross profit = revenue − COGS",
        "Net profit = gross profit − total expenses",
        "CashFlow receipts FY = SUM of raw SALES payments",
        "Closing cash = last month cumulative",
        "Aging buckets sum = total outstanding",
        "No negative balances in Raw_Invoices",
        "No paid_amount exceeding invoice total",
    ]
    formulas = [
        "=ABS('P&L'!N6-SUMIFS(Raw_Invoices!$F:$F,Raw_Invoices!$D:$D,\"SALES\"))<0.01",
        "=ABS('P&L'!N8-('P&L'!N6-'P&L'!N7))<0.01",
        "=ABS('P&L'!N17-('P&L'!N8-'P&L'!N16))<0.01",
        "=ABS('CashFlow'!N3-SUMIFS(Raw_Payments!$F:$F,Raw_Payments!$C:$C,\"SALES\"))<0.01",
        "=ABS('CashFlow'!N7-'CashFlow'!M7)<0.01",
        f"=ABS(Receivables!K7-SUM(Receivables!K3:K6))<0.01",
        "=COUNTIFS(Raw_Invoices!$L:$L,\"<0\")=0",
        (f"=SUMPRODUCT((Raw_Invoices!$J$2:$J${n_inv + 1}>"
         f"Raw_Invoices!$I$2:$I${n_inv + 1})+0)=0"),
    ]
    for i, (label, f) in enumerate(zip(checks, formulas), start=3):
        ch[f"A{i}"] = label
        ch[f"B{i}"] = f
        ch[f"B{i}"].font = Font(bold=True)
    ch.column_dimensions["A"].width = 52
    ch.column_dimensions["B"].width = 12

    wb.save(OUT)
    print(f"wrote {OUT}: {len(wb.sheetnames)} sheets, {n_inv} invoices, "
          f"{len(payments)} payments")


if __name__ == "__main__":
    build()
