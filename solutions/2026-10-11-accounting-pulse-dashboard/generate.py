#!/usr/bin/env python3
"""AccountingPulse - Excel accounting-automation factory.

Turns raw accounting CSVs (A/P bills, A/R invoices, payroll register, bank
statement) into a finance workbook with LIVE Excel formulas - no hard-coded
numbers anywhere:

  Raw_AP     vendor bills + formula columns: days overdue, aging bucket,
             outstanding amount
  Raw_AR     customer invoices + same formula columns (+ in-transit flag)
  Payroll    salary register + net-pay formula
  Bank       bank statement + net-amount formula
  Bank_Rec   bank reconciliation: adjusted bank vs adjusted book, live check
  AP_Aging   payable aging buckets (live SUMIFS) + totals + integrity check
  AR_Aging   receivable aging buckets (live SUMIFS) + totals + integrity check
  P&L        monthly profit & loss, Apr-26 .. Sep-26, all formulas
  Dashboard  inputs (as-of date, opening cash), KPIs + 3 native Excel charts

The refresh routine: append new rows to the CSVs and re-run this script, or
paste new rows into the Raw_* sheets and copy the formula columns down.
Everything downstream (aging, statements, KPIs, charts) is formula-driven.

Usage:
    python generate.py [--data DIR] [--out FILE]

Sample data in data/ covers Apr-Sep 2026 for fictional
'Rajesh Kumar & Sons, wholesale traders, Patna' (INR).
"""
import argparse
import csv
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

HERE = Path(__file__).resolve().parent
PERIOD_START = dt.date(2026, 4, 1)
N_MONTHS = 6
MONTHS = [dt.date(2026, 4 + i, 1) for i in range(N_MONTHS)]  # Apr..Sep 2026
AS_OF = dt.date(2026, 9, 30)
BOOK_OPEN = 500000

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
MONEY_FMT = '#,##0'
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BUCKETS = ["Current", "1-30", "31-60", "61-90", "90+"]


def load_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def add_named(wb, name, ref):
    wb.defined_names.add(DefinedName(name, attr_text=ref))


def bucket_formula(day_col, paid_col):
    """Aging-bucket formula given the day-overdue and paid-date column letters."""
    return (f'=IF({paid_col}{{r}}<>"","Paid",IF({day_col}{{r}}<=0,"Current",'
            f'IF({day_col}{{r}}<=30,"1-30",IF({day_col}{{r}}<=60,"31-60",'
            f'IF({day_col}{{r}}<=90,"61-90","90+")))))')


def build(data_dir, out_path):
    ap = load_csv(data_dir / "ap_ledger.csv")
    ar = load_csv(data_dir / "ar_ledger.csv")
    pay = load_csv(data_dir / "payroll.csv")
    st = load_csv(data_dir / "bank_statement.csv")
    n_ap, n_ar, n_pay, n_st = len(ap), len(ar), len(pay), len(st)

    wb = Workbook()

    # ------------------------------------------------- Raw_AP
    ws = wb.active
    ws.title = "Raw_AP"
    headers = ["Bill ID", "Vendor", "Date", "Due Date", "Amount", "Paid Date",
               "Pay Mode", "Cleared", "Category", "Days Overdue",
               "Aging Bucket", "Outstanding"]
    ws.append(headers)
    for i, r in enumerate(ap, start=2):
        ws.append([r["bill_id"], r["vendor"], dt.date.fromisoformat(r["date"]),
                   dt.date.fromisoformat(r["due_date"]), float(r["amount"]),
                   dt.date.fromisoformat(r["paid_date"]) if r["paid_date"] else None,
                   r["pay_mode"], int(r["cleared"]), r["category"],
                   f"=IF(F{i}=\"\",AsOf-D{i},F{i}-D{i})",
                   bucket_formula("J", "F").format(r=i),
                   f"=IF(F{i}=\"\",E{i},0)"])
    style_header(ws, len(headers))
    for col, w in zip("ABCDEFGHIJKL", [10, 18, 12, 12, 14, 12, 10, 9, 12, 13, 13, 14]):
        ws.column_dimensions[col].width = w
    for c in ws["E"][1:]:
        c.number_format = MONEY_FMT
    ws.freeze_panes = "A2"

    # ------------------------------------------------- Raw_AR
    ws = wb.create_sheet("Raw_AR")
    headers = ["Invoice ID", "Customer", "Date", "Due Date", "Amount",
               "Paid Date", "In Transit", "Days Overdue", "Aging Bucket",
               "Outstanding"]
    ws.append(headers)
    for i, r in enumerate(ar, start=2):
        ws.append([r["invoice_id"], r["customer"], dt.date.fromisoformat(r["date"]),
                   dt.date.fromisoformat(r["due_date"]), float(r["amount"]),
                   dt.date.fromisoformat(r["paid_date"]) if r["paid_date"] else None,
                   int(r["in_transit"]),
                   f"=IF(F{i}=\"\",AsOf-D{i},F{i}-D{i})",
                   bucket_formula("H", "F").format(r=i),
                   f"=IF(F{i}=\"\",E{i},0)"])
    style_header(ws, len(headers))
    for col, w in zip("ABCDEFGHIJ", [11, 16, 12, 12, 14, 12, 10, 13, 13, 14]):
        ws.column_dimensions[col].width = w
    for c in ws["E"][1:]:
        c.number_format = MONEY_FMT
    ws.freeze_panes = "A2"

    # ------------------------------------------------- Payroll
    ws = wb.create_sheet("Payroll")
    headers = ["Emp ID", "Name", "Department", "Pay Date", "Gross",
               "Deductions", "Net"]
    ws.append(headers)
    for i, r in enumerate(pay, start=2):
        ws.append([r["employee_id"], r["name"], r["department"],
                   dt.date.fromisoformat(r["pay_date"]), float(r["gross"]),
                   float(r["deductions"]), f"=E{i}-F{i}"])
    style_header(ws, len(headers))
    for col, w in zip("ABCDEFG", [9, 12, 13, 12, 13, 13, 13]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2, max_row=n_pay + 1, min_col=5, max_col=7):
        for c in row:
            c.number_format = MONEY_FMT
    ws.freeze_panes = "A2"

    # ------------------------------------------------- Bank
    ws = wb.create_sheet("Bank")
    headers = ["Date", "Description", "Ref", "Withdrawal", "Deposit", "Net"]
    ws.append(headers)
    for i, r in enumerate(st, start=2):
        ws.append([dt.date.fromisoformat(r["date"]), r["description"], r["ref"],
                   float(r["withdrawal"]), float(r["deposit"]), f"=E{i}-D{i}"])
    style_header(ws, len(headers))
    for col, w in zip("ABCDEF", [12, 30, 12, 14, 14, 14]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2, max_row=n_st + 1, min_col=4, max_col=6):
        for c in row:
            c.number_format = MONEY_FMT
    ws.freeze_panes = "A2"

    # ------------------------------------------------- Bank_Rec
    ws = wb.create_sheet("Bank_Rec")
    ws["A1"] = "Bank Reconciliation (INR) - all figures are live formulas"
    ws["A1"].font = TITLE_FONT
    rec = [
        ("Bank balance per statement", "=SUM(Bank!F:F)"),
        ("Add: deposits in transit", "=SUMIFS(Raw_AR!E:E,Raw_AR!G:G,1)"),
        ("Less: outstanding cheques",
         '=SUMIFS(Raw_AP!E:E,Raw_AP!G:G,"cheque",Raw_AP!F:F,"<>"&"",Raw_AP!H:H,0)'),
        ("ADJUSTED BANK BALANCE", "=B3+B4-B5", True),
        (None, None, None),
        ("Book balance per books",
         '=BookOpen'
         '+SUMIFS(Raw_AR!E:E,Raw_AR!F:F,"<>"&"",Raw_AR!C:C,">="&PeriodStart,Raw_AR!C:C,"<"&PeriodEnd)'
         '-SUMIFS(Raw_AP!E:E,Raw_AP!F:F,"<>"&"",Raw_AP!C:C,">="&PeriodStart,Raw_AP!C:C,"<"&PeriodEnd)'
         '-SUMIFS(Payroll!G:G,Payroll!D:D,">="&PeriodStart,Payroll!D:D,"<"&PeriodEnd)'),
        ("Add: interest credited (not in books)",
         '=SUMIFS(Bank!E:E,Bank!C:C,"INT-*")'),
        ("Less: bank charges (not in books)",
         '=SUMIFS(Bank!D:D,Bank!C:C,"CHG-*")'),
        ("ADJUSTED BOOK BALANCE", "=B8+B9-B10", True),
        (None, None, None),
        ("Difference (adjusted bank - adjusted book)", "=B6-B11"),
        ("Status",
         '=IF(ABS(B13)<0.01,"OK - reconciled","OUT OF BALANCE - investigate")'),
    ]
    for k, row in enumerate(rec):
        r = 3 + k
        label, formula, *rest = row
        if label is None:
            continue
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=formula).number_format = MONEY_FMT
        if rest and rest[0]:
            ws.cell(row=r, column=1).font = Font(bold=True)
            ws.cell(row=r, column=2).font = Font(bold=True)
    ws.conditional_formatting.add(
        "B14", FormulaRule(formula=['NOT(ISNUMBER(SEARCH("OK",B14)))'],
                           fill=RED_FILL, font=RED_FONT))
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 20

    # ------------------------------------------------- AP_Aging / AR_Aging
    def aging_sheet(title, bucket_col, out_col, data_n):
        ws = wb.create_sheet(title)
        ws["A1"] = f"{title.replace('_', ' ')} (INR) - live SUMIFS"
        ws["A1"].font = TITLE_FONT
        ws.append(["Bucket", "Outstanding", "Count"])
        for i, b in enumerate(BUCKETS, start=3):
            ws.cell(row=i, column=1, value=b)
            ws.cell(row=i, column=2,
                    value=f"=SUMIFS({out_col},{bucket_col},A{i})")
            ws.cell(row=i, column=3,
                    value=f'=COUNTIFS({bucket_col},A{i},{out_col},">0")')
        t = 3 + len(BUCKETS)
        ws.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
        ws.cell(row=t, column=2, value="=SUM(B3:B7)").font = Font(bold=True)
        ws.cell(row=t, column=3, value="=SUM(C3:C7)").font = Font(bold=True)
        ws.cell(row=t + 1, column=1, value="Integrity check")
        ws.cell(row=t + 1, column=2,
                value=f'=IF(ABS(B{t}-SUM({out_col}))<0.01,"OK - ties to raw data","DIFF")')
        style_header(ws, 3, row=2)
        for row in ws.iter_rows(min_row=3, max_row=t, min_col=2, max_col=2):
            for c in row:
                c.number_format = MONEY_FMT
        ws.conditional_formatting.add(
            "B7", CellIsRule(operator="greaterThan", formula=["0"],
                             fill=RED_FILL, font=RED_FONT))
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 10
        return ws

    ws_ap_age = aging_sheet("AP_Aging", "Raw_AP!K:K", "Raw_AP!L:L", n_ap)
    ws_ar_age = aging_sheet("AR_Aging", "Raw_AR!I:I", "Raw_AR!J:J", n_ar)

    # ------------------------------------------------- P&L
    ws = wb.create_sheet("P&L")
    ws["A1"] = "Profit & Loss (INR), Apr-Sep 2026"
    ws["A1"].font = TITLE_FONT
    headers = ["Line"] + [m.strftime("%b-%y") for m in MONTHS] + ["Total"]
    ws.append([])  # row 2 holds month start dates
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=j, value=h)
        if 1 < j <= 7:
            ws.cell(row=2, column=j, value=MONTHS[j - 2]).number_format = "mmm-yy"
    style_header(ws, len(headers), row=3)

    def msum_args(date_col):
        """date filter fragment shared by monthly SUMIFS."""
        return (f'{date_col},">="&B$2,{date_col},"<"&EDATE(B$2,1)')

    lines = [
        ("Revenue (invoiced)",
         lambda r: f"=SUMIFS(Raw_AR!E:E,{msum_args('Raw_AR!C:C')})"),
        ("Purchases",
         lambda r: f'=SUMIFS(Raw_AP!E:E,Raw_AP!I:I,"Purchases",{msum_args("Raw_AP!C:C")})'),
        ("Gross Profit", None),  # 4-5
        ("Payroll",
         lambda r: f"=SUMIFS(Payroll!G:G,{msum_args('Payroll!D:D')})"),
        ("Other operating expenses",
         lambda r: f'=SUMIFS(Raw_AP!E:E,Raw_AP!I:I,"<>"&"Purchases",{msum_args("Raw_AP!C:C")})'),
        ("Total operating expenses", None),  # 7+8
        ("NET INCOME", None),                # 6-9
        ("Total expenses (charts)", None),   # 5+9
    ]
    r0 = 4
    for k, (label, fn) in enumerate(lines):
        r = r0 + k
        ws.cell(row=r, column=1, value=label)
        if fn is not None:
            for j in range(2, 8):
                col = get_column_letter(j)
                ws.cell(row=r, column=j, value=fn(r).replace("B$2", f"{col}$2"))
            ws.cell(row=r, column=8, value=f"=SUM(B{r}:G{r})")
        else:
            for j in range(2, 9):
                col = get_column_letter(j)
                if k == 2:      # Gross = Rev - Purchases
                    ws.cell(row=r, column=j, value=f"={col}{r0}-{col}{r0+1}")
                elif k == 5:    # Opex total
                    ws.cell(row=r, column=j, value=f"={col}{r0+3}+{col}{r0+4}")
                elif k == 6:    # NI
                    ws.cell(row=r, column=j, value=f"={col}{r0+2}-{col}{r0+5}")
                elif k == 7:    # Total expenses
                    ws.cell(row=r, column=j, value=f"={col}{r0+1}+{col}{r0+5}")
    last = r0 + len(lines) - 1
    for row in ws.iter_rows(min_row=r0, max_row=last, min_col=2, max_col=8):
        for c in row:
            c.number_format = MONEY_FMT
            c.border = BORDER
    for bold_row in (r0 + 2, r0 + 5, r0 + 6, r0 + 7):
        for c in ws[bold_row]:
            c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 30
    for j in range(2, 9):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.freeze_panes = "B4"
    rev_row, ni_row, exp_row = r0, r0 + 6, r0 + 7

    # ------------------------------------------------- Dashboard
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "AccountingPulse - Rajesh Kumar & Sons (INR)"
    ws["A1"].font = Font(size=16, bold=True, color="1F4E78")
    ws["A2"] = "Apr-Sep 2026  |  All figures are live formulas - refresh = append rows + rerun generate.py"
    ws["A2"].font = Font(italic=True, color="808080")
    ws["A3"] = "As-of date"
    ws["B3"] = AS_OF
    ws["B3"].number_format = "dd-mmm-yy"
    ws["A4"] = "Opening book cash"
    ws["B4"] = BOOK_OPEN
    ws["B4"].number_format = MONEY_FMT
    ws["E1"] = PERIOD_START
    ws["E2"] = dt.date(2026, 10, 1)
    ws.column_dimensions["E"].hidden = True
    add_named(wb, "AsOf", "Dashboard!$B$3")
    add_named(wb, "BookOpen", "Dashboard!$B$4")
    add_named(wb, "PeriodStart", "Dashboard!$E$1")
    add_named(wb, "PeriodEnd", "Dashboard!$E$2")

    ws["A6"] = "Key figures"
    ws["A6"].font = Font(bold=True, size=12, color="1F4E78")
    kpis = [
        ("AP Outstanding", "='AP_Aging'!B8", MONEY_FMT),
        ("AR Outstanding", "='AR_Aging'!B8", MONEY_FMT),
        ("Adjusted bank balance", "='Bank_Rec'!B6", MONEY_FMT),
        ("Bank-book difference", "='Bank_Rec'!B13", MONEY_FMT),
        ("Reconciliation status", "='Bank_Rec'!B14", None),
        ("Monthly payroll run-rate", "=SUM(Payroll!G:G)/6", MONEY_FMT),
        ("Net income (6 mo)", f"='P&L'!H{ni_row}", MONEY_FMT),
        ("AR overdue 90+", "='AR_Aging'!B7", MONEY_FMT),
    ]
    for k, (label, formula, fmt) in enumerate(kpis):
        r = 7 + k
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=formula)
        if fmt:
            c.number_format = fmt
        c.font = Font(size=12, bold=True)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 22

    # dept payroll helper table for the pie chart
    ws["G6"] = "Department"
    ws["H6"] = "Payroll total"
    depts = sorted({r["department"] for r in pay})
    for k, dpt in enumerate(depts):
        r = 7 + k
        ws.cell(row=r, column=7, value=dpt)
        ws.cell(row=r, column=8,
                value=f'=SUMIFS(Payroll!G:G,Payroll!C:C,G{r})').number_format = MONEY_FMT
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 16

    # Chart 1: AP vs AR aging (bar)
    c1 = BarChart()
    c1.title = "AP vs AR Aging Buckets (INR)"
    c1.style = 10
    c1.height, c1.width = 7.5, 15
    c1.add_data(Reference(ws_ap_age, min_col=2, min_row=3, max_row=7),
                titles_from_data=False)
    c1.add_data(Reference(ws_ar_age, min_col=2, min_row=3, max_row=7),
                titles_from_data=False)
    c1.set_categories(Reference(ws_ap_age, min_col=1, min_row=3, max_row=7))
    c1.series[0].title = SeriesLabel(v="AP Outstanding")
    c1.series[1].title = SeriesLabel(v="AR Outstanding")
    c1.shape = 4
    ws.add_chart(c1, "D6")

    # Chart 2: monthly revenue vs total expenses (line)
    pls = wb["P&L"]
    c2 = LineChart()
    c2.title = "Monthly Revenue vs Expenses"
    c2.style = 2
    c2.height, c2.width = 7.5, 15
    data = Reference(pls, min_col=2, max_col=7, min_row=rev_row, max_row=exp_row)
    c2.add_data(data, from_rows=True, titles_from_data=False)
    c2.set_categories(Reference(pls, min_col=2, max_col=7, min_row=3))
    c2.series[0].title = SeriesLabel(v="Revenue")
    c2.series[1].title = SeriesLabel(v="Expenses")
    ws.add_chart(c2, "D22")

    # Chart 3: payroll by department (pie)
    c3 = PieChart()
    c3.title = "Payroll by Department (6 mo)"
    c3.style = 10
    c3.height, c3.width = 7.5, 15
    c3.add_data(Reference(ws, min_col=8, min_row=6, max_row=6 + len(depts)),
                titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=7, max_row=6 + len(depts)))
    ws.add_chart(c3, "D38")

    wb.save(out_path)
    print(f"wrote {out_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=str(HERE / "data"))
    ap.add_argument("--out", default=str(HERE / "accounting-pulse.xlsx"))
    a = ap.parse_args()
    build(Path(a.data), Path(a.out))


if __name__ == "__main__":
    main()
