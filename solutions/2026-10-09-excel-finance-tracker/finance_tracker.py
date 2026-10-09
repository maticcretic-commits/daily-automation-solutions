"""Excel Finance Tracker — sample transactions/invoices/budgets -> dashboard workbook.

Reads the three sample CSVs and builds Excel_Finance_Tracker.xlsx with:
  - Dashboard      : KPI cells with live formulas + 3 native charts
  - Transactions   : raw data with Month and TaxCat helper columns
  - Budget vs Actual: per-category 9-month budget vs actual, variance, Var%
  - Invoices       : invoice list with live aging (Days Outstanding, Payment Flag)
  - Tax Prep       : deductible totals grouped by tax category
  - ChartData      : helper series for the charts

All money cells are live formulas; the workbook refreshes when new rows are
pasted into Transactions / Invoices.
"""
import csv
import os
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Excel_Finance_Tracker.xlsx")

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
SECTION_FONT = Font(bold=True, size=11, color="1F3864")
TOTAL_FONT = Font(bold=True, size=11)
MONEY = "#,##0"
PCT = "0.0%"
DATEF = "yyyy-mm-dd"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONTHS = ["Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26",
          "Jun-26", "Jul-26", "Aug-26", "Sep-26"]
AS_OF = date(2026, 10, 9)

TAX_MAP = {  # expense category -> tax category
    "Rent": "Office",
    "Utilities": "Office",
    "Office Supplies": "Office",
    "Salaries": "Payroll",
    "Marketing": "Advertising",
    "Travel": "Travel",
    "Software": "Technology",
    "Insurance": "Insurance",
    "Professional Fees": "Professional Services",
}


def read_csv(name):
    with open(os.path.join(HERE, name), newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER


def border_all(ws, min_row, max_row, ncols):
    for r in range(min_row, max_row + 1):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER


# ---------------------------------------------------------------- data ----
def build_transactions(wb, tx_rows):
    ws = wb.active
    ws.title = "Transactions"
    headers = ["Date", "Type", "Category", "Description", "Amount",
               "Month", "TaxCat"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for r in tx_rows:
        d = datetime.strptime(r["Date"], "%Y-%m-%d").date()
        ws.append([d, r["Type"], r["Category"], r["Description"],
                   float(r["Amount"]), None, None])
    for i, _ in enumerate(tx_rows, start=2):
        ws[f"F{i}"] = f'=TEXT(A{i},"mmm-yy")'
        ws[f"G{i}"] = (f'=IF(B{i}="Income","",'
                       f'VLOOKUP(C{i},\'Tax Prep\'!$A$21:$B$29,2,FALSE))')
        ws[f"A{i}"].number_format = DATEF
        ws[f"E{i}"].number_format = MONEY
    widths = [13, 10, 20, 26, 13, 10, 22]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{ws.max_row}"
    border_all(ws, 1, ws.max_row, len(headers))
    return ws


# ------------------------------------------------------- budget vs actual --
def build_budget_vs_actual(wb, budgets, n_tx_rows):
    ws = wb.create_sheet("Budget vs Actual")
    headers = ["Category", "Monthly Budget", "9-mo Budget",
               "Actual", "Variance", "Var%"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    cats = list(budgets.keys())
    for i, cat in enumerate(cats, start=2):
        ws[f"A{i}"] = cat
        ws[f"B{i}"] = float(budgets[cat])
        ws[f"C{i}"] = f"=B{i}*9"
        ws[f"D{i}"] = (f"=SUMIFS(Transactions!$E:$E,Transactions!$C:$C,$A{i},"
                       f"Transactions!$B:$B,\"Expense\")")
        ws[f"E{i}"] = f"=C{i}-D{i}"
        ws[f"F{i}"] = f'=IF(C{i}=0,"",E{i}/C{i})'
        ws[f"F{i}"].number_format = PCT
        for col in "BCD":
            ws[f"{col}{i}"].number_format = MONEY
    t = len(cats) + 2
    ws[f"A{t}"] = "TOTAL"
    ws[f"A{t}"].font = TOTAL_FONT
    for col in "BCDE":
        ws[f"{col}{t}"] = f"=SUM({col}2:{col}{t - 1})"
        ws[f"{col}{t}"].font = TOTAL_FONT
        ws[f"{col}{t}"].number_format = MONEY
    ws[f"F{t}"] = f"=E{t}/C{t}"
    ws[f"F{t}"].font = TOTAL_FONT
    ws[f"F{t}"].number_format = PCT
    # red when over budget (variance < 0)
    ws.conditional_formatting.add(
        f"E2:E{t - 1}",
        CellIsRule(operator="lessThan", formula=["0"],
                   font=Font(color="9C0006"),
                   fill=PatternFill(start_color="FFC7CE", fill_type="solid")))
    for i, w in enumerate([22, 16, 14, 14, 14, 10], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    border_all(ws, 1, t, len(headers))
    return ws, len(cats)


# --------------------------------------------------------------- invoices --
def build_invoices(wb, inv_rows):
    ws = wb.create_sheet("Invoices")
    ws["A1"] = "As of:"
    ws["A1"].font = SECTION_FONT
    ws["B1"] = AS_OF
    ws["B1"].number_format = DATEF
    headers = ["InvoiceNo", "Client", "IssueDate", "DueDate", "Amount",
               "Status", "Days Outstanding", "Payment Flag"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=3, column=c, value=h)
    style_header(ws, 3, len(headers))
    for i, r in enumerate(inv_rows, start=4):
        ws[f"A{i}"] = r["InvoiceNo"]
        ws[f"B{i}"] = r["Client"]
        ws[f"C{i}"] = datetime.strptime(r["IssueDate"], "%Y-%m-%d").date()
        ws[f"D{i}"] = datetime.strptime(r["DueDate"], "%Y-%m-%d").date()
        ws[f"E{i}"] = float(r["Amount"])
        ws[f"F{i}"] = r["Status"]
        ws[f"G{i}"] = f'=IF(F{i}="Paid","",MAX(0,$B$1-D{i}))'
        ws[f"H{i}"] = (f'=IF(F{i}="Paid","Paid",'
                       f'IF(D{i}<=$B$1,"Overdue","Outstanding"))')
        for col in "CD":
            ws[f"{col}{i}"].number_format = DATEF
        ws[f"E{i}"].number_format = MONEY
    last = 3 + len(inv_rows)
    for i, w in enumerate([14, 16, 13, 13, 12, 10, 17, 13], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:H{last}"
    border_all(ws, 3, last, len(headers))
    return ws


# --------------------------------------------------------------- tax prep --
def build_tax_prep(wb):
    ws = wb.create_sheet("Tax Prep")
    ws["A1"] = "Deductible expenses by tax category (live from Transactions)"
    ws["A1"].font = TITLE_FONT
    headers = ["Tax Category", "Deductible Total", "Transactions"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=2, column=c, value=h)
    style_header(ws, 2, len(headers))
    tax_cats = sorted(set(TAX_MAP.values()))
    for i, tc in enumerate(tax_cats, start=3):
        ws[f"A{i}"] = tc
        ws[f"B{i}"] = (f"=SUMIFS(Transactions!$E:$E,Transactions!$G:$G,$A{i},"
                       f"Transactions!$B:$B,\"Expense\")")
        ws[f"C{i}"] = (f"=COUNTIFS(Transactions!$G:$G,$A{i},"
                       f"Transactions!$B:$B,\"Expense\")")
        ws[f"B{i}"].number_format = MONEY
    t = 3 + len(tax_cats)
    ws[f"A{t}"] = "TOTAL"
    ws[f"A{t}"].font = TOTAL_FONT
    ws[f"B{t}"] = f"=SUM(B3:B{t - 1})"
    ws[f"B{t}"].font = TOTAL_FONT
    ws[f"B{t}"].number_format = MONEY
    ws[f"C{t}"] = f"=SUM(C3:C{t - 1})"
    ws[f"C{t}"].font = TOTAL_FONT
    # mapping table used by the Transactions!G VLOOKUP (kept on this sheet)
    ws["A19"] = "Expense category -> tax category mapping"
    ws["A19"].font = SECTION_FONT
    ws["A20"] = "Expense Category"
    ws["B20"] = "Tax Category"
    for i, (exp, tc) in enumerate(sorted(TAX_MAP.items()), start=21):
        ws[f"A{i}"] = exp
        ws[f"B{i}"] = tc
    for i, w in enumerate([24, 18, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    border_all(ws, 2, t, len(headers))
    return ws


# -------------------------------------------------------------- chartdata --
def build_chartdata(wb):
    ws = wb.create_sheet("ChartData")
    headers = ["Month", "Income", "Expenses", "Cumulative Net"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, m in enumerate(MONTHS, start=2):
        ws[f"A{i}"] = m
        ws[f"B{i}"] = (f"=SUMIFS(Transactions!$E:$E,Transactions!$B:$B,\"Income\","
                       f"Transactions!$F:$F,$A{i})")
        ws[f"C{i}"] = (f"=SUMIFS(Transactions!$E:$E,Transactions!$B:$B,\"Expense\","
                       f"Transactions!$F:$F,$A{i})")
        ws[f"D{i}"] = f"=B{i}-C{i}" if i == 2 else f"=B{i}-C{i}+D{i - 1}"
        for col in "BCD":
            ws[f"{col}{i}"].number_format = MONEY
    for i, w in enumerate([10, 14, 14, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


# -------------------------------------------------------------- dashboard --
def build_dashboard(wb, n_budget_cats):
    ws = wb.create_sheet("Dashboard", 0)
    ws["A1"] = "Finance Tracker — Executive Dashboard"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:C1")
    kpis = [
        ("Total Income", '=SUMIF(Transactions!B:B,"Income",Transactions!E:E)'),
        ("Total Expenses", '=SUMIF(Transactions!B:B,"Expense",Transactions!E:E)'),
        ("Net Profit", "=B3-B4"),
        ("Cash on Hand", "=B3-B4"),
        ("Outstanding Receivables",
         '=SUMIFS(Invoices!E:E,Invoices!F:F,"Unpaid",Invoices!H:H,"Outstanding")'),
        ("Overdue Amount",
         '=SUMIFS(Invoices!E:E,Invoices!F:F,"Unpaid",Invoices!H:H,"Overdue")'),
    ]
    ws["A2"] = "KPI"
    ws["B2"] = "Amount"
    style_header(ws, 2, 2)
    for i, (label, formula) in enumerate(kpis, start=3):
        ws[f"A{i}"] = label
        ws[f"B{i}"] = formula
        ws[f"B{i}"].number_format = MONEY
        ws[f"B{i}"].font = Font(size=12)
    ws["A9"] = "Budget health: see 'Budget vs Actual' (red = over budget)"
    ws["A9"].font = Font(italic=True, color="808080")
    ws.merge_cells("A9:C9")
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 4

    n = len(MONTHS) + 1  # ChartData rows 2..10
    bar = BarChart()
    bar.title = "Monthly Income vs Expenses"
    bar.style = 10
    bar.add_data(Reference(wb["ChartData"], min_col=2, max_col=3,
                           min_row=1, max_row=n), titles_from_data=True)
    bar.set_categories(Reference(wb["ChartData"], min_col=1,
                                 min_row=2, max_row=n))
    bar.width, bar.height = 22, 13
    ws.add_chart(bar, "D3")

    line = LineChart()
    line.title = "Cumulative Net Cash"
    line.style = 10
    line.add_data(Reference(wb["ChartData"], min_col=4,
                            min_row=1, max_row=n), titles_from_data=True)
    line.set_categories(Reference(wb["ChartData"], min_col=1,
                                  min_row=2, max_row=n))
    line.width, line.height = 22, 13
    ws.add_chart(line, "D19")

    pie = PieChart()
    pie.title = "Expenses by Category"
    pie.add_data(Reference(wb["Budget vs Actual"], min_col=4,
                           min_row=2, max_row=n_budget_cats + 1),
                 titles_from_data=False)
    pie.set_categories(Reference(wb["Budget vs Actual"], min_col=1,
                                 min_row=2, max_row=n_budget_cats + 1))
    pie.width, pie.height = 20, 13
    ws.add_chart(pie, "L3")
    return ws


def main():
    tx_rows = read_csv("sample_transactions.csv")
    inv_rows = read_csv("sample_invoices.csv")
    budgets = {r["Category"]: r["MonthlyBudget"]
               for r in read_csv("sample_budgets.csv")}

    wb = Workbook()
    build_transactions(wb, tx_rows)
    ws_bva, n_cats = build_budget_vs_actual(wb, budgets, len(tx_rows))
    build_invoices(wb, inv_rows)
    build_tax_prep(wb)
    build_chartdata(wb)
    build_dashboard(wb, n_cats)

    wb.save(OUT)
    print(f"saved {OUT} "
          f"({len(tx_rows)} transactions, {len(inv_rows)} invoices)")


if __name__ == "__main__":
    main()
