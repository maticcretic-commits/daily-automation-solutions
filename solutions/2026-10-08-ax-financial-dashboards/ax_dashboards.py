"""AX 2012 Financial Dashboards.

Reads a Dynamics AX 2012-style GL trial-balance CSV export and builds a
self-contained Excel financial dashboard workbook with:

  - Data            : normalized export (plus a Signed amount column)
  - Income Statement: per-period P&L via live SUMIFS formulas
  - Balance Sheet   : per-period balance sheet via live SUMIFS formulas
  - KPI Dashboard   : annual KPIs (margins, current ratio, debt-to-equity)
  - ChartData       : helper series for charts
  - Charts          : 3 native Excel charts (revenue vs expenses, net-income
                      trend, expense breakdown)

All statement cells are formulas, so dropping a new period's export rows into
the Data sheet and extending the SUMIFS ranges refreshes everything.

Usage:
    python3 ax_dashboards.py [input.csv] [output.xlsx]
"""
import csv
import sys
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

IN_DEFAULT = "sample_ax_export.csv"
OUT_DEFAULT = "AX_2012_Financial_Dashboards.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SECTION_FONT = Font(bold=True, size=11, color="1F3864")
TOTAL_FONT = Font(bold=True, size=11)
MONEY = '"$"#,##0'
PCT = "0.0%"
RATIO = "0.00"
THIN = Side(style="thin", color="BFBFBF")
TOP_BORDER = Border(top=THIN)

# Presentation sign per account category: presentation = sign * (Debit - Credit)
SIGN = {"Revenue": -1, "Expense": 1, "Asset": 1, "Asset-Contra": 1,
        "Liability": -1, "Equity": -1}

PERIODS = list(range(1, 13))
# statement layout: A=account code (hidden), B=line item, C..N=periods, O=total
FIRST_P, LAST_P, TOTAL_C = "C", "N", "O"


def signed_amount(row):
    return SIGN[row["AccountCategory"]] * (row["Debit"] - row["Credit"])


def load_rows(path):
    rows = []
    with open(path, newline="") as f:
        for r in csv.DictReader(f):
            rows.append({
                "FiscalYear": int(r["FiscalYear"]),
                "Period": int(r["Period"]),
                "MainAccount": r["MainAccount"].strip(),
                "AccountName": r["AccountName"].strip(),
                "AccountCategory": r["AccountCategory"].strip(),
                "Debit": float(r["Debit"] or 0),
                "Credit": float(r["Credit"] or 0),
            })
    for r in rows:
        r["Signed"] = signed_amount(r)
    return rows


def paint_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def paint_totals(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = TOTAL_FONT
        cell.border = TOP_BORDER


def money_fmt(ws, row, first_col, last_col, pct_cols=()):
    for c in range(first_col, last_col + 1):
        ws.cell(row=row, column=c).number_format = MONEY


def statement_shell(wb, title):
    ws = wb.create_sheet(title)
    ws["A1"] = "Account"
    ws["B1"] = "Line Item"
    for i, p in enumerate(PERIODS):
        col = get_column_letter(3 + i)
        ws[f"{col}1"] = p
    ws[f"{TOTAL_C}1"] = "Total"
    ws.column_dimensions["A"].hidden = True
    ws.column_dimensions["B"].width = 30
    for i in range(12):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.column_dimensions[TOTAL_C].width = 16
    paint_header(ws, 1, 15)
    return ws


def sumifs_formula(period_col_letter, row):
    # =SUMIFS(Data!$H:$H, Data!$B:$B, <period header>, Data!$C:$C, <acct code>)
    return (f"=SUMIFS(Data!$H:$H,Data!$B:$B,{period_col_letter}$1,"
            f"Data!$C:$C,$A{row})")


def account_block(ws, start_row, accounts):
    """Write account rows; returns (next_row, first_row, last_row)."""
    row = start_row
    for main, name in accounts:
        ws[f"A{row}"] = main
        ws[f"B{row}"] = name
        for i, p in enumerate(PERIODS):
            col = get_column_letter(3 + i)
            ws[f"{col}{row}"] = sumifs_formula(col, row)
        ws[f"{TOTAL_C}{row}"] = f"=SUM(C{row}:{LAST_P}{row})"
        money_fmt(ws, row, 3, 15)
        row += 1
    return row, start_row, row - 1


def total_row(ws, row, label, first, last, ref_sheet=None):
    ws[f"B{row}"] = label
    for i in range(12):
        col = get_column_letter(3 + i)
        ws[f"{col}{row}"] = f"=SUM({col}{first}:{col}{last})"
    ws[f"{TOTAL_C}{row}"] = f"=SUM(C{row}:{LAST_P}{row})"
    money_fmt(ws, row, 3, 15)
    paint_totals(ws, row, 15)
    return row


def build(rows, out_path):
    wb = Workbook()

    # ---- Data sheet -----------------------------------------------------
    ws = wb.active
    ws.title = "Data"
    headers = ["FiscalYear", "Period", "MainAccount", "AccountName",
               "AccountCategory", "Debit", "Credit", "Signed"]
    ws.append(headers)
    for r in rows:
        ws.append([r["FiscalYear"], r["Period"], r["MainAccount"],
                   r["AccountName"], r["AccountCategory"],
                   r["Debit"], r["Credit"], round(r["Signed"], 2)])
    for c in range(6, 9):
        ws.cell(row=1, column=c).number_format = MONEY
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=6, max_col=8):
        for cell in row:
            cell.number_format = MONEY
    paint_header(ws, 1, 8)
    for w, letter in [(12, "A"), (8, "B"), (13, "C"), (24, "D"),
                      (16, "E"), (12, "F"), (12, "G"), (14, "H")]:
        ws.column_dimensions[letter].width = w

    # ---- Income Statement -----------------------------------------------
    is_ws = statement_shell(wb, "Income Statement")
    rev = sorted({(r["MainAccount"], r["AccountName"]) for r in rows
                  if r["AccountCategory"] == "Revenue"})
    exp = sorted({(r["MainAccount"], r["AccountName"]) for r in rows
                  if r["AccountCategory"] == "Expense"})
    is_ws["B2"] = "Revenue"
    is_ws["B2"].font = SECTION_FONT
    r, rev_first, rev_last = account_block(is_ws, 3, rev)
    total_row(is_ws, r, "Total Revenue", rev_first, rev_last)
    rev_total_row = r
    r += 2
    is_ws[f"B{r}"] = "Operating Expenses"
    is_ws[f"B{r}"].font = SECTION_FONT
    r += 1
    r, exp_first, exp_last = account_block(is_ws, r, exp)
    total_row(is_ws, r, "Total Expenses", exp_first, exp_last)
    exp_total_row = r
    ni_row = r + 1
    is_ws[f"B{ni_row}"] = "Net Income"
    for i in range(12):
        col = get_column_letter(3 + i)
        is_ws[f"{col}{ni_row}"] = f"={col}{rev_total_row}-{col}{exp_total_row}"
    is_ws[f"{TOTAL_C}{ni_row}"] = f"={TOTAL_C}{rev_total_row}-{TOTAL_C}{exp_total_row}"
    money_fmt(is_ws, ni_row, 3, 15)
    paint_totals(is_ws, ni_row, 15)

    # ---- Balance Sheet ----------------------------------------------------
    bs = statement_shell(wb, "Balance Sheet")
    assets = sorted({(r["MainAccount"], r["AccountName"]) for r in rows
                     if r["AccountCategory"] in ("Asset", "Asset-Contra")})
    liabs = sorted({(r["MainAccount"], r["AccountName"]) for r in rows
                    if r["AccountCategory"] == "Liability"})
    equity = sorted({(r["MainAccount"], r["AccountName"]) for r in rows
                     if r["AccountCategory"] == "Equity"})
    bs["B2"] = "Assets"
    bs["B2"].font = SECTION_FONT
    r, a_first, a_last = account_block(bs, 3, assets)
    total_row(bs, r, "Total Assets", a_first, a_last)
    assets_total_row = r
    r += 2
    bs[f"B{r}"] = "Liabilities"
    bs[f"B{r}"].font = SECTION_FONT
    r += 1
    r, l_first, l_last = account_block(bs, r, liabs)
    total_row(bs, r, "Total Liabilities", l_first, l_last)
    liab_total_row = r
    r += 2
    bs[f"B{r}"] = "Equity"
    bs[f"B{r}"].font = SECTION_FONT
    r += 1
    r, e_first, e_last = account_block(bs, r, equity)
    # Retained Earnings carries the OPENING balance in the export; roll the
    # current period's profit in when drawing the balance sheet (pre-close).
    re_row = e_last  # 302000 sorts last among equity accounts
    bs[f"B{re_row}"] = "Retained Earnings (incl. current profit)"
    for i in range(12):
        col = get_column_letter(3 + i)
        bs[f"{col}{re_row}"] = (f"={sumifs_formula(col, re_row)}"
                                f"+'Income Statement'!{col}{ni_row}")
    bs[f"{TOTAL_C}{re_row}"] = f"=SUM(C{re_row}:{LAST_P}{re_row})"
    total_row(bs, r, "Total Equity", e_first, e_last)
    equity_total_row = r
    le_row = r + 1
    bs[f"B{le_row}"] = "Total Liabilities & Equity"
    for i in range(12):
        col = get_column_letter(3 + i)
        bs[f"{col}{le_row}"] = f"={col}{liab_total_row}+{col}{equity_total_row}"
    bs[f"{TOTAL_C}{le_row}"] = f"={TOTAL_C}{liab_total_row}+{TOTAL_C}{equity_total_row}"
    money_fmt(bs, le_row, 3, 15)
    paint_totals(bs, le_row, 15)
    chk = le_row + 1
    bs[f"B{chk}"] = "Check (Assets - L&E) = 0"
    bs[f"B{chk}"].font = Font(italic=True)
    for i in range(12):
        col = get_column_letter(3 + i)
        bs[f"{col}{chk}"] = f"={col}{assets_total_row}-{col}{le_row}"
    bs[f"{TOTAL_C}{chk}"] = f"={TOTAL_C}{assets_total_row}-{TOTAL_C}{le_row}"
    money_fmt(bs, chk, 3, 15)

    # ---- KPI Dashboard ----------------------------------------------------
    kpi = wb.create_sheet("KPI Dashboard")
    kpi["A1"] = "KPI Dashboard — FY 2026 (annual)"
    kpi["A1"].font = Font(bold=True, size=13, color="1F3864")
    kpi.column_dimensions["A"].width = 26
    kpi.column_dimensions["B"].width = 18
    kpi_items = [
        ("Total Revenue", f"='Income Statement'!{TOTAL_C}{rev_total_row}", MONEY),
        ("Total Expenses", f"='Income Statement'!{TOTAL_C}{exp_total_row}", MONEY),
        ("Net Income", f"='Income Statement'!{TOTAL_C}{ni_row}", MONEY),
        ("Net Margin", "=B5/B3", PCT),
        ("Current Assets",
         f"='Balance Sheet'!{TOTAL_C}3+'Balance Sheet'!{TOTAL_C}4+"
         f"'Balance Sheet'!{TOTAL_C}5+'Balance Sheet'!{TOTAL_C}6", MONEY),
        ("Current Liabilities",
         f"='Balance Sheet'!{TOTAL_C}12+'Balance Sheet'!{TOTAL_C}13+"
         f"'Balance Sheet'!{TOTAL_C}14+'Balance Sheet'!{TOTAL_C}16", MONEY),
        ("Current Ratio", "=B7/B8", RATIO),
        ("Total Debt",
         f"='Balance Sheet'!{TOTAL_C}14+'Balance Sheet'!{TOTAL_C}15", MONEY),
        ("Total Equity", f"='Balance Sheet'!{TOTAL_C}{equity_total_row}", MONEY),
        ("Debt-to-Equity", "=B10/B11", RATIO),
    ]
    for i, (label, formula, fmt) in enumerate(kpi_items):
        row = 3 + i
        kpi[f"A{row}"] = label
        kpi[f"A{row}"].font = TOTAL_FONT
        kpi[f"B{row}"] = formula
        kpi[f"B{row}"].number_format = fmt
        kpi[f"B{row}"].font = Font(size=12)

    # ---- ChartData --------------------------------------------------------
    cd = wb.create_sheet("ChartData")
    cd["A1"] = "Period"
    cd["B1"] = "Revenue"
    cd["C1"] = "Expenses"
    cd["D1"] = "Net Income"
    for p in PERIODS:
        col = get_column_letter(2 + p)  # IS period p lives in column C..N
        r = p + 1
        cd[f"A{r}"] = p
        cd[f"B{r}"] = f"='Income Statement'!{col}{rev_total_row}"
        cd[f"C{r}"] = f"='Income Statement'!{col}{exp_total_row}"
        cd[f"D{r}"] = f"='Income Statement'!{col}{ni_row}"
        for c in "BCD":
            cd[f"{c}{r}"].number_format = MONEY
    paint_header(cd, 1, 4)

    # ---- Charts -----------------------------------------------------------
    ch = wb.create_sheet("Charts")

    bar = BarChart()
    bar.type = "col"
    bar.title = "Monthly Revenue vs Expenses"
    bar.y_axis.title = "USD"
    bar.x_axis.title = "Period"
    bar.style = 10
    data = Reference(cd, min_col=2, min_row=1, max_row=13, max_col=3)
    cats = Reference(cd, min_col=1, min_row=2, max_row=13)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.shape = 4
    ch.add_chart(bar, "A1")

    line = LineChart()
    line.title = "Net Income Trend"
    line.y_axis.title = "USD"
    line.style = 12
    data = Reference(cd, min_col=4, min_row=1, max_row=13)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    ch.add_chart(line, "A17")

    pie = PieChart()
    pie.title = "Annual Expense Breakdown"
    labels = Reference(is_ws, min_col=2, min_row=exp_first, max_row=exp_last)
    data = Reference(is_ws, min_col=15, min_row=exp_first, max_row=exp_last)
    pie.add_data(data, titles_from_data=False)
    pie.set_categories(labels)
    pie.dataLabels = None
    ch.add_chart(pie, "A33")

    wb.save(out_path)
    print(f"wrote {out_path}")
    return {
        "rev_total_row": rev_total_row, "exp_total_row": exp_total_row,
        "ni_row": ni_row, "assets_total_row": assets_total_row,
        "liab_total_row": liab_total_row, "equity_total_row": equity_total_row,
        "le_row": le_row, "chk_row": chk,
    }


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else IN_DEFAULT
    out = sys.argv[2] if len(sys.argv) > 2 else OUT_DEFAULT
    rows = load_rows(src)
    print(f"loaded {len(rows)} GL rows from {src}")
    build(rows, out)


if __name__ == "__main__":
    main()
