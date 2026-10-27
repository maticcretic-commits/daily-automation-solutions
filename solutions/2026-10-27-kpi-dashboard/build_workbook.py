#!/usr/bin/env python3
"""Build kpi-dashboard.xlsx from data/*.csv — every figure a live formula.

Sheets: Guide, Dashboard, Monthly, Channel, Raw, Targets, Checks.
Rules honoured (learned the hard way): plain A1 ranges only, range literals
repeat the column letter on both ends ($H$2:$H$5000), no Table objects,
chart series titles come from header cells (titles_from_data), never plain
strings. Raw Month column is a stored value so SUMIFS matches exactly.
"""
import csv
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
OUT = os.path.join(HERE, "kpi-dashboard.xlsx")

NAVY = "1F3864"
TEAL = "2E75B6"
LIGHT = "D9E2F3"
GREEN = "548235"
RED = "C00000"
AMBER = "BF8F00"


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.reader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)


def main():
    raw = read_csv("raw_transactions.csv")
    targets = read_csv("monthly_targets.csv")
    raw_header, raw_rows = raw[0], raw[1:]
    tgt_header, tgt_rows = targets[0], targets[1:]
    n_raw = len(raw_rows)
    raw_last = n_raw + 1  # header is row 1
    months = [r[0] for r in tgt_rows]
    n_m = len(months)
    m_last = n_m + 1

    wb = Workbook()

    # ---- Guide ----
    ws = wb.active
    ws.title = "Guide"
    guide = [
        ("KPI Pulse — Monthly KPI Dashboard (DEMO)", ""),
        ("Built around: Power BI KPI Dashboard Development (Freelancer, Nagpur client).", ""),
        ("Data: SYNTHETIC — fictional Nagpur Retail Co., Jan-Sep 2026, seeded.", ""),
        ("", ""),
        ("How to use:", ""),
        ("1. Paste new daily rows into Raw (keep the Month column = YYYY-MM).", ""),
        ("2. Monthly, Channel, Dashboard and Checks update by formula — no retyping.", ""),
        ("3. Edit Targets to change the RAG thresholds on Monthly.", ""),
        ("4. Power BI version: see powerbi/MEASURES.md for the DAX mapping.", ""),
        ("", ""),
        ("Sheets: Dashboard (KPI cards + charts) | Monthly (SUMIFS) | Channel", ""),
        ("(SUMIF) | Raw (daily rows) | Targets | Checks (must all read TRUE).", ""),
    ]
    for i, (a, _b) in enumerate(guide, start=1):
        ws.cell(row=i, column=1, value=a)
    ws.column_dimensions["A"].width = 95
    ws["A1"].font = Font(bold=True, size=14, color=NAVY)

    # ---- Raw ----
    ws = wb.create_sheet("Raw")
    ws.append(raw_header)
    style_header(ws, 1, len(raw_header))
    for r in raw_rows:
        ws.append([r[0], r[1], r[2], int(r[3]), int(r[4]), int(r[5]),
                   int(r[6]), float(r[7])])
    ws.auto_filter.ref = f"A1:H{raw_last}"
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFGH", [12, 9, 11, 9, 12, 12, 13, 10]):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["E"].number_format = '#,##0'
    ws.column_dimensions["F"].number_format = '#,##0'

    # ---- Targets ----
    ws = wb.create_sheet("Targets")
    ws.append(tgt_header)
    style_header(ws, 1, len(tgt_header))
    for r in tgt_rows:
        ws.append([r[0], int(r[1]), int(r[2])])
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14

    # ---- Monthly ----
    ws = wb.create_sheet("Monthly")
    headers = ["Month", "Revenue", "RevenueTarget", "Variance", "Variance%",
               "Orders", "AvgOrderValue", "Cost", "GrossProfit", "Margin%",
               "OnTime%", "CSAT", "RAG"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, month in enumerate(months):
        r = i + 2
        ws.cell(row=r, column=1, value=month)
        # B Revenue, F Orders, H Cost, K OnTime%, L CSAT via SUMIFS on Raw Month.
        ws.cell(row=r, column=2,
                value=f"=SUMIFS(Raw!$E$2:$E${raw_last},Raw!$B$2:$B${raw_last},$A{r})")
        ws.cell(row=r, column=3, value=f"=Targets!$B${r}")
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
        ws.cell(row=r, column=5, value=f"=IF(C{r}=0,0,D{r}/C{r})")
        ws.cell(row=r, column=6,
                value=f"=SUMIFS(Raw!$D$2:$D${raw_last},Raw!$B$2:$B${raw_last},$A{r})")
        ws.cell(row=r, column=7, value=f"=IF(F{r}=0,0,B{r}/F{r})")
        ws.cell(row=r, column=8,
                value=f"=SUMIFS(Raw!$F$2:$F${raw_last},Raw!$B$2:$B${raw_last},$A{r})")
        ws.cell(row=r, column=9, value=f"=B{r}-H{r}")
        ws.cell(row=r, column=10, value=f"=IF(B{r}=0,0,I{r}/B{r})")
        ws.cell(row=r, column=11,
                value=f"=IF(F{r}=0,0,SUMIFS(Raw!$G$2:$G${raw_last},Raw!$B$2:$B${raw_last},$A{r})/F{r})")
        ws.cell(row=r, column=12,
                value=f"=IF(F{r}=0,0,SUMIFS(Raw!$H$2:$H${raw_last},Raw!$B$2:$B${raw_last},$A{r})/F{r})")
        ws.cell(row=r, column=13,
                value=f'=IF(E{r}>=0,"On track",IF(E{r}>=-0.05,"Watch","At risk"))')
    ws.append(["TOTAL", f"=SUM(B2:B{m_last})", f"=SUM(C2:C{m_last})",
               f"=B{m_last + 1}-C{m_last + 1}", None,
               f"=SUM(F2:F{m_last})", None, f"=SUM(H2:H{m_last})",
               f"=B{m_last + 1}-H{m_last + 1}", None, None, None, None])
    ws.cell(row=m_last + 1, column=5,
            value=f"=IF(C{m_last + 1}=0,0,D{m_last + 1}/C{m_last + 1})")
    ws.cell(row=m_last + 1, column=7,
            value=f"=IF(F{m_last + 1}=0,0,B{m_last + 1}/F{m_last + 1})")
    ws.cell(row=m_last + 1, column=10,
            value=f"=IF(B{m_last + 1}=0,0,I{m_last + 1}/B{m_last + 1})")
    for c in ("B", "C", "D", "H", "I"):
        ws.column_dimensions[c].number_format = '#,##0'
    for c in ("E", "J", "K"):
        ws.column_dimensions[c].number_format = '0.0%'
    ws.column_dimensions["L"].number_format = '0.00'
    ws.column_dimensions["G"].number_format = '#,##0'
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["M"].width = 10
    ws.freeze_panes = "A2"

    # ---- Channel ----
    ws = wb.create_sheet("Channel")
    ch_headers = ["Channel", "Revenue", "Orders", "Cost", "GrossProfit",
                  "Margin%", "Share%"]
    ws.append(ch_headers)
    style_header(ws, 1, len(ch_headers))
    channels = ["Online", "Store", "Wholesale"]
    for i, ch in enumerate(channels):
        r = i + 2
        ws.cell(row=r, column=1, value=ch)
        ws.cell(row=r, column=2,
                value=f"=SUMIF(Raw!$C$2:$C${raw_last},$A{r},Raw!$E$2:$E${raw_last})")
        ws.cell(row=r, column=3,
                value=f"=SUMIF(Raw!$C$2:$C${raw_last},$A{r},Raw!$D$2:$D${raw_last})")
        ws.cell(row=r, column=4,
                value=f"=SUMIF(Raw!$C$2:$C${raw_last},$A{r},Raw!$F$2:$F${raw_last})")
        ws.cell(row=r, column=5, value=f"=B{r}-D{r}")
        ws.cell(row=r, column=6, value=f"=IF(B{r}=0,0,E{r}/B{r})")
        ws.cell(row=r, column=7, value=f"=IF(SUM($B$2:$B$4)=0,0,B{r}/SUM($B$2:$B$4))")
    ws.append(["TOTAL", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)",
               "=B5-D5", "=IF(B5=0,0,E5/B5)", "=SUM(G2:G4)"])
    for c in ("B", "C", "D", "E"):
        ws.column_dimensions[c].number_format = '#,##0'
    for c in ("F", "G"):
        ws.column_dimensions[c].number_format = '0.0%'

    # ---- Dashboard ----
    ws = wb.create_sheet("Dashboard", 1)  # after Guide
    ws["A1"] = "KPI Pulse — Nagpur Retail Co. (SYNTHETIC DEMO, Jan-Sep 2026)"
    ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    cards = [
        ("Latest month revenue (Sep 2026)", f"=Monthly!B{m_last}"),
        ("YTD revenue", f"=SUM(Monthly!B2:B{m_last})"),
        ("YTD gross margin", f"=(SUM(Monthly!B2:B{m_last})-SUM(Monthly!H2:H{m_last}))/SUM(Monthly!B2:B{m_last})"),
        ("Latest on-time delivery", f"=Monthly!K{m_last}"),
        ("Latest CSAT (out of 5)", f"=Monthly!L{m_last}"),
        ("YTD target attainment", f"=SUM(Monthly!B2:B{m_last})/SUM(Targets!B2:B{m_last})"),
    ]
    ws["A3"] = "KPI"
    ws["B3"] = "Value"
    style_header(ws, 3, 2)
    for i, (label, formula) in enumerate(cards):
        r = 4 + i
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=formula)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    ws["B4"].number_format = '#,##0'
    ws["B5"].number_format = '#,##0'
    ws["B6"].number_format = '0.0%'
    ws["B7"].number_format = '0.0%'
    ws["B8"].number_format = '0.00'
    ws["B9"].number_format = '0.0%'

    # Charts: series titles from header cells via titles_from_data.
    bar = BarChart()
    bar.type = "col"
    bar.title = "Revenue vs target by month"
    data = _chart_data(wb["Monthly"], "Monthly", 2, 3, m_last)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(_cats(wb["Monthly"], m_last))
    bar.height, bar.width = 7.5, 14
    ws.add_chart(bar, "A12")

    line = LineChart()
    line.title = "Gross margin % by month"
    data = _chart_data(wb["Monthly"], "Monthly", 10, 10, m_last)
    line.add_data(data, titles_from_data=True)
    line.set_categories(_cats(wb["Monthly"], m_last))
    line.height, line.width = 7.5, 14
    ws.add_chart(line, "A28")

    pie = PieChart()
    pie.title = "Revenue share by channel"
    data = _chart_data(wb["Channel"], "Channel", 2, 2, 4)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(_cats(wb["Channel"], 4, col=1))
    pie.height, pie.width = 7.5, 14
    ws.add_chart(pie, "A44")

    # ---- Checks ----
    ws = wb.create_sheet("Checks")
    ws.append(["Check", "Result"])
    style_header(ws, 1, 2)
    tot = m_last + 1
    checks = [
        ("Monthly revenue = Raw revenue",
         f"=ABS(SUM(Monthly!B2:B{m_last})-SUM(Raw!$E$2:$E${raw_last}))<0.01"),
        ("Monthly orders = Raw orders",
         f"=ABS(SUM(Monthly!F2:F{m_last})-SUM(Raw!$D$2:$D${raw_last}))<0.01"),
        ("Monthly cost = Raw cost",
         f"=ABS(SUM(Monthly!H2:H{m_last})-SUM(Raw!$F$2:$F${raw_last}))<0.01"),
        ("Channel revenue = Raw revenue",
         f"=ABS(SUM(Channel!B2:B4)-SUM(Raw!$E$2:$E${raw_last}))<0.01"),
        ("Channel share sums to 100%",
         "=ABS(SUM(Channel!G2:G4)-1)<0.0001"),
        ("Monthly TOTAL row = sum of months (revenue)",
         f"=ABS(Monthly!B{tot}-SUM(Monthly!B2:B{m_last}))<0.01"),
        ("Dashboard YTD revenue = Monthly total",
         f"=ABS(Dashboard!B5-Monthly!B{tot})<0.01"),
        ("Dashboard latest revenue = Sep Monthly revenue",
         f"=ABS(Dashboard!B4-Monthly!B{m_last})<0.01"),
    ]
    for label, formula in checks:
        ws.append([label, formula])
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 12

    wb.save(OUT)
    print(f"wrote {OUT} ({n_raw} raw rows, {n_m} months)")


def _chart_data(ws, title, min_col, max_col, max_row):
    from openpyxl.chart.data_source import NumDataSource  # noqa: F401 (kept for clarity)
    from openpyxl.chart import Reference
    return Reference(ws, min_col=min_col, min_row=1,
                     max_col=max_col, max_row=max_row)


def _cats(ws, max_row, col=1):
    from openpyxl.chart import Reference
    return Reference(ws, min_col=col, min_row=2, max_row=max_row)


if __name__ == "__main__":
    main()
