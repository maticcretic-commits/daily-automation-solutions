#!/usr/bin/env python3
"""PortfolioPulse — risk-adjusted portfolio performance dashboard.

Built around a real Freelancer brief: a dynamic Excel spreadsheet /
dashboard for portfolio performance analysis (risk-adjusted returns:
alpha, beta, Sharpe ratio, drawdowns, volatility).

Demo data is fictional and deterministic (seeded): a diversified
portfolio vs a market benchmark, monthly returns Jan 2025 - Oct 2026
(22 months). Every figure in the workbook is a live Excel formula off
the Returns sheet — nothing hard-coded except the monthly returns
themselves and the risk-free rate.

Sheets: Guide, Returns, Metrics, Dashboard (KPIs + 3 native charts),
Checks. No macros, no structured table references, plain A1 ranges
with the column letter repeated on both ends (LibreOffice-safe).
"""
import csv
import math
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
N = 22  # Jan 2025 - Oct 2026
FIRST, LAST = 2, 2 + N - 1  # data rows in Returns

NAVY = "1F3864"
TEAL = "0F7B7A"
LIGHT = "D9E2F3"
GREY = "F2F2F2"


def build_data():
    rng = random.Random(20261105)
    months = []
    y, m = 2025, 1
    for _ in range(N):
        months.append(f"{y:04d}-{m:02d}")
        m += 1
        if m == 13:
            y, m = y + 1, 1
    rows = []
    for month in months:
        bench = rng.gauss(0.007, 0.032)
        port = 0.9 * bench + rng.gauss(0.002, 0.018)
        rows.append((month, round(port, 6), round(bench, 6), 0.0045))
    return rows


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def build_workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Returns"
    headers = ["Month", "Portfolio return", "Benchmark return", "Risk-free (monthly)",
               "Portfolio excess", "Benchmark excess", "Active return",
               "Portfolio wealth", "Benchmark wealth",
               "Portfolio drawdown", "Benchmark drawdown"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, (month, p, b, rf) in enumerate(rows):
        r = FIRST + i
        ws.cell(row=r, column=1, value=month)
        ws.cell(row=r, column=2, value=p)
        ws.cell(row=r, column=3, value=b)
        ws.cell(row=r, column=4, value=rf)
        ws.cell(row=r, column=5, value=f"=B{r}-D{r}")
        ws.cell(row=r, column=6, value=f"=C{r}-D{r}")
        ws.cell(row=r, column=7, value=f"=B{r}-C{r}")
        if i == 0:
            ws.cell(row=r, column=8, value=f"=1+B{r}")
            ws.cell(row=r, column=9, value=f"=1+C{r}")
        else:
            ws.cell(row=r, column=8, value=f"=H{r-1}*(1+B{r})")
            ws.cell(row=r, column=9, value=f"=I{r-1}*(1+C{r})")
        ws.cell(row=r, column=10, value=f"=H{r}/MAX($H${FIRST}:H{r})-1")
        ws.cell(row=r, column=11, value=f"=I{r}/MAX($I${FIRST}:I{r})-1")
    for r in range(FIRST, LAST + 1):
        for c in (2, 3, 4, 5, 6, 7, 10, 11):
            ws.cell(row=r, column=c).number_format = "0.00%"
        for c in (8, 9):
            ws.cell(row=r, column=c).number_format = "0.0000"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{LAST}"
    widths = [10, 16, 16, 18, 16, 16, 14, 16, 16, 18, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- Metrics ----
    mt = wb.create_sheet("Metrics")
    mt.append(["Metric", "Value", "Formula (as entered)", "How to read it"])
    style_header(mt, 1, 4)
    B = f"Returns!$B${FIRST}:$B${LAST}"
    C = f"Returns!$C${FIRST}:$C${LAST}"
    E = f"Returns!$E${FIRST}:$E${LAST}"
    F = f"Returns!$F${FIRST}:$F${LAST}"
    G = f"Returns!$G${FIRST}:$G${LAST}"
    H = f"Returns!$H${FIRST}:$H${LAST}"
    I_ = f"Returns!$I${FIRST}:$I${LAST}"
    J = f"Returns!$J${FIRST}:$J${LAST}"
    K = f"Returns!$K${FIRST}:$K${LAST}"
    specs = [
        ("Months in sample", f"=COUNT({B})", "0", "Count of monthly observations."),
        ("Portfolio cumulative return", f"=Returns!H{LAST}-1", "0.00%", "Final wealth minus 1."),
        ("Benchmark cumulative return", f"=Returns!I{LAST}-1", "0.00%", "Final wealth minus 1."),
        ("Portfolio annualised return", f"=Returns!H{LAST}^(12/COUNT({B}))-1", "0.00%", "Geometric annualisation over the sample."),
        ("Benchmark annualised return", f"=Returns!I{LAST}^(12/COUNT({C}))-1", "0.00%", "Geometric annualisation over the sample."),
        ("Portfolio volatility (monthly)", f"=STDEV({B})", "0.00%", "Sample standard deviation of monthly returns."),
        ("Portfolio volatility (annualised)", f"=STDEV({B})*SQRT(12)", "0.00%", "Monthly volatility x sqrt(12)."),
        ("Benchmark volatility (annualised)", f"=STDEV({C})*SQRT(12)", "0.00%", "Monthly volatility x sqrt(12)."),
        ("Portfolio Sharpe ratio", f"=AVERAGE({E})/STDEV({B})*SQRT(12)", "0.00", "Mean excess return / volatility, annualised."),
        ("Benchmark Sharpe ratio", f"=AVERAGE({F})/STDEV({C})*SQRT(12)", "0.00", "Mean excess return / volatility, annualised."),
        ("Beta (vs benchmark)", f"=SLOPE({B},{C})", "0.00", "Sensitivity to benchmark moves; 1.0 = market-like."),
        ("Alpha (monthly)", f"=AVERAGE({E})-SLOPE({B},{C})*AVERAGE({F})", "0.00%", "Mean excess return not explained by beta."),
        ("Alpha (annualised)", f"=(AVERAGE({E})-SLOPE({B},{C})*AVERAGE({F}))*12", "0.00%", "Monthly alpha x 12."),
        ("Correlation (vs benchmark)", f"=CORREL({B},{C})", "0.00", "Co-movement, -1 to +1."),
        ("Portfolio max drawdown", f"=MIN({J})", "0.00%", "Worst peak-to-trough fall in portfolio wealth."),
        ("Benchmark max drawdown", f"=MIN({K})", "0.00%", "Worst peak-to-trough fall in benchmark wealth."),
        ("Best month (portfolio)", f"=MAX({B})", "0.00%", "Highest single monthly return."),
        ("Worst month (portfolio)", f"=MIN({B})", "0.00%", "Lowest single monthly return."),
        ("Win rate (portfolio)", f"=COUNTIF({B},\">0\")/COUNT({B})", "0.00%", "Share of months with a positive return."),
        ("Tracking error (annualised)", f"=STDEV({G})*SQRT(12)", "0.00%", "Volatility of active returns, annualised."),
        ("Information ratio", f"=AVERAGE({G})/STDEV({G})*SQRT(12)", "0.00", "Mean active return / tracking error, annualised."),
    ]
    for i, (label, formula, fmt, note) in enumerate(specs):
        r = 2 + i
        mt.cell(row=r, column=1, value=label)
        c = mt.cell(row=r, column=2, value=formula)
        c.number_format = fmt
        mt.cell(row=r, column=3, value=formula)
        mt.cell(row=r, column=4, value=note)
    mt.column_dimensions["A"].width = 34
    mt.column_dimensions["B"].width = 16
    mt.column_dimensions["C"].width = 46
    mt.column_dimensions["D"].width = 52
    # Metric row lookup for Dashboard: row = 2 + index
    idx = {label: 2 + i for i, (label, *_rest) in enumerate(specs)}

    # ---- Dashboard ----
    db = wb.create_sheet("Dashboard")
    db.append(["Portfolio Performance Dashboard — fictional demo data, Jan 2025-Oct 2026"])
    db["A1"].font = Font(bold=True, size=14, color=NAVY)
    db.append(["KPI", "Value"])
    style_header(db, 2, 2)
    kpis = [
        ("Portfolio cumulative return", idx["Portfolio cumulative return"], "0.00%"),
        ("Benchmark cumulative return", idx["Benchmark cumulative return"], "0.00%"),
        ("Portfolio annualised return", idx["Portfolio annualised return"], "0.00%"),
        ("Portfolio volatility (annualised)", idx["Portfolio volatility (annualised)"], "0.00%"),
        ("Portfolio Sharpe ratio", idx["Portfolio Sharpe ratio"], "0.00"),
        ("Beta (vs benchmark)", idx["Beta (vs benchmark)"], "0.00"),
        ("Alpha (annualised)", idx["Alpha (annualised)"], "0.00%"),
        ("Portfolio max drawdown", idx["Portfolio max drawdown"], "0.00%"),
        ("Information ratio", idx["Information ratio"], "0.00"),
        ("Win rate (portfolio)", idx["Win rate (portfolio)"], "0.00%"),
    ]
    for i, (label, mrow, fmt) in enumerate(kpis):
        r = 3 + i
        db.cell(row=r, column=1, value=label)
        c = db.cell(row=r, column=2, value=f"=Metrics!B{mrow}")
        c.number_format = fmt
    db.column_dimensions["A"].width = 36
    db.column_dimensions["B"].width = 16

    ret = wb["Returns"]
    cats = Reference(ret, min_col=1, min_row=FIRST, max_row=LAST)
    # Wealth growth line chart
    ch1 = LineChart()
    ch1.title = "Growth of 1.00 invested"
    ch1.y_axis.title = "Wealth"
    ch1.add_data(Reference(ret, min_col=8, min_row=1, max_row=LAST), titles_from_data=True)
    ch1.add_data(Reference(ret, min_col=9, min_row=1, max_row=LAST), titles_from_data=True)
    ch1.set_categories(cats)
    ch1.height, ch1.width = 8, 16
    db.add_chart(ch1, "D2")
    # Monthly returns column chart
    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "Monthly returns"
    ch2.y_axis.number_format = "0%"
    ch2.add_data(Reference(ret, min_col=2, min_row=1, max_row=LAST), titles_from_data=True)
    ch2.add_data(Reference(ret, min_col=3, min_row=1, max_row=LAST), titles_from_data=True)
    ch2.set_categories(cats)
    ch2.height, ch2.width = 8, 16
    db.add_chart(ch2, "D20")
    # Drawdown line chart
    ch3 = LineChart()
    ch3.title = "Drawdown from peak"
    ch3.y_axis.number_format = "0%"
    ch3.add_data(Reference(ret, min_col=10, min_row=1, max_row=LAST), titles_from_data=True)
    ch3.add_data(Reference(ret, min_col=11, min_row=1, max_row=LAST), titles_from_data=True)
    ch3.set_categories(cats)
    ch3.height, ch3.width = 8, 16
    db.add_chart(ch3, "D38")

    # ---- Checks ----
    ck = wb.create_sheet("Checks")
    ck.append(["Check", "Result"])
    style_header(ck, 1, 2)
    checks = [
        ("Months counted = 22", f"=IF(Metrics!B2=22,\"PASS\",\"FAIL\")"),
        ("All drawdowns <= 0", f"=IF(AND(MAX({J})<=0,MAX({K})<=0),\"PASS\",\"FAIL\")"),
        ("All wealth values > 0", f"=IF(MIN({H},{I_})>0,\"PASS\",\"FAIL\")"),
        ("Correlation within [-1, 1]", f"=IF(AND(Metrics!B{idx['Correlation (vs benchmark)']}>=-1,Metrics!B{idx['Correlation (vs benchmark)']}<=1),\"PASS\",\"FAIL\")"),
        ("Win rate within [0, 1]", f"=IF(AND(Metrics!B{idx['Win rate (portfolio)']}>=0,Metrics!B{idx['Win rate (portfolio)']}<=1),\"PASS\",\"FAIL\")"),
        ("Final wealth > 0 and finite", f"=IF(AND(Returns!H{LAST}>0,Returns!I{LAST}>0),\"PASS\",\"FAIL\")"),
        ("Max drawdown is negative or zero", f"=IF(Metrics!B{idx['Portfolio max drawdown']}<=0,\"PASS\",\"FAIL\")"),
        ("Beta is finite and non-zero", f"=IF(AND(Metrics!B{idx['Beta (vs benchmark)']}<>0,Metrics!B{idx['Beta (vs benchmark)']}>-5,Metrics!B{idx['Beta (vs benchmark)']}<5),\"PASS\",\"FAIL\")"),
    ]
    for i, (label, formula) in enumerate(checks):
        r = 2 + i
        ck.cell(row=r, column=1, value=label)
        ck.cell(row=r, column=2, value=formula)
    ck.column_dimensions["A"].width = 40
    ck.column_dimensions["B"].width = 12

    # ---- Guide ----
    gd = wb.create_sheet("Guide", 0)
    gd.append(["PortfolioPulse — how to use this workbook"])
    gd["A1"].font = Font(bold=True, size=14, color=NAVY)
    lines = [
        "1. Paste your own monthly returns into Returns columns B (portfolio) and C (benchmark);",
        "   keep column D as the monthly risk-free rate. Every metric, KPI and chart updates.",
        "2. To add a month, insert a row inside the data block (not below it) and copy the",
        "   formulas down, then extend the Metrics ranges to cover the new last row.",
        "3. Methods: Sharpe = mean excess return / return volatility x sqrt(12); beta is the",
        "   SLOPE of portfolio on benchmark returns; alpha = mean excess - beta x mean",
        "   benchmark excess; drawdown = wealth / running peak wealth - 1.",
        "4. Demo data is fictional (seeded, deterministic). Nothing here is investment advice;",
        "   it is a reporting template, not a recommendation.",
    ]
    for i, line in enumerate(lines, 2):
        gd.cell(row=i, column=1, value=line)
    gd.column_dimensions["A"].width = 110
    return wb


def main():
    rows = build_data()
    with open(BASE / "data" / "returns.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "portfolio", "benchmark", "risk_free"])
        w.writerows(rows)
    wb = build_workbook(rows)
    wb.save(BASE / "portfolio-performance.xlsx")
    print(f"Wrote {len(rows)} months, {BASE / 'portfolio-performance.xlsx'}")


if __name__ == "__main__":
    main()
