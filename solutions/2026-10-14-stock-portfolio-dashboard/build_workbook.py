#!/usr/bin/env python3
"""Build stock-portfolio-dashboard.xlsx (openpyxl, plain A1 ranges, no Table objects).

Reads data/*.csv (see generate.py) and produces a genuine 8-sheet OOXML workbook
where every figure is a live formula. Numeric formula columns guard blanks to 0,
never "" (Excel/LO #VALUE! rule).
"""
import csv
import os
from copy import copy

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import LineChart, PieChart, BarChart, Reference
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "stock-portfolio-dashboard.xlsx")
N = 30  # stocks
W = 33  # weekly closes W0..W32

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
KPI_FONT = Font(bold=True, size=12, color="1F4E79")
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
INR = '"₹"#,##0.00'
INR0 = '"₹"#,##0'
PCT = '0.00%'
PCT1 = '0.0%'


def read_csv(name):
    with open(os.path.join(DATA, name), newline="") as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN


def main():
    companies = read_csv("companies.csv")
    holdings = read_csv("holdings.csv")
    prices = read_csv("prices.csv")
    weeks = read_csv("weeks.csv")
    index = read_csv("index.csv")
    comp = {r["ticker"]: r for r in companies}
    hold = {r["ticker"]: (int(r["qty"]), float(r["avg_cost"])) for r in holdings}
    price = {r["ticker"]: [float(r[f"w{i}"]) for i in range(W)] for r in prices}
    wk_label = [r["week_label"] for r in weeks]
    wk_date = [r["date"] for r in weeks]
    bench = [float(r["nifty_like"]) for r in index]
    tickers = [r["ticker"] for r in companies]
    sectors = sorted({r["sector"] for r in companies})
    last_col = get_column_letter(2 + W - 1)   # AH
    prev_col = get_column_letter(2 + W - 2)   # AG
    wk4_col = get_column_letter(2 + W - 5)    # AD

    wb = Workbook()

    # ============ Guide ============
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "StockPulse — AI-Enabled Stock Portfolio Dashboard (DEMO)"
    ws["A1"].font = TITLE_FONT
    guide_lines = [
        ("What this is", "A working demo of an AI-enabled Excel stock dashboard for a 30-stock Indian portfolio: "
         "live-formula holdings ledger, 33 weekly closes, weekly portfolio rollup, sector view, KPI dashboard with "
         "3 native charts, a rule-based weekly AI report, and integrity checks."),
        ("Data", "All data is FICTIONAL, generated from a fixed seed (data/*.csv). Nothing here is market data. "
         "Replace Prices with your real weekly closes and Holdings with your real qty/avg cost to go live."),
        ("Refresh", "1) Paste new weekly closes into the next column on Prices. 2) Update the Last-Price column refs "
         "(F on Holdings) to the new column. 3) Add a week row on Weekly. Everything else recalculates."),
        ("Weekly AI report", "AI_Report is written by a rule-based generator (build_workbook.py) at build time from "
         "the workbook's own numbers: movers, allocation drift, benchmark comparison, watchlist. In production this "
         "step is a scheduled LLM call over the same KPI block."),
        ("Integrity", "The Checks sheet re-verifies every headline number. If any check is FALSE, do not trust the dashboard."),
    ]
    r = 3
    for title, body in guide_lines:
        ws.cell(row=r, column=1, value=title).font = Font(bold=True)
        c = ws.cell(row=r + 1, column=1, value=body)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=8)
        ws.row_dimensions[r + 1].height = 45
        r += 2
    ws.column_dimensions["A"].width = 110
    ws.sheet_view.showGridLines = False

    # ============ Holdings ============
    ws = wb.create_sheet("Holdings")
    hdr = ["Ticker", "Company", "Sector", "Qty", "Avg Cost", "Last Price",
           "Invested", "Market Value", "P&L", "P&L %", "Allocation %"]
    ws.append(hdr)
    style_header(ws, 1, len(hdr))
    for i, t in enumerate(tickers, start=2):
        q, c_ = hold[t]
        ws.cell(row=i, column=1, value=t)
        ws.cell(row=i, column=2, value=comp[t]["company"])
        ws.cell(row=i, column=3, value=comp[t]["sector"])
        ws.cell(row=i, column=4, value=q)
        ws.cell(row=i, column=5, value=c_).number_format = INR
        ws.cell(row=i, column=6,
                value=f"=INDEX(Prices!${last_col}$2:${last_col}${N+1},MATCH(A{i},Prices!$A$2:$A${N+1},0))").number_format = INR
        ws.cell(row=i, column=7, value=f"=D{i}*E{i}").number_format = INR
        ws.cell(row=i, column=8, value=f"=D{i}*F{i}").number_format = INR
        ws.cell(row=i, column=9, value=f"=H{i}-G{i}").number_format = INR
        ws.cell(row=i, column=10, value=f"=IF(G{i}=0,0,I{i}/G{i})").number_format = PCT1
        ws.cell(row=i, column=11, value=f"=IF($H${N+2}=0,0,H{i}/$H${N+2})").number_format = PCT1
    tr = N + 2
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, letter in [(7, "G"), (8, "H"), (9, "I")]:
        ws.cell(row=tr, column=col, value=f"=SUM({letter}2:{letter}{N+1})").number_format = INR
        ws.cell(row=tr, column=col).font = Font(bold=True)
    ws.cell(row=tr, column=10, value=f"=IF(G{tr}=0,0,I{tr}/G{tr})").number_format = PCT1
    ws.cell(row=tr, column=10).font = Font(bold=True)
    ws.cell(row=tr, column=11, value=f"=SUM(K2:K{N+1})").number_format = PCT1
    ws.cell(row=tr, column=11).font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{N+1}"
    widths = [10, 28, 16, 8, 12, 12, 14, 14, 14, 9, 12]
    for i, wd in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = wd
    ws.conditional_formatting.add(f"J2:J{tr}",
                                 CellIsRule(operator="lessThan", formula=["0"], font=Font(color="9C0006")))
    ws.conditional_formatting.add(f"J2:J{tr}",
                                 CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="006100")))

    # ============ Prices ============
    ws = wb.create_sheet("Prices")
    ws.append(["Ticker"] + wk_label + ["Last Wk Return %", "4W Return %"])
    style_header(ws, 1, 1 + W + 2)
    for i, t in enumerate(tickers, start=2):
        ws.cell(row=i, column=1, value=t)
        for j, p in enumerate(price[t]):
            c = ws.cell(row=i, column=2 + j, value=p)
            c.number_format = INR
        # AI = Last Wk Return %, AJ = 4W Return %
        ws.cell(row=i, column=2 + W, value=f"={last_col}{i}/{prev_col}{i}-1").number_format = PCT1
        ws.cell(row=i, column=3 + W, value=f"={last_col}{i}/{wk4_col}{i}-1").number_format = PCT1
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(1 + W + 2)}1"
    ws.column_dimensions["A"].width = 12
    for j in range(W):
        ws.column_dimensions[get_column_letter(2 + j)].width = 11

    # ============ Weekly ============
    ws = wb.create_sheet("Weekly")
    ws.append(["Week", "Date", "Portfolio Value", "Weekly Return", "Drawdown", "Benchmark"])
    style_header(ws, 1, 6)
    for i in range(W):
        r = i + 2
        pcol = get_column_letter(2 + i)
        ws.cell(row=r, column=1, value=wk_label[i])
        ws.cell(row=r, column=2, value=wk_date[i])
        ws.cell(row=r, column=3,
                value=f"=SUMPRODUCT(Holdings!$D$2:$D${N+1},Prices!${pcol}$2:${pcol}${N+1})").number_format = INR0
        if i == 0:
            ws.cell(row=r, column=4, value=0).number_format = PCT1
        else:
            ws.cell(row=r, column=4,
                    value=f"=IF(C{r-1}=0,0,C{r}/C{r-1}-1)").number_format = PCT1
        ws.cell(row=r, column=5, value=f"=C{r}/MAX($C$2:C{r})-1").number_format = PCT1
        ws.cell(row=r, column=6, value=bench[i]).number_format = INR0
    ws.freeze_panes = "A2"
    for wd, col in [(8, "A"), (12, "B"), (17, "C"), (13, "D"), (11, "E"), (13, "F")]:
        ws.column_dimensions[col].width = wd

    # ============ Sector ============
    ws = wb.create_sheet("Sector")
    ws.append(["Sector", "Stocks", "Invested", "Market Value", "P&L", "P&L %", "Allocation %"])
    style_header(ws, 1, 7)
    srows = len(sectors)
    for i, s in enumerate(sectors, start=2):
        ws.cell(row=i, column=1, value=s)
        ws.cell(row=i, column=2, value=f'=COUNTIF(Holdings!$C$2:$C${N+1},A{i})')
        ws.cell(row=i, column=3, value=f'=SUMIF(Holdings!$C$2:$C${N+1},A{i},Holdings!$G$2:$G${N+1})').number_format = INR
        ws.cell(row=i, column=4, value=f'=SUMIF(Holdings!$C$2:$C${N+1},A{i},Holdings!$H$2:$H${N+1})').number_format = INR
        ws.cell(row=i, column=5, value=f'=SUMIF(Holdings!$C$2:$C${N+1},A{i},Holdings!$I$2:$I${N+1})').number_format = INR
        ws.cell(row=i, column=6, value=f"=IF(C{i}=0,0,E{i}/C{i})").number_format = PCT1
        ws.cell(row=i, column=7, value=f"=IF($D${srows+2}=0,0,D{i}/$D${srows+2})").number_format = PCT1
    tr = srows + 2
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for col, letter in [(3, "C"), (4, "D"), (5, "E")]:
        ws.cell(row=tr, column=col, value=f"=SUM({letter}2:{letter}{srows+1})").number_format = INR
        ws.cell(row=tr, column=col).font = Font(bold=True)
    ws.cell(row=tr, column=6, value=f"=IF(C{tr}=0,0,E{tr}/C{tr})").number_format = PCT1
    ws.cell(row=tr, column=7, value=f"=SUM(G2:G{srows+1})").number_format = PCT1
    for wd, col in [(16, "A"), (8, "B"), (14, "C"), (14, "D"), (14, "E"), (9, "F"), (12, "G")]:
        ws.column_dimensions[col].width = wd

    # ============ Dashboard ============
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Portfolio Dashboard"
    ws["A1"].font = TITLE_FONT
    kpis = [
        ("Total Invested", f"=SUM(Holdings!G2:G{N+1})", INR0),
        ("Market Value", f"=SUM(Holdings!H2:H{N+1})", INR0),
        ("Total P&L", "=B4-B3", INR0),
        ("P&L %", "=IF(B3=0,0,B5/B3)", PCT1),
        ("Latest Weekly Return", f"=INDEX(Weekly!D2:D{W+1},{W})", PCT1),
        ("Best Performer", f'=INDEX(Holdings!A2:A{N+1},MATCH(MAX(Holdings!J2:J{N+1}),Holdings!J2:J{N+1},0))', None),
        ("Best P&L %", f"=MAX(Holdings!J2:J{N+1})", PCT1),
        ("Worst Performer", f'=INDEX(Holdings!A2:A{N+1},MATCH(MIN(Holdings!J2:J{N+1}),Holdings!J2:J{N+1},0))', None),
        ("Worst P&L %", f"=MIN(Holdings!J2:J{N+1})", PCT1),
        ("Largest Holding", f'=INDEX(Holdings!A2:A{N+1},MATCH(MAX(Holdings!H2:H{N+1}),Holdings!H2:H{N+1},0))', None),
        ("Largest Allocation", f"=MAX(Holdings!K2:K{N+1})", PCT1),
        ("Active Stocks", f"=COUNTA(Holdings!A2:A{N+1})", None),
        ("Positive Weeks", f'=COUNTIF(Weekly!D3:D{W+1},">0")', None),
    ]
    r = 3
    for label, formula, fmt in kpis:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=formula)
        if fmt:
            c.number_format = fmt
        c.font = KPI_FONT
        r += 1
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 20
    ws.sheet_view.showGridLines = False

    line = LineChart()
    line.title = "Portfolio Value — 33 Weeks"
    line.style = 10
    line.height, line.width = 7.5, 15
    data = Reference(wb["Weekly"], min_col=3, min_row=1, max_row=W + 1)
    cats = Reference(wb["Weekly"], min_col=1, min_row=2, max_row=W + 1)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    ws.add_chart(line, "D3")

    pie = PieChart()
    pie.title = "Sector Allocation"
    pie.height, pie.width = 7.5, 13
    pie.add_data(Reference(wb["Sector"], min_col=7, min_row=1, max_row=srows + 1), titles_from_data=True)
    pie.set_categories(Reference(wb["Sector"], min_col=1, min_row=2, max_row=srows + 1))
    ws.add_chart(pie, "D19")

    bar = BarChart()
    bar.type = "col"
    bar.title = "P&L % by Holding"
    bar.style = 10
    bar.height, bar.width = 7.5, 15
    bar.add_data(Reference(wb["Holdings"], min_col=10, min_row=1, max_row=N + 1), titles_from_data=True)
    bar.set_categories(Reference(wb["Holdings"], min_col=1, min_row=2, max_row=N + 1))
    ws.add_chart(bar, "D35")

    # ============ AI_Report (rule-based weekly generator) ============
    inv = sum(q * c for q, c in hold.values())
    last_px = {t: price[t][-1] for t in tickers}
    mv = sum(hold[t][0] * last_px[t] for t in tickers)
    pnl_pct = (mv - inv) / inv
    pvals = [sum(hold[t][0] * price[t][i] for t in tickers) for i in range(W)]
    wk_ret = pvals[-1] / pvals[-2] - 1
    bench_ret = bench[-1] / bench[-2] - 1
    pnl_by_t = {t: (last_px[t] / hold[t][1] - 1) for t in tickers}
    top3 = sorted(pnl_by_t, key=pnl_by_t.get, reverse=True)[:3]
    bot3 = sorted(pnl_by_t, key=pnl_by_t.get)[:3]
    alloc = {t: hold[t][0] * last_px[t] / mv for t in tickers}
    big = max(alloc, key=alloc.get)
    pos_weeks = sum(1 for i in range(1, W) if pvals[i] > pvals[i - 1])
    peak = pvals[0]
    max_dd = 0.0
    for v in pvals:
        peak = max(peak, v)
        max_dd = min(max_dd, v / peak - 1)
    sec_mv = {}
    for t in tickers:
        sec_mv[comp[t]["sector"]] = sec_mv.get(comp[t]["sector"], 0) + hold[t][0] * last_px[t]
    big_sec = max(sec_mv, key=sec_mv.get)

    def rs(x):
        return f"₹{x:,.0f}"

    ws = wb.create_sheet("AI_Report")
    ws["A1"] = f"Weekly AI Report — {wk_label[-1]} ({wk_date[-1]})"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Generated by StockPulse's rule-based weekly-report engine from this workbook's own numbers "
                "(demo of the AI weekly-reporting step; production = scheduled LLM call over the same KPI block).")
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 30

    paras = [
        ("Portfolio this week",
         f"Portfolio {'gained' if wk_ret >= 0 else 'fell'} {abs(wk_ret):.1%} this week to {rs(pvals[-1])}. "
         f"Since inception the book is {'up' if pnl_pct >= 0 else 'down'} {abs(pnl_pct):.1%} "
         f"({rs(mv - inv)} on {rs(inv)} invested). The fictional benchmark moved {bench_ret:+.1%} over the same week, "
         f"so the portfolio {'outperformed' if wk_ret - bench_ret >= 0 else 'underperformed'} by {abs(wk_ret - bench_ret):.1%}. "
         f"{pos_weeks} of the last {W - 1} weeks were positive; maximum drawdown to date is {max_dd:.1%}."),
        ("Movers",
         "Top contributors: " + ", ".join(f"{t} ({pnl_by_t[t]:+.1%})" for t in top3) + ". "
         "Biggest drags: " + ", ".join(f"{t} ({pnl_by_t[t]:+.1%})" for t in bot3) + ". " +
         ("Winners are doing the heavy lifting — consider trimming into strength. "
          if pnl_by_t[top3[0]] > 0.25 else "No extreme outliers this week; moves are broad-based. ")),
        ("Allocation & concentration",
         f"Largest position is {big} at {alloc[big]:.1%} of the book; largest sector is {big_sec} "
         f"at {sec_mv[big_sec] / mv:.1%}. " +
         ("Concentration is elevated — a single-name shock would move the needle; consider rebalancing toward target weights. "
          if alloc[big] > 0.12 else "Concentration looks reasonable; no single name dominates.")),
        ("Watchlist (rules)",
         " • ".join(
             [f"Review {t}: down {abs(pnl_by_t[t]):.0%} from cost — thesis check" for t in bot3 if pnl_by_t[t] < -0.15]
             + [f"Consider booking partial profit in {t} ({pnl_by_t[t]:+.0%})" for t in top3 if pnl_by_t[t] > 0.30]
             or ["No rule-based alerts this week."]
         )),
    ]
    r = 4
    for title, body in paras:
        ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=11)
        c = ws.cell(row=r + 1, column=1, value=body)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=8)
        ws.row_dimensions[r + 1].height = 60
        r += 2
    ws["A13"] = "Key figures (live formulas — always in sync with the dashboard)"
    ws["A13"].font = Font(bold=True)
    fig = [("Market Value", "=Dashboard!B4", INR0), ("Total P&L", "=Dashboard!B5", INR0),
           ("P&L %", "=Dashboard!B6", PCT1), ("Latest Weekly Return", "=Dashboard!B7", PCT1)]
    rr = 14
    for label, f_, fmt in fig:
        ws.cell(row=rr, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=rr, column=2, value=f_)
        c.number_format = fmt
        c.font = KPI_FONT
        rr += 1
    ws.column_dimensions["A"].width = 110
    ws.sheet_view.showGridLines = False

    # ============ Checks ============
    ws = wb.create_sheet("Checks")
    ws.append(["#", "Check", "Formula", "Expected", "Status"])
    style_header(ws, 1, 5)
    checks = [
        ("Allocations sum to 100%", f"=ABS(SUM(Holdings!K2:K{N+1})-1)<0.0001", "TRUE"),
        ("Dashboard MV = sum of Holdings MV", f"=ABS(Dashboard!B4-SUM(Holdings!H2:H{N+1}))<0.01", "TRUE"),
        ("P&L = MV − Invested", "=ABS(Dashboard!B5-(Dashboard!B4-Dashboard!B3))<0.01", "TRUE"),
        ("All 30 last prices resolve (>0)", f'=COUNTIF(Holdings!F2:F{N+1},">0")=30', "TRUE"),
        ("Weekly rollup ties to Holdings MV", f"=ABS(INDEX(Weekly!C2:C{W+1},{W})-Dashboard!B4)<0.01", "TRUE"),
        ("Sector MV sums to portfolio MV", f"=ABS(SUM(Sector!D2:D{srows+1})-Dashboard!B4)<0.01", "TRUE"),
        ("No negative quantities", f'=COUNTIF(Holdings!D2:D{N+1},"<0")=0', "TRUE"),
        ("Weekly return formula consistent", f"=ABS((INDEX(Weekly!C2:C{W+1},{W})/INDEX(Weekly!C2:C{W+1},{W-1})-1)-INDEX(Weekly!D2:D{W+1},{W}))<0.000001", "TRUE"),
    ]
    for i, (label, formula, _exp) in enumerate(checks, start=2):
        ws.cell(row=i, column=1, value=i - 1)
        ws.cell(row=i, column=2, value=label)
        ws.cell(row=i, column=3, value=f"'{formula}")
        ws.cell(row=i, column=4, value="TRUE")
        ws.cell(row=i, column=5, value=f'=IF({formula[1:]},"PASS","FAIL")')
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 62
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10

    wb.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
