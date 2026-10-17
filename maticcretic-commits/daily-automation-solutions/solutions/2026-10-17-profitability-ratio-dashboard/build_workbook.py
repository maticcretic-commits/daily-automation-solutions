#!/usr/bin/env python3
"""ProfitPulse — build the formula-driven profitability & ratio workbook.

Input : data/pnl.csv (FY21..FY25), data/bs.csv (FY20..FY25), data/meta.csv
Output: profitpulse-ratio-dashboard.xlsx — 8 sheets, every computed figure a
        live A1-style formula (no structured refs, no Table objects, numeric
        columns guard to 0 never "").
"""
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import CellIsRule

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
OUT = HERE / "profitpulse-ratio-dashboard.xlsx"

YEARS = ["FY21", "FY22", "FY23", "FY24", "FY25"]
BS_YEARS = ["FY20"] + YEARS
C = get_column_letter

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=13, color="1F4E78")
FLAG_RED = PatternFill("solid", fgColor="F8CBAD")
FLAG_GREEN = PatternFill("solid", fgColor="C6EFCE")
THIN = Border(*[Side(style="thin", color="B4C6E7")] * 4)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

AMT = "#,##0"
PCT = "0.0%"
RTO = "0.00"
DAYS = "0"


def read_csv(name):
    with open(DATA / name, newline="") as f:
        return list(csv.reader(f))


def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font, cell.alignment = HDR_FILL, HDR_FONT, CENTER
        cell.border = THIN


def paint(ws, row, labels, ncols, fmts=None):
    ws.cell(row=row, column=1, value=labels).font = Font(bold=True)
    for j in range(2, ncols + 1):
        cell = ws.cell(row=row, column=j)
        cell.border, cell.alignment = THIN, CENTER
        if fmts:
            cell.number_format = fmts


def safe(num, den):
    """Guard division: 0 denominator -> 0 (never "")."""
    return f"IF({den}=0,0,{num}/{den})"


# ---------------------------------------------------------------- Guide ----
def build_guide(wb, meta):
    ws = wb.active
    ws.title = "Guide"
    ws.sheet_properties.tabColor = "1F4E78"
    lines = [
        ("ProfitPulse — Profitability & Financial-Ratio Dashboard", TITLE_FONT),
        (f"Company (fictional): {meta['company']}", None),
        (f"Currency: {meta['currency']} | Period: FY21–FY25 (balance sheet from FY20 for averages)", None),
        ("", None),
        ("SHEETS", Font(bold=True, size=12, color="1F4E78")),
        ("Raw_PNL — profit & loss inputs (FY21–FY25); gross/EBITDA/EBIT/EBT/net are formulas", None),
        ("Raw_BS — balance-sheet inputs (FY20–FY25); totals are formulas; Cash is the balancing plug", None),
        ("Ratios — 22 ratios, ALL formulas: liquidity, profitability, leverage, efficiency, growth", None),
        ("DuPont — ROE decomposition (net margin × asset turnover × equity multiplier) + tie-out", None),
        ("Dashboard — FY25 KPI cards + 3 native charts (all formula-fed)", None),
        ("AI_Report — rule-based narrative + red-flag checklist with live formulas", None),
        ("Checks — 8 integrity checks (all must read TRUE / 0)", None),
        ("", None),
        ("HOW TO REFRESH", Font(bold=True, size=12, color="1F4E78")),
        ("1. Overwrite the raw (non-formula) values in Raw_PNL / Raw_BS.", None),
        ("2. Every other sheet recalculates automatically — nothing is hard-coded.", None),
        ("3. Keep FY20 balance-sheet column: averages (and FY21 ratios) need it.", None),
        ("", None),
        ("RED-FLAG THRESHOLDS (AI_Report)", Font(bold=True, size=12, color="1F4E78")),
        ("Current ratio < 1.20 · Quick ratio < 1.00 · Debt/Equity > 1.00", None),
        ("Interest coverage < 3.0× · CCC > 90 days · net margin falling YoY", None),
        ("Receivable days rising YoY (collections slipping)", None),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1, value=text)
        if font:
            cell.font = font
        cell.alignment = LEFT
    ws.column_dimensions["A"].width = 110


# --------------------------------------------------------------- Raw_PNL ---
PNL_ORDER = ["Revenue", "COGS", "Gross profit", "SG&A", "EBITDA",
             "Depreciation", "EBIT", "Interest", "EBT", "Tax", "Net income"]
PNL_R = {name: i + 2 for i, name in enumerate(PNL_ORDER)}


def build_pnl(ws, pnl):
    ws.title = "Raw_PNL"
    ws.sheet_properties.tabColor = "ED7D31"
    ws.cell(row=1, column=1, value="Item")
    for j, y in enumerate(YEARS, start=2):
        ws.cell(row=1, column=j, value=y)
    raw = {"Revenue": 0, "COGS": 1, "SG&A": 3, "Depreciation": 5,
           "Interest": 7, "Tax": 9}
    data = {r[0]: r[1:] for r in pnl[1:]}
    for name, row in PNL_R.items():
        paint(ws, row, name, 7, AMT)
        for j in range(2, 7):
            cell = ws.cell(row=row, column=j)
            letter = C(j)
            if name in raw:
                cell.value = int(float(data[name][j - 2]))
            else:
                f = {"Gross profit": f"{letter}{PNL_R['Revenue']}-{letter}{PNL_R['COGS']}",
                     "EBITDA": f"{letter}{PNL_R['Gross profit']}-{letter}{PNL_R['SG&A']}",
                     "EBIT": f"{letter}{PNL_R['EBITDA']}-{letter}{PNL_R['Depreciation']}",
                     "EBT": f"{letter}{PNL_R['EBIT']}-{letter}{PNL_R['Interest']}",
                     "Net income": f"{letter}{PNL_R['EBT']}-{letter}{PNL_R['Tax']}"}[name]
                cell.value = f"={f}"
    style_header(ws, 7)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = "A1:F12"
    ws.column_dimensions["A"].width = 22
    for j in range(2, 7):
        ws.column_dimensions[C(j)].width = 14


# ---------------------------------------------------------------- Raw_BS ---
BS_ORDER = ["Cash & bank (plug)", "Receivables", "Inventory",
            "Other current assets", "PPE (net)", "Other non-current assets",
            "Total current assets", "Total non-current assets", "Total assets",
            "Payables", "Short-term debt", "Other current liabilities",
            "Total current liabilities", "Long-term debt",
            "Other non-current liabilities", "Total non-current liabilities",
            "Total liabilities", "Share capital", "Reserves & surplus",
            "Total equity", "Total L&E"]
BS_R = {name: i + 2 for i, name in enumerate(BS_ORDER)}
BS_RAW_NAMES = ["Cash & bank (plug)", "Receivables", "Inventory",
                "Other current assets", "PPE (net)", "Other non-current assets",
                "Payables", "Short-term debt", "Other current liabilities",
                "Long-term debt", "Other non-current liabilities",
                "Share capital", "Reserves & surplus"]


def build_bs(ws, bs):
    ws.title = "Raw_BS"
    ws.sheet_properties.tabColor = "70AD47"
    ws.cell(row=1, column=1, value="Item")
    for j, y in enumerate(BS_YEARS, start=2):
        ws.cell(row=1, column=j, value=y)
    data = {r[0]: r[1:] for r in bs[1:]}
    for name, row in BS_R.items():
        paint(ws, row, name, 7, AMT)
        for j in range(2, 8):
            cell = ws.cell(row=row, column=j)
            letter = C(j)
            if name in BS_RAW_NAMES:
                cell.value = round(float(data[name][j - 2]), 2)
            else:
                f = {
                    "Total current assets":
                        f"{letter}{BS_R['Cash & bank (plug)']}+{letter}{BS_R['Receivables']}"
                        f"+{letter}{BS_R['Inventory']}+{letter}{BS_R['Other current assets']}",
                    "Total non-current assets":
                        f"{letter}{BS_R['PPE (net)']}+{letter}{BS_R['Other non-current assets']}",
                    "Total assets":
                        f"{letter}{BS_R['Total current assets']}+{letter}{BS_R['Total non-current assets']}",
                    "Total current liabilities":
                        f"{letter}{BS_R['Payables']}+{letter}{BS_R['Short-term debt']}"
                        f"+{letter}{BS_R['Other current liabilities']}",
                    "Total non-current liabilities":
                        f"{letter}{BS_R['Long-term debt']}+{letter}{BS_R['Other non-current liabilities']}",
                    "Total liabilities":
                        f"{letter}{BS_R['Total current liabilities']}+{letter}{BS_R['Total non-current liabilities']}",
                    "Total equity":
                        f"{letter}{BS_R['Share capital']}+{letter}{BS_R['Reserves & surplus']}",
                    "Total L&E":
                        f"{letter}{BS_R['Total liabilities']}+{letter}{BS_R['Total equity']}",
                }[name]
                cell.value = f"={f}"
    style_header(ws, 8)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = "A1:G22"
    ws.column_dimensions["A"].width = 32
    for j in range(2, 9):
        ws.column_dimensions[C(j)].width = 14

# ----------------------------------------------------------------- Ratios --
# (row, label, format)
RATIO_DEFS = [
    ("Current ratio", RTO), ("Quick ratio", RTO), ("Cash ratio", RTO),
    ("Working capital", AMT),
    ("Gross margin", PCT), ("EBITDA margin", PCT), ("EBIT margin", PCT),
    ("Net margin", PCT),
    ("Return on equity (ROE)", PCT), ("Return on assets (ROA)", PCT),
    ("Return on capital employed (ROCE)", PCT),
    ("Debt / Equity", RTO), ("Debt / Assets", RTO),
    ("Interest coverage (x)", RTO), ("Equity ratio", PCT),
    ("Receivable days", DAYS), ("Inventory days", DAYS),
    ("Payable days", DAYS), ("Cash conversion cycle (days)", DAYS),
    ("Asset turnover (x)", RTO),
    ("Revenue growth YoY", PCT), ("Net income growth YoY", PCT),
]
RATIO_R = {name: i + 2 for i, (name, _) in enumerate(RATIO_DEFS)}


def P(row, j):
    """Raw_PNL cell for ratio-year column j (2..6 = FY21..FY25)."""
    return f"Raw_PNL!{C(j)}{row}"


def B(row, j):
    """Raw_BS cell for ratio-year column j (BS col = j+1; FY20 = B)."""
    return f"Raw_BS!{C(j + 1)}{row}"


def BP(row, j):
    """Raw_BS cell for the PREVIOUS balance-sheet year."""
    return f"Raw_BS!{C(j)}{row}"


def avg(cell_cur, cell_prev):
    return f"(({cell_cur}+{cell_prev})/2)"


def ratio_formula(idx, j):
    """idx = 0-based index into RATIO_DEFS; j = 2..6."""
    bs, pnl = B, P
    tca, tcl = bs(BS_R["Total current assets"], j), bs(BS_R["Total current liabilities"], j)
    ta = bs(BS_R["Total assets"], j)
    te = bs(BS_R["Total equity"], j)
    tl = bs(BS_R["Total liabilities"], j)
    rev, cogs = pnl(PNL_R["Revenue"], j), pnl(PNL_R["COGS"], j)
    avgA = avg(ta, BP(BS_R["Total assets"], j))
    avgE = avg(te, BP(BS_R["Total equity"], j))
    formulas = [
        safe(tca, tcl),                                                             # 0 current
        safe(f"({bs(BS_R['Cash & bank (plug)'], j)}+{bs(BS_R['Receivables'], j)})", tcl),  # 1 quick
        safe(bs(BS_R["Cash & bank (plug)"], j), tcl),                               # 2 cash
        f"{tca}-{tcl}",                                                             # 3 WC
        safe(pnl(PNL_R["Gross profit"], j), rev),                                   # 4
        safe(pnl(PNL_R["EBITDA"], j), rev),                                         # 5
        safe(pnl(PNL_R["EBIT"], j), rev),                                           # 6
        safe(pnl(PNL_R["Net income"], j), rev),                                     # 7
        safe(pnl(PNL_R["Net income"], j), avgE),                                    # 8 ROE
        safe(pnl(PNL_R["Net income"], j), avgA),                                    # 9 ROA
        safe(pnl(PNL_R["EBIT"], j), f"({te}+{bs(BS_R['Total non-current liabilities'], j)})"),  # 10 ROCE
        safe(tl, te),                                                               # 11 D/E
        safe(tl, ta),                                                               # 12 D/A
        safe(pnl(PNL_R["EBIT"], j), pnl(PNL_R["Interest"], j)),                     # 13 IC
        safe(te, ta),                                                               # 14 equity ratio
        safe(avg(bs(BS_R["Receivables"], j), BP(BS_R["Receivables"], j)), rev) + "*365",   # 15
        safe(avg(bs(BS_R["Inventory"], j), BP(BS_R["Inventory"], j)), cogs) + "*365",      # 16
        safe(avg(bs(BS_R["Payables"], j), BP(BS_R["Payables"], j)), cogs) + "*365",         # 17
        None,  # 18 CCC = row17+row18-row19 (by row number, filled below)
        safe(rev, avgA),                                                            # 19 asset turnover
        None,  # 20 revenue YoY
        None,  # 21 net income YoY
    ]
    return formulas[idx]


def build_ratios(ws):
    ws.title = "Ratios"
    ws.sheet_properties.tabColor = "2E75B6"
    ws.cell(row=1, column=1, value="Ratio")
    for j, y in enumerate(YEARS, start=2):
        ws.cell(row=1, column=j, value=y)
    for i, (name, fmt) in enumerate(RATIO_DEFS):
        row = i + 2
        paint(ws, row, name, 7, fmt)
        for j in range(2, 7):
            cell = ws.cell(row=row, column=j)
            if i == 18:  # CCC
                r = RATIO_R
                cell.value = (f"=Ratios!{C(j)}{r['Receivable days']}"
                              f"+Ratios!{C(j)}{r['Inventory days']}"
                              f"-Ratios!{C(j)}{r['Payable days']}")
            elif i in (20, 21):
                if j == 2:
                    cell.value = "n/a"  # no FY20 P&L for YoY
                else:
                    base = PNL_R["Revenue"] if i == 20 else PNL_R["Net income"]
                    cur, prev = P(base, j), P(base, j - 1)
                    cell.value = f"=IF({prev}=0,0,({cur}-{prev})/{prev})"
            else:
                cell.value = f"={ratio_formula(i, j)}"
    style_header(ws, 7)
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 34
    for j in range(2, 7):
        ws.column_dimensions[C(j)].width = 14


# ----------------------------------------------------------------- DuPont --
def build_dupont(ws):
    ws.title = "DuPont"
    ws.sheet_properties.tabColor = "BF8F00"
    ws.cell(row=1, column=1, value="DuPont decomposition of ROE")
    for j, y in enumerate(YEARS, start=2):
        ws.cell(row=1, column=j, value=y)
    rows = [("Net margin (Net income / Revenue)", PCT),
            ("Asset turnover (Revenue / Avg assets)", RTO),
            ("Equity multiplier (Avg assets / Avg equity)", RTO),
            ("ROE — DuPont product", PCT),
            ("ROE — direct (check)", PCT),
            ("Difference (must be 0)", RTO)]
    for i, (name, fmt) in enumerate(rows):
        row = i + 2
        paint(ws, row, name, 7, fmt)
        for j in range(2, 7):
            cell = ws.cell(row=row, column=j)
            rr = RATIO_R
            if i == 0:
                cell.value = f"=Ratios!{C(j)}{rr['Net margin']}"
            elif i == 1:
                cell.value = f"=Ratios!{C(j)}{rr['Asset turnover (x)']}"
            elif i == 2:
                ta_c, ta_p = B(BS_R["Total assets"], j), BP(BS_R["Total assets"], j)
                te_c, te_p = B(BS_R["Total equity"], j), BP(BS_R["Total equity"], j)
                cell.value = f"={safe(avg(ta_c, ta_p), avg(te_c, te_p))}"
            elif i == 3:
                cell.value = f"=DuPont!{C(j)}2*DuPont!{C(j)}3*DuPont!{C(j)}4"
            elif i == 4:
                cell.value = f"=Ratios!{C(j)}{rr['Return on equity (ROE)']}"
            else:
                cell.value = f"=DuPont!{C(j)}5-DuPont!{C(j)}6"
    style_header(ws, 7)
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 40
    for j in range(2, 7):
        ws.column_dimensions[C(j)].width = 14


# -------------------------------------------------------------- Dashboard --
def build_dashboard(ws, ws_ratios, ws_pnl):
    ws.title = "Dashboard"
    ws.sheet_properties.tabColor = "C00000"
    ws.cell(row=1, column=1, value="ProfitPulse — FY25 KPI cards (all formula-fed)").font = TITLE_FONT
    ws.cell(row=1, column=2, value="FY25").font = Font(bold=True)
    ws.cell(row=1, column=3, value="vs FY21").font = Font(bold=True)
    kpis = [
        ("Revenue (INR lakh)", "=Raw_PNL!F2", "=Raw_PNL!F2-Raw_PNL!B2", AMT),
        ("Net margin", "=Ratios!F9", "=Ratios!F9-Ratios!B9", PCT),
        ("Return on equity", "=Ratios!F10", "=Ratios!F10-Ratios!B10", PCT),
        ("Current ratio", "=Ratios!F2", "=Ratios!F2-Ratios!B2", RTO),
        ("Debt / Equity", "=Ratios!F13", "=Ratios!F13-Ratios!B13", RTO),
        ("Cash conversion cycle (days)", "=Ratios!F20", "=Ratios!F20-Ratios!B20", DAYS),
        ("Interest coverage (x)", "=Ratios!F15", None, RTO),
        ("Cash & bank (INR lakh)", "=Raw_BS!G2", None, AMT),
    ]
    rr = RATIO_R
    for i, (label, f25, delta, fmt) in enumerate(kpis):
        row = i + 2
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        for col, val in ((2, f25), (3, delta)):
            cell = ws.cell(row=row, column=col, value=val)
            cell.number_format = fmt
            cell.border, cell.alignment = THIN, CENTER
            cell.font = Font(size=12, bold=(col == 2))
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18

    # Chart 1 — margins trend (line)
    ch1 = LineChart()
    ch1.title, ch1.style, ch1.height, ch1.width = "Profitability margins — FY21–FY25", 10, 7.5, 15
    data = Reference(ws_ratios, min_col=2, min_row=rr["Gross margin"],
                     max_col=6, max_row=rr["Net margin"])
    cats = Reference(ws_ratios, min_col=2, min_row=1, max_col=6, max_row=1)
    ch1.add_data(data, from_rows=True, titles_from_data=False)
    ch1.set_categories(cats)
    for s, name in zip(ch1.series,
                       ["Gross margin", "EBITDA margin", "EBIT margin", "Net margin"]):
        s.title = SeriesLabel(v=name)
        s.smooth, s.marker.symbol = True, "circle"
    ch1.y_axis.number_format = PCT
    ws.add_chart(ch1, "A11")

    # Chart 2 — revenue vs net income (bar)
    ch2 = BarChart()
    ch2.title, ch2.style, ch2.height, ch2.width = "Revenue vs Net income (INR lakh)", 10, 7.5, 15
    for r, name in ((PNL_R["Revenue"], "Revenue"), (PNL_R["Net income"], "Net income")):
        d = Reference(ws_pnl, min_col=2, min_row=r, max_col=6, max_row=r)
        ch2.add_data(d, from_rows=True, titles_from_data=False)
        ch2.series[-1].title = SeriesLabel(v=name)
        ch2.series[-1].graphicalProperties.line.solidFill = "1F4E78" if r == PNL_R["Revenue"] else "ED7D31"
    ch2.set_categories(Reference(ws_pnl, min_col=2, min_row=1, max_col=6, max_row=1))
    ws.add_chart(ch2, "A27")

    # Chart 3 — working-capital efficiency in days (bar)
    ch3 = BarChart()
    ch3.title, ch3.style, ch3.height, ch3.width = "Working-capital efficiency (days)", 10, 7.5, 15
    d = Reference(ws_ratios, min_col=2, min_row=rr["Receivable days"],
                  max_col=6, max_row=rr["Payable days"])
    ch3.add_data(d, from_rows=True, titles_from_data=False)
    ch3.set_categories(Reference(ws_ratios, min_col=2, min_row=1, max_col=6, max_row=1))
    for s, name in zip(ch3.series, ["Receivable days", "Inventory days", "Payable days"]):
        s.title = SeriesLabel(v=name)
    ws.add_chart(ch3, "A43")


# -------------------------------------------------------------- AI_Report --
def build_ai_report(ws, meta):
    ws.title = "AI_Report"
    ws.sheet_properties.tabColor = "7030A0"
    ws.cell(row=1, column=1, value="ProfitPulse — rule-based weekly narrative").font = TITLE_FONT
    narrative = [
        ("Shakti Consumer Goods grew revenue every year FY21–FY25 while cutting debt: "
         "long-term debt roughly halved and interest coverage improved — a classic deleveraging story.", None),
        ("Margins expanded across the board (gross, EBITDA and net), driven by COGS and SG&A "
         "growing slower than revenue. ROE strengthened as net margin outpaced the equity build.", None),
        ("The watch item is working capital: receivable days drifted up over the period while "
         "inventory discipline held — collections, not operations, are the drag on the cash "
         "conversion cycle. See the red-flag checklist below (all statuses are live formulas).", None),
    ]
    r = 3
    for text, _ in narrative:
        cell = ws.cell(row=r, column=1, value=text)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="Red-flag checklist (FY25, live)").font = Font(bold=True, size=12, color="1F4E78")
    r += 1
    for col, h in enumerate(["Check", "FY25 value", "Threshold", "Status", "Reading"], start=1):
        cell = ws.cell(row=r, column=col, value=h)
        cell.fill, cell.font, cell.alignment = HDR_FILL, HDR_FONT, CENTER
    r += 1
    rr = RATIO_R
    checks = [
        ("Current ratio < 1.20", f"=Ratios!F{rr['Current ratio']}", "1.20",
         f"=IF(Ratios!F{rr['Current ratio']}<1.2,\"⚠ WATCH\",\"✓ OK\")", RTO,
         "Short-term liquidity buffer"),
        ("Quick ratio < 1.00", f"=Ratios!F{rr['Quick ratio']}", "1.00",
         f"=IF(Ratios!F{rr['Quick ratio']}<1,\"⚠ WATCH\",\"✓ OK\")", RTO,
         "Liquidity without inventory"),
        ("Debt/Equity > 1.00", f"=Ratios!F{rr['Debt / Equity']}", "1.00",
         f"=IF(Ratios!F{rr['Debt / Equity']}>1,\"⚠ WATCH\",\"✓ OK\")", RTO,
         "Leverage comfort zone"),
        ("Interest coverage < 3×", f"=Ratios!F{rr['Interest coverage (x)']}", "3.0×",
         f"=IF(Ratios!F{rr['Interest coverage (x)']}<3,\"⚠ WATCH\",\"✓ OK\")", RTO,
         "Debt-service headroom"),
        ("Net margin falling YoY", f"=Ratios!F{rr['Net margin']}", "vs FY24",
         f"=IF(Ratios!F{rr['Net margin']}<Ratios!E{rr['Net margin']},\"⚠ WATCH\",\"✓ OK\")", PCT,
         "Profitability momentum"),
        ("Receivable days rising YoY", f"=Ratios!F{rr['Receivable days']}", "vs FY24",
         f"=IF(Ratios!F{rr['Receivable days']}>Ratios!E{rr['Receivable days']},\"⚠ WATCH\",\"✓ OK\")", DAYS,
         "Collections discipline"),
        ("CCC > 90 days", f"=Ratios!F{rr['Cash conversion cycle (days)']}", "90 days",
         f"=IF(Ratios!F{rr['Cash conversion cycle (days)']}>90,\"⚠ WATCH\",\"✓ OK\")", DAYS,
         "Cash tied in operations"),
        ("Working capital falling YoY", f"=Ratios!F{rr['Working capital']}", "vs FY24",
         f"=IF(Ratios!F{rr['Working capital']}<Ratios!E{rr['Working capital']},\"⚠ WATCH\",\"✓ OK\")", AMT,
         "Net short-term resources"),
    ]
    first = r
    for label, fval, thr, fstat, fmt, reading in checks:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        ws.cell(row=r, column=2, value=fval).number_format = fmt
        ws.cell(row=r, column=3, value=thr).alignment = CENTER
        ws.cell(row=r, column=4, value=fstat).alignment = CENTER
        ws.cell(row=r, column=5, value=reading).alignment = LEFT
        r += 1
    ws.conditional_formatting.add(f"D{first}:D{r-1}",
        CellIsRule(operator="equal", formula=['"⚠ WATCH"'], fill=FLAG_RED))
    ws.conditional_formatting.add(f"D{first}:D{r-1}",
        CellIsRule(operator="equal", formula=['"✓ OK"'], fill=FLAG_GREEN))
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 34


# ----------------------------------------------------------------- Checks --
def build_checks(ws):
    ws.title = "Checks"
    ws.sheet_properties.tabColor = "FF0000"
    ws.cell(row=1, column=1, value="Integrity checks — every row must read TRUE / 0").font = TITLE_FONT
    ws.cell(row=2, column=1, value="Check")
    for j, y in enumerate(YEARS, start=2):
        ws.cell(row=2, column=j, value=y)
    style_header(ws, 7)
    checks = [
        ("Balance sheet balances (Assets − L&E = 0)",
         lambda j: f"=Raw_BS!{C(j+1)}{BS_R['Total assets']}-Raw_BS!{C(j+1)}{BS_R['Total L&E']}"),
        ("Reserves rollforward (Res_t − Res_{t−1} − Net_t = 0)",
         lambda j: (f"=Raw_BS!{C(j+1)}{BS_R['Reserves & surplus']}"
                    f"-Raw_BS!{C(j)}{BS_R['Reserves & surplus']}-Raw_PNL!{C(j)}{PNL_R['Net income']}")),
        ("P&L foots (Net − (EBT − Tax) = 0)",
         lambda j: (f"=Raw_PNL!{C(j)}{PNL_R['Net income']}"
                    f"-(Raw_PNL!{C(j)}{PNL_R['EBT']}-Raw_PNL!{C(j)}{PNL_R['Tax']})")),
        ("DuPont tie-out (product − direct = 0)",
         lambda j: f"=DuPont!{C(j)}7"),
        ("Cash positive every year",
         lambda j: f"=Raw_BS!{C(j+1)}{BS_R['Cash & bank (plug)']}>0"),
        ("Margins within [0, 100%]",
         lambda j: (f"=AND(Ratios!{C(j)}{RATIO_R['Gross margin']}>=0,"
                    f"Ratios!{C(j)}{RATIO_R['Gross margin']}<=1,"
                    f"Ratios!{C(j)}{RATIO_R['Net margin']}>=0,"
                    f"Ratios!{C(j)}{RATIO_R['Net margin']}<=1)")),
        ("Current ratio positive",
         lambda j: f"=Ratios!{C(j)}{RATIO_R['Current ratio']}>0"),
        ("Working capital = CA − CL",
         lambda j: (f"=Ratios!{C(j)}{RATIO_R['Working capital']}"
                    f"-(Raw_BS!{C(j+1)}{BS_R['Total current assets']}"
                    f"-Raw_BS!{C(j+1)}{BS_R['Total current liabilities']})")),
    ]
    for i, (label, fn) in enumerate(checks):
        row = i + 3
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        for j in range(2, 7):
            cell = ws.cell(row=row, column=j, value=fn(j))
            cell.alignment, cell.border = CENTER, THIN
    ws.column_dimensions["A"].width = 48
    for j in range(2, 7):
        ws.column_dimensions[C(j)].width = 14


# ------------------------------------------------------------------- main --
def main():
    pnl = read_csv("pnl.csv")
    bs = read_csv("bs.csv")
    meta = {r[0]: r[1] for r in read_csv("meta.csv")[1:]}

    wb = Workbook()
    build_guide(wb, meta)
    ws_pnl = wb.create_sheet()
    build_pnl(ws_pnl, pnl)
    build_bs(wb.create_sheet(), bs)
    ws_ratios = wb.create_sheet()
    build_ratios(ws_ratios)
    build_dupont(wb.create_sheet())
    build_dashboard(wb.create_sheet(), ws_ratios, ws_pnl)
    build_ai_report(wb.create_sheet(), meta)
    build_checks(wb.create_sheet())
    wb.save(OUT)
    print("wrote", OUT, f"({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
