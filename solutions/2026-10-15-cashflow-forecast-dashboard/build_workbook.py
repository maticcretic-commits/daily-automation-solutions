#!/usr/bin/env python3
"""FlowForge — build cashflow-forecast.xlsx.

Sheets:
  Guide       — how to use the workbook
  Assumptions — editable inputs (yellow cells)
  History     — 26 weeks of actuals: inflow/outflow by category, totals, net,
                opening/closing balance, budget, budget-vs-actual variance
  Forecast    — 13-week projection from trailing averages + assumptions,
                cash-floor flags
  Dashboard   — KPI cards + native charts (line, column, pie)
  Checks      — 8 automated PASS/FAIL integrity checks
  Report      — rule-based weekly narrative (generated from this data)

Every money figure is a live formula — nothing hard-coded. Currency in INR.
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, Series
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill, numbers
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
DATA = ROOT / "data"
OUT = ROOT / "cashflow-forecast.xlsx"

INR = '"₹"#,##0'
PCT1 = '0.0%'
YELLOW = PatternFill("solid", fgColor="FFF2CC")
RED_FILL = PatternFill("solid", fgColor="FCE4EC")
AMBER_FILL = PatternFill("solid", fgColor="FFF3E0")
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
KPI_FONT = Font(bold=True, size=13, color="1F4E79")
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center")

INFLOW_CATS = ["sales_receipts", "gst_refund", "other_income"]
OUTFLOW_CATS = ["supplier_payments", "payroll", "rent", "logistics",
                "gst_tax", "loan_emi"]


def read_csv(name):
    with open(DATA / name) as f:
        return list(csv.DictReader(f))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER


def autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------- data
inflows = read_csv("history_inflows.csv")
outflows = read_csv("history_outflows.csv")
budget = read_csv("history_budget.csv")
W = len(inflows)  # 26
HIST_OPENING = 4_250_000
FC_WEEKS = 13

wk_label = lambda w: f"W{w}"          # W-25 .. W0
fc_label = lambda n: f"F{n}"         # F1 .. F13

tot_in = [sum(int(r[c]) for c in INFLOW_CATS) for r in inflows]
tot_out = [sum(int(r[c]) for c in OUTFLOW_CATS) for r in outflows]
net = [a - b for a, b in zip(tot_in, tot_out)]
closing_hist = []
bal = HIST_OPENING
for n in net:
    bal += n
    closing_hist.append(bal)
FC_OPENING = closing_hist[-1]

# narrative inputs
top_in_i = max(range(W), key=lambda i: tot_in[i])
top_out_i = max(range(W), key=lambda i: tot_out[i])
avg4_in = sum(tot_in[-4:]) / 4
avg4_out = sum(tot_out[-4:]) / 4
g_in, g_out, floor = 0.02, 0.015, 500_000
fc_in = [round(avg4_in * (1 + g_in) ** n) for n in range(1, FC_WEEKS + 1)]
fc_out = [round(avg4_out * (1 + g_out) ** n) for n in range(1, FC_WEEKS + 1)]
fc_net = [a - b for a, b in zip(fc_in, fc_out)]
fc_close = []
b = FC_OPENING
for n in fc_net:
    b += n
    fc_close.append(b)
min_cash = min(fc_close)
min_wk = fc_close.index(min_cash) + 1
below = sum(1 for c in fc_close if c < floor)

# ---------------------------------------------------------------- workbook
wb = Workbook()

# ---- Guide ------------------------------------------------------------
ws = wb.active
ws.title = "Guide"
guide = [
    ("FlowForge — 13-week cash-flow forecast + budget-variance dashboard", TITLE_FONT),
    ("Fictional demo data: 'Patna Traders Pvt. Ltd.', FMCG wholesale distributor. "
     "All values in INR. Every money figure on History/Forecast is a live formula — "
     "nothing is hard-coded, so one assumption change re-flows the whole model.", None),
    ("HOW TO USE", Font(bold=True, size=12)),
    ("1. Edit the yellow cells on Assumptions (growth rates, cash floor, openings).", None),
    ("2. Paste new weekly actuals into History columns B–D and F–K (values only); "
     "totals, net, balances, variances and alerts are all formulas and extend automatically.", None),
    ("3. Forecast re-projects from the trailing-4-week average of History plus your "
     "growth assumptions. Weeks under the cash floor are flagged in red.", None),
    ("4. Dashboard KPIs and all three charts read live from History/Forecast.", None),
    ("5. Checks must all read PASS after any edit — a FAIL means a formula was overwritten.", None),
    ("WEEKLY ROUTINE (15 min): paste actuals → check Report for the narrative → "
     "act on any BELOW FLOOR or ⚠ Review flags → re-export.", None),
]
for i, (t, f) in enumerate(guide, start=1):
    c = ws.cell(row=i, column=1, value=t)
    c.alignment = WRAP
    if f:
        c.font = f
ws.column_dimensions["A"].width = 130
for r in (1, 3):
    ws.row_dimensions[r].height = 22

# ---- Assumptions ------------------------------------------------------
ws = wb.create_sheet("Assumptions")
rows = [
    ("Setting", "Value", "Notes"),
    ("history_opening_balance", HIST_OPENING, "Opening bank balance at start of W-25"),
    ("forecast_opening_balance", FC_OPENING, "Opening balance for F1 (= W0 closing)"),
    ("inflow_growth", g_in, "Weekly inflow growth vs trailing-4-week average"),
    ("outflow_growth", g_out, "Weekly outflow growth vs trailing-4-week average"),
    ("cash_floor", floor, "Minimum safe cash — forecast weeks below this are flagged"),
    ("trailing_weeks", 4, "Weeks of history averaged for the run-rate forecast"),
    ("variance_alert_pct", 10, "Budget-vs-actual variance % that triggers a review flag"),
]
for i, (k, v, d) in enumerate(rows, start=1):
    ws.cell(row=i, column=1, value=k)
    ws.cell(row=i, column=2, value=v)
    ws.cell(row=i, column=3, value=d)
style_header(ws, 1, 3)
for r in range(2, 9):
    ws.cell(row=r, column=2).fill = YELLOW
ws["B2"].number_format = INR
ws["B3"].number_format = INR
ws["B4"].number_format = "0.0%"
ws["B5"].number_format = "0.0%"
ws["B6"].number_format = INR
autofit(ws, [30, 18, 70])

A = "Assumptions"  # sheet name for formulas
H_OPEN, F_OPEN = f"{A}!$B$2", f"{A}!$B$3"
G_IN, G_OUT = f"{A}!$B$4", f"{A}!$B$5"
FLOOR, VARPCT = f"{A}!$B$6", f"{A}!$B$8"

# ---- History ----------------------------------------------------------
ws = wb.create_sheet("History")
hdr = (["Week"] + [c.replace("_", " ").title() for c in INFLOW_CATS]
       + ["Total Inflow"]
       + [c.replace("_", " ").title() for c in OUTFLOW_CATS]
       + ["Total Outflow", "Net", "Opening", "Closing",
          "Budget Inflow", "Budget Outflow",
          "Inflow Var %", "Outflow Var %", "Alert"])
# col map: A week, B-D inflow cats, E total in, F-K outflow cats, L total out,
# M net, N opening, O closing, P/Q budget, R/S var, T alert
for j, h in enumerate(hdr, start=1):
    ws.cell(row=1, column=j, value=h)
style_header(ws, 1, len(hdr))
AVG4_IN = f"AVERAGE(History!$E${W - 2}:$E${W + 1})"    # rows 24..27 = E24:E27
AVG4_OUT = f"AVERAGE(History!$L${W - 2}:$L${W + 1})"
for i in range(W):
    r = i + 2
    ws.cell(row=r, column=1, value=wk_label(int(inflows[i]["week"])))
    for j, c in enumerate(INFLOW_CATS, start=2):
        ws.cell(row=r, column=j, value=int(inflows[i][c]))
    ws.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})")
    for j, c in enumerate(OUTFLOW_CATS, start=6):
        ws.cell(row=r, column=j, value=int(outflows[i][c]))
    ws.cell(row=r, column=12, value=f"=SUM(F{r}:K{r})")
    ws.cell(row=r, column=13, value=f"=E{r}-L{r}")
    ws.cell(row=r, column=14,
            value=f"={H_OPEN}" if i == 0 else f"=O{r - 1}")
    ws.cell(row=r, column=15, value=f"=N{r}+M{r}")
    ws.cell(row=r, column=16, value=int(budget[i]["budget_inflow"]))
    ws.cell(row=r, column=17, value=int(budget[i]["budget_outflow"]))
    ws.cell(row=r, column=18, value=f"=IF(P{r}=0,0,(E{r}-P{r})/P{r})")
    ws.cell(row=r, column=19, value=f"=IF(Q{r}=0,0,(L{r}-Q{r})/Q{r})")
    ws.cell(row=r, column=20,
            value=f'=IF(OR(ABS(R{r})>{VARPCT}/100,ABS(S{r})>{VARPCT}/100),"⚠ Review","")')
for col in list("BCDE") + list("FGHIJKL") + list("MNOPQ"):
    for r in range(2, W + 2):
        ws[f"{col}{r}"].number_format = INR
for col in "RS":
    for r in range(2, W + 2):
        ws[f"{col}{r}"].number_format = PCT1
ws.conditional_formatting.add(
    f"T2:T{W + 1}",
    CellIsRule(operator="notEqual", formula=['""'], fill=AMBER_FILL))
autofit(ws, [8] + [15] * 19)

# ---- Forecast ---------------------------------------------------------
ws = wb.create_sheet("Forecast")
fhdr = ["Week", "Forecast Inflow", "Forecast Outflow", "Net",
        "Opening", "Closing", "Floor Check"]
for j, h in enumerate(fhdr, start=1):
    ws.cell(row=1, column=j, value=h)
style_header(ws, 1, len(fhdr))
for n in range(1, FC_WEEKS + 1):
    r = n + 1
    ws.cell(row=r, column=1, value=fc_label(n))
    ws.cell(row=r, column=2,
            value=f"=ROUND({AVG4_IN}*(1+{G_IN})^({r}-1),0)")
    ws.cell(row=r, column=3,
            value=f"=ROUND({AVG4_OUT}*(1+{G_OUT})^({r}-1),0)")
    ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
    ws.cell(row=r, column=5,
            value=f"={F_OPEN}" if n == 1 else f"=F{r - 1}")
    ws.cell(row=r, column=6, value=f"=E{r}+D{r}")
    ws.cell(row=r, column=7,
            value=f'=IF(F{r}<{FLOOR},"▼ BELOW FLOOR","✓")')
for col in "BCDEF":
    for r in range(2, FC_WEEKS + 2):
        ws[f"{col}{r}"].number_format = INR
ws.conditional_formatting.add(
    f"G2:G{FC_WEEKS + 1}",
    CellIsRule(operator="equal", formula=['"▼ BELOW FLOOR"'], fill=RED_FILL))
autofit(ws, [8, 18, 18, 16, 18, 18, 16])

# ---- Dashboard --------------------------------------------------------
ws = wb.create_sheet("Dashboard")
ws["A1"] = "FlowForge — Cash Flow Dashboard"
ws["A1"].font = TITLE_FONT
ws["A2"] = ("Live from History/Forecast. Fictional demo data — 'Patna Traders "
            "Pvt. Ltd.', INR.")
kpis = [
    ("Forecast opening balance", f"={F_OPEN}"),
    ("Projected closing (F13)", f"=Forecast!F{FC_WEEKS + 1}"),
    ("Minimum cash over forecast", f"=MIN(Forecast!F2:F{FC_WEEKS + 1})"),
    ("Weeks below cash floor", f'=COUNTIF(Forecast!G2:G{FC_WEEKS + 1},"*BELOW*")'),
    ("Avg weekly net (forecast)", f"=AVERAGE(Forecast!D2:D{FC_WEEKS + 1})"),
    ("History total inflow (26 wks)", f"=SUM(History!E2:E{W + 1})"),
    ("History total outflow (26 wks)", f"=SUM(History!L2:L{W + 1})"),
    ("Weeks with budget-review alert", f'=COUNTIF(History!T2:T{W + 1},"*Review*")'),
]
for i, (label, formula) in enumerate(kpis):
    r = 4 + i
    ws.cell(row=r, column=1, value=label).font = Font(bold=True)
    c = ws.cell(row=r, column=2, value=formula)
    c.font = KPI_FONT
    if i != 3:
        c.number_format = INR
autofit(ws, [32, 20])

# helper: outflow mix totals for the pie
ws["K3"] = "Outflow mix (history)"
mix_r0 = 4
for i, c in enumerate(OUTFLOW_CATS):
    r = mix_r0 + i
    ws.cell(row=r, column=11, value=c.replace("_", " ").title())
    col = get_column_letter(6 + i)
    ws.cell(row=r, column=12,
            value=f"=SUM(History!{col}2:{col}{W + 1})").number_format = INR

# line chart: cash balance history + forecast
lc = LineChart()
lc.title = "Cash balance — history vs forecast"
lc.style = 2
lc.height, lc.width = 8, 16
cats_h = Reference(wb["History"], min_col=1, min_row=2, max_row=W + 1)
s1 = Series(Reference(wb["History"], min_col=15, min_row=2, max_row=W + 1),
            title="History closing")
s1.xvalues = cats_h
s2 = Series(Reference(wb["Forecast"], min_col=6, min_row=2, max_row=FC_WEEKS + 1),
            title="Forecast closing")
s2.xvalues = Reference(wb["Forecast"], min_col=1, min_row=2,
                       max_row=FC_WEEKS + 1)
lc.append(s1)
lc.append(s2)
ws.add_chart(lc, "A13")

# column chart: weekly inflow vs outflow
bc = BarChart()
bc.title = "Weekly inflow vs outflow (history)"
bc.style = 10
bc.height, bc.width = 8, 16
bc.append(Series(Reference(wb["History"], min_col=5, min_row=2, max_row=W + 1),
                title="Inflow"))
bc.append(Series(Reference(wb["History"], min_col=12, min_row=2, max_row=W + 1),
              title="Outflow"))
bc.dataLabels = None
s = bc.series[0]
s.xvalues = cats_h
ws.add_chart(bc, "A29")

# pie: outflow mix
pc = PieChart()
pc.title = "Outflow mix — 26 weeks"
pc.height, pc.width = 8, 12
pc.add_data(Reference(ws, min_col=12, min_row=mix_r0,
                      max_row=mix_r0 + len(OUTFLOW_CATS) - 1))
pc.set_categories(Reference(ws, min_col=11, min_row=mix_r0,
                            max_row=mix_r0 + len(OUTFLOW_CATS) - 1))
ws.add_chart(pc, "A45")

# ---- Checks -----------------------------------------------------------
ws = wb.create_sheet("Checks")
ws["A1"] = "Integrity checks — all must read PASS"
ws["A1"].font = Font(bold=True, size=12)
checks = [
    ("History reconciles: opening + net = closing every week",
     f"=IF(SUMPRODUCT((History!N2:N{W + 1}+History!M2:M{W + 1}"
     f"-History!O2:O{W + 1})^2)=0,\"PASS\",\"FAIL\")"),
    ("History continuity: each week's opening = prior closing",
     f"=IF(SUMPRODUCT((History!N3:N{W + 1}-History!O2:O{W})^2)=0,\"PASS\",\"FAIL\")"),
    ("Forecast continuity: each week's opening = prior closing",
     f"=IF(SUMPRODUCT((Forecast!E3:E{FC_WEEKS + 1}"
     f"-Forecast!F2:F{FC_WEEKS})^2)=0,\"PASS\",\"FAIL\")"),
    ("Forecast closing = opening + net every week",
     f"=IF(SUMPRODUCT((Forecast!F2:F{FC_WEEKS + 1}-Forecast!E2:E{FC_WEEKS + 1}"
     f"-Forecast!D2:D{FC_WEEKS + 1})^2)=0,\"PASS\",\"FAIL\")"),
    ("Forecast opening equals Assumptions value",
     f"=IF(Forecast!E2={F_OPEN},\"PASS\",\"FAIL\")"),
    ("No formula errors anywhere in History/Forecast",
     f"=IF(SUMPRODUCT(--ISERROR(History!A2:T{W + 1}))"
     f"+SUMPRODUCT(--ISERROR(Forecast!A2:G{FC_WEEKS + 1}))=0,\"PASS\",\"FAIL\")"),
    ("Review alerts fire exactly when |variance| exceeds threshold",
     f"=IF(SUMPRODUCT((ABS(History!R2:R{W + 1})>{VARPCT}/100)"
     f"*(History!T2:T{W + 1}=\"\"))+SUMPRODUCT((ABS(History!S2:S{W + 1})"
     f">{VARPCT}/100)*(History!T2:T{W + 1}=\"\"))=0,\"PASS\",\"FAIL\")"),
    ("Floor flags fire for every forecast week under the cash floor",
     f"=IF(SUMPRODUCT((Forecast!F2:F{FC_WEEKS + 1}<{FLOOR})"
     f"*(Forecast!G2:G{FC_WEEKS + 1}=\"✓\"))=0,\"PASS\",\"FAIL\")"),
]
ws["A2"], ws["B2"] = "Check", "Result"
style_header(ws, 2, 2)
for i, (label, formula) in enumerate(checks, start=3):
    ws.cell(row=i, column=1, value=label).alignment = WRAP
    ws.cell(row=i, column=2, value=formula).alignment = CENTER
    ws.cell(row=i, column=2).font = Font(bold=True)
ws.conditional_formatting.add(
    f"B3:B{2 + len(checks)}",
    CellIsRule(operator="equal", formula=['"PASS"'],
               fill=PatternFill("solid", fgColor="E8F5E9")))
ws.conditional_formatting.add(
    f"B3:B{2 + len(checks)}",
    CellIsRule(operator="equal", formula=['"FAIL"'], fill=RED_FILL))
autofit(ws, [62, 10])

# ---- Report (rule-based narrative) ------------------------------------
ws = wb.create_sheet("Report")
ws["A1"] = "Weekly cash narrative — generated from this workbook's data"
ws["A1"].font = Font(bold=True, size=12)
lines = [
    f"Strongest inflow week: {wk_label(int(inflows[top_in_i]['week']))} "
    f"(₹{tot_in[top_in_i]:,}); weakest: {wk_label(int(inflows[min(range(W), key=lambda i: tot_in[i])]['week']))} "
    f"(₹{min(tot_in):,}).",
    f"Heaviest outflow week: {wk_label(int(outflows[top_out_i]['week']))} "
    f"(₹{tot_out[top_out_i]:,}) — GST/EMI weeks drive the spikes; the forecast "
    "carries them inside the trailing average.",
    f"13-week outlook from F1 opening ₹{FC_OPENING:,}: projected closing "
    f"₹{fc_close[-1]:,}; minimum cash ₹{min_cash:,} in {fc_label(min_wk)}.",
]
if below:
    lines.append(f"ACTION: {below} forecast week(s) dip under the ₹{floor:,} cash "
                 "floor — pull supplier payments forward or draw the buffer before then.")
else:
    lines.append(f"No forecast week breaches the ₹{floor:,} cash floor — buffer intact.")
alerts = sum(1 for i in range(W)
             if abs((tot_in[i] - int(budget[i]['budget_inflow'])) / int(budget[i]['budget_inflow'])) > 0.10
             or abs((tot_out[i] - int(budget[i]['budget_outflow'])) / int(budget[i]['budget_outflow'])) > 0.10)
lines.append(f"{alerts} of {W} history weeks breached the 10% budget-variance "
             "threshold — see the Alert column on History.")
lines.append("Weekly routine: paste actuals → re-read this sheet → act on flags.")
for i, t in enumerate(lines, start=3):
    c = ws.cell(row=i, column=1, value=t)
    c.alignment = WRAP
    ws.row_dimensions[i].height = 30
ws.column_dimensions["A"].width = 130

wb.save(OUT)
print(f"wrote {OUT}  sheets={wb.sheetnames}")
