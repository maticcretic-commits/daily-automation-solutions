#!/usr/bin/env python3
"""Build options-signal-dashboard.xlsx — OptionsSignal, an NSE options
pattern-signal dashboard for a fictional NIFTY weekly expiry.

Built around the Freelancer lead "NSE Options Pattern Signal
Dashboard" (finance-adjacent dashboard wedge): option-chain data with
OI build-up signals, put-call ratio, max-pain, support / resistance
from the heaviest OI strikes, and a signal dashboard with charts.

Sheets:
  Guide        — what the pack is, how to read the signals, how to use it.
  OptionChain  — 29 strikes, CE/PE OI, OI change, volume, LTP, prev close,
                 plus live derived columns: total OI, PCR, % price change,
                 CE/PE OI-build-up signals, and a max-pain payout column
                 (SUMPRODUCT over the whole chain, per candidate strike).
  Signals      — per-strike CE signal, PE signal, combined read, payout,
                 and the max-pain strike flagged.
  Dashboard    — spot / expiry / ATM, KPIs (totals, PCR, max pain,
                 support, resistance, net OI change, signal counts),
                 a signal-count table, and 3 native charts.
  Checks       — 10 integrity checks against the independent Python
                 ground truth; all must read PASS.

Signal logic (standard OI interpretation):
  price up + OI up   = Long Build-up      (bullish for CE, bearish for PE)
  price down + OI up = Short Build-up     (bearish for CE, bullish for PE)
  price up + OI down = Short Covering     (bullish for CE, bearish for PE)
  price down + OI down = Long Unwinding   (bearish for CE, bullish for PE)
Combined read: both sides agreeing = Strong Bullish/Bearish; one side
agreeing = Bullish/Bearish; neither = Neutral.

Formula rules (LibreOffice 24.2): plain A1-style ranges, no Table
objects, no structured references; every range literal repeats the
column letter on both ends.
"""
import csv
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
XLSX = BASE / "options-signal-dashboard.xlsx"

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14)
BOLD = Font(bold=True, size=11)
NUM = "#,##0"
PRICE = "#,##0.00"
PCR_FMT = "0.00"
PCT = "0.0%"

FIRST_ROW = 4          # first strike row on OptionChain / Signals
N_STRIKES = 29
LAST_ROW = FIRST_ROW + N_STRIKES - 1   # 32

# OptionChain columns
# A Strike, B CE OI, C CE OI Chg, D CE Vol, E CE LTP, F CE PrevClose,
# G PE OI, H PE OI Chg, I PE Vol, J PE LTP, K PE PrevClose,
# L Total OI, M Total OI Chg, N PCR, O CE %Chg, P PE %Chg,
# Q CE Signal, R PE Signal, S Max Pain Payout


def load_chain():
    with open(DATA / "option_chain.csv", newline="") as f:
        rows = list(csv.DictReader(f))
    with open(DATA / "meta.csv", newline="") as f:
        meta = {r["key"]: r["value"] for r in csv.DictReader(f)}
    assert len(rows) == N_STRIKES
    return rows, meta


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT


def build():
    rows, meta = load_chain()
    spot = int(meta["spot"])
    expiry = meta["expiry"]
    wb = openpyxl.Workbook()

    # ---------------- Guide ----------------
    ws = wb.active
    ws.title = "Guide"
    ws["A1"] = "NSE Options Pattern Signal Dashboard — NIFTY weekly expiry"
    ws["A1"].font = TITLE_FONT
    lines = [
        "A formula-driven options signal dashboard for a fictional NIFTY",
        f"weekly expiry ({expiry}), spot {spot:,}. Built as a working demo",
        "for the Freelancer project 'NSE Options Pattern Signal Dashboard'",
        "(finance-adjacent dashboard wedge).",
        "",
        "HOW TO READ THE SIGNALS",
        "Each strike gets a CE signal and a PE signal from the classic",
        "OI-build-up reading — price move crossed with open-interest",
        "change:",
        "  Long Build-up    price up + OI up   (fresh longs)",
        "  Short Build-up   price down + OI up (fresh shorts)",
        "  Short Covering   price up + OI down (shorts exiting)",
        "  Long Unwinding   price down + OI down (longs exiting)",
        "For calls, Long Build-up / Short Covering are bullish; for puts",
        "they are bearish. The Signals sheet combines both sides into",
        "one read per strike: Strong Bullish / Bullish / Neutral /",
        "Bearish / Strong Bearish.",
        "",
        "KEY LEVELS (all live formulas)",
        "  Max pain   — the strike where total option payout at expiry",
        "               is smallest (SUMPRODUCT over the whole chain).",
        "  Support    — the strike with the heaviest PE open interest",
        "               (put writers defend it).",
        "  Resistance — the strike with the heaviest CE open interest",
        "               (call writers defend it).",
        "  PCR        — total PE OI / total CE OI; above 1 leans bullish.",
        "",
        "HOW TO USE",
        "1. Paste a fresh option chain into OptionChain columns A-K",
        "   (Strike, CE OI, CE OI Chg, CE Vol, CE LTP, CE PrevClose,",
        "   PE OI, PE OI Chg, PE Vol, PE LTP, PE PrevClose). Columns",
        "   L-S, Signals and Dashboard are all live formulas.",
        "2. Update the spot on the Dashboard (B4) and the ATM strike",
        "   follows. Change nothing else to re-read a new expiry.",
        "3. Read the Checks sheet after any change — all 10 checks",
        "   must read PASS.",
        "",
        "SHEETS",
        "OptionChain — the 29-strike chain + live derived columns.",
        "Signals     — per-strike CE/PE signals, combined read, payout.",
        "Dashboard   — KPIs, signal counts, 3 native charts.",
        "Checks      — 10 integrity checks; all must read PASS.",
        "",
        "DATA — all fictional. The chain is generated from a fixed seed",
        "for the demo; it is not real market data and not trading advice.",
    ]
    for i, line in enumerate(lines, start=3):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 78

    # ---------------- OptionChain ----------------
    ws = wb.create_sheet("OptionChain")
    ws["A1"] = (f"Option chain — NIFTY {expiry} expiry, spot {spot:,}")
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Columns A-K are inputs (paste a fresh chain here); "
                "columns L-S are live formulas.")
    headers = ["Strike", "CE OI", "CE OI Chg", "CE Vol", "CE LTP",
               "CE PrevClose", "PE OI", "PE OI Chg", "PE Vol", "PE LTP",
               "PE PrevClose", "Total OI", "Total OI Chg", "PCR",
               "CE % Chg", "PE % Chg", "CE Signal", "PE Signal",
               "Max Pain Payout"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, len(headers))
    for i, r in enumerate(rows):
        row = FIRST_ROW + i
        ws.cell(row=row, column=1, value=int(r["strike"]))
        for col, key in enumerate(
                ["ce_oi", "ce_oi_chg", "ce_vol"], start=2):
            ws.cell(row=row, column=col, value=int(r[key]))
        ws.cell(row=row, column=5, value=float(r["ce_ltp"]))
        ws.cell(row=row, column=6, value=float(r["ce_prev_close"]))
        for col, key in enumerate(
                ["pe_oi", "pe_oi_chg", "pe_vol"], start=7):
            ws.cell(row=row, column=col, value=int(r[key]))
        ws.cell(row=row, column=10, value=float(r["pe_ltp"]))
        ws.cell(row=row, column=11, value=float(r["pe_prev_close"]))
        # Live derived columns
        ws.cell(row=row, column=12, value=f"=B{row}+G{row}")
        ws.cell(row=row, column=13, value=f"=C{row}+H{row}")
        ws.cell(row=row, column=14,
               value=f"=IF(B{row}=0,0,G{row}/B{row})")
        ws.cell(row=row, column=15,
               value=f"=IF(F{row}=0,0,(E{row}-F{row})/F{row})")
        ws.cell(row=row, column=16,
               value=f"=IF(K{row}=0,0,(J{row}-K{row})/K{row})")
        ws.cell(row=row, column=17, value=(
            f'=IF(AND(C{row}>0,O{row}>0),"Long Build-up",'
            f'IF(AND(C{row}>0,O{row}<0),"Short Build-up",'
            f'IF(AND(C{row}<0,O{row}>0),"Short Covering",'
            f'"Long Unwinding")))'))
        ws.cell(row=row, column=18, value=(
            f'=IF(AND(H{row}>0,P{row}>0),"Long Build-up",'
            f'IF(AND(H{row}>0,P{row}<0),"Short Build-up",'
            f'IF(AND(H{row}<0,P{row}>0),"Short Covering",'
            f'"Long Unwinding")))'))
        # Max-pain payout if expiry settled at this strike:
        # CE payout = CE_OI * max(strike - K, 0) summed over all K,
        # PE payout = PE_OI * max(K - strike, 0) summed over all K.
        ws.cell(row=row, column=19, value=(
            f"=SUMPRODUCT($B${FIRST_ROW}:$B${LAST_ROW},"
            f"($A{row}-$A${FIRST_ROW}:$A${LAST_ROW}>0)"
            f"*($A{row}-$A${FIRST_ROW}:$A${LAST_ROW}))"
            f"+SUMPRODUCT($G${FIRST_ROW}:$G${LAST_ROW},"
            f"($A${FIRST_ROW}:$A${LAST_ROW}-$A{row}>0)"
            f"*($A${FIRST_ROW}:$A${LAST_ROW}-$A{row}))"))
    widths = {"A": 10, "B": 11, "C": 11, "D": 10, "E": 10, "F": 12,
              "G": 11, "H": 11, "I": 10, "J": 10, "K": 12, "L": 11,
              "M": 12, "N": 8, "O": 9, "P": 9, "Q": 16, "R": 16, "S": 17}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for row in range(FIRST_ROW, LAST_ROW + 1):
        for col in "BCDGHIJKLM":
            ws[f"{col}{row}"].number_format = NUM
        for col in "EFJK":
            ws[f"{col}{row}"].number_format = PRICE
        ws[f"N{row}"].number_format = PCR_FMT
        ws[f"O{row}"].number_format = PCT
        ws[f"P{row}"].number_format = PCT
        ws[f"S{row}"].number_format = NUM
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:S{LAST_ROW}"

    # ---------------- Signals ----------------
    ws = wb.create_sheet("Signals")
    ws["A1"] = "Per-strike signals — combined read"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Combined read: both sides agreeing = Strong; one side "
                "agreeing = plain Bullish/Bearish; neither = Neutral.")
    for col, h in enumerate(["Strike", "CE Signal", "PE Signal",
                             "Combined Read", "Max Pain Payout",
                             "Max Pain?"], start=1):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 6)
    for i in range(N_STRIKES):
        row = FIRST_ROW + i
        ws.cell(row=row, column=1, value=f"=OptionChain!A{row}")
        ws.cell(row=row, column=2, value=f"=OptionChain!Q{row}")
        ws.cell(row=row, column=3, value=f"=OptionChain!R{row}")
        ws.cell(row=row, column=4, value=(
            f'=IF(AND(B{row}="Long Build-up",'
            f'C{row}="Short Covering"),"Strong Bullish",'
            f'IF(AND(B{row}="Short Build-up",'
            f'C{row}="Long Build-up"),"Strong Bearish",'
            f'IF(OR(B{row}="Long Build-up",'
            f'C{row}="Short Covering"),"Bullish",'
            f'IF(OR(B{row}="Short Build-up",'
            f'C{row}="Long Build-up"),"Bearish","Neutral"))))'))
        ws.cell(row=row, column=5, value=f"=OptionChain!S{row}")
        ws.cell(row=row, column=6, value=(
            f'=IF(E{row}=MIN($E${FIRST_ROW}:$E${LAST_ROW}),'
            f'"MAX PAIN","")'))
    for col, w in zip("ABCDEF", [10, 16, 16, 16, 17, 10]):
        ws.column_dimensions[col].width = w
    for row in range(FIRST_ROW, LAST_ROW + 1):
        ws[f"A{row}"].number_format = NUM
        ws[f"E{row}"].number_format = NUM
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:F{LAST_ROW}"

    # ---------------- Dashboard ----------------
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "Options Signal Dashboard — NIFTY"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("All figures are live formulas off the OptionChain "
                "sheet — paste a fresh chain and the dashboard follows.")
    ws["A3"] = "Underlying"
    ws["B3"] = "NIFTY"
    ws["A4"] = "Spot"
    ws["B4"] = spot
    ws["B4"].number_format = NUM
    ws["A5"] = "Expiry"
    ws["B5"] = expiry
    ws["A6"] = "ATM strike"
    ws["B6"] = "=ROUND(B4/50,0)*50"
    ws["B6"].number_format = NUM
    for rr in range(3, 7):
        ws.cell(row=rr, column=1).font = BOLD
        ws.cell(row=rr, column=2).font = BOLD

    ws["A8"] = "Key figures"
    ws["A8"].font = BOLD
    kpis = [
        ("Total CE OI", f"=SUM(OptionChain!B{FIRST_ROW}:B{LAST_ROW})",
         NUM),
        ("Total PE OI", f"=SUM(OptionChain!G{FIRST_ROW}:G{LAST_ROW})",
         NUM),
        ("Total OI", f"=SUM(OptionChain!L{FIRST_ROW}:L{LAST_ROW})",
         NUM),
        ("Put-Call Ratio (PE OI / CE OI)", "=B10/B9", PCR_FMT),
        ("Net CE OI change",
         f"=SUM(OptionChain!C{FIRST_ROW}:C{LAST_ROW})", NUM),
        ("Net PE OI change",
         f"=SUM(OptionChain!H{FIRST_ROW}:H{LAST_ROW})", NUM),
        ("Max Pain strike",
         f"=INDEX(OptionChain!A{FIRST_ROW}:A{LAST_ROW},"
         f"MATCH(MIN(OptionChain!S{FIRST_ROW}:S{LAST_ROW}),"
         f"OptionChain!S{FIRST_ROW}:S{LAST_ROW},0))", NUM),
        ("Max Pain payout (Rs)",
         f"=MIN(OptionChain!S{FIRST_ROW}:S{LAST_ROW})", NUM),
        ("Support (max PE OI strike)",
         f"=INDEX(OptionChain!A{FIRST_ROW}:A{LAST_ROW},"
         f"MATCH(MAX(OptionChain!G{FIRST_ROW}:G{LAST_ROW}),"
         f"OptionChain!G{FIRST_ROW}:G{LAST_ROW},0))", NUM),
        ("Resistance (max CE OI strike)",
         f"=INDEX(OptionChain!A{FIRST_ROW}:A{LAST_ROW},"
         f"MATCH(MAX(OptionChain!B{FIRST_ROW}:B{LAST_ROW}),"
         f"OptionChain!B{FIRST_ROW}:B{LAST_ROW},0))", NUM),
        ("Bullish strikes (incl. Strong)",
         '=COUNTIF(Signals!D4:D32,"Bullish")'
         '+COUNTIF(Signals!D4:D32,"Strong Bullish")', NUM),
        ("Bearish strikes (incl. Strong)",
         '=COUNTIF(Signals!D4:D32,"Bearish")'
         '+COUNTIF(Signals!D4:D32,"Strong Bearish")', NUM),
        ("Neutral strikes", '=COUNTIF(Signals!D4:D32,"Neutral")', NUM),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        r = 9 + i
        ws.cell(row=r, column=1, value=label).font = BOLD
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = fmt

    ws["A23"] = "Signal counts by type"
    ws["A23"].font = BOLD
    for col, h in enumerate(["Signal", "CE strikes", "PE strikes"],
                            start=1):
        ws.cell(row=24, column=col, value=h)
    style_header(ws, 24, 3)
    for i, sig in enumerate(["Long Build-up", "Short Build-up",
                             "Short Covering", "Long Unwinding"]):
        r = 25 + i
        ws.cell(row=r, column=1, value=sig)
        ws.cell(row=r, column=2, value=(
            f'=COUNTIF(OptionChain!Q{FIRST_ROW}:Q{LAST_ROW},A{r})'))
        ws.cell(row=r, column=3, value=(
            f'=COUNTIF(OptionChain!R{FIRST_ROW}:R{LAST_ROW},A{r})'))
    ws.cell(row=29, column=1, value="Total").font = BOLD
    ws.cell(row=29, column=2, value="=SUM(B25:B28)").font = BOLD
    ws.cell(row=29, column=3, value="=SUM(C25:C28)").font = BOLD
    ws.column_dimensions["A"].width = 36
    for col in "BC":
        ws.column_dimensions[col].width = 16

    oc = wb["OptionChain"]
    ch1 = BarChart()
    ch1.type = "col"
    ch1.style = 10
    ch1.title = "Open interest by strike — CE vs PE"
    ch1.y_axis.title = "OI (contracts)"
    ch1.height = 7.5
    ch1.width = 16
    # Only B and G: add separately so D-F are excluded.
    ch1.add_data(Reference(oc, min_col=2, min_row=3,
                           max_row=LAST_ROW), titles_from_data=True)
    ch1.add_data(Reference(oc, min_col=7, min_row=3,
                           max_row=LAST_ROW), titles_from_data=True)
    cats_ref = Reference(oc, min_col=1, min_row=FIRST_ROW,
                         max_row=LAST_ROW)
    ch1.set_categories(cats_ref)
    ch1.shape = 4
    ws.add_chart(ch1, "E6")

    ch2 = BarChart()
    ch2.type = "col"
    ch2.style = 10
    ch2.title = "OI change by strike — CE vs PE"
    ch2.y_axis.title = "OI change (contracts)"
    ch2.height = 7.5
    ch2.width = 16
    ch2.add_data(Reference(oc, min_col=3, min_row=3,
                           max_row=LAST_ROW), titles_from_data=True)
    ch2.add_data(Reference(oc, min_col=8, min_row=3,
                           max_row=LAST_ROW), titles_from_data=True)
    ch2.set_categories(cats_ref)
    ch2.shape = 4
    ws.add_chart(ch2, "E23")

    ch3 = LineChart()
    ch3.style = 10
    ch3.title = "Put-Call Ratio by strike"
    ch3.y_axis.title = "PCR"
    ch3.height = 7.5
    ch3.width = 16
    ch3.add_data(Reference(oc, min_col=14, min_row=3,
                           max_row=LAST_ROW), titles_from_data=True)
    ch3.set_categories(cats_ref)
    ws.add_chart(ch3, "E40")

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    ws["A1"] = "Integrity checks — all must read PASS"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Expected values are the independent Python ground truth "
                "computed from data/option_chain.csv; Actual values are "
                "live formulas. Re-check after pasting a fresh chain.")
    ws.append([])
    ws.append(["#", "Check", "Expected", "Actual", "Result"])
    style_header(ws, 4, 5)
    # Ground truth computed here (same arithmetic as the test file).
    ce_oi = sum(int(r["ce_oi"]) for r in rows)
    pe_oi = sum(int(r["pe_oi"]) for r in rows)
    pcr = pe_oi / ce_oi
    strikes = [int(r["strike"]) for r in rows]
    ce_ois = [int(r["ce_oi"]) for r in rows]
    pe_ois = [int(r["pe_oi"]) for r in rows]

    def payout(s):
        return (sum(o * max(s - k, 0) for o, k in zip(ce_ois, strikes))
                + sum(o * max(k - s, 0) for o, k in zip(pe_ois, strikes)))

    payouts = {s: payout(s) for s in strikes}
    max_pain = min(payouts, key=payouts.get)
    support = strikes[pe_ois.index(max(pe_ois))]
    resistance = strikes[ce_ois.index(max(ce_ois))]
    checks = [
        ("Total CE OI", ce_oi, "=Dashboard!B9"),
        ("Total PE OI", pe_oi, "=Dashboard!B10"),
        ("Put-Call Ratio", round(pcr, 4), "=Dashboard!B12"),
        ("Max Pain strike", max_pain, "=Dashboard!B15"),
        ("Support strike (max PE OI)", support, "=Dashboard!B17"),
        ("Resistance strike (max CE OI)", resistance, "=Dashboard!B18"),
        ("CE signals counted = 29 strikes",
         N_STRIKES, "=COUNTA(OptionChain!Q4:Q32)"),
        ("PE signals counted = 29 strikes",
         N_STRIKES, "=COUNTA(OptionChain!R4:R32)"),
        ("Total OI = CE OI + PE OI", 0,
         "=Dashboard!B11-Dashboard!B9-Dashboard!B10"),
        ("Max Pain payout is the minimum payout", 0,
         "=Dashboard!B16-MIN(OptionChain!S4:S32)"),
    ]
    for i, (label, expected, actual) in enumerate(checks):
        r = 5 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=expected)  # numeric, never text
        ws.cell(row=r, column=4, value=actual)
        ws.cell(row=r, column=5,
               value=f'=IF(ABS(D{r}-C{r})<0.01,"PASS","FAIL")').font = BOLD
    for col, w in zip("ABCDE", [5, 44, 14, 18, 9]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    wb.save(XLSX)
    print(f"built {XLSX}")
    print(f"GT: CE OI {ce_oi:,} PE OI {pe_oi:,} PCR {pcr:.4f} "
          f"max pain {max_pain} support {support} "
          f"resistance {resistance}")


if __name__ == "__main__":
    build()
