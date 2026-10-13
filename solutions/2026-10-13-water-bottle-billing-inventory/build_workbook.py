#!/usr/bin/env python3
"""Build the AquaBill workbook: billing + sales + inventory for AquaPure Distributors.

Reads data/*.csv (from generate.py) and writes water-bottle-billing-inventory.xlsx.
Every computed figure is a live Excel formula (absolute A1 refs, full $COL$ROW:$COL$ROW
ranges — no structured refs, no Table objects, so LibreOffice headless recalc works).

Run: python3 build_workbook.py [--datadir data] [--out water-bottle-billing-inventory.xlsx]
"""
import argparse
import csv
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

NL, NI, NP, ND, NS = 4000, 500, 200, 400, 20  # headroom row counts
GST_PCT = 18
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF")
THIN = Border(*[Side(style="thin", color="B0B0B0")] * 4)
MONEY = "#,##0.00"
DTFMT = "YYYY-MM-DD"


def load(datadir):
    def rows(name):
        with open(f"{datadir}/{name}") as f:
            return list(csv.DictReader(f))
    return {
        "skus": rows("skus.csv"),
        "invoices": rows("invoices.csv"),
        "lines": rows("invoice_lines.csv"),
        "purchases": rows("purchases.csv"),
    }


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN


def build(datadir, outpath):
    d = load(datadir)
    skus, invoices, lines, purchases = (d["skus"], d["invoices"], d["lines"],
                                       d["purchases"])
    wb = Workbook()

    # ---------------- InvoiceLines (raw + line totals) ----------------
    ws = wb.active
    ws.title = "InvoiceLines"
    hdr = ["Inv No", "Date", "Line", "SKU", "Qty", "Rate (ex-GST)", "Line Total"]
    ws.append(hdr)
    for ln in lines:
        y, m, dd = map(int, ln["date"].split("-"))
        ws.append([ln["inv_no"], date(y, m, dd), int(ln["line_no"]), ln["sku"],
                   int(ln["qty"]), float(ln["rate"]), None])
    for r in range(2, NL + 2):
        ws.cell(row=r, column=7).value = (
            f'=IF($A{r}="",0,ROUND($E{r}*$F{r},2))')
        ws.cell(row=r, column=2).number_format = DTFMT
        ws.cell(row=r, column=6).number_format = MONEY
        ws.cell(row=r, column=7).number_format = MONEY
    style_header(ws, 7)
    ws.freeze_panes = "A2"
    for w, col in zip([16, 12, 6, 26, 8, 14, 12], range(1, 8)):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ---------------- Invoices (register, all computed) ----------------
    wi = wb.create_sheet("Invoices")
    hdr = ["Inv No", "Date", "Customer", "Subtotal", "GST 18%",
           "Grand Total", "Paid", "Balance", "Status"]
    wi.append(hdr)
    for inv in invoices:
        y, m, dd = map(int, inv["date"].split("-"))
        wi.append([inv["inv_no"], date(y, m, dd), inv["customer"],
                   None, None, None, float(inv["paid"]), None, None])
    for r in range(2, NI + 2):
        wi.cell(row=r, column=4).value = (
            f'=IF($A{r}="",0,ROUND(SUMIFS(InvoiceLines!$G$2:$G${NL+1},'
            f'InvoiceLines!$A$2:$A${NL+1},$A{r}),2))')
        wi.cell(row=r, column=5).value = f'=IF($A{r}="",0,ROUND($D{r}*{GST_PCT}%,2))'
        wi.cell(row=r, column=6).value = f'=IF($A{r}="",0,$D{r}+$E{r})'
        wi.cell(row=r, column=8).value = f'=IF($A{r}="",0,$F{r}-$G{r})'
        wi.cell(row=r, column=9).value = (
            f'=IF($A{r}="","",IF($H{r}=0,"Paid",IF($G{r}=0,"Unpaid","Partial")))')
        for c in (2,):
            wi.cell(row=r, column=c).number_format = DTFMT
        for c in (4, 5, 6, 7, 8):
            wi.cell(row=r, column=c).number_format = MONEY
    style_header(wi, 9)
    wi.freeze_panes = "A2"
    for w, col in zip([16, 12, 22, 12, 12, 13, 12, 12, 10], range(1, 10)):
        wi.column_dimensions[get_column_letter(col)].width = w

    # ---------------- Inventory (SKU master, all computed) ----------------
    wv = wb.create_sheet("Inventory")
    hdr = ["SKU", "Rate (ex-GST)", "Opening Stock", "Purchased In",
           "Sold Out", "Closing Stock", "Reorder Level", "Status", "Stock Value"]
    wv.append(hdr)
    for s in skus:
        wv.append([s["sku"], float(s["rate"]), int(s["opening_stock"]),
                   None, None, None, int(s["reorder_level"]), None, None])
    for r in range(2, NS + 2):
        wv.cell(row=r, column=4).value = (
            f'=IF($A{r}="",0,ROUND(SUMIFS(Purchases!$D$2:$D${NP+1},'
            f'Purchases!$B$2:$B${NP+1},$A{r}),2))')
        wv.cell(row=r, column=5).value = (
            f'=IF($A{r}="",0,ROUND(SUMIFS(InvoiceLines!$E$2:$E${NL+1},'
            f'InvoiceLines!$D$2:$D${NL+1},$A{r}),2))')
        wv.cell(row=r, column=6).value = f'=IF($A{r}="",0,$C{r}+$D{r}-$E{r})'
        wv.cell(row=r, column=8).value = (
            f'=IF($A{r}="","",IF($F{r}<=0,"OUT OF STOCK",'
            f'IF($F{r}<=$G{r},"REORDER","OK")))')
        wv.cell(row=r, column=9).value = f'=IF($A{r}="",0,ROUND($F{r}*$B{r},2))'
        wv.cell(row=r, column=2).number_format = MONEY
        wv.cell(row=r, column=9).number_format = MONEY
    style_header(wv, 9)
    for w, col in zip([26, 14, 14, 13, 12, 13, 13, 14, 13], range(1, 10)):
        wv.column_dimensions[get_column_letter(col)].width = w

    # ---------------- Purchases (raw stock-in) ----------------
    wp = wb.create_sheet("Purchases")
    wp.append(["Date", "SKU", "Lot No", "Qty", "Rate"])
    for p in purchases:
        y, m, dd = map(int, p["date"].split("-"))
        wp.append([date(y, m, dd), p["sku"], p["lot_no"],
                   int(p["qty"]), float(p["rate"])])
    for r in range(2, NP + 2):
        wp.cell(row=r, column=1).number_format = DTFMT
        wp.cell(row=r, column=5).number_format = MONEY
    style_header(wp, 5)
    wp.freeze_panes = "A2"
    for w, col in zip([12, 26, 10, 8, 12], range(1, 6)):
        wp.column_dimensions[get_column_letter(col)].width = w

    # ---------------- GST_Summary (monthly, 18% split) ----------------
    wg = wb.create_sheet("GST_Summary")
    wg.append(["Month", "Taxable Value", "CGST 9%", "SGST 9%", "Total GST"])
    months = [(2026, 4), (2026, 5), (2026, 6), (2026, 7), (2026, 8), (2026, 9)]
    for i, (yy, mm) in enumerate(months):
        r = i + 2
        nxt_m, nxt_y = (mm + 1, yy) if mm < 12 else (1, yy + 1)
        wg.cell(row=r, column=1, value=date(yy, mm, 1)).number_format = "MMM-YYYY"
        wg.cell(row=r, column=2).value = (
            f'=ROUND(SUMIFS(Invoices!$D$2:$D${NI+1},Invoices!$B$2:$B${NI+1},'
            f'">="&DATE({yy},{mm},1),Invoices!$B$2:$B${NI+1},'
            f'"<"&DATE({nxt_y},{nxt_m},1)),2)')
        wg.cell(row=r, column=3).value = f"=ROUND($B{r}*9%,2)"
        wg.cell(row=r, column=4).value = f"=ROUND($B{r}*9%,2)"
        wg.cell(row=r, column=5).value = f"=$C{r}+$D{r}"
        for c in range(2, 6):
            wg.cell(row=r, column=c).number_format = MONEY
    tr = len(months) + 2
    wg.cell(row=tr, column=1, value="TOTAL").font = Font(bold=True)
    for c, col in zip("BCDE", range(2, 6)):
        wg.cell(row=tr, column=col,
                value=f"=SUM({c}2:{c}{tr-1})").font = Font(bold=True)
        wg.cell(row=tr, column=col).number_format = MONEY
    style_header(wg, 5)
    for w, col in zip([12, 15, 13, 13, 13], range(1, 6)):
        wg.column_dimensions[get_column_letter(col)].width = w

    # ---------------- Daily (daily sales trend) ----------------
    wd = wb.create_sheet("Daily")
    wd.append(["Date", "Units Sold", "Revenue (ex-GST)"])
    start = date(2026, 4, 1)
    ndays = (date(2026, 9, 30) - start).days + 1
    for i in range(ndays):
        wd.cell(row=i + 2, column=1, value=start + timedelta(days=i)
                ).number_format = DTFMT
    for r in range(2, ND + 2):
        if r > ndays + 1:
            wd.cell(row=r, column=1).value = (
                f'=IF(A{r-1}="","",A{r-1}+1)')
            wd.cell(row=r, column=1).number_format = DTFMT
        wd.cell(row=r, column=2).value = (
            f'=IF($A{r}="",0,SUMIFS(InvoiceLines!$E$2:$E${NL+1},'
            f'InvoiceLines!$B$2:$B${NL+1},$A{r}))')
        wd.cell(row=r, column=3).value = (
            f'=IF($A{r}="",0,ROUND(SUMIFS(InvoiceLines!$G$2:$G${NL+1},'
            f'InvoiceLines!$B$2:$B${NL+1},$A{r}),2))')
        wd.cell(row=r, column=3).number_format = MONEY
    style_header(wd, 3)
    wd.freeze_panes = "A2"
    for w, col in zip([12, 12, 18], range(1, 4)):
        wd.column_dimensions[get_column_letter(col)].width = w

    # ---------------- Dashboard ----------------
    wb2 = wb.create_sheet("Dashboard", 0)
    wb2.sheet_properties.tabColor = "1F4E78"
    wb2["A1"] = "AquaBill — Billing, Sales & Inventory Dashboard"
    wb2["A1"].font = Font(bold=True, size=16, color="1F4E78")
    wb2["A2"] = "AquaPure Distributors, Patna  •  Apr–Sep 2026  •  All figures are live formulas"
    wb2["A2"].font = Font(italic=True, color="595959")
    kpis = [
        ("Total revenue (incl GST)", f"=SUM(Invoices!$F$2:$F${NI+1})", MONEY),
        ("Invoices issued", f"=COUNTA(Invoices!$A$2:$A${NI+1})", "0"),
        ("Units sold", f"=SUM(InvoiceLines!$E$2:$E${NL+1})", "0"),
        ("Avg invoice value", f'=IFERROR(AVERAGEIF(Invoices!$A$2:$A${NI+1},"<>",Invoices!$F$2:$F${NI+1}),0)', MONEY),
        ("GST collected (18%)", f"=SUM(Invoices!$E$2:$E${NI+1})", MONEY),
        ("Outstanding receivables", f"=SUM(Invoices!$H$2:$H${NI+1})", MONEY),
        ("SKUs at / below reorder level",
         f"=SUMPRODUCT((Inventory!$H$2:$H${NS+1}=\"REORDER\")*1)"
         f"+SUMPRODUCT((Inventory!$H$2:$H${NS+1}=\"OUT OF STOCK\")*1)", "0"),
        ("Stock value on hand", f"=SUM(Inventory!$I$2:$I${NS+1})", MONEY),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        r = 4 + i
        wb2.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = wb2.cell(row=r, column=2, value=formula)
        c.font = Font(bold=True, size=13, color="1F4E78")
        c.number_format = fmt
    wb2.column_dimensions["A"].width = 32
    wb2.column_dimensions["B"].width = 20

    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "Monthly taxable sales"
    ch1.y_axis.title = "INR"
    data = Reference(wg, min_col=2, min_row=1, max_row=7)
    cats = Reference(wg, min_col=1, min_row=2, max_row=7)
    ch1.add_data(data, titles_from_data=True)
    ch1.set_categories(cats)
    ch1.height, ch1.width = 7, 14
    wb2.add_chart(ch1, "A13")

    ch2 = BarChart()
    ch2.type = "bar"
    ch2.title = "Units sold by SKU"
    data2 = Reference(wv, min_col=5, min_row=1, max_row=7)
    cats2 = Reference(wv, min_col=1, min_row=2, max_row=7)
    ch2.add_data(data2, titles_from_data=True)
    ch2.set_categories(cats2)
    ch2.height, ch2.width = 7, 14
    wb2.add_chart(ch2, "A29")

    # ---------------- Checks (integrity) ----------------
    wc = wb.create_sheet("Checks")
    wc.append(["Check", "Result", "What it proves"])
    checks = [
        ("Invoice grand totals = subtotal + GST",
         f"=SUM(Invoices!$F$2:$F${NI+1})=SUM(Invoices!$D$2:$D${NI+1})+SUM(Invoices!$E$2:$E${NI+1})",
         "Register math ties out"),
        ("Every line total = ROUND(qty × rate, 2)",
         f"=SUMPRODUCT((InvoiceLines!$G$2:$G${NL+1}<>ROUND(InvoiceLines!$E$2:$E${NL+1}*InvoiceLines!$F$2:$F${NL+1},2))*(InvoiceLines!$A$2:$A${NL+1}<>\"\"))=0",
         "No line-item pricing errors"),
        ("GST ≈ 18% of taxable (within rounding)",
         f"=ABS(SUM(Invoices!$E$2:$E${NI+1})-SUM(Invoices!$D$2:$D${NI+1})*0.18)<0.5",
         "Tax math sane"),
        ("No negative closing stock",
         f"=SUMPRODUCT((Inventory!$F$2:$F${NS+1}<0)*1)=0",
         "Stock counts consistent"),
        ("Invoice numbers unique",
         f"=SUMPRODUCT((COUNTIF(Invoices!$A$2:$A${NI+1},Invoices!$A$2:$A${NI+1})>1)*(Invoices!$A$2:$A${NI+1}<>\"\"))=0",
         "No duplicate invoices"),
        ("Balance = grand − paid on every invoice",
         f"=SUMPRODUCT(ABS(Invoices!$H$2:$H${NI+1}-(Invoices!$F$2:$F${NI+1}-Invoices!$G$2:$G${NI+1})))<0.01",
         "Receivables reconcile"),
        ("GST summary ties to invoice register",
         f"=ABS(SUM(GST_Summary!$B$2:$B$7)-SUM(Invoices!$D$2:$D${NI+1}))<0.01",
         "Monthly tax report reconciles"),
        ("Daily revenue ties to invoice lines",
         f"=ABS(SUM(Daily!$C$2:$C${ND+1})-SUM(InvoiceLines!$G$2:$G${NL+1}))<0.01",
         "Daily trend reconciles"),
    ]
    for name, formula, note in checks:
        wc.append([name, formula, note])
    for r in range(2, 2 + len(checks)):
        wc.cell(row=r, column=2).alignment = Alignment(horizontal="center")
    style_header(wc, 3)
    for w, col in zip([44, 10, 34], range(1, 4)):
        wc.column_dimensions[get_column_letter(col)].width = w

    # ---------------- Guide ----------------
    wgd = wb.create_sheet("Guide")
    guide = [
        ("AquaBill — operator guide (AquaPure Distributors, Patna)", ""),
        ("", ""),
        ("What this is",
         "One workbook that runs billing, sales tracking and inventory for a packaged "
         "drinking-water distributor. Every figure is a live formula over the raw tables — "
         "paste new rows and the dashboard, GST summary and stock levels update."),
        ("Add an invoice",
         "1) Append one row per line item on InvoiceLines (Inv No, Date, Line, SKU, Qty, Rate). "
         "2) Append one row on Invoices (Inv No, Date, Customer, Paid). Subtotal, GST, Grand, "
         "Balance and Status compute themselves."),
        ("Record stock-in",
         "Append rows on Purchases (Date, SKU, Lot No, Qty, Rate). Inventory 'Purchased In' "
         "and 'Closing Stock' update automatically."),
        ("Record a payment",
         "Update the Paid column on the invoice's row in Invoices. Balance and Status "
         "(Paid / Partial / Unpaid) recalc; Outstanding receivables on the dashboard follows."),
        ("Change the GST rate",
         "Find/replace '18%' in the Invoices sheet formulas (and 9% in GST_Summary). "
         "The Checks sheet will confirm the new rate ties out."),
        ("Trust but verify",
         "Open the Checks sheet after every paste — all 8 checks must read TRUE. "
         "They reconcile the register, tax, stock and daily trend against each other."),
        ("Note",
         "All customers, transactions and figures are fictional demo data generated "
         "deterministically (generate.py --seed 20261013) for evaluation."),
    ]
    for title, body in guide:
        wgd.append([title, body])
    wgd.column_dimensions["A"].width = 34
    wgd.column_dimensions["B"].width = 130
    for r in range(1, len(guide) + 1):
        wgd.cell(row=r, column=1).font = Font(bold=True)
        wgd.cell(row=r, column=1).alignment = Alignment(vertical="top", wrap_text=True)
        wgd.cell(row=r, column=2).alignment = Alignment(vertical="top", wrap_text=True)
    wgd["A1"].font = Font(bold=True, size=14, color="1F4E78")

    wb.save(outpath)
    print(f"wrote {outpath}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--datadir", default="data")
    ap.add_argument("--out", default="water-bottle-billing-inventory.xlsx")
    args = ap.parse_args()
    build(args.datadir, args.out)


if __name__ == "__main__":
    main()
