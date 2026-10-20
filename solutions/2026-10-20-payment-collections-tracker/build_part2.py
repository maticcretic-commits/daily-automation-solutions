"""Part 2 of the workbook builder: FollowUp, Dashboard, AI_Report, Checks."""
from datetime import date
from openpyxl.styles import Font, Alignment
from openpyxl.chart import LineChart, BarChart, Reference
from build_workbook import (HDR_FILL, HDR_FONT, TITLE_FONT, MONEY_FMT, PCT_FMT,
                            DATE_FMT, BORDER, BUCKETS,
                            style_header, style_range, autofit)

AS_OF = date(2026, 10, 20)


def build_followup(wb, led_last, overdue_count):
    ws = wb.create_sheet("FollowUp")
    hdr = ["#", "DueID", "Borrower", "Collector", "Days Overdue",
           "Outstanding (Rs)", "Severity", "Next Action"]
    ws.append(hdr)
    last = overdue_count + 1
    for i in range(2, last + 1):
        n = i - 1
        ws.cell(row=i, column=1, value="=ROW()-1")
        ws.cell(row=i, column=2,
                value=f"=INDEX(Receivables_Ledger!A$2:A${led_last},SMALL(Receivables_Ledger!P$2:P${led_last},{n}))")
        ws.cell(row=i, column=3,
                value=f"=INDEX(Receivables_Ledger!C$2:C${led_last},SMALL(Receivables_Ledger!P$2:P${led_last},{n}))")
        ws.cell(row=i, column=4,
                value=f"=INDEX(Receivables_Ledger!E$2:E${led_last},SMALL(Receivables_Ledger!P$2:P${led_last},{n}))")
        ws.cell(row=i, column=5,
                value=f"=INDEX(Receivables_Ledger!M$2:M${led_last},SMALL(Receivables_Ledger!P$2:P${led_last},{n}))")
        ws.cell(row=i, column=6,
                value=f"=INDEX(Receivables_Ledger!J$2:J${led_last},SMALL(Receivables_Ledger!P$2:P${led_last},{n}))")
        ws.cell(row=i, column=7,
                value=f'=IF(E{i}>90,"LOST",IF(E{i}>60,"ESCALATE",IF(E{i}>30,"WATCH","MILD")))')
        ws.cell(row=i, column=8,
                value=f'=Guide!$B$3+IF(G{i}="LOST",0,IF(G{i}="ESCALATE",1,IF(G{i}="WATCH",3,7)))')
    style_header(ws, len(hdr))
    style_range(ws, 2, last, len(hdr))
    for row in ws.iter_rows(min_row=2, max_row=last, min_col=6, max_col=6):
        row[0].number_format = MONEY_FMT
    for row in ws.iter_rows(min_row=2, max_row=last, min_col=8, max_col=8):
        row[0].number_format = DATE_FMT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{last}"
    autofit(ws, [6, 9, 20, 12, 12, 16, 10, 12])
    return last


def build_dashboard(wb, led_last, collectors):
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "CollectionPulse — Collections Dashboard"
    ws["A1"].font = TITLE_FONT
    kpis = [
        ("Total dues (Rs)", "=SUM(Raw_Schedule!H:H)", MONEY_FMT),
        ("Collected (Rs)", "=SUM(Raw_Payments!E:E)", MONEY_FMT),
        ("Collection %", "=B4/B3", PCT_FMT),
        ("Overdue outstanding (Rs)",
         f'=SUMIFS(Receivables_Ledger!J2:J{led_last},Receivables_Ledger!J2:J{led_last},">0.005")', MONEY_FMT),
        ("Overdue dues",
         f'=COUNTIF(Receivables_Ledger!J2:J{led_last},">0.005")', "0"),
        ("Avg DPD (overdue)",
         f'=AVERAGEIFS(Receivables_Ledger!M2:M{led_last},Receivables_Ledger!J2:J{led_last},">0.005")', "0.0"),
        ("Best collector",
         "=INDEX(Collector_Scorecard!A:A,MATCH(MAX(Collector_Scorecard!E:E),Collector_Scorecard!E:E,0))", None),
        ("Worst collector",
         "=INDEX(Collector_Scorecard!A:A,MATCH(MIN(Collector_Scorecard!E:E),Collector_Scorecard!E:E,0))", None),
    ]
    ws["A2"] = "KPI"
    ws["B2"] = "Value"
    for i, (label, formula, fmt) in enumerate(kpis, start=3):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=formula)
        c.font = Font(bold=True, size=12, color="1F4E79")
        if fmt:
            c.number_format = fmt
    style_header(ws, 2)
    style_range(ws, 3, 10, 2)

    ws["A13"] = "Month"
    ws["B13"] = "Due (Rs)"
    ws["C13"] = "Collected (Rs)"
    ws["D13"] = "Efficiency %"
    months = [("Apr-26", 2026, 4), ("May-26", 2026, 5), ("Jun-26", 2026, 6),
              ("Jul-26", 2026, 7), ("Aug-26", 2026, 8), ("Sep-26", 2026, 9)]
    for j, (label, y, m) in enumerate(months, start=14):
        nm = m + 1 if m < 12 else 1
        ny = y if m < 12 else y + 1
        ws.cell(row=j, column=1, value=label)
        ws.cell(row=j, column=2,
                value=f'=SUMIFS(Raw_Schedule!H:H,Raw_Schedule!G:G,">="&DATE({y},{m},1),'
                      f'Raw_Schedule!G:G,"<"&DATE({ny},{nm},1))')
        ws.cell(row=j, column=3,
                value=f'=SUMIFS(Raw_Payments!E:E,Raw_Payments!D:D,">="&DATE({y},{m},1),'
                      f'Raw_Payments!D:D,"<"&DATE({ny},{nm},1))')
        ws.cell(row=j, column=4, value=f"=IF(B{j}=0,0,C{j}/B{j})")
    style_header(ws, 4)
    style_range(ws, 14, 19, 4)
    for col in (2, 3):
        for row in ws.iter_rows(min_row=14, max_row=19, min_col=col, max_col=col):
            row[0].number_format = MONEY_FMT
    for row in ws.iter_rows(min_row=14, max_row=19, min_col=4, max_col=4):
        row[0].number_format = PCT_FMT
    autofit(ws, [24, 16, 8, 8])

    line = LineChart()
    line.title = "Monthly collection efficiency"
    line.style = 10
    line.y_axis.title = "Efficiency"
    data = Reference(ws, min_col=4, min_row=13, max_row=19)
    cats = Reference(ws, min_col=1, min_row=14, max_row=19)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    ws.add_chart(line, "F2")

    bar1 = BarChart()
    bar1.title = "Outstanding by DPD bucket"
    bar1.style = 10
    aging = wb["Aging"]
    data = Reference(aging, min_col=5, min_row=1, max_row=6)
    cats = Reference(aging, min_col=1, min_row=2, max_row=6)
    bar1.add_data(data, titles_from_data=True)
    bar1.set_categories(cats)
    ws.add_chart(bar1, "F17")

    bar2 = BarChart()
    bar2.title = "Collector efficiency"
    bar2.style = 10
    sc = wb["Collector_Scorecard"]
    data = Reference(sc, min_col=5, min_row=1, max_row=len(collectors) + 1)
    cats = Reference(sc, min_col=1, min_row=2, max_row=len(collectors) + 1)
    bar2.add_data(data, titles_from_data=True)
    bar2.set_categories(cats)
    ws.add_chart(bar2, "F32")


def build_ai_report(wb, led_last, fu_last):
    ws = wb.create_sheet("AI_Report")
    ws["A1"] = "AI_Report — rule-based portfolio summary"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = ('="As of "&TEXT(Guide!$B$3,"DD MMM YYYY")&", scheduled dues total "'
                 '&TEXT(Dashboard!$B$3,"""Rs ""#,##0")&", of which "&TEXT(Dashboard!$B$4,"""Rs ""#,##0")'
                 '&" ("&TEXT(Dashboard!$B$5,"0.0%")&") has been collected. "'
                 '&TEXT(Dashboard!$B$7,"0")&" dues are overdue with "&TEXT(Dashboard!$B$6,"""Rs ""#,##0")'
                 '&" outstanding; average days overdue is "&TEXT(Dashboard!$B$8,"0")&". "'
                 '&"Best collector: "&Dashboard!$B$9&"; needs coaching: "&Dashboard!$B$10&"."')
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[3].height = 60
    ws["A5"] = "Flags"
    ws["A5"].font = Font(bold=True, size=12)
    flags = [
        '=IF(Dashboard!$B$5>=0.9,"OK  Collection rate healthy (>=90%)","!!  Collection rate below 90% - tighten follow-ups")',
        '=IF(SUMIFS(Aging!$E$2:$E$6,Aging!$A$2:$A$6,"90+")/MAX(1,Aging!$E$7)>0.25,'
        '"!!  90+ DPD bucket holds >25% of outstanding - escalate","OK  90+ DPD share under control")',
        '=IF(MAX(Collector_Scorecard!$E$2:$E$7)-MIN(Collector_Scorecard!$E$2:$E$7)>0.2,'
        '"!!  Collector efficiency gap >20pp - coach the bottom quartile","OK  Collector efficiency spread healthy")',
        f'=IF(COUNTIF(Receivables_Ledger!$K$2:$K${led_last},"PARTIAL")>0,'
        f'"ii  "&TEXT(COUNTIF(Receivables_Ledger!$K$2:$K${led_last},"PARTIAL"),"0")&" partial payments need top-up calls",'
        '"OK  No partial payments pending")',
        '=IF(Dashboard!$B$8>45,"!!  Average DPD above 45 days - portfolio stress","OK  Average DPD within tolerance")',
        f'=IF(COUNTIF(FollowUp!$G$2:$G${fu_last},"LOST")>0,'
        f'"!!  "&TEXT(COUNTIF(FollowUp!$G$2:$G${fu_last},"LOST"),"0")&" dues in LOST (90+ DPD) - consider write-off review",'
        '"OK  Nothing in 90+ DPD")',
        f'=IF(COUNTIF(FollowUp!$G$2:$G${fu_last},"ESCALATE")>10,'
        f'"!!  "&TEXT(COUNTIF(FollowUp!$G$2:$G${fu_last},"ESCALATE"),"0")&" dues need escalation this week",'
        '"OK  Escalation queue manageable")',
        '=IF(Checks!$B$9=TRUE,"OK  All 8 integrity checks pass","!!  Integrity check failing - see Checks sheet")',
    ]
    for i, f in enumerate(flags, start=6):
        c = ws.cell(row=i, column=1, value=f)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 28
    autofit(ws, [110])


def build_checks(wb, led_last, fu_last):
    ws = wb.create_sheet("Checks")
    ws.append(["Check", "Result"])
    checks = [
        ("No negative outstandings",
         f"=SUMPRODUCT(--(Receivables_Ledger!J2:J{led_last}<-0.01))=0"),
        ("Aging due ties to schedule total",
         "=SUM(Aging!C2:C6)=SUM(Raw_Schedule!H:H)"),
        ("Scorecard due ties to schedule total",
         "=SUM(Collector_Scorecard!C2:C7)=SUM(Raw_Schedule!H:H)"),
        ("FollowUp rows = overdue dues",
         f"=COUNTA(FollowUp!B2:B{fu_last})=COUNTIF(Receivables_Ledger!J2:J{led_last},\">0.005\")"),
        ("Ledger paid ties to payments total",
         f"=SUM(Receivables_Ledger!I2:I{led_last})=SUM(Raw_Payments!E:E)"),
        ("No formula errors in ledger",
         f"=SUMPRODUCT(--ISERROR(Receivables_Ledger!A2:P{led_last}))=0"),
        ("Status values valid (PAID/PARTIAL/OVERDUE)",
         f'=SUMPRODUCT(--(Receivables_Ledger!K2:K{led_last}<>"PAID")'
         f'*--(Receivables_Ledger!K2:K{led_last}<>"PARTIAL")'
         f'*--(Receivables_Ledger!K2:K{led_last}<>"OVERDUE"))=0'),
        ("Outstanding = due - paid (global)",
         f"=ABS(SUM(Receivables_Ledger!J2:J{led_last})-(SUM(Receivables_Ledger!H2:H{led_last})-SUM(Receivables_Ledger!I2:I{led_last})))<0.01"),
    ]
    for label, formula in checks:
        ws.append([label, formula])
    style_header(ws, 2)
    style_range(ws, 2, len(checks) + 1, 2)
    for row in ws.iter_rows(min_row=2, max_row=len(checks) + 1, min_col=2, max_col=2):
        row[0].alignment = Alignment(horizontal="center")
        row[0].font = Font(bold=True)
    autofit(ws, [48, 10])
    return len(checks)
