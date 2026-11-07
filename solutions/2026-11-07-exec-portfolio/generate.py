#!/usr/bin/env python3
"""ExecPortfolio — executive portfolio reporting factory (Excel side of an
M365 Power BI brief).

Built around a real Freelancer brief ("Executive Portfolio Power BI
Dashboards -- 2", project 40806608): end-to-end M365 executive portfolio
reporting — Power BI dashboards fed from Excel, Planner and SharePoint via
Power Query, with Power Automate refresh and alerts.

What this builds: the Excel foundation that feeds that Power BI model.
Source tables mimic the three M365 feeds (Excel project master, Planner task
export, SharePoint risk list) plus a monthly spend staging table, and every
reporting figure is a live Excel formula — nothing hard-coded except the
source rows themselves and the as-of date. In Power BI these same tables
load through Power Query (one query per sheet) and the Portfolio rollup
becomes the fact table; the Dashboard KPIs map 1:1 to DAX measures.

Demo portfolio: Meridian Group, a fictional company — 12 projects across
4 programmes, 60 Planner-style tasks, 24 SharePoint-style risks, monthly
spend Jun–Nov 2026. All data is fictional and deterministic (seeded).

Sheets: Guide, Projects, Tasks, Risks, Spend, Portfolio (per-project
rollup), Monthly (spend by programme), Dashboard (9 KPIs + 3 native
charts), Checks. No macros, no structured table references, plain A1
ranges with the column letter repeated on both ends (LibreOffice-safe).

Model rules (all visible in the formulas):
- Portfolio % complete is the AVERAGEIF of its tasks' % complete.
- A task is OVERDUE when its due date is before the as-of date
  (2026-11-07, cell Portfolio!M1) and its status is not Completed.
- RAG: Green when spend variance <= 10% over budget AND no overdue
  tasks; Red when variance > 25% over budget OR 5+ overdue tasks;
  Amber otherwise.
- Risk score = Likelihood (1-5) x Impact (1-5).
- Monthly spend is SUMIFS over the Spend staging table, which carries
  the programme denormalised — exactly what a Power Query merge of the
  Excel master into the spend extract produces.
"""
import csv
import datetime
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
AS_OF = datetime.date(2026, 11, 7)
MONTHS = ["2026-06", "2026-07", "2026-08", "2026-09", "2026-10", "2026-11"]

NAVY = "1F3864"
TEAL = "0F7B7A"
LIGHT = "D9E2F3"
MONEY = "$#,##0"
PCT1 = "0.0%"
PCT0 = "0%"

PROGRAMMES = [
    "Digital Transformation",
    "Operations Excellence",
    "Growth & Markets",
    "Compliance & Risk",
]

PROJECTS = [
    # (id, name, programme, sponsor, owner, phase, start, end, budget $)
    ("P01", "CRM Consolidation", "Digital Transformation", "A. Mehta", "R. Iyer",
     "Execution", "2026-04-01", "2026-12-15", 240_000),
    ("P02", "SharePoint Intranet Rebuild", "Digital Transformation", "S. Rao", "K. Nair",
     "Execution", "2026-05-01", "2027-01-31", 180_000),
    ("P03", "Power Automate Approvals", "Digital Transformation", "A. Mehta", "T. Das",
     "Execution", "2026-06-01", "2026-11-30", 95_000),
    ("P04", "Warehouse Scanning Rollout", "Operations Excellence", "V. Kapoor", "M. Ali",
     "Execution", "2026-03-15", "2026-10-31", 210_000),
    ("P05", "Fleet Telematics", "Operations Excellence", "V. Kapoor", "P. Shah",
     "Planning", "2026-07-01", "2027-03-31", 160_000),
    ("P06", "Supplier Scorecard", "Operations Excellence", "N. Bose", "R. Iyer",
     "Execution", "2026-05-15", "2026-12-31", 75_000),
    ("P07", "New Market Entry — East", "Growth & Markets", "D. Roy", "S. Patel",
     "Execution", "2026-04-01", "2027-02-28", 250_000),
    ("P08", "Partner Portal", "Growth & Markets", "D. Roy", "K. Nair",
     "Planning", "2026-08-01", "2027-04-30", 130_000),
    ("P09", "Pricing Engine Revamp", "Growth & Markets", "S. Rao", "T. Das",
     "Execution", "2026-06-15", "2026-12-31", 145_000),
    ("P10", "DPDP Compliance Programme", "Compliance & Risk", "N. Bose", "M. Ali",
     "Execution", "2026-05-01", "2026-11-30", 110_000),
    ("P11", "Vendor Risk Register", "Compliance & Risk", "N. Bose", "P. Shah",
     "Execution", "2026-07-01", "2027-01-31", 65_000),
    ("P12", "Audit Trail Automation", "Compliance & Risk", "A. Mehta", "S. Patel",
     "Initiation", "2026-09-01", "2027-05-31", 85_000),
]

TASK_TITLES = [
    "Scope sign-off", "Data extract & cleanse", "Stakeholder workshop",
    "Build first draft", "UAT round 1", "Fix UAT defects", "Go-live checklist",
    "Training session", "Handover document", "Steering committee update",
    "Vendor evaluation", "Pilot with two teams", "Rollout plan",
    "Budget review", "Risk review", "Cutover rehearsal",
]

RISK_TITLES = [
    "Key sponsor availability", "Vendor delivery slips", "Data quality gaps",
    "Scope creep from steering group", "Integration with legacy system",
    "Staff turnover in delivery team", "Regulatory change mid-project",
    "Third-party API limits",
]

TASK_STATUSES = ["Not started", "In progress", "Completed", "Blocked"]


def build_data():
    rng = random.Random(20261107)
    tasks = []
    tid = 1
    for pid, *_ in PROJECTS:
        titles = rng.sample(TASK_TITLES, 5)
        for title in titles:
            roll = rng.random()
            if roll < 0.45:
                status, pct = "Completed", 1.0
            elif roll < 0.80:
                status, pct = "In progress", rng.choice([0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9])
            elif roll < 0.95:
                status, pct = "Not started", 0.0
            else:
                status, pct = "Blocked", rng.choice([0.1, 0.25, 0.4])
            # Due dates Oct 1 – Dec 15 2026: some land before the as-of date.
            # Real date values (not text) so COUNTIFS "<"&as-of compares
            # numerically in both Excel and LibreOffice.
            month = rng.choice([10, 10, 11, 11, 12])
            day = rng.randint(1, 28) if month != 12 else rng.randint(1, 15)
            due = datetime.date(2026, month, day)
            owner = rng.choice(["R. Iyer", "K. Nair", "T. Das", "M. Ali", "P. Shah", "S. Patel"])
            priority = rng.choice(["High", "Medium", "Medium", "Low"])
            tasks.append((f"T{tid:03d}", pid, title, owner, status, due, pct, priority))
            tid += 1
    risks = []
    rid = 1
    for pid, *_ in PROJECTS:
        for title in rng.sample(RISK_TITLES, 2):
            likelihood = rng.randint(1, 5)
            impact = rng.randint(1, 5)
            mitigation = rng.choice(["Open", "In progress", "In progress", "Closed"])
            risks.append((f"R{rid:02d}", pid, title, likelihood, impact, mitigation))
            rid += 1
    spend = []
    for pid, _, prog, *_rest, budget in [(p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8]) for p in PROJECTS]:
        monthly_base = budget / 6.0
        for m in MONTHS:
            # Early months run light for late-starting projects.
            factor = rng.uniform(0.7, 1.3)
            amount = round(monthly_base * factor / 100) * 100
            spend.append((m, pid, prog, amount))
    return tasks, risks, spend


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def money_range(ws, rows, cols, fmt=MONEY):
    for r in rows:
        for c in cols:
            ws.cell(row=r, column=c).number_format = fmt


def build_workbook(tasks, risks, spend):
    wb = Workbook()
    n_proj = len(PROJECTS)
    n_task = len(tasks)
    n_risk = len(risks)
    n_spend = len(spend)
    p_first, p_last = 2, 1 + n_proj          # Projects rows
    t_first, t_last = 2, 1 + n_task          # Tasks rows
    r_first, r_last = 2, 1 + n_risk          # Risks rows
    s_first, s_last = 2, 1 + n_spend         # Spend rows
    pf_first, pf_last = 3, 2 + n_proj         # Portfolio rows (row 2 = header)
    m_first, m_last = 2, 1 + len(MONTHS)     # Monthly rows

    # ---- Projects (Excel master — inputs) ----
    ws = wb.active
    ws.title = "Projects"
    headers = ["Project ID", "Project", "Programme", "Sponsor", "Owner", "Phase",
               "Start", "End", "Budget ($)"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, p in enumerate(PROJECTS):
        r = p_first + i
        for c, v in enumerate(p, 1):
            cell = ws.cell(row=r, column=c, value=v)
            if c in (7, 8):  # Start / End as real dates, not text.
                cell.value = datetime.date.fromisoformat(v)
                cell.number_format = "yyyy-mm-dd"
    money_range(ws, range(p_first, p_last + 1), [9])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{p_last}"
    for i, w in enumerate([11, 28, 22, 11, 11, 12, 12, 12, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- Tasks (Planner export — inputs) ----
    tk = wb.create_sheet("Tasks")
    tk_headers = ["Task ID", "Project ID", "Task", "Owner", "Status", "Due date",
                  "% Complete", "Priority"]
    tk.append(tk_headers)
    style_header(tk, 1, len(tk_headers))
    for i, t in enumerate(tasks):
        r = t_first + i
        for c, v in enumerate(t, 1):
            tk.cell(row=r, column=c, value=v)
    for r in range(t_first, t_last + 1):
        tk.cell(row=r, column=7).number_format = PCT0
        tk.cell(row=r, column=6).number_format = "yyyy-mm-dd"
    tk.freeze_panes = "A2"
    tk.auto_filter.ref = f"A1:H{t_last}"
    for i, w in enumerate([10, 11, 26, 11, 13, 12, 12, 10], 1):
        tk.column_dimensions[get_column_letter(i)].width = w

    # ---- Risks (SharePoint list — inputs) ----
    rk = wb.create_sheet("Risks")
    rk_headers = ["Risk ID", "Project ID", "Risk", "Likelihood (1-5)",
                  "Impact (1-5)", "Risk score", "Mitigation status"]
    rk.append(rk_headers)
    style_header(rk, 1, len(rk_headers))
    for i, risk in enumerate(risks):
        r = r_first + i
        for c, v in enumerate(risk, 1):
            rk.cell(row=r, column=c, value=v)
        rk.cell(row=r, column=6, value=f"=D{r}*E{r}")
    rk.freeze_panes = "A2"
    rk.auto_filter.ref = f"A1:G{r_last}"
    for i, w in enumerate([10, 11, 36, 15, 12, 11, 17], 1):
        rk.column_dimensions[get_column_letter(i)].width = w

    # ---- Spend (staging — inputs, programme denormalised by the merge) ----
    sp = wb.create_sheet("Spend")
    sp_headers = ["Month", "Project ID", "Programme", "Actual spend ($)"]
    sp.append(sp_headers)
    style_header(sp, 1, len(sp_headers))
    for i, s in enumerate(spend):
        r = s_first + i
        for c, v in enumerate(s, 1):
            sp.cell(row=r, column=c, value=v)
    money_range(sp, range(s_first, s_last + 1), [4])
    sp.freeze_panes = "A2"
    sp.auto_filter.ref = f"A1:D{s_last}"
    for i, w in enumerate([10, 11, 22, 17], 1):
        sp.column_dimensions[get_column_letter(i)].width = w

    # ---- Portfolio (rollup — every cell a formula) ----
    pf = wb.create_sheet("Portfolio")
    pf["L1"] = "As-of date"
    pf["M1"] = AS_OF
    pf["M1"].number_format = "yyyy-mm-dd"
    pf["L1"].font = Font(bold=True)
    pf_headers = ["Project ID", "Project", "Programme", "Budget ($)", "Spent ($)",
                  "Variance %", "% Complete", "Tasks", "Done", "Overdue",
                  "RAG", "Top risk score"]
    pf.append(pf_headers)
    style_header(pf, 2, len(pf_headers))
    for i in range(n_proj):
        r = pf_first + i
        pr = p_first + i  # same row index on Projects
        pf.cell(row=r, column=1, value=f"=Projects!A{pr}")
        pf.cell(row=r, column=2, value=f"=Projects!B{pr}")
        pf.cell(row=r, column=3, value=f"=Projects!C{pr}")
        pf.cell(row=r, column=4, value=f"=Projects!I{pr}")
        pf.cell(row=r, column=5,
               value=f"=SUMIFS(Spend!$D${s_first}:$D${s_last},Spend!$B${s_first}:$B${s_last},$A{r})")
        pf.cell(row=r, column=6, value=f"=IF(D{r}=0,0,(E{r}-D{r})/D{r})")
        pf.cell(row=r, column=7,
               value=f"=AVERAGEIF(Tasks!$B${t_first}:$B${t_last},$A{r},Tasks!$G${t_first}:$G${t_last})")
        pf.cell(row=r, column=8,
               value=f"=COUNTIF(Tasks!$B${t_first}:$B${t_last},$A{r})")
        pf.cell(row=r, column=9,
               value=f"=COUNTIFS(Tasks!$B${t_first}:$B${t_last},$A{r},Tasks!$E${t_first}:$E${t_last},\"Completed\")")
        pf.cell(row=r, column=10,
               value=f"=COUNTIFS(Tasks!$B${t_first}:$B${t_last},$A{r},Tasks!$F${t_first}:$F${t_last},\"<\"&$M$1,Tasks!$E${t_first}:$E${t_last},\"<>Completed\")")
        pf.cell(row=r, column=11,
               value=f"=IF(AND(F{r}<=0.1,J{r}=0),\"Green\",IF(OR(F{r}>0.25,J{r}>=5),\"Red\",\"Amber\"))")
        # MAXIFS is #NAME? in this LibreOffice build and AGGREGATE does not
        # force array evaluation here; SUMPRODUCT(MAX((cond)*range)) is the
        # LO-safe conditional-max pattern.
        pf.cell(row=r, column=12,
               value=f"=SUMPRODUCT(MAX((Risks!$B${r_first}:$B${r_last}=$A{r})*Risks!$F${r_first}:$F${r_last}))")
    total_row = pf_last + 1
    pf.cell(row=total_row, column=1, value="TOTAL")
    pf.cell(row=total_row, column=1).font = Font(bold=True)
    for c in (4, 5, 8, 9, 10):
        col = get_column_letter(c)
        pf.cell(row=total_row, column=c,
               value=f"=SUM({col}{pf_first}:{col}{pf_last})").font = Font(bold=True)
    pf.cell(row=total_row, column=6,
           value=f"=IF(D{total_row}=0,0,(E{total_row}-D{total_row})/D{total_row})").font = Font(bold=True)
    pf.cell(row=total_row, column=7,
           value=f"=AVERAGE(G{pf_first}:G{pf_last})").font = Font(bold=True)
    money_range(pf, list(range(pf_first, pf_last + 1)) + [total_row], [4, 5])
    for r in list(range(pf_first, pf_last + 1)) + [total_row]:
        pf.cell(row=r, column=6).number_format = PCT1
        pf.cell(row=r, column=7).number_format = PCT0
    pf.freeze_panes = "A3"
    pf.auto_filter.ref = f"A2:L{pf_last}"
    for i, w in enumerate([11, 28, 22, 13, 13, 11, 11, 8, 8, 9, 9, 13], 1):
        pf.column_dimensions[get_column_letter(i)].width = w

    # ---- Monthly (spend by programme — formulas) ----
    mo = wb.create_sheet("Monthly")
    mo_headers = ["Month", "Total spend ($)"] + PROGRAMMES
    mo.append(mo_headers)
    style_header(mo, 1, len(mo_headers))
    for i, m in enumerate(MONTHS):
        r = m_first + i
        mo.cell(row=r, column=1, value=m)
        mo.cell(row=r, column=2,
               value=f"=SUMIFS(Spend!$D${s_first}:$D${s_last},Spend!$A${s_first}:$A${s_last},$A{r})")
        for j, prog in enumerate(PROGRAMMES):
            mo.cell(row=r, column=3 + j,
                   value=f"=SUMIFS(Spend!$D${s_first}:$D${s_last},Spend!$A${s_first}:$A${s_last},$A{r},Spend!$C${s_first}:$C${s_last},\"{prog}\")")
    m_total = m_last + 1
    mo.cell(row=m_total, column=1, value="TOTAL")
    mo.cell(row=m_total, column=1).font = Font(bold=True)
    for c in range(2, 2 + len(PROGRAMMES) + 1):
        col = get_column_letter(c)
        mo.cell(row=m_total, column=c,
               value=f"=SUM({col}{m_first}:{col}{m_last})").font = Font(bold=True)
    money_range(mo, list(range(m_first, m_last + 1)) + [m_total], range(2, 7))
    mo.freeze_panes = "A2"
    for i, w in enumerate([10, 16, 22, 22, 22, 22], 1):
        mo.column_dimensions[get_column_letter(i)].width = w

    # ---- Dashboard ----
    db = wb.create_sheet("Dashboard")
    db.append(["Meridian Group — Executive Portfolio Dashboard (fictional demo data)"])
    db["A1"].font = Font(bold=True, size=14, color=NAVY)
    db.append(["As-of date", f"=Portfolio!M1"])
    db["B2"].number_format = "yyyy-mm-dd"
    db.append(["KPI", "Value"])
    style_header(db, 3, 2)
    pf_tot = pf_last + 1
    kpis = [
        ("Total portfolio budget", f"=Portfolio!D{pf_tot}", MONEY),
        ("Total spent to date", f"=Portfolio!E{pf_tot}", MONEY),
        ("Spend vs budget", f"=Portfolio!F{pf_tot}", PCT1),
        ("Projects on track (Green)", f"=COUNTIF(Portfolio!K{pf_first}:K{pf_last},\"Green\")", "0"),
        ("Projects at risk (Red)", f"=COUNTIF(Portfolio!K{pf_first}:K{pf_last},\"Red\")", "0"),
        ("Overdue tasks", f"=Portfolio!J{pf_tot}", "0"),
        ("Tasks completed", f"=Portfolio!I{pf_tot}&\" / \"&Portfolio!H{pf_tot}", "@"),
        ("Task completion", f"=IF(Portfolio!H{pf_tot}=0,0,Portfolio!I{pf_tot}/Portfolio!H{pf_tot})", PCT0),
        ("Average project completion", f"=Portfolio!G{pf_tot}", PCT0),
        ("Highest open risk score", f"=MAX(Risks!F{r_first}:F{r_last})", "0"),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        r = 4 + i
        db.cell(row=r, column=1, value=label)
        c = db.cell(row=r, column=2, value=formula)
        c.number_format = fmt
    db.column_dimensions["A"].width = 30
    db.column_dimensions["B"].width = 20

    # Programme summary table (drives chart 1).
    prog_head_row = 4 + len(kpis) + 2
    db.cell(row=prog_head_row, column=1, value="Programme").font = Font(bold=True)
    db.cell(row=prog_head_row, column=2, value="Projects").font = Font(bold=True)
    db.cell(row=prog_head_row, column=3, value="Budget ($)").font = Font(bold=True)
    db.cell(row=prog_head_row, column=4, value="Spent ($)").font = Font(bold=True)
    db.cell(row=prog_head_row, column=5, value="Avg % complete").font = Font(bold=True)
    for j, prog in enumerate(PROGRAMMES):
        r = prog_head_row + 1 + j
        db.cell(row=r, column=1, value=prog)
        db.cell(row=r, column=2,
               value=f"=COUNTIF(Portfolio!$C${pf_first}:$C${pf_last},$A{r})")
        db.cell(row=r, column=3,
               value=f"=SUMIF(Portfolio!$C${pf_first}:$C${pf_last},$A{r},Portfolio!$D${pf_first}:$D${pf_last})")
        db.cell(row=r, column=4,
               value=f"=SUMIF(Portfolio!$C${pf_first}:$C${pf_last},$A{r},Portfolio!$E${pf_first}:$E${pf_last})")
        db.cell(row=r, column=5,
               value=f"=AVERAGEIF(Portfolio!$C${pf_first}:$C${pf_last},$A{r},Portfolio!$G${pf_first}:$G${pf_last})")
        db.cell(row=r, column=5).number_format = PCT0
        db.cell(row=r, column=3).number_format = MONEY
        db.cell(row=r, column=4).number_format = MONEY
    for i, w in enumerate([30, 20, 14, 14, 16], 1):
        db.column_dimensions[get_column_letter(i)].width = w
    prog_first = prog_head_row + 1
    prog_last = prog_head_row + len(PROGRAMMES)

    cats_prog = Reference(db, min_col=1, min_row=prog_first, max_row=prog_last)
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "Budget vs spent by programme"
    ch1.y_axis.number_format = "$#,##0"
    ch1.add_data(Reference(db, min_col=3, min_row=prog_head_row, max_row=prog_last),
                 titles_from_data=True)
    ch1.add_data(Reference(db, min_col=4, min_row=prog_head_row, max_row=prog_last),
                 titles_from_data=True)
    ch1.set_categories(cats_prog)
    ch1.height, ch1.width = 8, 16
    db.add_chart(ch1, "G3")

    cats_proj = Reference(pf, min_col=2, min_row=pf_first, max_row=pf_last)
    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "% complete by project"
    ch2.y_axis.number_format = "0%"
    ch2.add_data(Reference(pf, min_col=7, min_row=2, max_row=pf_last),
                 titles_from_data=True)
    ch2.set_categories(cats_proj)
    ch2.height, ch2.width = 8, 16
    db.add_chart(ch2, "G21")

    cats_month = Reference(mo, min_col=1, min_row=m_first, max_row=m_last)
    ch3 = LineChart()
    ch3.title = "Monthly spend by programme"
    ch3.y_axis.number_format = "$#,##0"
    for j in range(len(PROGRAMMES)):
        ch3.add_data(Reference(mo, min_col=3 + j, min_row=1, max_row=m_last),
                     titles_from_data=True)
    ch3.set_categories(cats_month)
    ch3.height, ch3.width = 8, 16
    db.add_chart(ch3, "G39")

    # ---- Checks ----
    ck = wb.create_sheet("Checks")
    ck.append(["Check", "Result"])
    style_header(ck, 1, 2)
    checks = [
        ("Portfolio spent total = Spend staging total",
         f"=IF(Portfolio!E{pf_tot}-SUM(Spend!D{s_first}:D{s_last})=0,\"PASS\",\"FAIL\")"),
        ("Portfolio task count = Tasks rows",
         f"=IF(Portfolio!H{pf_tot}-COUNTA(Tasks!A{t_first}:A{t_last})=0,\"PASS\",\"FAIL\")"),
        ("Monthly total = Spend staging total",
         f"=IF(Monthly!B{m_total}-SUM(Spend!D{s_first}:D{s_last})=0,\"PASS\",\"FAIL\")"),
        ("Every RAG is Green, Amber or Red",
         f"=IF(COUNTIF(Portfolio!K{pf_first}:K{pf_last},\"Green\")+COUNTIF(Portfolio!K{pf_first}:K{pf_last},\"Amber\")+COUNTIF(Portfolio!K{pf_first}:K{pf_last},\"Red\")={n_proj},\"PASS\",\"FAIL\")"),
        ("All risk scores are 1-25",
         f"=IF(AND(MIN(Risks!F{r_first}:F{r_last})>=1,MAX(Risks!F{r_first}:F{r_last})<=25),\"PASS\",\"FAIL\")"),
        ("No project % complete above 100%",
         f"=IF(MAX(Portfolio!G{pf_first}:G{pf_last})<=1,\"PASS\",\"FAIL\")"),
        ("Dashboard budget KPI ties to Portfolio total",
         f"=IF(Dashboard!B4-Portfolio!D{pf_tot}=0,\"PASS\",\"FAIL\")"),
        ("Dashboard spent KPI ties to Portfolio total",
         f"=IF(Dashboard!B5-Portfolio!E{pf_tot}=0,\"PASS\",\"FAIL\")"),
        ("Every task's project exists in the master list",
         f"=IF(SUMPRODUCT(--ISNA(MATCH(Tasks!B{t_first}:B{t_last},Projects!A{p_first}:A{p_last},0)))=0,\"PASS\",\"FAIL\")"),
        ("Every spend row's project exists in the master list",
         f"=IF(SUMPRODUCT(--ISNA(MATCH(Spend!B{s_first}:B{s_last},Projects!A{p_first}:A{p_last},0)))=0,\"PASS\",\"FAIL\")"),
    ]
    for i, (label, formula) in enumerate(checks):
        r = 2 + i
        ck.cell(row=r, column=1, value=label)
        ck.cell(row=r, column=2, value=formula)
    ck.column_dimensions["A"].width = 58
    ck.column_dimensions["B"].width = 12

    # ---- Guide ----
    gd = wb.create_sheet("Guide", 0)
    gd.append(["ExecPortfolio — how to use this workbook"])
    gd["A1"].font = Font(bold=True, size=14, color=NAVY)
    lines = [
        "1. This is the Excel foundation for an M365 executive portfolio report:",
        "   Projects = the Excel project master; Tasks = a Planner task export;",
        "   Risks = a SharePoint risk list; Spend = monthly actuals. In Power BI,",
        "   load one query per sheet through Power Query; the Portfolio sheet is",
        "   the fact table and the Dashboard KPIs map 1:1 to DAX measures.",
        "2. Paste your own rows into Projects / Tasks / Risks / Spend and extend",
        "   the formula ranges on Portfolio and Monthly to match. The as-of date",
        "   (Portfolio!M1, currently 2026-11-07) drives the Overdue count — change",
        "   it to re-cut the report for any date.",
        "3. RAG rule: Green when spend is within 10% over budget AND no task is",
        "   overdue; Red when spend is over 25% above budget OR 5+ tasks overdue;",
        "   Amber otherwise. Risk score = Likelihood x Impact (1-5 each).",
        "4. Refresh discipline (mirrors the Power Automate flow): refresh the",
        "   source extracts, paste them in, check the Checks sheet — every check",
        "   must read PASS before the pack goes to executives.",
        "5. Demo data is fictional (seeded, deterministic) for Meridian Group.",
        "   Nothing here is financial advice; it is a reporting template.",
    ]
    for i, line in enumerate(lines, 2):
        gd.cell(row=i, column=1, value=line)
    gd.column_dimensions["A"].width = 110
    return wb


def write_csvs(tasks, risks, spend):
    with open(BASE / "data" / "tasks.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["task_id", "project_id", "task", "owner", "status", "due", "pct", "priority"])
        w.writerows(tasks)
    with open(BASE / "data" / "risks.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["risk_id", "project_id", "risk", "likelihood", "impact", "mitigation"])
        w.writerows(risks)
    with open(BASE / "data" / "spend.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "project_id", "programme", "amount"])
        w.writerows(spend)
    with open(BASE / "data" / "projects.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["project_id", "project", "programme", "sponsor", "owner", "phase",
                    "start", "end", "budget"])
        w.writerows(PROJECTS)


def main():
    tasks, risks, spend = build_data()
    write_csvs(tasks, risks, spend)
    wb = build_workbook(tasks, risks, spend)
    wb.save(BASE / "exec-portfolio.xlsx")
    print(f"Wrote {len(PROJECTS)} projects, {len(tasks)} tasks, "
          f"{len(risks)} risks, {len(spend)} spend rows -> "
          f"{BASE / 'exec-portfolio.xlsx'}")


if __name__ == "__main__":
    main()
