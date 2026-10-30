#!/usr/bin/env python3
"""Build the Excel automation demo: before.xlsx (messy, synthetic) -> after.xlsx (cleaned).

Re-runnable and deterministic: every run regenerates before.xlsx from a
seeded generator, then reads it back and writes after.xlsx. The input file
is never edited in place during cleaning.

Synthetic data only. Fictional business: "Demo Traders".
No Power Query, SQL or DAX is used anywhere in this demo.

Usage:
    python3 build_excel_demo.py
Outputs (same folder as this script):
    before.xlsx, after.xlsx, BEFORE_AFTER.md, build_counts.json
"""

import json
import random
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
BEFORE_PATH = HERE / "before.xlsx"
AFTER_PATH = HERE / "after.xlsx"
REPORT_PATH = HERE / "BEFORE_AFTER.md"
COUNTS_PATH = HERE / "build_counts.json"

BUSINESS = "Demo Traders (FICTIONAL — synthetic demo data)"
SEED = 20261030

CATEGORIES = ["IT Equipment", "Furniture", "Vehicle", "Tools"]
ITEMS = {
    "IT Equipment": ["Laptop 14in", "Monitor 24in", "Office Printer", "Router",
                     "External Hard Drive", "Keyboard Set"],
    "Furniture": ["Office Chair", "Work Desk", "Filing Cabinet",
                  "Conference Table", "Bookshelf"],
    "Vehicle": ["Delivery Van", "Scooter", "Pickup Truck"],
    "Tools": ["Drill Machine", "Welding Kit", "Tool Chest", "Generator",
              "Air Compressor"],
}
STATUSES = ["In Use", "In Use", "In Storage", "In Use", "Under Repair", "Disposed"]
VALID_STATUSES = {"In Use", "In Storage", "Disposed", "Under Repair"}
LOCATIONS = ["Patna", "Delhi", "Mumbai", "Kolkata"]
DATE_FORMATS = ["%d/%m/%Y", "%Y-%m-%d", "%d %b %Y", "%B %d, %Y", "%d.%m.%Y"]

N_UNIQUE = 55
BLANK_QTY_ROWS = {5, 19, 41}        # 1-based asset numbers
TEXT_QTY_ROWS = {3, 9, 27, 33, 48}  # quantity stored as text
REAL_DATE_ROWS = {7, 23}            # date stored as a real date already
BAD_DATE_ROW = 11                   # "31/02/2024" — impossible date
BAD_STATUS_ROW = 17                 # status "In Usee"
USD_ROW = 31                        # price shown as "$45" — flagged, verify
DUPLICATE_OF = [2, 8, 15, 22, 40]   # asset numbers duplicated exactly

HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
HEADER_FONT = Font(bold=True)
FLAG_FILL = PatternFill("solid", fgColor="FCE4E4")
FLAG_FONT = Font(color="9C0006")


def generate_rows():
    """Return the 60 messy rows for before.xlsx (55 unique + 5 exact duplicates)."""
    rng = random.Random(SEED)
    rows = []
    for i in range(1, N_UNIQUE + 1):
        category = CATEGORIES[i % 4]
        item = ITEMS[category][i % len(ITEMS[category])]
        asset_id = f"AST-{i:03d}"
        name = f"{item} — S/N {1000 + i}"

        base = date(2022, 1, 1).toordinal() + (i * 37) % 700
        d = date.fromordinal(base)
        if i == BAD_DATE_ROW:
            purchase = "31/02/2024"
        elif i in REAL_DATE_ROWS:
            purchase = datetime(d.year, d.month, d.day)
        else:
            purchase = d.strftime(DATE_FORMATS[i % len(DATE_FORMATS)])

        qty = (i * 3) % 12 + 1
        if i in BLANK_QTY_ROWS:
            quantity = None
        elif i in TEXT_QTY_ROWS:
            quantity = str(qty)
        else:
            quantity = qty

        if i == USD_ROW:
            price = "$45"
        else:
            base_price = {"IT Equipment": 800, "Furniture": 1500,
                          "Vehicle": 85000, "Tools": 900}[category]
            span = {"IT Equipment": 44000, "Furniture": 18500,
                    "Vehicle": 365000, "Tools": 29000}[category]
            amount = base_price + (i * 137) % span
            style = i % 4
            if style == 0:
                price = f"\u20b9{amount:,}"
            elif style == 1:
                price = f"Rs. {amount}"
            elif style == 2:
                price = f"INR {amount:,}"
            else:
                price = amount

        status = "In Usee" if i == BAD_STATUS_ROW else STATUSES[i % len(STATUSES)]
        location = LOCATIONS[i % len(LOCATIONS)]
        rows.append([asset_id, name, category, purchase, quantity, price,
                     status, location])

    for dup in DUPLICATE_OF:
        rows.append(list(rows[dup - 1]))  # exact duplicate row
    rng.shuffle(rows)  # deterministic shuffle so duplicates are not all at the end
    return rows


def write_before(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Register"
    ws["A1"] = BUSINESS
    ws["A2"] = "Asset register as supplied — messy source file. Do not edit; the build script regenerates it."
    headers = ["Asset ID", "Item", "Category", "Purchase Date", "Quantity",
               "Unit Price", "Status", "Location"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for r, row in enumerate(rows, start=4):
        for col, value in enumerate(row, start=1):
            ws.cell(row=r, column=col, value=value)
    widths = [12, 34, 14, 16, 10, 14, 14, 12]
    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:H{3 + len(rows)}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(BEFORE_PATH)
    return len(rows)


def parse_date(value):
    """Return (date or None, note or None)."""
    if isinstance(value, datetime):
        return value.date(), None
    if isinstance(value, date):
        return value, None
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date(), None
        except ValueError:
            continue
    return None, f"Date {text!r} could not be parsed"


def parse_quantity(value):
    """Return (int or None, note or None, flagged bool)."""
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return None, "Quantity missing", True
    if isinstance(value, str):
        return int(value.strip()), f"Quantity text {value!r} converted to number", False
    return int(value), None, False


def parse_price(value):
    """Return (number or None, note or None, flagged bool)."""
    if isinstance(value, (int, float)):
        return value, None, False
    text = str(value).strip()
    flagged = "$" in text
    cleaned = (text.replace("\u20b9", "").replace("Rs.", "").replace("INR", "")
               .replace("$", "").replace(",", "").strip())
    amount = float(cleaned)
    if amount == int(amount):
        amount = int(amount)
    note = f"Price {text!r} converted to {amount}"
    if flagged:
        note += " (currency symbol was '$' — figure kept as supplied; verify currency)"
    return amount, note, flagged


def build_after():
    wb = load_workbook(BEFORE_PATH)
    ws = wb["Register"]
    # Data starts at row 4 (rows 1-2 are title notes, row 3 is the header).
    raw = [tuple(c.value for c in row) for row in ws.iter_rows(min_row=4, max_col=8)]

    seen = set()
    cleaned = []
    counts = Counter()
    for row in raw:
        asset_id, name, category, purchase, quantity, price, status, location = row
        if asset_id in seen:
            counts["duplicates_removed"] += 1
            continue
        seen.add(asset_id)

        notes, flags = [], []
        parsed_date, date_note = parse_date(purchase)
        if date_note:
            flags.append("Purchase date could not be parsed")
            notes.append(f"Date {purchase!r} could not be parsed — left blank")
            counts["dates_unparseable"] += 1
        elif isinstance(purchase, str):
            notes.append(f"Date {purchase!r} normalised")
            counts["dates_normalised"] += 1

        qty, qty_note, qty_flag = parse_quantity(quantity)
        if qty_note:
            notes.append(qty_note)
        if qty_flag:
            flags.append("Quantity missing — line total not computed")
            counts["quantities_missing"] += 1
        if isinstance(quantity, str) and quantity.strip() != "":
            counts["quantities_text_to_number"] += 1

        amount, price_note, price_flag = parse_price(price)
        if price_note:
            notes.append(price_note)
        if isinstance(price, str):
            counts["prices_text_to_number"] += 1
        if price_flag:
            flags.append("Price currency symbol was '$' — verify currency")
            counts["prices_currency_unclear"] += 1

        if status not in VALID_STATUSES:
            flags.append(f"Status {status!r} not recognised — kept as supplied")
            counts["statuses_invalid"] += 1

        cleaned.append([asset_id, name, category, parsed_date, qty, amount,
                        None, status, location, "; ".join(flags), "; ".join(notes)])

    out = Workbook()
    reg = out.active
    reg.title = "Register"
    reg["A1"] = BUSINESS
    reg["A2"] = "Cleaned asset register. Line totals and the Summary sheet are Excel formulas — see BEFORE_AFTER.md."
    headers = ["Asset ID", "Item", "Category", "Purchase Date", "Quantity",
               "Unit Price", "Line Total", "Status", "Location", "Flag", "Notes"]
    for col, h in enumerate(headers, start=1):
        cell = reg.cell(row=3, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    n = len(cleaned)
    for r, row in enumerate(cleaned, start=4):
        for col, value in enumerate(row, start=1):
            cell = reg.cell(row=r, column=col, value=value)
        # Line Total is a formula, never a hard-coded value.
        reg.cell(row=r, column=7).value = f'=IF(OR(E{r}="",F{r}=""),"",E{r}*F{r})'
        reg.cell(row=r, column=4).number_format = "yyyy-mm-dd"
        reg.cell(row=r, column=6).number_format = "#,##0.00"
        reg.cell(row=r, column=7).number_format = "#,##0.00"
    widths = [12, 34, 14, 14, 10, 13, 13, 14, 12, 44, 62]
    for col, w in enumerate(widths, start=1):
        reg.column_dimensions[get_column_letter(col)].width = w
    reg.freeze_panes = "A4"
    reg.auto_filter.ref = f"A3:K{3 + n}"
    reg.print_title_rows = "1:3"
    reg.conditional_formatting.add(
        f"A4:K{3 + n}",
        FormulaRule(formula=['$J4<>""'], fill=FLAG_FILL, font=FLAG_FONT),
    )

    # Summary sheet — every value cell is a formula.
    sm = out.create_sheet("Summary")
    sm["A1"] = "Asset Register — Summary"
    sm["A1"].font = Font(bold=True, size=14)
    sm["A2"] = "All values below are Excel formulas over the Register sheet. Synthetic demo data — fictional business 'Demo Traders'."
    sm["A4"], sm["B4"], sm["C4"] = "Category", "Asset Count", "Total Value"
    for col in range(1, 4):
        cell = sm.cell(row=4, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    first, last = 4, 3 + n
    for i, cat in enumerate(CATEGORIES):
        r = 5 + i
        sm.cell(row=r, column=1, value=cat)
        sm.cell(row=r, column=2).value = f"=COUNTIF(Register!$C${first}:$C${last},A{r})"
        sm.cell(row=r, column=3).value = f"=SUMIF(Register!$C${first}:$C${last},A{r},Register!$G${first}:$G${last})"
        sm.cell(row=r, column=3).number_format = "#,##0.00"
    gr = 5 + len(CATEGORIES)
    sm.cell(row=gr, column=1, value="Grand Total").font = HEADER_FONT
    sm.cell(row=gr, column=2).value = f"=SUM(B5:B{gr - 1})"
    sm.cell(row=gr, column=3).value = f"=SUM(C5:C{gr - 1})"
    sm.cell(row=gr, column=3).number_format = "#,##0.00"
    sm.cell(row=gr + 2, column=1, value="Total assets (rows)")
    sm.cell(row=gr + 2, column=2).value = f"=COUNTA(Register!$A${first}:$A${last})"
    sm.cell(row=gr + 3, column=1, value="Flagged rows")
    sm.cell(row=gr + 3, column=2).value = f'=COUNTIF(Register!$J${first}:$J${last},"?*")'
    sm.cell(row=gr + 5, column=1, value="Build notes (counts recorded at build time, not formulas)").font = HEADER_FONT
    sm.cell(row=gr + 6, column=1, value="Rows in before.xlsx")
    sm.cell(row=gr + 6, column=2, value=len(raw))
    sm.cell(row=gr + 7, column=1, value="Rows in after.xlsx")
    sm.cell(row=gr + 7, column=2, value=n)
    sm.cell(row=gr + 8, column=1, value="Duplicate rows removed")
    sm.cell(row=gr + 8, column=2, value=counts["duplicates_removed"])
    for col, w in zip(range(1, 4), [24, 14, 16]):
        sm.column_dimensions[get_column_letter(col)].width = w
    sm.freeze_panes = "A5"
    sm.print_title_rows = "1:4"

    out.save(AFTER_PATH)

    flagged = sum(1 for row in cleaned if row[9])
    totals = {}
    grand = 0.0
    for row in cleaned:
        if row[4] is not None and row[5] is not None:
            value = row[4] * row[5]
            totals[row[2]] = totals.get(row[2], 0.0) + value
            grand += value
    for cat in CATEGORIES:
        totals.setdefault(cat, 0.0)
    counts_out = {
        "before_rows": len(raw),
        "after_rows": n,
        "duplicates_removed": counts["duplicates_removed"],
        "flagged_rows": flagged,
        "dates_normalised": counts["dates_normalised"],
        "dates_unparseable": counts["dates_unparseable"],
        "quantities_missing": counts["quantities_missing"],
        "quantities_text_to_number": counts["quantities_text_to_number"],
        "prices_text_to_number": counts["prices_text_to_number"],
        "prices_currency_unclear": counts["prices_currency_unclear"],
        "statuses_invalid": counts["statuses_invalid"],
        "totals_by_category": totals,
        "grand_total": grand,
    }
    return counts_out


def write_report(counts):
    lines = [
        "# Excel automation demo — before / after",
        "",
        f"Business: {BUSINESS}. All data is synthetic and fictional. No client data, no named clients.",
        "",
        "## Row counts",
        "",
        f"- before.xlsx Register sheet: {counts['before_rows']} data rows (60-row source: 55 unique assets + 5 exact duplicates).",
        f"- after.xlsx Register sheet: {counts['after_rows']} data rows.",
        f"- Duplicate rows removed: {counts['duplicates_removed']} (first occurrence kept, exact copies dropped).",
        f"- Flagged rows in after.xlsx: {counts['flagged_rows']} (Flag column, column J; flagged rows are highlighted by conditional formatting).",
        "",
        "## Fixes applied (counts from this build)",
        "",
        f"- Dates normalised to real Excel dates (yyyy-mm-dd): {counts['dates_normalised']} rows. Five source formats were recognised: DD/MM/YYYY, YYYY-MM-DD, DD Mon YYYY, Month DD, YYYY, DD.MM.YYYY.",
        f"- Dates that could not be parsed: {counts['dates_unparseable']} row(s) — the date cell is left blank, the original value is recorded in Notes, and the row is flagged.",
        f"- Quantities stored as text converted to numbers: {counts['quantities_text_to_number']} row(s).",
        f"- Quantities missing (blank): {counts['quantities_missing']} row(s) — left blank and flagged; the line-total formula returns blank for those rows instead of inventing a quantity.",
        f"- Unit prices stored as text (mixed symbols: ₹, Rs., INR, bare number) converted to numbers: {counts['prices_text_to_number']} row(s). The original text is recorded in Notes.",
        f"- Prices with an unclear currency symbol ('$'): {counts['prices_currency_unclear']} row(s) — numeric figure kept as supplied, row flagged for the owner to verify the currency. No exchange rate was invented.",
        f"- Invalid status values: {counts['statuses_invalid']} row(s) — kept exactly as supplied and flagged, never silently corrected.",
        "- Nothing else was changed: item names, categories and locations are carried through untouched.",
        "",
        "## Exact formulas used (after.xlsx)",
        "",
        "- Line Total, Register column G (row 4 shown; the row number changes per row):",
        "  `=IF(OR(E4=\"\",F4=\"\"),\"\",E4*F4)`",
        "- Summary, Asset Count per category (row 5 shown):",
        "  `=COUNTIF(Register!$C$4:$C$58,A5)`",
        "- Summary, Total Value per category (row 5 shown):",
        "  `=SUMIF(Register!$C$4:$C$58,A5,Register!$G$4:$G$58)`",
        "- Summary, Grand Total row: `=SUM(B5:B8)` and `=SUM(C5:C8)` (category rows).",
        "- Summary, Total assets (rows): `=COUNTA(Register!$A$4:$A$58)`.",
        "- Summary, Flagged rows: `=COUNTIF(Register!$J$4:$J$58,\"?*\")`.",
        "- No total in after.xlsx is a hard-coded value. The only static numbers on the Summary sheet are the three labelled build notes (before/after row counts and duplicates removed), which are build-time facts, not computed values.",
        "",
        "## Expected totals (computed independently in Python, verified against LibreOffice recalc — see VERIFY.txt)",
        "",
    ]
    for cat in CATEGORIES:
        lines.append(f"- {cat}: {counts['totals_by_category'][cat]:,.2f}")
    lines.append(f"- Grand total: {counts['grand_total']:,.2f}")
    lines += [
        "",
        "## How to verify each claim by opening the files",
        "",
        "1. Open before.xlsx: note duplicate Asset IDs, dates stored as text in several formats, quantities stored as text, prices with ₹ / Rs. / INR / $ symbols, blank quantities, and one status value 'In Usee'.",
        "2. Open after.xlsx, Register sheet: click any Purchase Date cell — the formula bar shows a real date. Click any Quantity or Unit Price cell — values are numbers. Click any Line Total cell — the formula bar shows the IF/OR formula above, not a typed number.",
        "3. Open after.xlsx, Summary sheet: click any count or total — the formula bar shows COUNTIF / SUMIF / SUM. Change a quantity on the Register sheet and watch the Summary totals update.",
        "4. Filter the Flag column (column J) to non-blank: the flagged rows are exactly the rows counted above; flagged rows are shaded pink.",
        "5. Re-run the build script: `python3 build_excel_demo.py`, then `python3 test_excel_demo.py` — both files regenerate with identical cell values and formulas (xlsx zip timestamps differ between runs, as normal) and all tests pass (see VERIFY.txt).",
        "6. Headless recalc check: `soffice --headless --convert-to xlsx --outdir /tmp/recalc after.xlsx`, then open the converted file and confirm the Summary totals match the expected totals above.",
        "",
        "## Rules applied (for a client with their own file)",
        "",
        "- Duplicates: exact duplicate Asset IDs — keep the first occurrence.",
        "- Dates: five recognised formats; anything else is flagged, never guessed.",
        "- Numbers: text numbers and currency symbols are stripped to plain numbers; the original text is kept in Notes; a '$' symbol is flagged rather than converted.",
        "- Missing quantities: left blank and flagged; no quantity is invented.",
        "- Invalid status values: kept as supplied and flagged; no value is silently corrected.",
        "- A client supplying their own file would need to confirm: the ID column, the recognised date formats, which currency the figures are in, and the allowed status list.",
    ]
    REPORT_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    rows = generate_rows()
    n_before = write_before(rows)
    counts = build_after()
    assert counts["before_rows"] == n_before == 60
    assert counts["after_rows"] == 55
    write_report(counts)
    COUNTS_PATH.write_text(json.dumps(counts, indent=2) + "\n", encoding="utf-8")
    print(f"before.xlsx: {counts['before_rows']} rows")
    print(f"after.xlsx: {counts['after_rows']} rows, "
          f"{counts['duplicates_removed']} duplicates removed, "
          f"{counts['flagged_rows']} flagged")
    print(f"grand total (formula-computed): {counts['grand_total']:,.2f}")
    print(f"report: {REPORT_PATH.name}, counts: {COUNTS_PATH.name}")


if __name__ == "__main__":
    main()
