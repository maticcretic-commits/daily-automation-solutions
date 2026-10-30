# Excel automation demo — before / after

Business: Demo Traders (FICTIONAL — synthetic demo data). All data is synthetic and fictional. No client data, no named clients.

## Row counts

- before.xlsx Register sheet: 60 data rows (60-row source: 55 unique assets + 5 exact duplicates).
- after.xlsx Register sheet: 55 data rows.
- Duplicate rows removed: 5 (first occurrence kept, exact copies dropped).
- Flagged rows in after.xlsx: 6 (Flag column, column J; flagged rows are highlighted by conditional formatting).

## Fixes applied (counts from this build)

- Dates normalised to real Excel dates (yyyy-mm-dd): 52 rows. Five source formats were recognised: DD/MM/YYYY, YYYY-MM-DD, DD Mon YYYY, Month DD, YYYY, DD.MM.YYYY.
- Dates that could not be parsed: 1 row(s) — the date cell is left blank, the original value is recorded in Notes, and the row is flagged.
- Quantities stored as text converted to numbers: 5 row(s).
- Quantities missing (blank): 3 row(s) — left blank and flagged; the line-total formula returns blank for those rows instead of inventing a quantity.
- Unit prices stored as text (mixed symbols: ₹, Rs., INR, bare number) converted to numbers: 42 row(s). The original text is recorded in Notes.
- Prices with an unclear currency symbol ('$'): 1 row(s) — numeric figure kept as supplied, row flagged for the owner to verify the currency. No exchange rate was invented.
- Invalid status values: 1 row(s) — kept exactly as supplied and flagged, never silently corrected.
- Nothing else was changed: item names, categories and locations are carried through untouched.

## Exact formulas used (after.xlsx)

- Line Total, Register column G (row 4 shown; the row number changes per row):
  `=IF(OR(E4="",F4=""),"",E4*F4)`
- Summary, Asset Count per category (row 5 shown):
  `=COUNTIF(Register!$C$4:$C$58,A5)`
- Summary, Total Value per category (row 5 shown):
  `=SUMIF(Register!$C$4:$C$58,A5,Register!$G$4:$G$58)`
- Summary, Grand Total row: `=SUM(B5:B8)` and `=SUM(C5:C8)` (category rows).
- Summary, Total assets (rows): `=COUNTA(Register!$A$4:$A$58)`.
- Summary, Flagged rows: `=COUNTIF(Register!$J$4:$J$58,"?*")`.
- No total in after.xlsx is a hard-coded value. The only static numbers on the Summary sheet are the three labelled build notes (before/after row counts and duplicates removed), which are build-time facts, not computed values.

## Expected totals (computed independently in Python, verified against LibreOffice recalc — see VERIFY.txt)

- IT Equipment: 60,268.00
- Furniture: 253,936.00
- Vehicle: 8,705,928.00
- Tools: 596,170.00
- Grand total: 9,616,302.00

## How to verify each claim by opening the files

1. Open before.xlsx: note duplicate Asset IDs, dates stored as text in several formats, quantities stored as text, prices with ₹ / Rs. / INR / $ symbols, blank quantities, and one status value 'In Usee'.
2. Open after.xlsx, Register sheet: click any Purchase Date cell — the formula bar shows a real date. Click any Quantity or Unit Price cell — values are numbers. Click any Line Total cell — the formula bar shows the IF/OR formula above, not a typed number.
3. Open after.xlsx, Summary sheet: click any count or total — the formula bar shows COUNTIF / SUMIF / SUM. Change a quantity on the Register sheet and watch the Summary totals update.
4. Filter the Flag column (column J) to non-blank: the flagged rows are exactly the rows counted above; flagged rows are shaded pink.
5. Re-run the build script: `python3 build_excel_demo.py`, then `python3 test_excel_demo.py` — both files regenerate with identical cell values and formulas (xlsx zip timestamps differ between runs, as normal) and all tests pass (see VERIFY.txt).
6. Headless recalc check: `soffice --headless --convert-to xlsx --outdir /tmp/recalc after.xlsx`, then open the converted file and confirm the Summary totals match the expected totals above.

## Rules applied (for a client with their own file)

- Duplicates: exact duplicate Asset IDs — keep the first occurrence.
- Dates: five recognised formats; anything else is flagged, never guessed.
- Numbers: text numbers and currency symbols are stripped to plain numbers; the original text is kept in Notes; a '$' symbol is flagged rather than converted.
- Missing quantities: left blank and flagged; no quantity is invented.
- Invalid status values: kept as supplied and flagged; no value is silently corrected.
- A client supplying their own file would need to confirm: the ID column, the recognised date formats, which currency the figures are in, and the allowed status list.
