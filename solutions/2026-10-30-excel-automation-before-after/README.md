# Excel automation demo — messy asset register, cleaned

A real, runnable Excel clean-up: a Python script (openpyxl) reads a messy
source workbook and writes a cleaned workbook where every total is a live
Excel formula. Nothing is faked, and nothing is hard-coded where a formula
belongs.

**Synthetic data only.** The business is "Demo Traders", clearly marked
fictional in both workbooks. No client data, no named clients, no delivered
client project is claimed. This demo shows what the same treatment would do
to a client's own file.

## What it proves

- Reading a messy `.xlsx` exactly as supplied and writing a cleaned copy
  without touching the source file.
- Normalising dates stored as text in five different formats into real
  Excel dates, without guessing the ones that cannot be parsed.
- Converting text numbers and mixed currency symbols (₹, Rs., INR, bare
  numbers) into real numbers, while keeping the original text in a Notes
  column for audit.
- Removing exact duplicate rows by Asset ID (first occurrence kept).
- Flagging problems instead of inventing data: blank quantities, an
  unparseable date, an unrecognised status value, and a `$` price are all
  flagged and highlighted — never silently corrected.
- Writing Excel formulas for line totals and a summary sheet (COUNTIF /
  SUMIF / SUM), plus filters, frozen panes, print titles, and conditional
  formatting on flagged rows.

## Files

| File | What it is |
|---|---|
| `build_excel_demo.py` | The build script. Generates `before.xlsx` (seeded, deterministic), then reads it and writes `after.xlsx`, `BEFORE_AFTER.md`, and `build_counts.json`. Re-running it regenerates identical cell values and formulas. |
| `before.xlsx` | Messy source: 60 rows (55 unique assets + 5 exact duplicates), text dates, text quantities, mixed currency symbols, 3 blank quantities, 1 unparseable date, 1 invalid status. |
| `after.xlsx` | Cleaned: 55 rows, `Register` sheet (formulas in column G) + `Summary` sheet (all values are formulas). 6 rows flagged. |
| `BEFORE_AFTER.md` | Row counts, every fix with counts, the exact formulas used, and how to verify each claim by opening the files. |
| `build_counts.json` | Machine-readable counts from the build (used by the tests). |
| `test_excel_demo.py` | 33 checks, including a LibreOffice headless recalc check of the formula totals. |
| `VERIFY.txt` | Output of the last test run — all checks passed. |

## How to run

```bash
python3 build_excel_demo.py   # regenerates before.xlsx, after.xlsx, BEFORE_AFTER.md
python3 test_excel_demo.py    # runs all checks (needs LibreOffice for the recalc check)
```

Requires Python 3.10+ and openpyxl (`pip install openpyxl`). No other
dependencies. LibreOffice (`soffice`) is only needed for the recalc check;
without it, the test script says so plainly and verifies the formulas by
inspection instead.

## Before / after — 5 sample rows (real values from the generated files)

"Total" in the After column is the value the Line Total formula computes
(the cell itself holds `=IF(OR(E<n>="",F<n>=""),"",E<n>*F<n>)`).

| Asset | Before (as supplied) | After (cleaned) |
|---|---|---|
| AST-003 Generator | Date "April 22, 2022" (text); Qty "10" (text); Price 1311 | Date 2022-04-22; Qty 10; Price 1311; Total 13,110 (formula) |
| AST-005 Office Chair | Date "05/07/2022" (text); Qty blank; Price "Rs. 2185" (text) | Date 2022-07-05; Qty blank — flagged "Quantity missing"; Price 2185; Total blank (formula, nothing invented) |
| AST-011 Welding Kit | Date "31/02/2024" (impossible date) | Date blank — flagged "Purchase date could not be parsed"; original value kept in Notes; Qty 10; Price 2407; Total 24,070 (formula) |
| AST-017 Filing Cabinet | Status "In Usee"; Price "Rs. 3829" (text) | Status kept as "In Usee" — flagged, never silently corrected; Price 3829; Total 15,316 (formula) |
| AST-031 Welding Kit | Price "$45" | Price 45 — flagged "verify currency" (figure kept as supplied, no exchange rate invented); Total 450 (formula) |

Duplicates removed: AST-002, AST-008, AST-015, AST-022, AST-040 each
appeared twice in `before.xlsx` and appear once in `after.xlsx`.

Totals (formula-computed, recalc-verified): IT Equipment 60,268.00;
Furniture 253,936.00; Vehicle 8,705,928.00; Tools 596,170.00;
grand total 9,616,302.00. See `BEFORE_AFTER.md` for the full verification
steps.

## What a client would need to supply

- Their own workbook (or CSV), with the columns they actually use.
- The rules: which column is the ID, which date formats appear, what
  currency the figures are in, and the allowed list of status values.
- Confirmation of the judgement calls: keep-first for duplicates, flag
  (don't fix) invalid values, and what a blank quantity means in their
  process.

The script would then be adapted to their columns and rules — the demo
script is written to make those rules explicit and easy to change, not
buried in the code.

## Limitations — stated plainly

- **Synthetic data.** Every row is invented for demonstration. Results on
  a real file depend on that file's actual mess.
- **No Power Query.** The cleaning is done in Python/openpyxl and delivered
  as formulas in the workbook. Power Query is not evidenced in Nitesh's
  CV and is not claimed here.
- **No SQL or DAX.** None is used or implied. The Summary sheet uses only
  COUNTIF / SUMIF / SUM.
- **Date formats.** Only the five formats listed in `BEFORE_AFTER.md` are
  recognised. Ambiguous formats (for example, a date that could be either
  DD/MM or MM/DD) are not guessed — an unparseable date is flagged.
- **Currency.** Figures are treated as plain numbers in one currency. A
  `$` symbol is flagged for the owner to verify; no exchange rate is
  applied or invented.
- **Excel version behaviour may vary.** Formulas use only IF / OR /
  COUNTIF / SUMIF / SUM / COUNTA, which are long-standard functions, and
  the totals were verified by a LibreOffice headless recalculation — but
  rendering (for example, conditional-formatting colours) can differ
  slightly between Excel and LibreOffice.
- **Duplicates are exact Asset ID matches.** Near-duplicates (same item,
  different ID, or an ID with a typo) are not detected; that would need a
  matching rule agreed with the client.
