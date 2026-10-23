# NSEPulse — NSE Market Dashboard (Excel demo)

**SYNTHETIC SAMPLE DATA** — every figure in this workbook is deterministic
fictional demo data (`generate.py`, seed `20261023`). It is NOT real NSE data.
Built 2026-10-23 as a reference implementation for an NSE open-interest /
volume / price dashboard brief.

## The problem
A trader-style NSE dashboard brief asks for: open interest, % change, volume
and LTP pulled into Excel, plus a sectoral heatmap — refreshed daily, with
every figure traceable to its source snapshot.

## The approach
- `generate.py` — deterministic synthetic snapshots for 60 tickers across
  8 sectors (T-1 previous close, T today): LTP, volume, open interest,
  20-day average volume.
- `build_workbook.py` — builds `nsepulse.xlsx` with openpyxl. **Every derived
  figure is a live Excel formula** (VLOOKUPs into the previous snapshot,
  price-OI signal logic, LARGE/SMALL extraction tables, AVERAGEIF/SUMIF/
  COUNTIFS sector heatmap, native bar charts). Paste fresh snapshots into
  `Raw_Today` / `Raw_PrevDay` and the whole dashboard recomputes.
- Price-OI interpretation (standard): price up + OI up = Long Buildup,
  price down + OI up = Short Buildup, price down + OI down = Long Unwinding,
  price up + OI down = Short Covering.
- `Checks` sheet — 8 formula self-tests (OI walk, breadth adds up, KPI ties,
  sector totals tie, signal categories exhaustive). All read TRUE.
- Production refresh path: point Power Query at the NSE snapshot source and
  refresh the two raw sheets — no rebuild needed.

Sheets: `Guide`, `Dashboard` (6 KPI cards + 2 charts), `Raw_Today`,
`Raw_PrevDay`, `Top_Movers`, `OI_Signals`, `Sector_Heatmap`, `Checks`.

## How to run
```bash
python3 main.py          # generate data + build nsepulse.xlsx
python3 -m pytest tests/ -q
```

## Tests — 11/11 pass
LibreOffice headless recalculation of the real workbook is compared against
an independent Python ground truth computed straight from the CSVs: top
gainer / top loser / strongest long buildup / biggest volume shocker tickers
and values, market breadth, the full top-10 gainers table, sector volume
totals, all 8 Checks TRUE, and a scan proving zero formula errors anywhere.
