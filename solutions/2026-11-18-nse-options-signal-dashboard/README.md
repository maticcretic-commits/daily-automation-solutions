# OptionsSignal — NSE Options Pattern Signal Dashboard

A formula-driven options signal dashboard for a fictional NIFTY weekly
expiry (26 Nov 2026, spot 25,480). Built as a working demo for the
Freelancer project **"NSE Options Pattern Signal Dashboard"**
(finance-adjacent dashboard wedge — an options-chain dashboard with
OI build-up signals, put-call ratio, max pain and support/resistance).

## What it does

- **OptionChain** — 29 strikes (24,800–26,200) with CE/PE open interest,
  OI change, volume, LTP and previous close. Live derived columns:
  total OI, PCR per strike, % price change, CE/PE OI-build-up signals,
  and a max-pain payout column (SUMPRODUCT over the whole chain, per
  candidate expiry strike).
- **Signals** — per-strike CE signal, PE signal, combined read
  (Strong Bullish / Bullish / Neutral / Bearish / Strong Bearish),
  payout, and the max-pain strike flagged.
- **Dashboard** — spot / expiry / ATM, KPIs (total CE/PE OI, PCR 0.934,
  max pain 25,500, support 25,450, resistance 25,550, net OI change,
  bullish 9 / bearish 12 / neutral 8 strikes), a signal-count table,
  and 3 native charts (OI by strike, OI change by strike, PCR by strike).
- **Checks** — 10 integrity checks against the independent Python ground
  truth; all read PASS.

## Signal logic (standard OI interpretation)

| Price | OI | Signal | CE read | PE read |
|---|---|---|---|---|
| up | up | Long Build-up | bullish | bearish |
| down | up | Short Build-up | bearish | bullish |
| up | down | Short Covering | bullish | bearish |
| down | down | Long Unwinding | bearish | bullish |

Combined read: both sides agreeing = Strong Bullish/Bearish; one side
agreeing = Bullish/Bearish; neither = Neutral.

## Key levels (all live formulas)

- **Max pain** — the strike where total option payout at expiry is
  smallest (25,500 in this demo chain).
- **Support** — the strike with the heaviest PE open interest
  (25,450 — put writers defend it).
- **Resistance** — the strike with the heaviest CE open interest
  (25,550 — call writers defend it).
- **PCR** — total PE OI / total CE OI (0.934 here).

## How to use

1. Paste a fresh option chain into OptionChain columns A–K (Strike,
   CE OI, CE OI Chg, CE Vol, CE LTP, CE PrevClose, PE OI, PE OI Chg,
   PE Vol, PE LTP, PE PrevClose). Columns L–S, Signals and Dashboard
   are all live formulas.
2. Update the spot on the Dashboard (B4); the ATM strike follows.
3. Read the Checks sheet after any change — all 10 checks must PASS.

## Files

- `generate.py` — builds the workbook from `data/option_chain.csv`.
- `make_data.py` — generates the fictional chain (seed 20261118).
- `data/option_chain.csv` — the 29-strike fictional chain.
- `data/meta.csv` — spot, expiry, underlying, seed.
- `options-signal-dashboard.xlsx` — the built workbook.
- `tests/test_build.py` — rebuilds, then verifies via LibreOffice
  headless recalc against an independent Python ground truth
  (per-strike derived columns, combined reads, all KPIs, signal
  counts, zero formula errors, all 10 Checks PASS).

## Run the tests

```sh
python3 tests/test_build.py
```

## Data

All fictional. The chain is generated from a fixed seed for the demo;
it is not real market data and not trading advice.
