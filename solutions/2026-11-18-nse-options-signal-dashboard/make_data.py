#!/usr/bin/env python3
"""Generate the fictional NIFTY option-chain CSV for OptionsSignal.

Deterministic (seed 20261118). Underlying NIFTY, spot 25,480, weekly
expiry 26 Nov 2026, strikes 24,800-26,200 in steps of 50 (29 strikes).

Open interest follows a bell curve around the at-the-money strike, with
the CE peak deliberately at 25,600 (resistance: heavy call writing) and
the PE peak at 25,400 (support: heavy put writing) — the classic
max-pain / support-resistance picture a trader reads off an option
chain. OI changes and price moves are randomised but seeded, and the
seed was chosen so all four OI-build-up signal types occur on both
sides of the chain.

All data is fictional — invented for the demo, not market data.
"""
import csv
import math
import random
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"

SPOT = 25480
EXPIRY = "2026-11-26"
STRIKES = list(range(24800, 26201, 50))
SEED = 20261118


def bell(strike, centre, width, peak):
    return peak * math.exp(-0.5 * ((strike - centre) / width) ** 2)


def option_price(strike, is_call, spot=SPOT):
    """Crude Black-Scholes-free premium: intrinsic + time value that
    decays with distance from the spot. Good enough for a demo chain."""
    intrinsic = max(spot - strike, 0) if is_call else max(strike - spot, 0)
    dist = abs(strike - spot)
    time_value = 120 * math.exp(-dist / 350.0) + 8
    return round(intrinsic + time_value, 2)


def main():
    rng = random.Random(SEED)
    rows = []
    for strike in STRIKES:
        ce_oi = int(bell(strike, 25600, 260, 185000)
                    + rng.uniform(0, 12000))
        pe_oi = int(bell(strike, 25400, 260, 172000)
                    + rng.uniform(0, 12000))
        ce_oi = max(ce_oi, 800)
        pe_oi = max(pe_oi, 800)
        ce_chg = int(rng.uniform(-0.35, 0.45) * ce_oi)
        pe_chg = int(rng.uniform(-0.35, 0.45) * pe_oi)
        ce_ltp = option_price(strike, True)
        pe_ltp = option_price(strike, False)
        # Previous close: a small random move either way, so price-change
        # signs vary and all four signal types occur.
        ce_prev = round(ce_ltp * (1 + rng.uniform(-0.12, 0.12)), 2)
        pe_prev = round(pe_ltp * (1 + rng.uniform(-0.12, 0.12)), 2)
        rows.append({
            "strike": strike,
            "ce_oi": ce_oi,
            "ce_oi_chg": ce_chg,
            "ce_vol": int(rng.uniform(0.4, 1.6) * abs(ce_chg) + 500),
            "ce_ltp": ce_ltp,
            "ce_prev_close": ce_prev,
            "pe_oi": pe_oi,
            "pe_oi_chg": pe_chg,
            "pe_vol": int(rng.uniform(0.4, 1.6) * abs(pe_chg) + 500),
            "pe_ltp": pe_ltp,
            "pe_prev_close": pe_prev,
        })

    with open(DATA / "option_chain.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)

    with open(DATA / "meta.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["key", "value"])
        w.writerow(["spot", SPOT])
        w.writerow(["expiry", EXPIRY])
        w.writerow(["underlying", "NIFTY"])
        w.writerow(["seed", SEED])

    # Sanity: every signal type must occur on both sides.
    def signal(oi_chg, ltp, prev):
        up = ltp > prev
        if oi_chg > 0 and up:
            return "Long Build-up"
        if oi_chg > 0 and not up:
            return "Short Build-up"
        if oi_chg < 0 and up:
            return "Short Covering"
        return "Long Unwinding"

    ce_signals = {signal(r["ce_oi_chg"], r["ce_ltp"],
                         r["ce_prev_close"]) for r in rows}
    pe_signals = {signal(r["pe_oi_chg"], r["pe_ltp"],
                         r["pe_prev_close"]) for r in rows}
    print(f"wrote {len(rows)} strikes, spot {SPOT}, expiry {EXPIRY}")
    print(f"CE signals present: {sorted(ce_signals)}")
    print(f"PE signals present: {sorted(pe_signals)}")
    assert len(ce_signals) == 4, "not all CE signal types occur"
    assert len(pe_signals) == 4, "not all PE signal types occur"


if __name__ == "__main__":
    main()
