#!/usr/bin/env python3
"""
forecast_drivers.py  —  OPTIONAL add-on for the Equity Valuation Cockpit.

Uses Google's TimesFM (zero-shot time-series foundation model) to forecast a
company's REVENUE — a *driver* of intrinsic value — not its stock price.

It turns hand-picked DCF growth assumptions into data-driven ones:

    quarterly revenue history  ->  TimesFM (20q / 5yr ahead, P10/P50/P90)
        -> annual revenue path -> YoY growth  g1..g5   (ASSUMPTIONS inputs)
        -> P10/P90 spread       -> growth sigma         (Monte-Carlo input)

IMPORTANT (read the project README): this forecasts the *driver*, not the price.
Treat the output as ONE input / a sanity-check on your own growth view — a
foundation model cannot see product cycles, recessions, or competition.

Usage:
    python3 forecast_drivers.py                 # runs the embedded Apple demo
    python3 forecast_drivers.py --csv rev.csv   # your data (one column of quarterly revenue)
    python3 forecast_drivers.py --selftest      # validate the math WITHOUT loading the model

Install the optional deps first (see forecast/requirements.txt):
    pip install "timesfm[torch]" numpy
And always run the preflight from the timesfm skill before first model load.
"""
import argparse, json, math, sys
import numpy as np

# 80% prediction interval -> sigma:  z for the 90th percentile of a normal
Z80 = 1.2815515594  # P90 = mean + Z80*sigma  (so sigma = (P90-P10)/(2*Z80))

# Apple fiscal-quarter revenue, FY2015 Q1 .. FY2025 Q4 ($ billions, approximate).
# Annual sums reconcile to the reported totals (FY21 365.8, FY23 383.3, FY25 416.2).
APPLE_QUARTERLY = [
    74.6, 58.0, 49.6, 51.5,   # FY2015
    75.9, 50.6, 42.4, 46.9,   # FY2016
    78.4, 52.9, 45.4, 52.6,   # FY2017
    88.3, 61.1, 53.3, 62.9,   # FY2018
    84.3, 58.0, 53.8, 64.0,   # FY2019
    91.8, 58.3, 59.7, 64.7,   # FY2020
    111.4, 89.6, 81.4, 83.4,  # FY2021
    123.9, 97.3, 83.0, 90.1,  # FY2022
    117.2, 94.8, 81.8, 89.5,  # FY2023
    119.6, 90.8, 85.8, 94.9,  # FY2024
    124.3, 95.4, 94.0, 102.5, # FY2025
]
LAST_FISCAL_YEAR = 2025


# ----------------------------------------------------------------------------
# TimesFM forecast (loads the ~800 MB model — run the preflight first!)
# ----------------------------------------------------------------------------
def timesfm_forecast(values, horizon=20):
    """Return (point[H], q10[H], q90[H]) for the next `horizon` quarters."""
    import torch, timesfm
    torch.set_float32_matmul_precision("high")
    model = timesfm.TimesFM_2p5_200M_torch.from_pretrained(
        "google/timesfm-2.5-200m-pytorch"
    )
    model.compile(timesfm.ForecastConfig(
        max_context=max(64, len(values)),
        max_horizon=horizon,
        normalize_inputs=True,
        use_continuous_quantile_head=True,
        force_flip_invariance=True,
        infer_is_positive=True,      # revenue is strictly positive
        fix_quantile_crossing=True,
    ))
    point, quantiles = model.forecast(
        horizon=horizon, inputs=[np.asarray(values, dtype=np.float32)]
    )
    # quantiles[..., k]: 0=mean, 1=q10, 5=q50(=point), 9=q90
    return point[0], quantiles[0, :, 1], quantiles[0, :, 9]


# ----------------------------------------------------------------------------
# Convert a quarterly forecast -> annual DCF drivers (pure Python, no model)
# ----------------------------------------------------------------------------
def quarterly_to_drivers(history_q, point_q, q10_q, q90_q, n_years=5):
    """
    history_q : list of past quarterly revenue (length multiple of 4 preferred)
    point_q/q10_q/q90_q : forecast quarters (length >= 4*n_years)
    Returns a list of per-year dicts with revenue path, YoY growth, and growth sigma.
    Note: annual P10/P90 sum the quarterly quantiles (a deliberately CONSERVATIVE
    / wider band — quantiles are not strictly additive), which is the safe choice
    for a Monte-Carlo sigma (more humility, not less).
    """
    prev = float(sum(history_q[-4:]))           # last actual fiscal-year revenue
    out = []
    for i in range(n_years):
        sl = slice(4 * i, 4 * i + 4)
        rev_p50 = float(np.sum(point_q[sl]))
        rev_p10 = float(np.sum(q10_q[sl]))
        rev_p90 = float(np.sum(q90_q[sl]))
        g50 = rev_p50 / prev - 1.0
        g10 = rev_p10 / prev - 1.0
        g90 = rev_p90 / prev - 1.0
        sigma = max(0.0, (g90 - g10) / (2.0 * Z80))
        out.append({
            "year": LAST_FISCAL_YEAR + i + 1,
            "rev_p50": round(rev_p50, 1),
            "rev_p10": round(rev_p10, 1),
            "rev_p90": round(rev_p90, 1),
            "growth": round(g50, 4),
            "growth_sigma": round(sigma, 4),
        })
        prev = rev_p50                          # chain off the median path
    return out


def render(drivers, source):
    g = [d["growth"] for d in drivers]
    sig = sum(d["growth_sigma"] for d in drivers) / len(drivers)
    print(f"\n=== Revenue-driven DCF assumptions  ({source}) ===")
    print(f"{'FY':>6} {'P50 rev':>10} {'P10..P90':>16} {'YoY growth':>11} {'sigma':>8}")
    for d in drivers:
        band = f"{d['rev_p10']:.0f}..{d['rev_p90']:.0f}"
        print(f"{d['year']:>6} {d['rev_p50']:>10.1f} {band:>16} "
              f"{d['growth']*100:>9.1f}% {d['growth_sigma']*100:>6.1f}%")
    print("\n-> Paste into the cockpit ASSUMPTIONS sheet / Inputs drawer:")
    print(f"   Revenue growth yr 1..5 = {', '.join(f'{x*100:.1f}%' for x in g)}")
    print(f"   SD revenue growth (pts) = {sig*100:.1f}%   (Monte-Carlo)")
    print("\n   Reminder: this informs the *inputs* — it is not a price forecast,")
    print("   and a foundation model can't see product cycles or recessions.\n")
    return {"source": source, "drivers": drivers,
            "growth_g1_g5": g, "mc_growth_sigma": round(sig, 4)}


def _synthetic_forecast(history_q, horizon=20):
    """A deterministic stand-in used ONLY by --selftest (no model needed),
    so the conversion math and output format can be validated offline."""
    last_year = np.array(history_q[-4:], dtype=float)
    seasonal = last_year / last_year.mean()          # quarterly shape
    base = float(history_q[-4:][0]) / seasonal[0]     # ~avg quarter level
    pt, lo, hi = [], [], []
    for i in range(horizon):
        yr = i // 4
        level = base * (1.05 ** (yr + 1)) / 4 * 4     # ~5% annual growth
        q = level * seasonal[i % 4]
        pt.append(q); lo.append(q * 0.93); hi.append(q * 1.08)  # ~+-7% band
    return np.array(pt), np.array(lo), np.array(hi)


def main():
    ap = argparse.ArgumentParser(description="TimesFM revenue forecast -> DCF drivers")
    ap.add_argument("--csv", help="CSV with one column of quarterly revenue")
    ap.add_argument("--col", default=None, help="column name (default: first column)")
    ap.add_argument("--years", type=int, default=5)
    ap.add_argument("--out", default="driver_forecast.json")
    ap.add_argument("--selftest", action="store_true",
                    help="validate conversion math WITHOUT loading TimesFM")
    a = ap.parse_args()

    if a.csv:
        import pandas as pd
        df = pd.read_csv(a.csv)
        col = a.col or df.columns[0]
        history = df[col].dropna().astype(float).tolist()
        source = f"{col} from {a.csv}"
    else:
        history = APPLE_QUARTERLY
        source = "Apple FY2015-FY2025 quarterly revenue (approx.)"

    horizon = 4 * a.years
    if a.selftest:
        pt, lo, hi = _synthetic_forecast(history, horizon)
        source += "  [SELFTEST: synthetic forecast, NOT TimesFM]"
    else:
        if len(history) < 32:
            print(f"WARNING: only {len(history)} points; TimesFM is most reliable "
                  f"with >=32 (use quarterly data).", file=sys.stderr)
        pt, lo, hi = timesfm_forecast(history, horizon)

    drivers = quarterly_to_drivers(history, pt, lo, hi, n_years=a.years)
    payload = render(drivers, source)
    with open(a.out, "w") as f:
        json.dump(payload, f, indent=2)
    print(f"wrote {a.out}")


if __name__ == "__main__":
    main()
