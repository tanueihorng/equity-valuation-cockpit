# Optional add-on — data-driven growth assumptions (TimesFM)

This module uses Google's [TimesFM](https://github.com/google-research/timesfm)
foundation model to forecast a company's **revenue** (a *driver* of value) and
convert it into the cockpit's DCF growth inputs (`g1..g5`) and a Monte-Carlo
growth sigma.

> ⚠️ **It forecasts the driver, not the price.** Stock prices are ~random walks;
> no model (TimesFM included) reliably predicts them. This turns your *hand-picked*
> growth assumptions into *data-informed* ones with uncertainty bands — treat the
> output as one input / a sanity check, never as a price oracle.

It is **fully optional** and kept separate from the core tool: the HTML stays
zero-dependency and the Excel only needs `openpyxl`.

## Setup

```bash
pip install -r forecast/requirements.txt
# Recommended: run the TimesFM preflight first — it blocks if RAM is too low
# (the 200M model needs ~1.6 GB free).
```

## Use

```bash
# Validate the conversion math without downloading the model:
python3 forecast/forecast_drivers.py --selftest

# Real forecast on the built-in Apple demo (loads the model):
python3 forecast/forecast_drivers.py

# Your own company (one column of quarterly revenue):
python3 forecast/forecast_drivers.py --csv my_revenue.csv --col revenue
```

It prints the 5-year revenue path with P10/P50/P90 bands, the implied `g1..g5`,
and the Monte-Carlo sigma — ready to paste into the cockpit's **ASSUMPTIONS**
sheet / **Inputs** drawer. It also writes `driver_forecast.json`.

**Why quarterly?** TimesFM is most reliable with ≥32 context points; annual
revenue (~10 points) is too short. Quarterly data also lets the model learn
seasonality, which it then folds into the forecast.
