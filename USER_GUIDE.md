# Equity Valuation Cockpit — User Guide

A two-part tool for deciding **what a stock is worth** and **when it's safe to buy**.

- **`Equity_Valuation_Cockpit.xlsx`** — a 9-sheet Excel model with live formulas.
- **`equity-cockpit.html`** — a single, self-contained web page with the same engine, live charts, and optional live-data fetch. Double-click to open; works offline.
- **`Equity_Valuation_Cockpit_AAPL.xlsx`** — the same model pre-filled with Apple as a worked example.

Both files run the **identical math** and reconcile to the cent. Use whichever you prefer — Excel for tinkering with formulas, HTML for the visual dashboard.

> ⚠️ **This tool quantifies *your assumptions*. It is a thinking aid, not a price target or investment advice.** Garbage in, garbage out. Always sanity-check the inputs and the verdict against your own judgement.

---

## 1. The one rule

**You only ever type into one place:**

- **Excel** → the **`ASSUMPTIONS`** sheet (every editable cell has a gold left-border). Everything else is live formulas.
- **HTML** → the **Inputs** drawer (top-right button). Editing any field recomputes the whole page instantly.

Change one input and the verdict, fair value, charts, entry price, and Monte Carlo all update together.

---

## 2. Quick start (60 seconds)

1. Open either file. It loads pre-filled with a sample company (or Apple, for the `_AAPL` file).
2. Read the **verdict chip** top-left: `STRONG BUY → BUY → HOLD/WATCH → OVERVALUED/AVOID` (or `HOLD (RISK FLAG)`).
3. Overwrite the inputs with a real company's numbers (see §6 for where to get them).
4. Re-read the verdict, the **fair value**, and the **Buy-below** price.

---

## 3. How to read the verdict (the core logic)

The verdict is **not** just "is it below fair value." It compares your **margin of safety** to a **required bar that adapts to quality, market regime, and your own uncertainty**:

```
Margin of safety (MoS) = (Blended fair value − Price) ÷ Blended fair value

Required MoS (effective) = Quality-tier base
                         + Market-regime adjustment
                         + Valuation-uncertainty adjustment
                         + Portfolio-concentration adjustment
```

**Verdict:**

| Condition | Verdict |
|---|---|
| Altman Z < 1.81 **or** Beneish M > −1.78 | 🟠 **HOLD (RISK FLAG)** — capped regardless of cheapness |
| MoS ≥ Required + 15% | 🟢 **STRONG BUY** |
| MoS ≥ Required | 🟢 **BUY — MEETS BAR** |
| MoS ≥ Required − 15% | 🟠 **HOLD / WATCH** |
| otherwise | 🔴 **OVERVALUED / AVOID** |

### The required-MoS bar, explained

| Driver | Logic |
|---|---|
| **Quality tier** | Wide-moat **10%** · Solid **20%** · Average **35%** · Low-quality **50%**. (Set by Moat rating + ROIC-vs-WACC + Piotroski.) A wonderful business needs only a small discount; a weak one needs a big one. |
| **Market regime** | +5% when the market is **expensive** (high CAPE or thin implied ERP), −5% when **cheap**. |
| **Uncertainty** | +7.5% per level (1 = utility-stable, 3 = highly uncertain). |
| **Concentration** | +5% if you hold ≤6 names, +2.5% if ≤12. Concentrated portfolios demand more safety. |

This is the answer to *"by how many percent is it safe to buy"* — and it's exactly how Graham, Buffett/Munger, and Damodaran describe calibrating margin of safety.

---

## 4. Tour of the sheets / cards

| Sheet (Excel) / Card (HTML) | What it gives you |
|---|---|
| **DASHBOARD / Verdict** | The headline: verdict, blended fair value, upside, margin of safety, entry zone, score chips, one-line thesis. |
| **CONTEXT / Context & Regime** | Valuation vs the stock's **own 10-yr range**, quality tier, market regime, and the full **required-MoS build**. |
| **DCF** | 2-stage FCFF + FCFE discounted cash flow, WACC builder, dual terminal value (Gordon + exit multiple), and a WACC × terminal-growth sensitivity grid. |
| **RELATIVE** | Implied price from 6 peer multiples (P/E, fwd P/E, EV/EBITDA, EV/Sales, P/B, P/FCF) + PEG. |
| **VALUE_HEURISTICS** | Graham Number, Graham Formula, EPV, Owner Earnings, and a **Reverse-DCF** (the growth the market is pricing in). |
| **QUALITY** | ROIC vs WACC, Piotroski F-Score (0–9), Altman Z (distress), Beneish M (earnings manipulation). |
| **TIMING** | Fibonacci ladder, 50/200-day MAs, RSI, 52-week position, margin-of-safety bands, and a 4-tranche DCA planner. |
| **SCENARIO_MC** | Bull/Base/Bear DCF, prob-weighted fair value, a 2,000-trial **Monte Carlo** distribution, and **Kelly** position sizing. |

### The blended fair value

The headline fair value is a weighted average of the methods (weights editable on `ASSUMPTIONS` / Inputs):

| Method | Default weight |
|---|---|
| DCF — FCFF | 30% |
| DCF — FCFE | 15% |
| Relative (P/E & EV/EBITDA) | 25% |
| EPV | 20% |
| Graham Number | 5% |
| Graham Formula | 5% |

> **Tip for growth stocks:** Graham/EPV are brutal on high-growth companies (they assume away growth). For a fast-growing name, lower their weights and lean on DCF + Relative.

---

## 5. The two questions, and where each is answered

**① Is it worth it? (valuation)** → the **football field** chart / valuation summary. When the blended fair-value line sits well above the price line, it's cheap on multiple methods.

**② When do I buy? (timing & price)** → bottom-up:
- **Buy-below price** = fair value × (1 − required MoS). This is your trigger price.
- **Fibonacci ladder + MoS bands** give the entry zone.
- **DCA tranche planner** lays out a laddered 4-step entry instead of going all-in.
- **Reverse-DCF** tells you what growth the *current* price already assumes — a reality check on the bull case.

---

## 6. How to value a real stock (step by step)

1. **Gather the latest fundamentals** (most recent 10-K/10-Q or company release):
   - Market: price, shares outstanding, 52-week high/low, beta, recent swing high/low, 50/200-day MA, RSI.
   - Income/cash-flow: revenue, EBIT (operating) margin, tax rate, D&A %, CapEx %, net income, interest expense.
   - Balance sheet: total debt, cash + marketable securities, book equity, total assets, current assets/liabilities, retained earnings, total liabilities.
2. **Set your view:** 5-year growth path, terminal growth, peer multiples, exit multiple, moat rating (0–2), and the stock's **own 10-year P/E and EV/EBITDA range** (low/median/high — find on macrotrends/stockanalysis).
3. **Set risk calibration:** estimate uncertainty (1–3) and your portfolio's number of holdings.
4. **Read the verdict** and the buy-below price.

**Where to get data:** company SEC filings / investor-relations releases (authoritative), or macrotrends.net, stockanalysis.com, finviz, your brokerage. The HTML can also auto-fill via a live API (see §7).

### Worked examples (validated)

| Ticker | Price | Fair value | Read | Verdict |
|---|---|---|---|---|
| **GOOGL** | $366 | $184 | 34× P/E, top of its own range; heavy AI capex | OVERVALUED / AVOID* |
| **COST** | $953 | $323 | ~52× P/E — wonderful business, indefensible price | OVERVALUED / AVOID |
| **VZ** | $47 | $81 | Cheap (11× P/E) but $131B debt trips the distress gate | HOLD (RISK FLAG)** |

\* GOOGL is assumption-sensitive: the AVOID is driven by holding AI capex at 15% of revenue and including Graham/EPV. Normalize capex and de-weight Graham/EPV and it moves toward fair.
\** VZ shows a real limitation — see §8.

---

## 7. Live data (HTML only)

Click **Live data**, choose a provider (default: Financial Modeling Prep), and paste **your own free API key** (nothing is stored or hardcoded). It auto-fills price, shares, revenue, EBIT, debt, cash, EPS, BVPS, beta, and 52-week range, then recomputes. If the key/quota/CORS fails, it falls back to manual mode — the page never breaks. Free-tier keys (FMP ~250 calls/day, Finnhub, Alpha Vantage) are plenty for checking a few names.

---

## 8. Known limitations & caveats (read this)

- **Garbage in, garbage out.** The output is only as good as your inputs and growth assumptions. The tool makes those explicit and stress-tests them — that's its job.
- **Altman Z over-penalizes telecoms/utilities.** It was built for manufacturers and flags stable, heavily-levered, investment-grade businesses (e.g., Verizon) as "distressed" purely on leverage. Treat the risk flag as *"check the balance sheet,"* not *"bankruptcy imminent."*
- **DCF punishes heavy-capex growth.** Companies in a big investment phase (e.g., AI data centers) look expensive on near-term free cash flow. Consider normalizing CapEx % over the forecast as the build-out matures.
- **Graham Number / EPV are deep-value lenses.** They will call almost any growth stock "overvalued." Re-weight them down for compounders.
- **Market regime is a dial, not a timer.** CAPE and ERP predict long-run returns but are useless for short-term timing — they nudge the required bar, they don't issue buy/sell signals.
- **Monte Carlo is illustrative** and seed/trial-dependent (Excel 2,000 trials, HTML 10,000); percentiles will vary run to run.
- **Reverse-DCF caps at ±the solver grid** (−10% to +30%). For extremely cheap/expensive names the implied growth reads at the grid edge ("≥30%" or "≤−10%").

---

## 9. Methodology reference (formulas)

- **WACC** = (E/V)·Ke + (D/V)·Kd·(1−tax); Ke = Rf + β·ERP; Kd pre-tax = interest/debt (if available) else manual.
- **DCF (FCFF)**: project revenue → EBIT → NOPAT → FCFF (NOPAT + D&A − CapEx − ΔNWC); discount at WACC; terminal = average of Gordon-growth and exit-multiple; EV − net debt = equity ÷ shares.
- **Graham Number** = √(22.5 × EPS × BVPS). **Graham Formula** = EPS × (8.5 + 2g) × 4.4 / AAA-yield.
- **EPV** = NOPAT ÷ WACC − net debt, per share. **Owner Earnings** = NI + D&A − maintenance capex (~70% of capex).
- **Reverse-DCF**: solves the growth that makes the full DCF equal the current price.
- **ROIC** = NOPAT ÷ (debt + equity − cash). **Altman Z** = 1.2A+1.4B+3.3C+0.6D+1.0E. **Beneish M** = 8-ratio model.
- **Implied ERP** = market earnings yield + growth − 10-yr Treasury (preferred over the academically-flawed Fed model, which is shown only as a rough proxy).
- **Kelly** = (p·b − q) ÷ b, with a half-Kelly recommendation and a 25% cap.

---

## 10. Regenerating / editing the Excel

The Excel is generated by `build_excel.py` (run `python3 build_excel.py`). To produce a ticker-specific copy, you can edit the `ASSUMPTIONS` cells directly, or adapt `build_apple.py` (which overrides the defaults programmatically). The HTML is a single editable file — just open it in any editor.

---

*Built as a decision-support tool. Not investment advice. Verify everything before risking capital.*
