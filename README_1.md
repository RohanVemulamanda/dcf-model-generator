# DCF Model Generator

A tool that pulls a public company's real historical financials from SEC filings and automatically builds a full FCFF discounted cash flow valuation — forecast, WACC via CAPM, Gordon Growth terminal value, sensitivity analysis, and audit checks — delivered as both a live web app and a downloadable, formula-driven Excel model.

**[Try it live](#)** — replace with your deployed Streamlit URL once it's live.

---

## Why this exists

This started as a hand-built FCFF DCF model for Johnson & Johnson (see that project [here](#) — link your `jnj-dcf-valuation` repo). Building it by hand meant working through every step of the methodology myself: pulling 10-Ks off SEC EDGAR, sourcing WACC inputs from FRED and Damodaran's data, building the forecast, discounting the cash flows, and stress-testing the output. This project takes that exact methodology and turns it into reusable code, so the same rigor applies to any company, not just one.

## What it does

Give it a ticker and a set of cost-of-capital assumptions, and it will:

1. Pull that company's historical revenue, margins, and balance sheet data directly from SEC's public XBRL data — no manual filing-hunting.
2. Build a 5-year forecast using a historical-trend rule: revenue growth tapers toward the terminal growth rate, margins hold at historical averages.
3. Compute WACC from CAPM (risk-free rate, equity risk premium, beta, cost of debt, capital structure weights).
4. Build unlevered free cash flow (FCFF), discount it, and apply a Gordon Growth terminal value.
5. Bridge enterprise value to equity value to an implied price per share.
6. Run the valuation across a WACC × terminal-growth sensitivity grid.
7. Run 8 automated audit checks (WACC sanity bounds, positive values, no blank cells, no errors) before handing back a result.
8. Export everything to a formatted `.xlsx` with live Excel formulas — not static numbers — following standard IB conventions: blue font for hardcoded inputs, black for same-sheet formulas, green for cross-sheet links.

## Architecture

```
sec_financials_fetcher.py   Ticker -> historical financials (SEC XBRL API)
dcf_engine.py                Historicals -> forecast -> WACC -> FCFF -> DCF -> sensitivity -> audit
excel_builder.py             Engine output -> formatted, formula-driven .xlsx
app.py                       Streamlit web interface wrapping all three
```

A deliberate design choice runs through all of it: the forecast-generation logic is plain, deterministic Python — the same historical-trend rule used for the JNJ model by hand. Nothing here uses an AI model to pick assumptions (yet — see Roadmap). That split matters: it means every number in the output is fully explainable and reproducible, which is the standard a real financial model has to meet.

## Try it yourself

```
pip install -r requirements.txt
streamlit run app.py
```

Fill in the sidebar (ticker, terminal growth rate, cost-of-capital inputs, market data) and click **Generate DCF**. It pulls live data, runs the full valuation, and gives you a downloadable Excel model.

To deploy it publicly: push this repo to GitHub, then deploy it free at [share.streamlit.io](https://share.streamlit.io) with `app.py` as the entry point.

## Validation

Before shipping this, I checked it two ways:

- **Against the manual JNJ build.** Feeding the engine the same inputs used in the hand-built model reproduced cost of equity, after-tax cost of debt, and WACC to the decimal, with enterprise value, equity value, and implied share price all within ~1% of the original.
- **Excel formulas against the underlying Python math.** The generated workbook was run through a real recalculation engine; the Excel-computed enterprise value, equity value, and share price matched the Python engine's numbers exactly, with zero formula errors across the workbook.

## Sources

- SEC EDGAR — XBRL company facts API (`data.sec.gov`)
- Federal Reserve Bank of St. Louis (FRED) — risk-free rate (series DGS10)
- Aswath Damodaran (NYU Stern) — equity risk premium, industry beta, cost of debt data

## Roadmap

The one piece not yet built: having an AI model *propose* the forecast assumptions — with reasoning, grounded in the real historical data — instead of the fixed taper-to-terminal-growth rule. The architecture is already set up for this: a new assumption-generating function would slot in as a drop-in replacement for the current forecast step, with the rest of the pipeline (WACC, discounting, Excel output, the web interface) unchanged. The core design principle for that addition: the AI proposes, the code still computes — an LLM should never be responsible for the arithmetic in a financial model.

## Disclaimer

Built as an academic/portfolio project to demonstrate DCF methodology and software engineering applied to finance. Not investment advice. Verify all figures independently before relying on them for any purpose.
