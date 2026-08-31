"""
app.py
--------------------------------
Phase 5: the web interface.

Turns the pipeline (SEC fetcher -> DCF engine -> Excel builder) into a
click-a-button tool -- no terminal, no coding knowledge required to use it.

This runs entirely on the deterministic forecast from Phase 2 -- NO AI /
API key needed yet. That's deliberate: it means this is a real, working,
publicly-deployable tool today. When Phase 3 (the AI assumption layer) is
ready, swap the call to generate_forecast() below for the AI-assisted
version and nothing else in this file has to change.

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Deploy for free, so anyone can use it via a public link:
    1. Push this whole folder to a public GitHub repo.
    2. Go to share.streamlit.io, sign in with GitHub, click "New app",
       point it at the repo, set the main file to app.py.
    3. It builds automatically and gives you a public URL to put on your resume.
"""

import streamlit as st

from sec_financials_fetcher import build_financial_table
from dcf_engine import (
    compute_historical_metrics, generate_forecast, WaccInputs, compute_wacc,
    build_fcff, discount_cash_flows, bridge_to_share_price,
    build_sensitivity_table, run_audit_checks,
)
from excel_builder import build_workbook

st.set_page_config(page_title="DCF Model Generator", page_icon=":material/bar_chart:", layout="wide")

# Simple abuse guard for a public deployment: caps how many times ONE
# visitor's browser session can run a fetch+build. This doesn't touch your
# wallet directly (see the API key discussion for that), but it keeps one
# person from hammering the SEC API or the server. Refreshing the page
# resets it, which is a fine tradeoff for a resume-demo tool.
MAX_RUNS_PER_SESSION = 8

if "run_count" not in st.session_state:
    st.session_state.run_count = 0

st.title("DCF Model Generator")
st.caption(
    "Pulls real historical financials from SEC filings and builds a full FCFF DCF valuation: "
    "forecast, WACC via CAPM, Gordon Growth terminal value, sensitivity table, and audit checks."
)
st.info(
    "Educational project, not investment advice. The forecast here uses a fixed historical-trend "
    "rule, not human or AI judgment -- treat the output as a starting point for further research, "
    "not a recommendation.",
    icon=":material/info:",
)

with st.sidebar:
    st.header("Inputs")
    ticker = st.text_input("Ticker", value="JNJ").strip().upper()
    terminal_growth = st.number_input(
        "Terminal growth rate", value=0.04, step=0.0025, format="%.4f",
        help="Long-run growth assumption used in the Gordon Growth terminal value. "
             "Often set near long-run nominal GDP growth (~3-4%).",
    )

    st.subheader("Cost of capital (CAPM)")
    st.caption("Risk-free rate: FRED (series DGS10). ERP, beta, cost of debt: A. Damodaran (NYU Stern).")
    risk_free_rate = st.number_input("Risk-free rate (10-yr UST)", value=0.0469, step=0.0001, format="%.4f")
    erp = st.number_input("Equity risk premium", value=0.0446, step=0.0001, format="%.4f")
    beta = st.number_input(
        "Relevered beta", value=1.00, step=0.01, format="%.3f",
        help="Don't leave this at 1.00 -- look up the company's own beta or relever an industry beta to its capital structure.",
    )
    pretax_kd = st.number_input("Pre-tax cost of debt", value=0.055, step=0.001, format="%.4f")
    tax_rate = st.number_input("Tax rate", value=0.25, step=0.01, format="%.2f")

    st.subheader("Market data")
    st.caption("Pull these from a source like a market data site or the company's latest 10-K.")
    market_price = st.number_input("Current share price ($)", value=0.0, step=0.01, min_value=0.0)
    shares_outstanding = st.number_input("Diluted shares outstanding (mm)", value=0.0, step=1.0, min_value=0.0)
    total_debt = st.number_input("Total debt ($mm)", value=0.0, step=1.0, min_value=0.0)
    cash = st.number_input("Cash & equivalents ($mm)", value=0.0, step=1.0, min_value=0.0)

    run_clicked = st.button("Generate DCF", type="primary", width='stretch')


def run_pipeline():
    """Runs the whole fetch -> forecast -> WACC -> DCF -> Excel pipeline and renders the results."""
    historicals = build_financial_table(ticker, years=5)

    if historicals.loc["Revenue"].dropna().empty:
        st.error(f"No revenue data found for '{ticker}'. Is it a US-listed company that files 10-Ks?")
        return

    hist = compute_historical_metrics(historicals)
    forecast = generate_forecast(hist, years=5, terminal_growth=terminal_growth)
    fcff_df = build_fcff(forecast, hist, tax_rate=tax_rate)

    market_value_equity = shares_outstanding * market_price
    wacc_inputs = WaccInputs(
        risk_free_rate=risk_free_rate, equity_risk_premium=erp, beta=beta,
        pretax_cost_of_debt=pretax_kd, tax_rate=tax_rate,
        market_value_equity=market_value_equity, total_debt=total_debt,
    )
    wacc_result = compute_wacc(wacc_inputs)
    dcf_result = discount_cash_flows(fcff_df, wacc_result.wacc, terminal_growth)  # raises if WACC <= g
    bridge = bridge_to_share_price(dcf_result.enterprise_value, total_debt, cash, shares_outstanding)
    checks = run_audit_checks(forecast, wacc_result, dcf_result, bridge, shares_outstanding)

    # --- Headline metrics ---
    st.subheader(f"{ticker} — DCF Results")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("WACC", f"{wacc_result.wacc:.2%}")
    c2.metric("Enterprise Value", f"${dcf_result.enterprise_value:,.0f}mm")
    c3.metric("Equity Value", f"${bridge['equity_value']:,.0f}mm")
    delta = f"{(bridge['price_per_share'] / market_price - 1):+.1%} vs market" if market_price > 0 else None
    c4.metric("Implied Price / Share", f"${bridge['price_per_share']:.2f}", delta=delta)

    tab1, tab2, tab3, tab4 = st.tabs(["Forecast", "FCFF & DCF", "Sensitivity", "Audit"])

    with tab1:
        st.dataframe(forecast.set_index("Year").T, width='stretch')

    with tab2:
        display_cols = ["FCFF", "Discount Factor", "PV of FCFF"]
        st.dataframe(dcf_result.fcff_table.set_index("Year")[display_cols].T, width='stretch')
        tv_pct = dcf_result.pv_terminal_value / dcf_result.enterprise_value
        st.write(
            f"Terminal value: ${dcf_result.terminal_value:,.0f}mm &nbsp;|&nbsp; "
            f"PV of terminal value: ${dcf_result.pv_terminal_value:,.0f}mm &nbsp;|&nbsp; "
            f"Terminal value as % of EV: {tv_pct:.1%}"
        )

    with tab3:
        sens = build_sensitivity_table(fcff_df, wacc_result.wacc, terminal_growth, total_debt, cash, shares_outstanding)
        st.dataframe(sens.style.format("${:.2f}", na_rep="n/a"), width='stretch')
        st.caption("Rows = WACC, columns = terminal growth rate. Center cell matches the headline price above.")

    with tab4:
        for check in checks:
            icon = ":white_check_mark:" if check["status"] == "PASS" else ":x:"
            st.write(f"{icon} **{check['check']}** — {check['detail']}")

    # --- Excel download ---
    with st.spinner("Building formatted Excel workbook..."):
        xlsx_path = f"/tmp/{ticker}_dcf.xlsx"
        build_workbook(
            ticker=ticker, company_name=ticker, historicals=historicals, forecast=forecast,
            tax_rate=tax_rate, wacc_inputs=wacc_inputs, terminal_growth=terminal_growth,
            total_debt=total_debt, cash=cash, shares_outstanding=shares_outstanding,
            market_price=market_price if market_price > 0 else None, output_path=xlsx_path,
        )
        with open(xlsx_path, "rb") as f:
            st.download_button(
                "Download formatted Excel model",
                data=f.read(),
                file_name=f"{ticker}_DCF_Model.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    st.caption(
        "Note: open the downloaded file and edit any cell (or just save it) once in Excel/Google "
        "Sheets/LibreOffice so it recalculates -- openpyxl writes formula text but not cached values."
    )


if run_clicked:
    if st.session_state.run_count >= MAX_RUNS_PER_SESSION:
        st.error(f"Limit of {MAX_RUNS_PER_SESSION} runs per session reached — refresh the page to reset.")
        st.stop()
    if not ticker:
        st.error("Enter a ticker.")
        st.stop()
    if shares_outstanding <= 0 or market_price <= 0 or total_debt < 0:
        st.error("Fill in a positive share price and diluted shares outstanding under Market data.")
        st.stop()

    st.session_state.run_count += 1

    try:
        with st.spinner(f"Pulling {ticker}'s financials from SEC..."):
            run_pipeline()
    except ValueError as e:
        st.error(str(e))
    except Exception as e:  # noqa: BLE001 -- surfaced deliberately for a public-facing app
        st.error(f"Something went wrong building this model: {e}")
else:
    st.write("Fill in the sidebar and click **Generate DCF** to build a model.")
