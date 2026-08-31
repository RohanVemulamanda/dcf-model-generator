"""
dcf_engine.py
--------------------------------
Phase 2 of the "AI DCF model generator" project.

This module encodes the exact valuation methodology used to build the
Johnson & Johnson model by hand, as reusable functions that work for ANY
company's historical financials -- not just JNJ.

Pipeline stages, matching the JNJ build step for step:
    1. Historical margin/growth analysis   -> compute_historical_metrics()
    2. Five-year revenue & margin forecast  -> generate_forecast()
    3. Beta unlevering/relevering            -> unlever_beta() / relever_beta()
    4. WACC via CAPM                         -> compute_wacc()
    5. FCFF build                            -> build_fcff()
    6. Gordon Growth terminal value + DCF     -> discount_cash_flows()
    7. Enterprise value -> equity -> price    -> bridge_to_share_price()
    8. WACC x terminal-growth sensitivity     -> build_sensitivity_table()
    9. Audit checks                          -> run_audit_checks()

Design principle (the reason this file is separate from any AI code):
    Everything here is plain, deterministic, testable Python. Given the
    same inputs, it always produces the same output -- exactly like an
    Excel model with real formulas. generate_forecast() picks assumptions
    the same conservative way we did by hand for JNJ: take historical
    averages, taper growth toward the terminal rate. That's the "safe
    default." An AI layer can later PROPOSE different assumption inputs
    to feed into this same engine, but it never touches the arithmetic --
    that split is what makes the tool's output trustworthy instead of a
    black box.
"""

from dataclasses import dataclass
import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Stage 1: Historical metrics
# ---------------------------------------------------------------------------

@dataclass
class HistoricalMetrics:
    avg_revenue_growth: float
    latest_revenue: float
    avg_gross_margin: float
    avg_sga_pct: float
    avg_rd_pct: float
    avg_da_pct: float
    avg_capex_pct: float
    avg_nwc_pct: float
    latest_nwc_pct: float


def _row_has_data(financials: pd.DataFrame, row_name: str) -> bool:
    """
    True only if `row_name` exists AND contains at least one real (non-null)
    value. A row can exist but be entirely blank -- sec_financials_fetcher.py
    always creates a row for every line item it looks for, even when SEC has
    no data under any of its known tags for that company (e.g. R&D Expense
    for a retailer like Costco, which genuinely doesn't report R&D). Checking
    only "is the row present" was the bug: a present-but-empty row silently
    turned into NaN and poisoned every downstream calculation (EBIT, FCFF,
    enterprise value, share price) instead of being handled explicitly.
    """
    return row_name in financials.index and financials.loc[row_name].notna().any()


def _pct_of_revenue(financials: pd.DataFrame, revenue: pd.Series, row_name: str,
                     required: bool, use_abs: bool = False) -> pd.Series:
    """
    Returns `row_name` as a % of revenue. Optional line items (like R&D,
    which most non-tech/pharma companies genuinely don't report) default to
    0 when absent -- that's a real, valid economic state, not missing data.
    Required line items (Revenue, COGS, SG&A, D&A, Capex) raise a clear,
    actionable error instead of silently propagating NaN through the whole
    model -- a company whose filings tag these under a name this project
    doesn't know about needs a new tag added to LINE_ITEMS in
    sec_financials_fetcher.py, not a model that quietly outputs "$nan".
    """
    if not _row_has_data(financials, row_name):
        if required:
            raise ValueError(
                f"No data found for required line item '{row_name}'. This company's SEC "
                f"filings may tag it under a name not in LINE_ITEMS (sec_financials_fetcher.py) "
                f"-- add the tag it actually uses, or this ticker can't be modeled yet."
            )
        return pd.Series(0.0, index=revenue.index)

    values = financials.loc[row_name].astype(float)
    if use_abs:
        values = values.abs()
    return values / revenue


def compute_historical_metrics(financials: pd.DataFrame) -> HistoricalMetrics:
    """
    `financials`: DataFrame with fiscal years as columns (oldest -> newest,
    left to right) and rows for at least: Revenue, Cost of Goods Sold,
    SG&A Expense, Depreciation & Amortization, Capital Expenditures.
    R&D Expense and the working-capital rows (Accounts Receivable,
    Inventory, Accounts Payable) are optional -- if missing (or present but
    empty), those assumptions default to 0. This is exactly the shape
    produced by sec_financials_fetcher.py's build_financial_table().
    """
    if not _row_has_data(financials, "Revenue"):
        raise ValueError("No revenue data found -- this ticker can't be modeled.")
    revenue = financials.loc["Revenue"].astype(float)
    growth = revenue.pct_change().dropna()

    gross_profit = revenue - _pct_of_revenue(financials, revenue, "Cost of Goods Sold", required=True) * revenue
    gross_margin = gross_profit / revenue

    sga_pct = _pct_of_revenue(financials, revenue, "SG&A Expense", required=True)
    rd_pct = _pct_of_revenue(financials, revenue, "R&D Expense", required=False)
    da_pct = _pct_of_revenue(financials, revenue, "Depreciation & Amortization", required=True)
    capex_pct = _pct_of_revenue(financials, revenue, "Capital Expenditures", required=True, use_abs=True)

    wc_rows = ["Accounts Receivable", "Inventory", "Accounts Payable"]
    if all(_row_has_data(financials, r) for r in wc_rows):
        nwc = (financials.loc["Accounts Receivable"].astype(float)
               + financials.loc["Inventory"].astype(float)
               - financials.loc["Accounts Payable"].astype(float))
        nwc_pct = nwc / revenue
    else:
        nwc_pct = pd.Series(0.0, index=revenue.index)

    return HistoricalMetrics(
        avg_revenue_growth=float(growth.mean()),
        latest_revenue=float(revenue.iloc[-1]),
        avg_gross_margin=float(gross_margin.mean()),
        avg_sga_pct=float(sga_pct.mean()),
        avg_rd_pct=float(rd_pct.mean()),
        avg_da_pct=float(da_pct.mean()),
        avg_capex_pct=float(capex_pct.mean()),
        avg_nwc_pct=float(nwc_pct.mean()),
        latest_nwc_pct=float(nwc_pct.iloc[-1]),
    )


# ---------------------------------------------------------------------------
# Stage 2: Forecast
# ---------------------------------------------------------------------------

def generate_forecast(hist: HistoricalMetrics, years: int = 5, terminal_growth: float = 0.04,
                       start_growth: float | None = None) -> pd.DataFrame:
    """
    Builds the 5-year revenue and margin forecast the same way we did for
    JNJ by hand: revenue growth tapers linearly from a starting rate down
    to the terminal growth rate; margins are held at their historical
    averages; NWC drifts from its most recent actual value toward its
    historical average.
    """
    if start_growth is None:
        # Same rule of thumb used for JNJ: start near the historical
        # average, but never below the terminal rate.
        start_growth = max(hist.avg_revenue_growth, terminal_growth + 0.01)

    growth_path = np.linspace(start_growth, terminal_growth, years)

    revenue = []
    rev = hist.latest_revenue
    for g in growth_path:
        rev *= (1 + g)
        revenue.append(rev)

    forecast = pd.DataFrame({
        "Year": [f"Yr {i + 1}" for i in range(years)],
        "Revenue Growth %": growth_path,
        "Revenue": revenue,
        "Gross Margin %": [hist.avg_gross_margin] * years,
        "SG&A % of Revenue": [hist.avg_sga_pct] * years,
        "R&D % of Revenue": [hist.avg_rd_pct] * years,
        "D&A % of Revenue": [hist.avg_da_pct] * years,
        "Capex % of Revenue": [hist.avg_capex_pct] * years,
        "NWC % of Revenue": np.linspace(hist.latest_nwc_pct, hist.avg_nwc_pct, years),
    })
    return forecast


# ---------------------------------------------------------------------------
# Stage 3: Beta unlever / relever (Hamada equation)
# ---------------------------------------------------------------------------

def unlever_beta(levered_beta: float, tax_rate: float, debt: float, equity: float) -> float:
    """Strips a peer/industry beta of its capital structure effect."""
    de = debt / equity
    return levered_beta / (1 + (1 - tax_rate) * de)


def relever_beta(unlevered_beta: float, tax_rate: float, debt: float, equity: float) -> float:
    """Re-applies leverage using the TARGET company's own capital structure."""
    de = debt / equity
    return unlevered_beta * (1 + (1 - tax_rate) * de)


# ---------------------------------------------------------------------------
# Stage 4: WACC (CAPM)
# ---------------------------------------------------------------------------

@dataclass
class WaccInputs:
    risk_free_rate: float
    equity_risk_premium: float
    beta: float                    # already relevered to the target company's own D/E
    pretax_cost_of_debt: float
    tax_rate: float
    market_value_equity: float
    total_debt: float


@dataclass
class WaccResult:
    cost_of_equity: float
    after_tax_cost_of_debt: float
    weight_equity: float
    weight_debt: float
    wacc: float


def compute_wacc(inp: WaccInputs) -> WaccResult:
    cost_of_equity = inp.risk_free_rate + inp.beta * inp.equity_risk_premium
    after_tax_kd = inp.pretax_cost_of_debt * (1 - inp.tax_rate)

    total_capital = inp.market_value_equity + inp.total_debt
    we = inp.market_value_equity / total_capital
    wd = inp.total_debt / total_capital
    wacc = we * cost_of_equity + wd * after_tax_kd

    return WaccResult(cost_of_equity, after_tax_kd, we, wd, wacc)


# ---------------------------------------------------------------------------
# Stage 5: FCFF build
# ---------------------------------------------------------------------------

def build_fcff(forecast: pd.DataFrame, hist: HistoricalMetrics, tax_rate: float) -> pd.DataFrame:
    df = forecast.copy()
    df["Gross Profit"] = df["Revenue"] * df["Gross Margin %"]
    df["SG&A"] = df["Revenue"] * df["SG&A % of Revenue"]
    df["R&D"] = df["Revenue"] * df["R&D % of Revenue"]
    df["EBIT"] = df["Gross Profit"] - df["SG&A"] - df["R&D"]
    df["NOPAT"] = df["EBIT"] * (1 - tax_rate)
    df["D&A"] = df["Revenue"] * df["D&A % of Revenue"]
    df["Capex"] = df["Revenue"] * df["Capex % of Revenue"]
    df["NWC Balance"] = df["Revenue"] * df["NWC % of Revenue"]

    # Change in NWC needs an anchor: the last ACTUAL (historical) NWC
    # balance, not just the forecast years compared to each other.
    prior_nwc_balance = hist.latest_revenue * hist.latest_nwc_pct
    balances = [prior_nwc_balance] + df["NWC Balance"].tolist()
    df["Change in NWC"] = [balances[i + 1] - balances[i] for i in range(len(df))]

    df["FCFF"] = df["NOPAT"] + df["D&A"] - df["Capex"] - df["Change in NWC"]
    return df


# ---------------------------------------------------------------------------
# Stage 6: Discounting + Gordon Growth terminal value
# ---------------------------------------------------------------------------

@dataclass
class DcfResult:
    pv_explicit_fcff: float
    terminal_value: float
    pv_terminal_value: float
    enterprise_value: float
    fcff_table: pd.DataFrame


def discount_cash_flows(fcff_df: pd.DataFrame, wacc: float, terminal_growth: float) -> DcfResult:
    if wacc <= terminal_growth:
        raise ValueError(
            "WACC must exceed the terminal growth rate -- otherwise Gordon Growth "
            "produces a negative or infinite terminal value. Got WACC="
            f"{wacc:.4f}, g={terminal_growth:.4f}."
        )

    n = len(fcff_df)
    discount_factors = [1 / (1 + wacc) ** (i + 1) for i in range(n)]

    df = fcff_df.copy()
    df["Discount Factor"] = discount_factors
    df["PV of FCFF"] = df["FCFF"] * df["Discount Factor"]

    pv_explicit = float(df["PV of FCFF"].sum())

    terminal_fcff = df["FCFF"].iloc[-1] * (1 + terminal_growth)
    terminal_value = terminal_fcff / (wacc - terminal_growth)
    pv_terminal_value = terminal_value * discount_factors[-1]

    enterprise_value = pv_explicit + pv_terminal_value

    return DcfResult(pv_explicit, float(terminal_value), float(pv_terminal_value),
                      float(enterprise_value), df)


# ---------------------------------------------------------------------------
# Stage 7: Enterprise value -> equity value -> implied share price
# ---------------------------------------------------------------------------

def bridge_to_share_price(enterprise_value: float, total_debt: float, cash: float,
                           shares_outstanding: float) -> dict:
    equity_value = enterprise_value - total_debt + cash
    price_per_share = equity_value / shares_outstanding
    return {
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
        "price_per_share": price_per_share,
    }


# ---------------------------------------------------------------------------
# Stage 8: Sensitivity table
# ---------------------------------------------------------------------------

def build_sensitivity_table(fcff_df: pd.DataFrame, wacc_center: float, g_center: float,
                             total_debt: float, cash: float, shares_outstanding: float,
                             wacc_step: float = 0.005, g_step: float = 0.0025,
                             grid_size: int = 5) -> pd.DataFrame:
    half = grid_size // 2
    wacc_values = [wacc_center + (i - half) * wacc_step for i in range(grid_size)]
    g_values = [g_center + (i - half) * g_step for i in range(grid_size)]

    table = pd.DataFrame(
        index=[f"{w:.2%}" for w in wacc_values],
        columns=[f"{g:.2%}" for g in g_values],
        dtype=float,
    )

    for w in wacc_values:
        for g in g_values:
            row, col = f"{w:.2%}", f"{g:.2%}"
            if w <= g:
                table.loc[row, col] = np.nan
                continue
            result = discount_cash_flows(fcff_df, w, g)
            bridge = bridge_to_share_price(result.enterprise_value, total_debt, cash, shares_outstanding)
            table.loc[row, col] = bridge["price_per_share"]

    return table


# ---------------------------------------------------------------------------
# Stage 9: Audit checks
# ---------------------------------------------------------------------------

def run_audit_checks(forecast: pd.DataFrame, wacc_result: WaccResult, dcf_result: DcfResult,
                      bridge: dict, shares_outstanding: float) -> list[dict]:
    checks = []

    def check(name: str, passed: bool, detail: str = ""):
        checks.append({"check": name, "status": "PASS" if passed else "FAIL", "detail": detail})

    check("WACC within a reasonable band (3%-15%)",
          0.03 <= wacc_result.wacc <= 0.15, f"WACC = {wacc_result.wacc:.2%}")
    check("Capital structure weights sum to 100%",
          abs((wacc_result.weight_equity + wacc_result.weight_debt) - 1.0) < 1e-6)
    check("No negative revenue in forecast", bool((forecast["Revenue"] > 0).all()))
    check("Gross margin between 0% and 100%",
          bool(forecast["Gross Margin %"].between(0, 1).all()))
    check("Enterprise value is positive",
          dcf_result.enterprise_value > 0, f"EV = ${dcf_result.enterprise_value:,.0f}")
    check("Equity value is positive",
          bridge["equity_value"] > 0, f"Equity value = ${bridge['equity_value']:,.0f}")
    check("Shares outstanding is positive", shares_outstanding > 0)
    check("Implied share price is positive and finite",
          bridge["price_per_share"] > 0 and np.isfinite(bridge["price_per_share"]),
          f"${bridge['price_per_share']:.2f}")
    tv_pct_of_ev = dcf_result.pv_terminal_value / dcf_result.enterprise_value
    check("Terminal value is not an unreasonable share of EV (<90%)",
          tv_pct_of_ev < 0.90, f"TV = {tv_pct_of_ev:.1%} of EV")
    check("No missing (NaN) values in the FCFF forecast",
          not dcf_result.fcff_table["FCFF"].isna().any())

    return checks
