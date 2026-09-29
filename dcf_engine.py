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
    with a fixed, explainable rule: margins held at the latest actual year,
    revenue growth starting from the historical CAGR (or an analyst
    override) and tapering to the terminal rate. That's the "safe
    default." An AI layer can later PROPOSE different assumption inputs
    to feed into this same engine, but it never touches the arithmetic --
    that split is what makes the tool's output trustworthy instead of a
    black box.
"""

from dataclasses import dataclass, field
import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Stage 1: Historical metrics
# ---------------------------------------------------------------------------

# Line items the model cannot work without. Missing entirely -> clear error.
REQUIRED_ITEMS = ["Revenue", "Cost of Goods Sold", "SG&A Expense",
                  "Depreciation & Amortization", "Capital Expenditures"]

# Line items where "never reported" genuinely means zero (a retailer with no
# R&D, a company with no debt or no marketable securities). If SEC has NO
# data for these in any year, they are set to 0 -- which is exactly what the
# Excel model shows and assumes, so the workbook and this engine agree.
OPTIONAL_ZERO_ITEMS = ["R&D Expense", "Cash & Equivalents", "Marketable Securities",
                       "Accounts Receivable", "Inventory", "Accounts Payable",
                       "Short-Term Debt", "Long-Term Debt"]

# Above this, the historical-trend rule is extrapolating hypergrowth and the
# audit flags it for REVIEW (override Year-1 growth with guidance/consensus).
MAX_MECHANICAL_GROWTH = 0.30


def _row_has_data(financials: pd.DataFrame, row_name: str) -> bool:
    """True only if `row_name` exists AND contains at least one real (non-null) value."""
    return row_name in financials.index and financials.loc[row_name].notna().any()


def prepare_historicals(financials: pd.DataFrame) -> pd.DataFrame:
    """
    The single cleaned version of the historicals that BOTH this engine and
    excel_builder.py use, so the Python numbers and the Excel formulas can't
    drift apart:
      * required rows that are completely empty -> ValueError (a company whose
        filings use a tag LINE_ITEMS doesn't know about can't be modeled yet)
      * optional rows that are completely empty -> 0 (a real economic state)
      * gaps in individual years stay NaN (shown as "N/A" in Excel and
        excluded from any calculation that needs them)
    """
    df = financials.copy().astype(float)
    for item in REQUIRED_ITEMS:
        if not _row_has_data(df, item):
            raise ValueError(
                f"No data found for required line item '{item}'. This company's SEC "
                f"filings may tag it under a name not in LINE_ITEMS (sec_financials_fetcher.py) "
                f"-- add the tag it actually uses, or this ticker can't be modeled yet."
            )
    for item in OPTIONAL_ZERO_ITEMS:
        if not _row_has_data(df, item):
            df.loc[item] = 0.0
    if pd.isna(df.loc["Revenue"].iloc[-1]):
        raise ValueError("The most recent fiscal year has no revenue, so there's nothing to grow "
                         "the forecast from. Check the Historicals table for misaligned years.")
    df.attrs = dict(financials.attrs)
    return df


@dataclass
class HistoricalMetrics:
    revenue_cagr: float
    latest_revenue: float
    gross_margin: float        # latest actual
    sga_pct: float             # latest actual
    rd_pct: float              # latest actual
    da_pct: float              # latest actual
    capex_pct: float           # latest actual
    avg_nwc_pct: float
    latest_nwc_pct: float
    # Which fiscal period (column) each "latest actual" came from, and the
    # span the CAGR covers -- used to write honest notes in the Excel model.
    basis: dict = field(default_factory=dict)


def _latest(series: pd.Series) -> tuple[float, str]:
    valid = series.dropna()
    if valid.empty:
        return float("nan"), ""
    return float(valid.iloc[-1]), str(valid.index[-1])


def compute_historical_metrics(financials: pd.DataFrame) -> HistoricalMetrics:
    """
    `financials`: the table from sec_financials_fetcher.build_financial_table()
    (fiscal years as columns, oldest -> newest).

    Forecast rule inputs:
      * Revenue growth starts from the historical CAGR (the compounding-
        consistent growth rate; the old arithmetic mean of yearly growth rates
        overstates it when growth is volatile).
      * Margins and cost ratios use the MOST RECENT actual year. Averaging five
        years mixed stale periods into the forecast -- for NVIDIA it pulled the
        EBIT margin down to ~47% vs. 60% actually reported in FY2026, while the
        growth rule pushed revenue up, two errors partly cancelling out.
    """
    df = prepare_historicals(financials)
    revenue = df.loc["Revenue"]

    valid_rev = revenue.dropna()
    first_i = list(revenue.index).index(valid_rev.index[0])
    last_i = len(revenue) - 1
    periods = last_i - first_i
    cagr = (valid_rev.iloc[-1] / valid_rev.iloc[0]) ** (1 / periods) - 1 if periods > 0 else 0.0

    ratios = {
        "gross_margin": 1 - df.loc["Cost of Goods Sold"] / revenue,
        "sga_pct": df.loc["SG&A Expense"] / revenue,
        "rd_pct": df.loc["R&D Expense"] / revenue,
        "da_pct": df.loc["Depreciation & Amortization"] / revenue,
        "capex_pct": df.loc["Capital Expenditures"].abs() / revenue,
    }
    latest_vals, basis = {}, {"cagr_from": str(valid_rev.index[0]), "cagr_to": str(valid_rev.index[-1])}
    for name, series in ratios.items():
        latest_vals[name], basis[name] = _latest(series)

    nwc = df.loc["Accounts Receivable"] + df.loc["Inventory"] - df.loc["Accounts Payable"]
    nwc_pct = (nwc / revenue).dropna()
    if nwc_pct.empty:
        nwc_pct = pd.Series(0.0, index=revenue.index)
    basis["nwc_latest"] = str(nwc_pct.index[-1])
    basis["nwc_avg_from"], basis["nwc_avg_to"] = str(nwc_pct.index[0]), str(nwc_pct.index[-1])

    metrics = HistoricalMetrics(
        revenue_cagr=float(cagr),
        latest_revenue=float(revenue.iloc[-1]),
        avg_nwc_pct=float(nwc_pct.mean()),
        latest_nwc_pct=float(nwc_pct.iloc[-1]),
        basis=basis,
        **latest_vals,
    )

    # Last line of defense: one NaN here turns EV, equity value and share
    # price into "$nan" with no explanation. Fail loudly and say which input.
    for name, value in vars(metrics).items():
        if name != "basis" and not np.isfinite(value):
            raise ValueError(
                f"Couldn't compute '{name}' from this company's SEC data -- the line items "
                f"behind it don't overlap with the revenue years. Check the Historicals table."
            )
    return metrics


def balance_sheet_bridge(financials: pd.DataFrame) -> tuple[float, float]:
    """
    (total debt, cash & marketable securities) from the latest fiscal year's
    balance sheet, in $mm. Total debt = short-term + long-term borrowings --
    NOT total liabilities, which would double-count operating items like
    accounts payable that already sit in working capital. A missing latest-
    year value counts as 0 (the Excel bridge uses N() for the same behavior,
    and the audit tab flags any N/A cells).
    """
    df = prepare_historicals(financials)
    latest = df.iloc[:, -1].fillna(0.0)
    total_debt = float(latest["Short-Term Debt"] + latest["Long-Term Debt"])
    cash = float(latest["Cash & Equivalents"] + latest["Marketable Securities"])
    return total_debt, cash


# ---------------------------------------------------------------------------
# Stage 2: Forecast
# ---------------------------------------------------------------------------

def generate_forecast(hist: HistoricalMetrics, years: int = 5, terminal_growth: float = 0.04,
                       start_growth: float | None = None) -> pd.DataFrame:
    """
    Builds the 5-year forecast: revenue growth tapers linearly from a starting
    rate down to the terminal growth rate; margins and cost ratios are held at
    their most recent actual values; NWC drifts from its most recent actual
    toward its historical average.

    `start_growth` is the analyst override for Year 1 (e.g. from company
    guidance or consensus). If omitted, the historical revenue CAGR is used,
    never below terminal growth + 1pp.
    """
    override = start_growth is not None
    if not override:
        start_growth = max(hist.revenue_cagr, terminal_growth + 0.01)

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
        "Gross Margin %": [hist.gross_margin] * years,
        "SG&A % of Revenue": [hist.sga_pct] * years,
        "R&D % of Revenue": [hist.rd_pct] * years,
        "D&A % of Revenue": [hist.da_pct] * years,
        "Capex % of Revenue": [hist.capex_pct] * years,
        "NWC % of Revenue": np.linspace(hist.latest_nwc_pct, hist.avg_nwc_pct, years),
    })
    forecast.attrs["growth_basis"] = "override" if override else "cagr"
    forecast.attrs["terminal_growth"] = terminal_growth
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

MIN_WACC_G_SPREAD = 0.01  # below this, Gordon Growth blows up -> shown as n/a


def sensitivity_axes(wacc_center: float, g_center: float) -> tuple[list[float], list[float]]:
    """
    Clean, rounded grid values (same convention as the hand-built JNJ model):
    7 WACC rows in 0.50% steps and 5 terminal-growth columns in 0.50% steps,
    centered on the base case rounded to the nearest 0.50%. The live base case
    is shown separately, so it doesn't need to sit exactly on a grid line.
    """
    step = 0.005
    w0 = round(wacc_center / step) * step
    g0 = round(g_center / step) * step
    waccs = [round(w0 + (i - 3) * step, 4) for i in range(7)]
    gs = [round(g0 + (i - 2) * step, 4) for i in range(5)]
    return waccs, gs


def build_sensitivity_table(fcff_df: pd.DataFrame, wacc_center: float, g_center: float,
                             total_debt: float, cash: float, shares_outstanding: float) -> pd.DataFrame:
    wacc_values, g_values = sensitivity_axes(wacc_center, g_center)
    table = pd.DataFrame(
        index=[f"{w:.2%}" for w in wacc_values],
        columns=[f"{g:.2%}" for g in g_values],
        dtype=float,
    )
    for w in wacc_values:
        for g in g_values:
            row, col = f"{w:.2%}", f"{g:.2%}"
            if w - g < MIN_WACC_G_SPREAD:
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

    def check(name: str, passed: bool, detail: str = "", soft: bool = False):
        # soft=True -> a rule-of-thumb flag ("REVIEW"), not an error ("FAIL")
        status = "PASS" if passed else ("REVIEW" if soft else "FAIL")
        checks.append({"check": name, "status": status, "detail": detail})

    check("WACC within a reasonable band (3%-15%)",
          0.03 <= wacc_result.wacc <= 0.15, f"WACC = {wacc_result.wacc:.2%}", soft=True)
    y1 = float(forecast["Revenue Growth %"].iloc[0])
    check(f"Year-1 revenue growth <= {MAX_MECHANICAL_GROWTH:.0%} (else override with guidance/consensus)",
          y1 <= MAX_MECHANICAL_GROWTH, f"Year-1 growth = {y1:.1%}", soft=True)
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
    check("Terminal value is 50%-90% of EV (rule of thumb)",
          0.50 <= tv_pct_of_ev <= 0.90, f"TV = {tv_pct_of_ev:.1%} of EV", soft=True)
    check("No missing (NaN) values in the FCFF forecast",
          not dcf_result.fcff_table["FCFF"].isna().any())

    return checks
