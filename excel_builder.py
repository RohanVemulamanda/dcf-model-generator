"""
excel_builder.py
--------------------------------
Phase 2b of the "AI DCF model generator" project.

Takes the outputs of dcf_engine.py for ANY company and writes an actual,
formula-driven Excel workbook -- not a static report. Every projection and
valuation cell is a real Excel formula that recalculates if you change an
input. Layout and styling deliberately mirror the hand-built Johnson &
Johnson model (JNJ_DCF_Model.xlsx), tab for tab:

    Cover | Historicals | WACC | Forecast | DCF | Sensitivity | Audit

Formatting conventions (same as the JNJ model):
    blue  = hardcoded input (pulled from a filing, or typed into the app)
    black = a formula that only uses cells on its own sheet
    green = a formula that pulls a value from a DIFFERENT sheet
    yellow fill = key outputs (FCFF, WACC, implied value per share)
    navy section bars, Arial throughout, no gridlines, notes under each tab,
    hover comments on inputs explaining their source or the rule behind them.

openpyxl writes formula TEXT but not computed values, so the file needs one
recalculation pass (open + save in Excel/Sheets/LibreOffice) before cached
values show up in previews.
"""

import math
import re
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from dcf_engine import (MAX_MECHANICAL_GROWTH, MIN_WACC_G_SPREAD, HistoricalMetrics,
                        prepare_historicals, sensitivity_axes)

# ---------------------------------------------------------------------------
# Style constants (taken from the JNJ model)
# ---------------------------------------------------------------------------
FONT_NAME = "Arial"
NAVY, NAVY_2 = "1F3864", "2E5395"
BLUE, BLACK, GREEN, WHITE = "0000FF", "000000", "008000", "FFFFFF"
NOTE_GRAY, NA_GRAY = "595959", "808080"
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")

FMT_MM = '#,##0;\\(#,##0\\);\\-'
FMT_PCT1 = '0.0%;\\(0.0%\\);\\-'
FMT_PCT2 = '0.00%;\\(0.00%\\);\\-'
FMT_PRICE = '\\$#,##0.00;"($"#,##0.00\\);\\-'
FMT_SHARES = '#,##0.0;\\(#,##0.0\\);\\-'
FMT_BETA = '0.0000'
FMT_INT = '0'

THIN = Side(style="thin")
DOUBLE = Side(style="double")
COMMENT_AUTHOR = "DCF Model Generator"
NA = "N/A"


def _font(color=BLACK, size=10, bold=False, italic=False):
    return Font(name=FONT_NAME, size=size, bold=bold, italic=italic, color=color)


def _put(ws, row, col, value, *, color=BLACK, size=10, bold=False, italic=False, fmt=None,
         fill=None, top=False, bottom=False, double_bottom=False, box=False, h=None, wrap=False,
         indent=0):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = _font(color, size, bold, italic)
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if box:
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    elif top or bottom or double_bottom:
        cell.border = Border(top=THIN if top else Side(),
                             bottom=DOUBLE if double_bottom else (THIN if bottom else Side()))
    if h or wrap or indent:
        cell.alignment = Alignment(horizontal=h or ("left" if indent else None), wrap_text=wrap, indent=indent,
                                   vertical="top" if wrap else None)
    return cell


def _na(ws, row, col, top=False, bold=False):
    """The JNJ model's convention for a missing data point: grey italic 'N/A', right-aligned."""
    return _put(ws, row, col, NA, color=NA_GRAY, italic=True, h="right", top=top)


def _comment(cell, text):
    cell.comment = Comment(text, COMMENT_AUTHOR)


def _setup(ws, widths: dict[str, float]):
    ws.sheet_view.showGridLines = False
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _title(ws, text, subtitle, last_col):
    ws.merge_cells(f"A1:{last_col}1")
    ws.merge_cells(f"A2:{last_col}2")
    _put(ws, 1, 1, text, color=NAVY, size=14, bold=True)
    _put(ws, 2, 1, subtitle, color=NOTE_GRAY, italic=True)
    ws.row_dimensions[1].height = 17.35


def _section(ws, row, text, last_col):
    ws.merge_cells(f"A{row}:{last_col}{row}")
    _put(ws, row, 1, text, color=WHITE, size=11, bold=True, fill=PatternFill("solid", fgColor=NAVY))


def _col_header(ws, row, labels, first_label="($ in millions)", first_bold=True):
    _put(ws, row, 1, first_label, bold=first_bold, bottom=True)
    for j, label in enumerate(labels):
        _put(ws, row, 2 + j, label, bold=True, bottom=True, h="right")


def _notes(ws, row, notes, first_col="A", last_col="H", width_units=170.0):
    """'Notes:' header, then one merged, wrapped, 8pt grey italic row per note."""
    _put(ws, row, 1, "Notes:", bold=True)
    chars_per_line = width_units * 1.45
    for i, text in enumerate(notes, start=1):
        r = row + i
        ws.merge_cells(f"{first_col}{r}:{last_col}{r}")
        _put(ws, r, ws[f"{first_col}1"].column, text, color=NOTE_GRAY, size=8, italic=True, wrap=True)
        lines = max(2, math.ceil(len(text) / chars_per_line))
        ws.row_dimensions[r].height = round(11.5 * lines + 2.5, 2)


def _finish(ws):
    for r in range(1, ws.max_row + 1):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = 15.0


def fiscal_year_label(period_end: str) -> str:
    """
    FY label from a fiscal period end date. A 52/53-week year that ends in the
    first days of January belongs to the prior year (JNJ's year ended Jan 2,
    2022 is FY2021); otherwise the label is the calendar year the period ends
    in (NVIDIA's year ended Jan 25, 2026 is FY2026, matching how it reports).
    """
    d = date.fromisoformat(str(period_end)[:10])
    return f"FY{d.year - 1 if (d.month == 1 and d.day <= 7) else d.year}"


def _nice_date(period_end: str) -> str:
    d = date.fromisoformat(str(period_end)[:10])
    return f"{d:%b} {d.day}, {d.year}"


def _long_date(d: date) -> str:
    return f"{d:%B} {d.day}, {d.year}"


def build_workbook(
    ticker: str,
    company_name: str,
    historicals: pd.DataFrame,     # rows = line items, columns = fiscal period ends (sec_financials_fetcher)
    forecast: pd.DataFrame,        # dcf_engine.generate_forecast output (the blue assumption values)
    hist: HistoricalMetrics,       # dcf_engine.compute_historical_metrics output (for assumption notes)
    wacc_inputs,                   # dcf_engine.WaccInputs (rates, beta, tax rate typed into the app)
    terminal_growth: float,
    shares_outstanding: float,     # diluted, millions
    market_price: float,
    output_path: str,
    cik: int | None = None,
    model_date: date | None = None,
):
    model_date = model_date or date.today()
    hdf = prepare_historicals(historicals)  # the same cleaned table the engine used
    ends = [str(c)[:10] for c in hdf.columns]
    n_hist, n_fcst = len(ends), len(forecast)
    hist_fy = [fiscal_year_label(e) for e in ends]
    last_fy = hist_fy[-1]
    last_fy_num = int(last_fy[2:])
    fcst_fy = [f"FY{last_fy_num + i + 1}E" for i in range(n_fcst)]
    first_fcst, last_fcst = fcst_fy[0], fcst_fy[-1]
    L = get_column_letter(1 + n_hist)               # last historical column letter
    FL = get_column_letter(1 + n_fcst)               # last forecast column letter
    tax_rate = wacc_inputs.tax_rate
    name = company_name or ticker
    cik_txt = f"CIK {cik:010d}" if cik else "CIK n/a"

    wb = Workbook()
    wb.remove(wb.active)
    ws_cover, ws_hist, ws_wacc = wb.create_sheet("Cover"), wb.create_sheet("Historicals"), wb.create_sheet("WACC")
    ws_fcst, ws_dcf = wb.create_sheet("Forecast"), wb.create_sheet("DCF")
    ws_sens, ws_audit = wb.create_sheet("Sensitivity"), wb.create_sheet("Audit")

    # =======================================================================
    # HISTORICALS
    # =======================================================================
    ws = ws_hist
    _setup(ws, {"A": 46, **{get_column_letter(2 + j): 13 for j in range(max(n_hist, 5))},
                get_column_letter(2 + max(n_hist, 5)): 2, get_column_letter(3 + max(n_hist, 5)): 60})
    _title(ws, f"{name} ({ticker}) — Historical Financial Statements",
           f"$ in millions  |  Source: SEC EDGAR XBRL company facts, Form 10-K filings ({cik_txt})", "F")

    H: dict[str, int] = {}          # key -> row
    is_na: dict[tuple[str, int], bool] = {}
    row = 4

    def value_of(item, j):
        if item not in hdf.index:
            return float("nan")
        return hdf.loc[item].iloc[j]

    def input_row(key, label, item=None, *, bold_label=False, bold=False, indent=0, top=False,
                  double_bottom=False, fmt=FMT_MM):
        nonlocal row
        item = item or key
        _put(ws, row, 1, label, bold=bold_label or bold, indent=indent)
        for j in range(n_hist):
            v = value_of(item, j)
            if pd.isna(v):
                _na(ws, row, 2 + j, top=top)
                is_na[(key, j)] = True
            else:
                _put(ws, row, 2 + j, float(v), color=BLUE, bold=bold, fmt=fmt, top=top,
                     double_bottom=double_bottom)
                is_na[(key, j)] = False
        H[key] = row
        row += 1

    token = re.compile(r"\{([^}]+)\}")

    def formula_row(key, label, template, *, bold=False, indent=0, top=False, fmt=FMT_MM,
                    label_italic=False):
        nonlocal row
        _put(ws, row, 1, label, bold=bold, italic=label_italic, indent=indent)
        deps = token.findall(template)
        for j in range(n_hist):
            col = get_column_letter(2 + j)
            if any(is_na.get((d, j), False) for d in deps):
                _na(ws, row, 2 + j, top=top)
                is_na[(key, j)] = True
                continue
            formula = "=" + token.sub(lambda m: f"{col}{H[m.group(1)]}", template)
            _put(ws, row, 2 + j, formula, bold=bold, fmt=fmt, top=top)
            is_na[(key, j)] = False
        H[key] = row
        row += 1

    def subheader(text):
        nonlocal row
        _put(ws, row, 1, text, bold=True, italic=True)
        row += 1

    def blank():
        nonlocal row
        row += 1

    # --- Income statement ---
    _section(ws, row, "INCOME STATEMENT — SELECTED ITEMS", "F"); row += 1
    _put(ws, row, 1, "Fiscal year ended", bold=True, italic=True)
    for j, e in enumerate(ends):
        _put(ws, row, 2 + j, _nice_date(e), color=BLUE, italic=True, h="right")
    row += 1
    _col_header(ws, row, hist_fy); row += 1
    input_row("Revenue", "Revenue", bold=True)
    input_row("Cost of Goods Sold", "Cost of revenue")
    formula_row("gp", "Gross profit", "{Revenue}-{Cost of Goods Sold}", bold=True, top=True)
    formula_row("gm", "  Gross margin %", "{gp}/{Revenue}", indent=1, fmt=FMT_PCT1)
    input_row("SG&A Expense", "Selling, general & administrative expense")
    formula_row("sga_pct", "  % of revenue", "{SG&A Expense}/{Revenue}", indent=1, fmt=FMT_PCT1)
    input_row("R&D Expense", "Research & development expense")
    formula_row("rd_pct", "  % of revenue", "{R&D Expense}/{Revenue}", indent=1, fmt=FMT_PCT1)
    formula_row("ebit", "EBIT (Gross profit − SG&A − R&D)", "{gp}-{SG&A Expense}-{R&D Expense}",
                bold=True, top=True)
    formula_row("ebit_m", "  EBIT margin %", "{ebit}/{Revenue}", indent=1, fmt=FMT_PCT1)
    input_row("Operating Income", "  Operating income, as reported (10-K) [memo]", indent=1)
    input_row("Net Income", "Net income", bold_label=True)
    formula_row("ni_m", "  Net margin %", "{Net Income}/{Revenue}", indent=1, fmt=FMT_PCT1)
    blank()

    # --- Cash flow ---
    _section(ws, row, "CASH FLOW STATEMENT — SELECTED ITEMS", "F"); row += 1
    _col_header(ws, row, hist_fy); row += 1
    input_row("Depreciation & Amortization", "Depreciation & amortization")
    formula_row("da_pct", "  % of revenue", "{Depreciation & Amortization}/{Revenue}", indent=1, fmt=FMT_PCT1)
    input_row("Capital Expenditures", "Capital expenditures (purchases of PP&E)")
    formula_row("capex_pct", "  % of revenue", "ABS({Capital Expenditures})/{Revenue}", indent=1, fmt=FMT_PCT1)
    blank()

    # --- Balance sheet ---
    _section(ws, row, "BALANCE SHEET — SELECTED ITEMS", "F"); row += 1
    _col_header(ws, row, hist_fy); row += 1
    subheader("Current assets")
    input_row("Cash & Equivalents", "Cash & cash equivalents", indent=1)
    input_row("Marketable Securities", "Marketable securities (current)", indent=1)
    input_row("Accounts Receivable", "Accounts receivable, net", indent=1)
    input_row("Inventory", "Inventories", indent=1)
    input_row("Total Current Assets", "Total current assets (as reported)", bold=True, top=True)
    input_row("Total Assets", "Total assets (as reported)", bold=True)
    blank()
    subheader("Current liabilities")
    input_row("Short-Term Debt", "Short-term debt (incl. current portion of LT debt)", indent=1)
    input_row("Accounts Payable", "Accounts payable", indent=1)
    input_row("Total Current Liabilities", "Total current liabilities (as reported)", bold=True, top=True)
    input_row("Long-Term Debt", "Long-term debt (non-current)")
    blank()
    subheader("Memo items")
    formula_row("tdebt", "Total debt (short-term + long-term)", "{Short-Term Debt}+{Long-Term Debt}",
                bold=True, indent=1)
    formula_row("cashms", "Cash & marketable securities", "{Cash & Equivalents}+{Marketable Securities}",
                bold=True, indent=1)
    formula_row("netdebt", "Net debt (Total debt − Cash & mkt. securities)", "{tdebt}-{cashms}",
                bold=True, indent=1)
    formula_row("nwc", "Non-cash working capital (AR + Inventory − AP)",
                "{Accounts Receivable}+{Inventory}-{Accounts Payable}", bold=True, indent=1)
    formula_row("nwc_pct", "  NWC % of revenue", "{nwc}/{Revenue}", indent=1, fmt=FMT_PCT1)
    hist_last_row = row - 1
    blank()
    _notes(ws, row, [
        f"Source: SEC EDGAR XBRL company facts API (data.sec.gov), {name} Form 10-K filings, {cik_txt}. "
        f"Each value is the most recently filed figure for that fiscal year, so later restatements "
        f"supersede originals.",
        "Only full-year (~365-day) 10-K values are used. Where the company changed US-GAAP tags over "
        "time, each year uses the highest-priority tag that reported it (see LINE_ITEMS in "
        "sec_financials_fetcher.py).",
        "N/A = no value found in SEC data for that year under any known tag; it is excluded from every "
        "calculation that needs it. An optional item the company never reports at all (e.g. R&D, "
        "marketable securities, short-term debt) is shown as zero, which is also what the model assumes.",
        "EBIT is defined as Gross profit − SG&A − R&D, the same definition used in the forecast. It can "
        "differ from reported operating income when the company books other operating items (e.g. "
        "impairments or one-time charges); the reported figure is shown as a memo line and compared on "
        "the Audit tab.",
    ], width_units=46 + 13 * 5 + 2 + 60)
    _finish(ws)

    # Bridge references: link the memo totals like the JNJ model does; if the
    # latest memo cell is N/A, fall back to N() so a gap counts as zero (the
    # engine's balance_sheet_bridge() does the same) instead of erroring.
    last_j = n_hist - 1

    def bridge_ref(memo_key, parts):
        if not is_na[(memo_key, last_j)]:
            return f"Historicals!{L}{H[memo_key]}"
        return "(" + "+".join(f"N(Historicals!{L}{H[p]})" for p in parts) + ")"

    DEBT_REF = bridge_ref("tdebt", ["Short-Term Debt", "Long-Term Debt"])
    CASH_REF = bridge_ref("cashms", ["Cash & Equivalents", "Marketable Securities"])

    # =======================================================================
    # WACC
    # =======================================================================
    ws = ws_wacc
    _setup(ws, {"A": 50, "B": 16, "C": 2, "D": 60})
    _title(ws, f"{name} — Weighted Average Cost of Capital (WACC)",
           "CAPM cost of equity, after-tax cost of debt, and market-value capital structure weights", "D")
    W = {}
    r = 4
    _section(ws, r, "COST OF EQUITY — CAPM", "B"); r += 1

    def w_row(key, label, value, fmt, *, color=BLACK, bold=False, top=False, comment=None, label_bold=None,
              size=10, fill=None):
        nonlocal r
        lab = _put(ws, r, 1, label, bold=bold if label_bold is None else label_bold, size=size)
        cell = _put(ws, r, 2, value, color=color, bold=bold, fmt=fmt, top=top, size=size, fill=fill)
        if comment:
            _comment(cell, comment)
        W[key] = r
        r += 1
        return lab

    entered = f"Entered in the DCF Model Generator app on {_long_date(model_date)}."
    w_row("rf", "Risk-free rate (10-yr U.S. Treasury yield)", wacc_inputs.risk_free_rate, FMT_PCT2,
          color=BLUE, comment=f"{entered} Suggested source: FRED series DGS10 (10-Year Treasury Constant "
                              f"Maturity), https://fred.stlouisfed.org/series/DGS10")
    w_row("erp", "Equity risk premium", wacc_inputs.equity_risk_premium, FMT_PCT2, color=BLUE,
          comment=f"{entered} Suggested source: A. Damodaran (NYU Stern), implied equity risk premium "
                  f"for the US market.")
    w_row("beta", "Levered beta (company-specific)", wacc_inputs.beta, FMT_BETA, color=BLUE,
          comment=f"{entered} Used as-is, so it should already reflect this company's own capital "
                  f"structure (e.g. a Damodaran industry unlevered beta relevered to the company's D/E, "
                  f"or a regression beta).")
    w_row("coe", "Cost of equity  [ = Rf + β×ERP ]", f"=B{W['rf']}+B{W['beta']}*B{W['erp']}", FMT_PCT2,
          bold=True, top=True)
    r += 1
    _put(ws, r, 1, f"{ticker} capital structure (market values)", bold=True, italic=True); r += 1
    w_row("px", f"{ticker} current share price", market_price, FMT_PRICE, color=BLUE,
          comment=f"{entered} Pull from a market data source on the same day as the risk-free rate.")
    w_row("sh", "Diluted shares outstanding (millions)", shares_outstanding, FMT_SHARES, color=BLUE,
          comment=f"{entered} Suggested source: diluted weighted-average shares from the latest 10-Q/10-K, "
                  f"or current shares outstanding from a market data source.")
    w_row("mve", "Market value of equity ($mm)", f"=B{W['px']}*B{W['sh']}", FMT_MM, bold=True)
    w_row("debt", f"Total debt ($mm) — {last_fy}A (short-term + long-term)", f"={DEBT_REF}", FMT_MM,
          color=GREEN)
    w_row("de", f"{ticker} D/E (market value of equity)", f"=B{W['debt']}/B{W['mve']}", FMT_PCT2)
    r += 1
    _section(ws, r, "COST OF DEBT", "B"); r += 1
    w_row("kd", "Pre-tax cost of debt", wacc_inputs.pretax_cost_of_debt, FMT_PCT2, color=BLUE,
          comment=f"{entered} Suggested sources: yield on the company's own long-dated bonds, or "
                  f"Damodaran's industry-average / synthetic-rating cost of debt.")
    w_row("tax", "Marginal tax rate (used for cost of debt and taxes on EBIT)", tax_rate, FMT_PCT2,
          color=BLUE, comment=f"{entered} Also drives 'Taxes on EBIT' on the Forecast tab. Common choices: "
                              f"~25% US federal + state marginal rate, or company tax-rate guidance.")
    w_row("atkd", "After-tax cost of debt  [ = Pre-tax × (1 − t) ]", f"=B{W['kd']}*(1-B{W['tax']})",
          FMT_PCT2, bold=True, top=True)
    r += 1
    _section(ws, r, "CAPITAL STRUCTURE WEIGHTS (Market-Value)", "B"); r += 1
    w_row("mve2", "Market value of equity ($mm)", f"=B{W['mve']}", FMT_MM)
    w_row("debt2", "Total debt ($mm)", f"=B{W['debt']}", FMT_MM)
    w_row("cap", "Total capital ($mm)", f"=B{W['mve2']}+B{W['debt2']}", FMT_MM, bold=True, top=True)
    w_row("we", "Weight of equity", f"=B{W['mve2']}/B{W['cap']}", FMT_PCT2)
    w_row("wd", "Weight of debt", f"=B{W['debt2']}/B{W['cap']}", FMT_PCT2)
    r += 1
    _section(ws, r, "WACC", "B"); r += 1
    w_row("wacc", "WACC  [ = We×Ke + Wd×Kd(1−t) ]",
          f"=B{W['we']}*B{W['coe']}+B{W['wd']}*B{W['atkd']}", FMT_PCT2, bold=True, top=True, size=12,
          fill=YELLOW_FILL)
    ws.cell(row=W["wacc"], column=1).font = _font(size=12, bold=True)
    wacc_last_row = r - 1
    r += 2
    _notes(ws, r, [
        f"Risk-free rate, equity risk premium, beta, pre-tax cost of debt, tax rate, share price and diluted "
        f"shares were entered by the user in the app on {_long_date(model_date)}; hover each blue input for "
        f"suggested sources.",
        f"Capital structure weights use {name}'s market value of equity (current share price × diluted shares) "
        f"and book value of total debt (short-term + long-term borrowings from the {last_fy}A balance sheet) as "
        f"a proxy for its market value. Operating leases and non-debt liabilities are excluded.",
        "Cost of equity uses CAPM. Because the beta input is used as-is, it should already be levered to this "
        "company's own D/E; the WACC is only as good as the inputs typed into the app.",
    ], last_col="D", width_units=50 + 16 + 2 + 60)
    _finish(ws)
    WACC_CELL, TAX_CELL = f"WACC!$B${W['wacc']}", f"WACC!$B${W['tax']}"
    PX_CELL, SH_CELL, MVE_CELL = f"WACC!$B${W['px']}", f"WACC!$B${W['sh']}", f"WACC!$B${W['mve']}"

    # =======================================================================
    # FORECAST
    # =======================================================================
    ws = ws_fcst
    _setup(ws, {"A": 54, **{get_column_letter(2 + j): 12 for j in range(max(n_hist, n_fcst, 5))},
                "G": 2, "H": 60})
    _title(ws, f"{name} — Revenue, Margin & FCFF Forecast",
           f"$ in millions, except %  |  Five-year forecast ({first_fcst}–{last_fcst}) built from "
           f"{last_fy}A actuals", "F")
    F = {}
    r = 4
    _section(ws, r, "HISTORICAL TREND SUMMARY (reference — linked from Historicals tab)", "F"); r += 1
    _col_header(ws, r, hist_fy, first_label=""); r += 1

    def trend_row(key, label, hist_key):
        nonlocal r
        _put(ws, r, 1, label)
        for j in range(n_hist):
            col = get_column_letter(2 + j)
            if is_na[(hist_key, j)]:
                _na(ws, r, 2 + j)
            else:
                _put(ws, r, 2 + j, f"=Historicals!{col}{H[hist_key]}", color=GREEN, fmt=FMT_PCT1)
        F[key] = r
        r += 1

    _put(ws, r, 1, "Revenue growth %")
    _na(ws, r, 2)
    for j in range(1, n_hist):
        col, prev = get_column_letter(2 + j), get_column_letter(1 + j)
        if is_na[("Revenue", j)] or is_na[("Revenue", j - 1)]:
            _na(ws, r, 2 + j)
        else:
            _put(ws, r, 2 + j, f"=Historicals!{col}{H['Revenue']}/Historicals!{prev}{H['Revenue']}-1",
                 color=GREEN, fmt=FMT_PCT1)
    F["trend_growth"] = r; r += 1
    trend_row("trend_gm", "Gross margin %", "gm")
    trend_row("trend_sga", "SG&A % of revenue", "sga_pct")
    trend_row("trend_rd", "R&D % of revenue", "rd_pct")
    trend_row("trend_da", "D&A % of revenue", "da_pct")
    trend_row("trend_capex", "Capex % of revenue", "capex_pct")
    trend_row("trend_nwc", "NWC % of revenue", "nwc_pct")
    r += 1

    growth_basis = forecast.attrs.get("growth_basis", "cagr")
    _section(ws, r, f"FORECAST ASSUMPTIONS — {first_fcst}–{last_fcst} (rule-based starting point; blue = "
                    f"override freely)", "F"); r += 1
    _col_header(ws, r, fcst_fy, first_label=""); r += 1

    b = hist.basis
    fy_of = lambda end: fiscal_year_label(end) if end else last_fy  # noqa: E731
    g_first = float(forecast["Revenue Growth %"].iloc[0])
    if growth_basis == "override":
        growth_note = (f"Year 1 = analyst override entered in the app ({g_first:.1%}), e.g. from company "
                       f"guidance or consensus. Years 2–5 taper linearly to the {terminal_growth:.1%} "
                       f"terminal growth rate.")
    else:
        growth_note = (f"Year 1 = {fy_of(b.get('cagr_from'))}–{fy_of(b.get('cagr_to'))} revenue CAGR "
                       f"({hist.revenue_cagr:.1%}), floored at terminal growth + 1pp; Years 2–5 taper "
                       f"linearly to the {terminal_growth:.1%} terminal growth rate. For companies whose "
                       f"growth is changing fast, override Year 1 with guidance or consensus.")

    def held(label, value, end):
        return (f"Held flat at the {fy_of(end)} actual ({value:.1%}). The most recent year reflects current "
                f"pricing, mix and scale; a multi-year average would blend in older periods. Override if you "
                f"expect this to change.")

    def assumption_row(key, label, series, note):
        nonlocal r
        lab = _put(ws, r, 1, label)
        _comment(lab, note)
        for j in range(n_fcst):
            _put(ws, r, 2 + j, float(series.iloc[j]), color=BLUE, fmt=FMT_PCT1)
        F[key] = r
        r += 1

    assumption_row("g", "Revenue growth %", forecast["Revenue Growth %"], growth_note)
    assumption_row("gm", "Gross margin %", forecast["Gross Margin %"],
                   held("Gross margin", hist.gross_margin, b.get("gross_margin")))
    assumption_row("sga", "SG&A % of revenue", forecast["SG&A % of Revenue"],
                   held("SG&A", hist.sga_pct, b.get("sga_pct")))
    assumption_row("rd", "R&D % of revenue", forecast["R&D % of Revenue"],
                   held("R&D", hist.rd_pct, b.get("rd_pct")))
    _put(ws, r, 1, "Marginal tax rate (linked from WACC tab)")
    for j in range(n_fcst):
        _put(ws, r, 2 + j, f"={TAX_CELL}", color=GREEN, fmt=FMT_PCT1)
    F["tax"] = r; r += 1
    assumption_row("da", "D&A % of revenue", forecast["D&A % of Revenue"],
                   held("D&A", hist.da_pct, b.get("da_pct")))
    assumption_row("capex", "Capex % of revenue", forecast["Capex % of Revenue"],
                   held("Capex", hist.capex_pct, b.get("capex_pct")))
    assumption_row("nwc", "NWC % of revenue", forecast["NWC % of Revenue"],
                   f"Starts at the {fy_of(b.get('nwc_latest'))} actual ({hist.latest_nwc_pct:.1%}) and drifts "
                   f"linearly to the {fy_of(b.get('nwc_avg_from'))}–{fy_of(b.get('nwc_avg_to'))} average "
                   f"({hist.avg_nwc_pct:.1%}). NWC = AR + Inventory − AP.")
    r += 1

    _section(ws, r, "FORECAST BUILD — UNLEVERED FREE CASH FLOW (FCFF)", "F"); r += 1
    _col_header(ws, r, fcst_fy, first_label=""); r += 1

    def build_row(key, label, make, *, bold=False, top=False, fmt=FMT_MM, label_italic=False, indent=0,
                  fill=None, first_color=BLACK):
        nonlocal r
        _put(ws, r, 1, label, bold=bold, italic=label_italic, indent=indent, fill=fill)
        for j in range(n_fcst):
            c, p = get_column_letter(2 + j), get_column_letter(1 + j)
            _put(ws, r, 2 + j, make(c, p, j), bold=bold, top=top, fmt=fmt, fill=fill,
                 color=first_color if j == 0 else BLACK)
        F[key] = r
        r += 1

    rev_r = r
    build_row("rev", "Revenue", lambda c, p, j: (f"=Historicals!{L}{H['Revenue']}*(1+{c}{F['g']})" if j == 0
                                                 else f"={p}{rev_r}*(1+{c}{F['g']})"),
              bold=True, first_color=GREEN)
    build_row("gp", "Gross profit", lambda c, p, j: f"={c}{F['rev']}*{c}{F['gm']}")
    build_row("sga_d", "SG&A", lambda c, p, j: f"={c}{F['rev']}*{c}{F['sga']}")
    build_row("rd_d", "R&D", lambda c, p, j: f"={c}{F['rev']}*{c}{F['rd']}")
    build_row("ebit", "EBIT", lambda c, p, j: f"={c}{F['gp']}-{c}{F['sga_d']}-{c}{F['rd_d']}", bold=True, top=True)
    build_row("ebit_m", "  EBIT margin %", lambda c, p, j: f"={c}{F['ebit']}/{c}{F['rev']}", fmt=FMT_PCT1,
              label_italic=True, indent=1)
    build_row("taxes", "Less: Taxes on EBIT", lambda c, p, j: f"={c}{F['ebit']}*{c}{F['tax']}")
    build_row("nopat", "NOPAT", lambda c, p, j: f"={c}{F['ebit']}-{c}{F['taxes']}", bold=True, top=True)
    build_row("da_d", "Plus: D&A", lambda c, p, j: f"={c}{F['rev']}*{c}{F['da']}")
    build_row("capex_d", "Less: Capex", lambda c, p, j: f"={c}{F['rev']}*{c}{F['capex']}")
    build_row("nwc_bal", "  (memo) NWC balance", lambda c, p, j: f"={c}{F['rev']}*{c}{F['nwc']}",
              label_italic=True, indent=1)
    # Year-1 change in NWC is anchored on the last ACTUAL balance.
    if not is_na[("nwc", last_j)]:
        prior_nwc = f"Historicals!{L}{H['nwc']}"
    else:  # latest NWC missing: latest revenue x last available NWC % (same as the engine)
        lv = max(j for j in range(n_hist) if not is_na[("nwc_pct", j)]) if any(
            not is_na[("nwc_pct", j)] for j in range(n_hist)) else None
        prior_nwc = (f"Historicals!{L}{H['Revenue']}*Historicals!{get_column_letter(2 + lv)}{H['nwc_pct']}"
                     if lv is not None else "0")
    build_row("d_nwc", "Less: Increase in NWC",
              lambda c, p, j: (f"={c}{F['nwc_bal']}-{prior_nwc}" if j == 0
                               else f"={c}{F['nwc_bal']}-{p}{F['nwc_bal']}"), first_color=GREEN)
    build_row("fcff", "Unlevered Free Cash Flow (FCFF)",
              lambda c, p, j: f"={c}{F['nopat']}+{c}{F['da_d']}-{c}{F['capex_d']}-{c}{F['d_nwc']}",
              bold=True, top=True, fill=YELLOW_FILL)
    fcst_last_row = r - 1
    r += 2
    _notes(ws, r, [
        "Blue assumption cells are rule-based starting points generated by the DCF Model Generator (a fixed "
        "historical-trend rule, not human or AI judgment); hover each assumption label for the rule used. "
        "Override any blue cell and the whole model recalculates.",
        f"Year-1 ({first_fcst}) revenue and increase-in-NWC formulas reference {last_fy}A actuals on the "
        f"Historicals tab (green links) as their base; all subsequent years build off the prior forecast column.",
        "EBIT is defined consistently with the Historicals tab: Gross profit less SG&A less R&D. Non-operating "
        "items (interest, other income/expense) and one-time charges are not forecast (treated as $0 going "
        "forward, a standard simplifying assumption).",
        "Taxes on EBIT use the marginal tax rate on the WACC tab. NWC is non-cash working capital "
        "(accounts receivable + inventory − accounts payable) modeled as a percentage of revenue.",
    ], width_units=54 + 12 * 5 + 2 + 60)
    _finish(ws)
    FCFF_ROW = F["fcff"]

    # =======================================================================
    # DCF
    # =======================================================================
    ws = ws_dcf
    _setup(ws, {"A": 48, **{get_column_letter(2 + j): 13 for j in range(max(n_fcst, 5))}, "G": 2, "H": 60})
    val_date = f"{_nice_date(ends[-1])} ({last_fy} fiscal year-end)"
    _title(ws, f"{name} — DCF Valuation (FCFF / Enterprise Value Approach)",
           f"Valuation date: {val_date}  |  $ in millions, except per-share data", "F")
    D = {}
    r = 4
    _section(ws, r, "PRESENT VALUE OF FORECAST FREE CASH FLOW", "F"); r += 1
    _col_header(ws, r, fcst_fy, first_label="", first_bold=False); r += 1

    def d_row(key, label, make, *, bold=False, top=False, fmt=FMT_MM, color=BLACK):
        nonlocal r
        _put(ws, r, 1, label, bold=bold)
        for j in range(n_fcst):
            c = get_column_letter(2 + j)
            _put(ws, r, 2 + j, make(c, j), bold=bold, top=top, fmt=fmt, color=color)
        D[key] = r
        r += 1

    d_row("period", "Discount period (years)", lambda c, j: j + 1, fmt=FMT_INT)
    d_row("fcff", "Unlevered Free Cash Flow (FCFF)", lambda c, j: f"=Forecast!{c}{FCFF_ROW}", bold=True,
          color=GREEN)
    d_row("wacc", "WACC", lambda c, j: f"={WACC_CELL}", fmt=FMT_PCT2, color=GREEN)
    d_row("df", "Discount factor  [ = 1 / (1+WACC)^period ]", lambda c, j: f"=1/(1+{c}{D['wacc']})^{c}{D['period']}",
          fmt=FMT_BETA)
    d_row("pv", "PV of FCFF", lambda c, j: f"={c}{D['fcff']}*{c}{D['df']}", bold=True, top=True)
    r += 1

    def single(key, label, value, fmt=FMT_MM, *, color=BLACK, bold=False, top=False, fill=None, comment=None):
        nonlocal r
        _put(ws, r, 1, label, bold=bold, fill=fill)
        cell = _put(ws, r, 2, value, color=color, bold=bold, top=top, fmt=fmt, fill=fill)
        if comment:
            _comment(cell, comment)
        D[key] = r
        r += 1

    _section(ws, r, "TERMINAL VALUE — GORDON GROWTH METHOD", "B"); r += 1
    single("g", "Terminal growth rate (g)", terminal_growth, FMT_PCT2, color=BLUE,
           comment=f"Entered in the app. Often set near long-run nominal GDP growth (~3–4%); it must stay below "
                   f"WACC for the Gordon Growth formula to work.")
    single("tfcff", f"Terminal year FCFF  [ = {last_fcst} FCFF × (1+g) ]", f"={FL}{D['fcff']}*(1+B{D['g']})")
    single("tv", "Terminal value  [ = Terminal FCFF / (WACC − g) ]", f"=B{D['tfcff']}/({FL}{D['wacc']}-B{D['g']})",
           bold=True, top=True)
    single("pvtv", "PV of terminal value  [ = TV × Year-5 discount factor ]", f"=B{D['tv']}*{FL}{D['df']}",
           bold=True, top=True)
    r += 1
    _section(ws, r, "ENTERPRISE VALUE TO IMPLIED SHARE PRICE", "B"); r += 1
    single("sumpv", f"Sum of PV of forecast FCFF ({first_fcst}–{last_fcst})", f"=SUM(B{D['pv']}:{FL}{D['pv']})")
    single("pvtv2", "PV of terminal value", f"=B{D['pvtv']}")
    single("ev", "Enterprise value", f"=B{D['sumpv']}+B{D['pvtv2']}", bold=True, top=True)
    single("tvpct", "  Terminal value % of enterprise value", f"=B{D['pvtv2']}/B{D['ev']}", FMT_PCT1)
    single("debt", f"Less: Total debt ({last_fy}A)", f"=-{DEBT_REF}", color=GREEN)
    single("cash", f"Plus: Cash & marketable securities ({last_fy}A)", f"={CASH_REF}", color=GREEN)
    single("eq", "Equity value", f"=B{D['ev']}+B{D['debt']}+B{D['cash']}", bold=True, top=True)
    single("sh", "Diluted shares outstanding (millions)", f"={SH_CELL}", FMT_SHARES, color=GREEN)
    single("px", "Implied value per share", f"=B{D['eq']}/B{D['sh']}", FMT_PRICE, bold=True, top=True,
           fill=YELLOW_FILL)
    single("mkt", "Current share price", f"={PX_CELL}", FMT_PRICE, color=GREEN)
    single("updown", "Implied upside / (downside)", f"=B{D['px']}/B{D['mkt']}-1", FMT_PCT1, bold=True)
    dcf_last_row = r - 1
    r += 2
    _notes(ws, r, [
        f"FCFF for {first_fcst}–{last_fcst} is linked from the Forecast tab. Discounting uses the standard "
        f"end-of-year (not mid-year) convention, with Year 1 = {first_fcst}.",
        f"Valuation date is set at the {last_fy}A fiscal year-end ({_nice_date(ends[-1])}), the last completed "
        f"balance sheet date, to keep the discount periods and the debt/cash bridge internally consistent. The "
        f"current share price (entered {_long_date(model_date)}) is a later comparison point, so the implied "
        f"upside/(downside) also spans the time in between.",
        "Terminal value uses the Gordon Growth (perpetuity growth) method: Terminal FCFF / (WACC − g). See the "
        "Sensitivity tab for how the implied share price moves with WACC and g.",
        f"The bridge uses {last_fy}A total debt (short-term + long-term borrowings, not total liabilities) and "
        f"cash & marketable securities, linked from the Historicals tab; diluted shares and the current share "
        f"price are linked from the WACC tab for consistency with the market-value weights used there.",
    ], width_units=48 + 13 * 5 + 2 + 60)
    _finish(ws)

    # =======================================================================
    # SENSITIVITY
    # =======================================================================
    ws = ws_sens
    _setup(ws, {"A": 20, **{get_column_letter(c): 12 for c in range(2, 8)}, "H": 2, "I": 60})
    _title(ws, f"{name} — Sensitivity Analysis: Implied Share Price",
           "Implied per-share value as a function of WACC (rows) and terminal growth rate g (columns)", "G")
    wacc_base = _engine_wacc(wacc_inputs)
    wacc_vals, g_vals = sensitivity_axes(wacc_base, terminal_growth)
    _section(ws, 4, "IMPLIED SHARE PRICE ($ / SHARE)", "G")
    _put(ws, 5, 1, "WACC ↓  /  g →", size=9, bold=True, italic=True, bottom=True, h="center", wrap=True)
    for j, g in enumerate(g_vals):
        _put(ws, 5, 3 + j, g, color=BLUE, bold=True, fmt=FMT_PCT2, bottom=True, h="center")
    _put(ws, 6, 1, "Base case (live)", bold=True, italic=True)
    _put(ws, 6, 2, f"={WACC_CELL}", color=GREEN, italic=True, fmt=FMT_PCT2)
    ws.merge_cells("C6:G6")
    _put(ws, 6, 3, f"=DCF!$B${D['px']}", color=GREEN, bold=True, fmt=FMT_PRICE)
    fr = f"Forecast!$B${FCFF_ROW}:${FL}${FCFF_ROW}"
    last_fcff = f"Forecast!${FL}${FCFF_ROW}"
    for i, w in enumerate(wacc_vals):
        rr = 7 + i
        _put(ws, rr, 2, w, color=BLUE, bold=True, fmt=FMT_PCT2, box=True)
        for j in range(len(g_vals)):
            c = get_column_letter(3 + j)
            formula = (f'=IF($B{rr}-{c}$5<{MIN_WACC_G_SPREAD},"n/a",(NPV($B{rr},{fr})'
                       f"+({last_fcff}*(1+{c}$5))/($B{rr}-{c}$5)/(1+$B{rr})^{n_fcst}"
                       f"+DCF!$B${D['debt']}+DCF!$B${D['cash']})/DCF!$B${D['sh']})")
            _put(ws, rr, 3 + j, formula, fmt=FMT_PRICE, box=True, h="right")
    sens_last_row = 6 + len(wacc_vals)
    _notes(ws, sens_last_row + 2, [
        f"Each grid cell independently recomputes the full DCF bridge at the row's WACC and the column's "
        f"terminal growth rate g: PV of {first_fcst}–{last_fcst} FCFF via NPV(), plus the PV of a Gordon Growth "
        f"terminal value, less {last_fy}A total debt, plus {last_fy}A cash & marketable securities, divided by "
        f"diluted shares outstanding.",
        "The \"Base case (live)\" row above the grid recalculates automatically from the WACC and DCF tabs and "
        "will not generally land exactly on a grid line, since the grid uses clean, rounded WACC/g increments "
        "for readability.",
        f"Cells show n/a where WACC is less than {MIN_WACC_G_SPREAD * 100:.0f} percentage point above g, where "
        f"the Gordon Growth "
        f"denominator (WACC − g) gets small enough to make the terminal value unstable.",
    ], last_col="I", width_units=20 + 12 * 6 + 2 + 60)
    _finish(ws)

    # =======================================================================
    # AUDIT
    # =======================================================================
    ws = ws_audit
    _setup(ws, {"A": 4, "B": 62, "C": 16, "D": 12, "E": 2, "F": 60})
    _title(ws, f"{name} DCF Model — Audit & Integrity Checks",
           "Every check below is a live formula; REVIEW flags a rule-of-thumb result worth a second look, "
           "FAIL flags an error", "F")
    _put(ws, 4, 1, "#", bold=True, bottom=True, h="center")
    _put(ws, 4, 2, "Check", bold=True, bottom=True, h="left")
    _put(ws, 4, 3, "Result", bold=True, bottom=True, h="left")
    _put(ws, 4, 4, "Status", bold=True, bottom=True, h="center")

    hL = lambda key: f"Historicals!{L}{H[key]}"  # noqa: E731
    dfs = [f"DCF!{get_column_letter(2 + j)}{D['df']}" for j in range(n_fcst)]
    checks = [
        ("WACC within a reasonable band (3%–15%)", f"={WACC_CELL}", FMT_PCT2,
         'IF(AND({C}>=0.03,{C}<=0.15),"PASS","REVIEW")'),
        ("WACC exceeds terminal growth rate g (required for Gordon Growth)",
         f"={WACC_CELL}-DCF!$B${D['g']}", FMT_PCT2, 'IF({C}>0,"PASS","FAIL")'),
        ("Cost of equity exceeds after-tax cost of debt (equity is the riskier claim)",
         f"=WACC!$B${W['coe']}-WACC!$B${W['atkd']}", FMT_PCT2, 'IF({C}>0,"PASS","FAIL")'),
        ("Capital structure weights (equity + debt) sum to 100%",
         f"=WACC!$B${W['we']}+WACC!$B${W['wd']}", FMT_PCT2, 'IF(ABS({C}-1)<0.0001,"PASS","FAIL")'),
        ("Discount factors decline monotonically from Year 1 to Year 5",
         "=IF(AND(" + ",".join(f"{a}>{b}" for a, b in zip(dfs, dfs[1:])) + "),1,0)", FMT_INT,
         'IF({C}=1,"PASS","FAIL")'),
        ("Forecast revenue is positive in all five forecast years",
         f"=SUMPRODUCT(--(Forecast!B{F['rev']}:{FL}{F['rev']}>0))", FMT_INT, f'IF({{C}}={n_fcst},"PASS","FAIL")'),
        (f"Year-1 revenue growth ≤ {MAX_MECHANICAL_GROWTH:.0%} (else use guidance/consensus)",
         f"=Forecast!$B${F['g']}", FMT_PCT1,
         f'IF({{C}}<={MAX_MECHANICAL_GROWTH},"PASS","REVIEW")'),
        ("All five forecast-year FCFF values are positive",
         f"=SUMPRODUCT(--(Forecast!B{FCFF_ROW}:{FL}{FCFF_ROW}>0))", FMT_INT, f'IF({{C}}={n_fcst},"PASS","REVIEW")'),
        ("Terminal value is 50%–90% of enterprise value (rule of thumb)", f"=DCF!$B${D['tvpct']}", FMT_PCT1,
         'IF(AND({C}>=0.5,{C}<=0.9),"PASS","REVIEW")'),
        (f"{last_fy} EBIT ties to reported operating income (within 1% of revenue)",
         f'=IF(AND(ISNUMBER({hL("ebit")}),ISNUMBER({hL("Operating Income")})),'
         f'{hL("ebit")}-{hL("Operating Income")},"N/A")', FMT_MM,
         f'IF(ISNUMBER({{C}}),IF(ABS({{C}})<=0.01*{hL("Revenue")},"PASS","REVIEW"),"INFO")'),
        ("Equity value is positive", f"=DCF!$B${D['eq']}", FMT_MM, 'IF({C}>0,"PASS","FAIL")'),
        ("Implied share price and diluted share count are both positive",
         f"=IF(AND(DCF!$B${D['px']}>0,DCF!$B${D['sh']}>0),1,0)", FMT_INT, 'IF({C}=1,"PASS","FAIL")'),
        ("Historical data completeness (N/A cells on the Historicals tab)",
         f'=COUNTIF(Historicals!B7:{L}{hist_last_row},"N/A")', FMT_INT, 'IF({C}=0,"PASS","REVIEW")'),
        ("No formula errors across Historicals, WACC, Forecast and DCF",
         f"=SUMPRODUCT(--ISERROR(Historicals!B7:{L}{hist_last_row}))"
         f"+SUMPRODUCT(--ISERROR(WACC!B5:B{wacc_last_row}))"
         f"+SUMPRODUCT(--ISERROR(Forecast!B6:{FL}{fcst_last_row}))"
         f"+SUMPRODUCT(--ISERROR(DCF!B6:{FL}{dcf_last_row}))", FMT_INT, 'IF({C}=0,"PASS","FAIL")'),
    ]
    for i, (label, result, fmt, status) in enumerate(checks, start=1):
        rr = 4 + i
        _put(ws, rr, 1, i, fmt=FMT_INT, h="center")
        _put(ws, rr, 2, label)
        _put(ws, rr, 3, result, color=GREEN, fmt=fmt, h="right")
        _put(ws, rr, 4, "=IFERROR(" + status.replace("{C}", f"C{rr}") + ',"FAIL")', bold=True, h="center")
    first_chk, last_chk = 5, 4 + len(checks)
    status_row = last_chk + 2
    _put(ws, status_row, 1, "Overall model status:", bold=True)
    _put(ws, status_row, 3,
         f'=IF(COUNTIF(D{first_chk}:D{last_chk},"FAIL")>0,"REVIEW NEEDED",'
         f'IF(COUNTIF(D{first_chk}:D{last_chk},"REVIEW")>0,"PASS — SEE REVIEW ITEMS","ALL CHECKS PASS"))',
         color=None, size=11, bold=True, h="right")
    ws.cell(row=status_row, column=3).font = Font(name=FONT_NAME, size=11, bold=True)
    _notes(ws, status_row + 2, [
        "This tab is fully formula-driven: every \"Result\" and \"Status\" cell recalculates automatically if "
        "any input elsewhere in the workbook changes.",
        "REVIEW is a rule-of-thumb flag, not an error: it marks a result that warrants a second look at the "
        "forecast or WACC/g assumptions (e.g. WACC outside 3%–15%, Year-1 growth above "
        f"{MAX_MECHANICAL_GROWTH:.0%}, or terminal value outside 50%–90% of EV). FAIL marks a result that is "
        "mathematically or logically wrong.",
        "The EBIT tie-out compares this model's EBIT definition to reported operating income for the latest "
        "year; a gap usually means one-time operating charges that the forecast does not carry forward.",
    ], first_col="B", last_col="F", width_units=62 + 16 + 12 + 2 + 60)
    _finish(ws)

    # =======================================================================
    # COVER (written last so it can link to cells on every other tab)
    # =======================================================================
    ws = ws_cover
    _setup(ws, {"A": 34, "B": 20, "C": 4, "D": 46, "E": 60})
    ws.merge_cells("A1:E1")
    ws.merge_cells("A2:E2")
    c1 = _put(ws, 1, 1, f"{name.upper()} ({ticker})", color=WHITE, size=20, bold=True,
              fill=PatternFill("solid", fgColor=NAVY))
    c1.alignment = Alignment(horizontal="left", vertical="center")
    _put(ws, 2, 1, "Discounted Cash Flow Valuation — Unlevered Free Cash Flow (FCFF) Approach", color=WHITE,
         size=12, italic=True, fill=PatternFill("solid", fgColor=NAVY_2))
    ws.row_dimensions[1].height = 33.75
    ws.row_dimensions[2].height = 19.5

    info = [
        ("Model date:", _long_date(model_date)),
        ("Valuation (discounting) date:", val_date),
        ("Ticker / SEC CIK:", f"{ticker} / {cik:010d}" if cik else ticker),
        ("Forecast basis:", "Year-1 growth: analyst override" if growth_basis == "override"
         else "Historical-trend rule (auto-generated)"),
    ]
    for i, (label, value) in enumerate(info):
        _put(ws, 4 + i, 1, label, bold=True)
        _put(ws, 4 + i, 2, value, color=BLUE)

    ws.merge_cells("A9:B9")
    _put(ws, 9, 1, "VALUATION SUMMARY", color=WHITE, size=11, bold=True, fill=PatternFill("solid", fgColor=NAVY))
    summary = [
        (10, "Current share price", f"={PX_CELL}", FMT_PRICE, {}),
        (11, "Implied value per share (DCF)", f"=DCF!$B${D['px']}", FMT_PRICE, {"bold": True, "fill": YELLOW_FILL}),
        (12, "Implied upside / (downside)", f"=DCF!$B${D['updown']}", FMT_PCT1, {"bold": True}),
        (14, "Enterprise value ($mm)", f"=DCF!$B${D['ev']}", FMT_MM, {}),
        (15, "Equity value ($mm)", f"=DCF!$B${D['eq']}", FMT_MM, {}),
        (16, "Market capitalization ($mm)", f"={MVE_CELL}", FMT_MM, {}),
        (17, "Diluted shares outstanding (mm)", f"={SH_CELL}", FMT_SHARES, {}),
        (19, "WACC", f"={WACC_CELL}", FMT_PCT2, {}),
        (20, "Terminal growth rate (g)", f"=DCF!$B${D['g']}", FMT_PCT2, {}),
        (21, "Terminal value % of enterprise value", f"=DCF!$B${D['tvpct']}", FMT_PCT1, {}),
    ]
    for rr, label, formula, fmt, style in summary:
        _put(ws, rr, 1, label, bold=style.get("bold", False), fill=style.get("fill"))
        _put(ws, rr, 2, formula, color=GREEN, fmt=fmt, **style)
    _put(ws, 23, 1, "Audit status")
    audit_cell = _put(ws, 23, 2, f"=Audit!$C${status_row}", bold=True)
    audit_cell.font = Font(name=FONT_NAME, size=10, bold=True)

    ws.merge_cells("A26:B26")
    _put(ws, 26, 1, "MODEL CONTENTS", color=WHITE, size=11, bold=True, fill=PatternFill("solid", fgColor=NAVY))
    contents = [
        ("Historicals", f"{hist_fy[0]}–{last_fy} income statement, cash flow and balance sheet items pulled from "
                        f"SEC 10-K XBRL data, with margin ratios and memo items"),
        ("Forecast", f"Historical trend summary, rule-based assumptions, and five-year ({first_fcst}–{last_fcst}) "
                     f"revenue, margin and FCFF build"),
        ("WACC", "CAPM cost of equity, after-tax cost of debt, and market-value capital structure weights"),
        ("DCF", "Discounting of forecast FCFF, Gordon Growth terminal value, and enterprise value to implied "
                "share price bridge"),
        ("Sensitivity", "Implied share price sensitized to WACC and terminal growth rate"),
        ("Audit", "Formula-driven integrity and sanity checks across the full model"),
    ]
    for i, (tab, desc) in enumerate(contents):
        rr = 27 + i
        ws.merge_cells(f"B{rr}:E{rr}")
        _put(ws, rr, 1, tab, bold=True)
        cell = _put(ws, rr, 2, desc, size=9, italic=True, wrap=True)
        cell.font = Font(name=FONT_NAME, size=9, italic=True)
        ws.row_dimensions[rr].height = 24.0

    _put(ws, 34, 1, "Sources:", bold=True)
    sources = [
        f"{name} Form 10-K filings via the SEC EDGAR XBRL company facts API ({cik_txt}); fiscal years "
        f"{hist_fy[0]}–{last_fy}.",
        f"Cost-of-capital inputs, current share price and diluted shares: entered by the user in the DCF Model "
        f"Generator app on {_long_date(model_date)}.",
        "Suggested sources for those inputs: FRED series DGS10 (risk-free rate); Aswath Damodaran, NYU Stern "
        "(equity risk premium, industry betas, cost of debt).",
        "Educational project, not investment advice. Forecast assumptions come from a fixed historical-trend rule "
        "unless overridden; treat the output as a starting point for further research.",
    ]
    for i, text in enumerate(sources):
        rr = 35 + i
        ws.merge_cells(f"A{rr}:E{rr}")
        _put(ws, rr, 1, text, color=NOTE_GRAY, size=8, italic=True, wrap=True)
        ws.row_dimensions[rr].height = 24.0

    _put(ws, 40, 1, "Color legend:", bold=True)
    _put(ws, 41, 1, "Blue = hardcoded input", color=BLUE)
    _put(ws, 42, 1, "Black = formula (same sheet)")
    _put(ws, 43, 1, "Green = link to another sheet", color=GREEN)
    _finish(ws)

    wb.save(output_path)
    return output_path


def _engine_wacc(inp) -> float:
    """Base-case WACC, for centering the sensitivity grid (same formula as dcf_engine.compute_wacc)."""
    ke = inp.risk_free_rate + inp.beta * inp.equity_risk_premium
    kd = inp.pretax_cost_of_debt * (1 - inp.tax_rate)
    cap = inp.market_value_equity + inp.total_debt
    return inp.market_value_equity / cap * ke + inp.total_debt / cap * kd
