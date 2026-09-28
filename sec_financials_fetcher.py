"""
sec_financials_fetcher.py
--------------------------------
Phase 1 of the "AI DCF Model Generator" project.

What this does:
    Given a stock ticker (e.g. "JNJ"), this script automatically pulls that
    company's last several years of financials directly from the SEC --
    no manual 10-K hunting required. This is the same underlying data we
    pulled BY HAND for the Johnson & Johnson DCF project, but generalized
    so it works for (almost) any US public company.

How it works (read this before you touch the code -- you should be able
to explain this part in an interview):
    1. SEC publishes a free, public, no-signup-required API called
       "XBRL frames / company concept" data. Every number in a 10-K is
       tagged with a standardized label (e.g. "Revenues", "Assets",
       "NetIncomeLoss") under the US-GAAP taxonomy -- that's what makes
       filings machine-readable instead of just PDF-like text.
    2. We first look up the company's CIK (SEC's internal company ID)
       from its ticker, using SEC's own ticker-to-CIK mapping file.
    3. For each financial line item we care about (revenue, COGS, SG&A,
       etc.), we ask SEC's "companyconcept" endpoint for every historical
       value ever reported under that tag, then keep only values from annual
       reports ("10-K") that actually cover a full year (~365 days). Note
       that SEC's "fp" field describes the FILING, not the fact -- a 10-K can
       contain Q4-only numbers that still say fp="FY", so we check the
       reported start/end dates instead.
    4. Different companies use different tag names for the same concept,
       and many companies SWITCH tags over time (e.g. revenue moved to
       "RevenueFromContractWithCustomerExcludingAssessedTax" when ASC 606
       took effect in 2018; some companies later moved again). So for each
       fiscal year we take the value from the highest-priority tag that has
       one, falling back to the next tag for years the first one doesn't
       cover -- rather than committing to whichever tag we find first.
    5. Revenue defines the fiscal years in the table (its N most recent
       annual periods). Every other line item is lined up to those exact
       years, so every column is the same fiscal year for every row.
       Everything gets assembled into one clean table: line items as rows,
       fiscal years as columns. That table is what Phase 2 (the Excel
       builder) and Phase 3 (the AI assumption generator) will consume.

Design note that matters for the bigger project:
    This script does ZERO forecasting and ZERO judgment calls -- it only
    fetches and organizes real, reported historical numbers. That's
    deliberate. The "AI" part of this project should only ever touch the
    forecast assumptions, never the historical facts or the arithmetic.

Requirements:
    pip install requests pandas

Usage:
    python sec_financials_fetcher.py JNJ
    python sec_financials_fetcher.py AAPL --years 5 --out aapl_financials.csv

IMPORTANT -- before running this:
    SEC requires every script that calls its API to identify itself with
    a real contact email in the User-Agent header (this is their policy,
    not a technical requirement -- they will block you without it).
    Edit SEC_USER_AGENT below to use your own name/email.
"""

import argparse
import sys
import time
from datetime import date
from typing import Optional

import requests
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG -- edit this before running
# ---------------------------------------------------------------------------
SEC_USER_AGENT = "Rohan Vemulamanda - student project rohanvemulamanda@gmail.com"

TICKER_MAP_URL = "https://www.sec.gov/files/company_tickers.json"
CONCEPT_URL_TMPL = "https://data.sec.gov/api/xbrl/companyconcept/CIK{cik:010d}/us-gaap/{tag}.json"

HEADERS = {"User-Agent": SEC_USER_AGENT}

# SEC asks for no more than ~10 requests/second; we go much slower than that
# to be a polite, well-behaved client.
REQUEST_DELAY_SECONDS = 0.3

# ---------------------------------------------------------------------------
# Line items we want, each with a fallback list of US-GAAP tags.
# Order matters: we try the first tag, then fall back to the next if the
# company never reported under that exact tag name.
#
# A fallback entry is normally a single tag name (a string). It can also be a
# COMPOSITE fallback -- a list of tags -- for companies that don't report one
# combined line for this item at all, and instead split it across several
# tags. Microsoft is the case that surfaced this: it has no
# "SellingGeneralAndAdministrativeExpense" tag anywhere in its filings --
# it reports "Selling and marketing" and "General and administrative"
# separately. When a candidate is itself a list, every tag in it is fetched
# and summed together period-by-period, and that combination is only used if
# every one of its sub-tags actually has data.
# ---------------------------------------------------------------------------
LINE_ITEMS: dict[str, list[str | list[str]]] = {
    "Revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
    ],
    "Cost of Goods Sold": [
        "CostOfGoodsAndServicesSold",
        "CostOfRevenue",
        "CostOfGoodsSold",
    ],
    "SG&A Expense": [
        "SellingGeneralAndAdministrativeExpense",
        ["SellingAndMarketingExpense", "GeneralAndAdministrativeExpense"],
    ],
    "R&D Expense": [
        "ResearchAndDevelopmentExpense",
    ],
    "Operating Income": [
        "OperatingIncomeLoss",
    ],
    "Net Income": [
        "NetIncomeLoss",
        "ProfitLoss",
    ],
    "Depreciation & Amortization": [
        "DepreciationDepletionAndAmortization",
        "DepreciationAmortizationAndAccretionNet",
        "DepreciationAndAmortization",
        ["Depreciation", "AmortizationOfIntangibleAssets"],
    ],
    "Capital Expenditures": [
        "PaymentsToAcquirePropertyPlantAndEquipment",
        "PaymentsToAcquireProductiveAssets",  # e.g. NVIDIA: "purchases of property, equipment and intangibles"
        "PaymentsForCapitalImprovements",
    ],
    "Total Assets": [
        "Assets",
    ],
    "Total Current Assets": [
        "AssetsCurrent",
    ],
    "Total Current Liabilities": [
        "LiabilitiesCurrent",
    ],
    "Cash & Equivalents": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    ],
    "Inventory": [
        "InventoryNet",
    ],
    "Accounts Receivable": [
        "AccountsReceivableNetCurrent",
        "ReceivablesNetCurrent",
    ],
    "Accounts Payable": [
        "AccountsPayableCurrent",
        "AccountsPayableTradeCurrent",
    ],
    "Long-Term Debt": [
        "LongTermDebtNoncurrent",
    ],
}

# A handful of items (like Total Assets) are point-in-time "instant" facts
# rather than "duration" facts -- they don't have a "start" date, only "end".
INSTANT_ITEMS = {
    "Total Assets",
    "Total Current Assets",
    "Total Current Liabilities",
    "Cash & Equivalents",
    "Inventory",
    "Accounts Receivable",
    "Accounts Payable",
    "Long-Term Debt",
}


def get_cik_for_ticker(ticker: str) -> int:
    """Look up a company's SEC CIK number from its stock ticker."""
    resp = requests.get(TICKER_MAP_URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()  # dict keyed by row index -> {"cik_str", "ticker", "title"}

    ticker = ticker.upper().strip()
    for row in data.values():
        if row["ticker"] == ticker:
            return int(row["cik_str"])

    raise ValueError(f"Could not find a CIK for ticker '{ticker}'. Is it a US-listed company?")


def fetch_concept_series(cik: int, tag: str) -> Optional[list[dict]]:
    """
    Pull every historical value SEC has for one US-GAAP tag for one company.
    Returns None if the company has never reported under this exact tag.
    """
    url = CONCEPT_URL_TMPL.format(cik=cik, tag=tag)
    resp = requests.get(url, headers=HEADERS, timeout=30)
    time.sleep(REQUEST_DELAY_SECONDS)

    if resp.status_code == 404:
        return None  # this company doesn't use this tag -- try the next fallback
    resp.raise_for_status()

    payload = resp.json()
    return payload.get("units", {}).get("USD", [])


# A 52/53-week fiscal year (NVIDIA, Apple, Costco...) runs 364 or 371 days.
# Anything far outside this band is a quarter, a stub/transition period, or
# a multi-year cumulative figure -- not a fiscal year.
MIN_ANNUAL_DAYS = 340
MAX_ANNUAL_DAYS = 390


def _is_full_year(entry: dict) -> bool:
    start, end = entry.get("start"), entry.get("end")
    if not start or not end:
        return False
    days = (date.fromisoformat(end) - date.fromisoformat(start)).days
    return MIN_ANNUAL_DAYS <= days <= MAX_ANNUAL_DAYS


def annual_values_from_series(raw_series: list[dict], instant: bool) -> dict[str, float]:
    """
    Filter a raw SEC concept series down to clean, one-value-per-fiscal-year
    annual figures from 10-K filings, keyed by period END date.

    Two filters matter here:
      * Duration items (revenue, capex...) must actually span ~one year.
        SEC's "fp" field is the fiscal period of the FILING, not of the
        number -- every fact in a 10-K says fp="FY", including any Q4-only
        figures, so "fp == FY" alone can let a quarterly number through.
      * Instant items (balance sheet) are point-in-time and have no start date.

    Companies sometimes restate prior years in a later filing -- when that
    happens we keep the most recently FILED value for a given period, since
    that's the most up-to-date figure.

    Deliberately NOT using SEC's "fy" integer field to pick or order years:
    "fy" is also the fiscal year of the FILING, so the prior-year comparative
    columns in a 10-K carry the newer filing's year. A period END date has no
    such ambiguity, whatever fiscal calendar the company uses.

    Returns ALL annual periods found; build_financial_table() decides which
    fiscal years to keep so every line item lines up on the same years.
    """
    by_period_end: dict[str, dict] = {}

    for entry in raw_series:
        if entry.get("form") != "10-K":
            continue
        end = entry.get("end")
        if end is None:
            continue
        if instant:
            if entry.get("start"):
                continue
        elif not _is_full_year(entry):
            continue

        existing = by_period_end.get(end)
        if existing is None or entry["filed"] > existing["filed"]:
            by_period_end[end] = entry

    return {end: e["val"] for end, e in by_period_end.items()}


def _candidate_values(cik: int, candidate: str | list[str], instant: bool) -> dict[str, float]:
    """Annual values for one fallback candidate: a single tag, or a composite (list) of tags summed."""
    if isinstance(candidate, str):
        raw = fetch_concept_series(cik, candidate)
        return annual_values_from_series(raw, instant) if raw else {}

    # Composite candidate: only use periods where EVERY sub-tag has data -- a
    # partial sum (e.g. marketing but not G&A) would understate the line item.
    parts = []
    for sub_tag in candidate:
        raw = fetch_concept_series(cik, sub_tag)
        if not raw:
            return {}
        parts.append(annual_values_from_series(raw, instant))
    common_ends = set.intersection(*(set(p) for p in parts))
    return {end: sum(p[end] for p in parts) for end in common_ends}


def fetch_line_item(cik: int, line_item: str, instant: bool,
                    needed_ends: list[str] | None = None) -> dict[str, float]:
    """
    Merge a line item's fallback tags period by period: for each fiscal year,
    use the highest-priority tag that reported it.

    This is the fix for the "$nan" bug. The old logic committed to the FIRST
    tag that existed at all -- even if the company stopped using it years
    ago -- and then kept that tag's 5 most recent periods. For a company that
    switched tags (NVIDIA was the case that surfaced this), Revenue came back
    covering old fiscal years while the balance sheet covered recent ones, so
    the "latest revenue" the forecast grows from was blank, and every number
    downstream became NaN.

    If `needed_ends` is given, stop fetching once all of those periods are
    covered (saves SEC requests for the common case where tag #1 is complete).
    """
    merged: dict[str, float] = {}
    for candidate in LINE_ITEMS[line_item]:
        values = _candidate_values(cik, candidate, instant)
        added = [end for end in values if end not in merged]
        for end in added:
            merged[end] = values[end]
        if added:
            name = candidate if isinstance(candidate, str) else " + ".join(candidate)
            print(f"  -> '{name}' supplied {len(added)} period(s)")
        if needed_ends is not None and all(end in merged for end in needed_ends):
            break
    return merged


def build_financial_table(ticker: str, years: int = 5) -> pd.DataFrame:
    """
    Orchestrates the full pull: ticker -> CIK -> every line item -> one
    clean DataFrame (rows = line items, columns = fiscal years).

    Revenue's N most recent annual periods define the columns; every other
    line item is aligned to exactly those period end dates.
    """
    print(f"Looking up CIK for {ticker}...")
    cik = get_cik_for_ticker(ticker)
    print(f"  -> CIK {cik}")

    print("Fetching 'Revenue'...")
    revenue = fetch_line_item(cik, "Revenue", instant=False)
    if not revenue:
        raise ValueError(f"No annual revenue found in SEC filings for '{ticker}'. "
                         f"Is it a US-listed company that files 10-Ks?")
    fiscal_year_ends = sorted(revenue)[-years:]
    print(f"  -> fiscal years (period end): {', '.join(fiscal_year_ends)}")

    rows: dict[str, dict[str, float]] = {"Revenue": {e: revenue[e] for e in fiscal_year_ends}}

    for line_item in LINE_ITEMS:
        if line_item == "Revenue":
            continue
        print(f"Fetching '{line_item}'...")
        values = fetch_line_item(cik, line_item, instant=line_item in INSTANT_ITEMS,
                                 needed_ends=fiscal_year_ends)
        rows[line_item] = {e: values[e] for e in fiscal_year_ends if e in values}
        missing = [e for e in fiscal_year_ends if e not in values]
        if len(missing) == len(fiscal_year_ends):
            print(f"  -> WARNING: no data found for '{line_item}' under any known tag")
        elif missing:
            print(f"  -> WARNING: '{line_item}' missing for {', '.join(missing)}")

    df = pd.DataFrame(rows).T  # line items as rows, fiscal years as columns
    df = df.reindex(index=list(LINE_ITEMS), columns=fiscal_year_ends).astype(float)  # oldest -> newest

    # SEC reports every dollar figure in RAW dollars (e.g. 88240950000).
    # The rest of this project (dcf_engine.py, excel_builder.py, app.py) is
    # built and labeled in MILLIONS of dollars throughout -- that's the
    # standard convention for a model like this, and it's what every "$mm"
    # label in the Excel output and the app assumes. Convert here, once, at
    # the source, so nothing downstream has to think about units again.
    df = df / 1_000_000

    df.index.name = "Line Item"
    return df


def main():
    parser = argparse.ArgumentParser(description="Pull historical financials for a US public company from SEC.")
    parser.add_argument("ticker", help="Stock ticker, e.g. JNJ")
    parser.add_argument("--years", type=int, default=5, help="Number of fiscal years to pull (default 5)")
    parser.add_argument("--out", default=None, help="Output CSV filename (default: <TICKER>_financials.csv)")
    args = parser.parse_args()

    if "example.com" in SEC_USER_AGENT:
        print("ERROR: Edit SEC_USER_AGENT at the top of this file to use your real name/email before running.")
        sys.exit(1)

    df = build_financial_table(args.ticker, years=args.years)

    out_path = args.out or f"{args.ticker.upper()}_financials.csv"
    df.to_csv(out_path)
    print(f"\nSaved {out_path}")
    print(df)


if __name__ == "__main__":
    main()
