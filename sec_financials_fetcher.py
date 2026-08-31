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
       value ever reported under that tag, then keep only the full-year
       ("FY") values that came from annual reports ("10-K").
    4. Different companies sometimes use slightly different tag names for
       the same concept (e.g. some tag revenue as "Revenues", others as
       "RevenueFromContractWithCustomerExcludingAssessedTax"). We handle
       this with a fallback list per line item -- we try the first tag,
       and if the company never used it, we try the next one.
    5. Everything gets assembled into one clean table: line items as rows,
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
# ---------------------------------------------------------------------------
LINE_ITEMS: dict[str, list[str]] = {
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
    ],
    "Capital Expenditures": [
        "PaymentsToAcquirePropertyPlantAndEquipment",
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


def annual_values_from_series(raw_series: list[dict], years: int) -> dict[int, float]:
    """
    Filter a raw SEC concept series down to clean, one-value-per-fiscal-year
    annual figures, keeping only full-year ("FY") values reported in 10-Ks.

    Companies sometimes restate prior years in a later filing -- when that
    happens we keep the most recently FILED value for a given fiscal year,
    since that's the most up-to-date figure.
    """
    by_fiscal_year: dict[int, dict] = {}

    for entry in raw_series:
        if entry.get("form") != "10-K":
            continue
        if entry.get("fp") != "FY":
            continue

        fy = entry.get("fy")
        if fy is None:
            continue

        existing = by_fiscal_year.get(fy)
        if existing is None or entry["filed"] > existing["filed"]:
            by_fiscal_year[fy] = entry

    # Keep only the most recent `years` fiscal years
    sorted_years = sorted(by_fiscal_year.keys())[-years:]
    return {fy: by_fiscal_year[fy]["val"] for fy in sorted_years}


def build_financial_table(ticker: str, years: int = 5) -> pd.DataFrame:
    """
    Orchestrates the full pull: ticker -> CIK -> every line item -> one
    clean DataFrame (rows = line items, columns = fiscal years).
    """
    print(f"Looking up CIK for {ticker}...")
    cik = get_cik_for_ticker(ticker)
    print(f"  -> CIK {cik}")

    rows: dict[str, dict[int, float]] = {}

    for line_item, tag_candidates in LINE_ITEMS.items():
        print(f"Fetching '{line_item}'...")
        found = False

        for tag in tag_candidates:
            raw_series = fetch_concept_series(cik, tag)
            if raw_series:
                rows[line_item] = annual_values_from_series(raw_series, years)
                print(f"  -> using tag '{tag}' ({len(rows[line_item])} years found)")
                found = True
                break

        if not found:
            print(f"  -> WARNING: no data found for '{line_item}' under any known tag")
            rows[line_item] = {}

    df = pd.DataFrame(rows).T  # line items as rows, fiscal years as columns
    df = df.sort_index(axis=1)  # oldest year first, left to right

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
