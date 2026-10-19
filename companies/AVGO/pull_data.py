"""Raw data pull for the AVGO sector teardown (AI Infra & Semis). As-of: 2026-10-19.

Rerun:  python3 pull_data.py   (run from ~/workspace/sector-teardown-20/companies/AVGO/)
Outputs raw CSVs + info JSONs into ./data/ with the as-of stamp.
Own annual + quarterly financials are pulled; peers' info fields cover the comps.
Requires: pip install yfinance pandas
"""
import json
import time
from pathlib import Path

import pandas as pd
import yfinance as yf

AS_OF = "2026-10-19"
OUT = Path(__file__).resolve().parent / "data"
OUT.mkdir(parents=True, exist_ok=True)

TICKER = "AVGO"
PEERS = ["NVDA", "AMD", "ALAB"]

INFO_FIELDS = [
    "longName", "currentPrice", "previousClose", "marketCap", "enterpriseValue",
    "beta", "sharesOutstanding", "floatShares", "trailingEps", "forwardEps",
    "trailingPE", "forwardPE", "priceToBook", "enterpriseToRevenue",
    "enterpriseToEbitda", "totalRevenue", "ebitda", "totalDebt",
    "bookValue", "revenueGrowth", "earningsGrowth", "sector", "industry",
    "fiftyTwoWeekHigh", "fiftyTwoWeekLow", "currency",
]


def fetch_ticker(ticker: str, retries: int = 1):
    for attempt in range(retries + 1):
        try:
            t = yf.Ticker(ticker)
            _ = t.info.get("longName")
            return t
        except Exception as e:  # noqa: BLE001
            print(f"WARN: {ticker} info attempt {attempt + 1} failed: {e}")
            if attempt < retries:
                time.sleep(5)
    return None


def pull_info(ticker: str) -> dict:
    t = fetch_ticker(ticker)
    info = {k: (t.info.get(k) if t else None) for k in INFO_FIELDS}
    fn = OUT / f"{ticker}_info_{AS_OF}.json"
    with open(fn, "w") as f:
        json.dump({"as_of": AS_OF, "ticker": ticker, "info": info}, f, indent=2, default=str)
    print("Wrote", fn)
    return info


def pull_financials(ticker: str):
    t = yf.Ticker(ticker)
    frames = {
        "income_annual": t.income_stmt,
        "balance_annual": t.balance_sheet,
        "cashflow_annual": t.cashflow,
        "income_quarterly": t.quarterly_income_stmt,
    }
    for name, df in frames.items():
        if df is None or df.empty:
            print(f"WARN: {ticker} {name} empty")
            continue
        fn = OUT / f"{ticker}_{name}_{AS_OF}.csv"
        df.to_csv(fn)
        print("Wrote", fn)


def main():
    manifest = {"as_of": AS_OF, "ticker": TICKER, "peers": PEERS, "infos": {}}
    all_tickers = [TICKER] + [p for p in PEERS if p != TICKER]
    for i, tk in enumerate(all_tickers):
        manifest["infos"][tk] = pull_info(tk)
        if tk == TICKER:
            pull_financials(tk)
        if i < len(all_tickers) - 1:
            time.sleep(2)
    with open(OUT / f"manifest_{AS_OF}.json", "w") as f:
        json.dump(manifest, f, indent=2, default=str)
    print("Wrote manifest:", OUT / f"manifest_{AS_OF}.json")


if __name__ == "__main__":
    main()
