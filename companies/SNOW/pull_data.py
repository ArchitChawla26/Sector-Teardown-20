"""Raw data pull for the SNOW enterprise-SaaS teardown. As-of: 2026-10-19.

Rerun:  python3 pull_data.py   (run from the SNOW company folder)
Outputs raw CSVs + info JSON into ./data/ with the as-of stamp.
Requires: pip install yfinance pandas
"""
import json
import time
import pandas as pd
import yfinance as yf
from pathlib import Path

AS_OF = "2026-10-19"
OUT = Path(__file__).resolve().parent / "data"
OUT.mkdir(parents=True, exist_ok=True)

TICKER = "SNOW"
PEERS = ['CRM', 'NOW', 'PLTR']

INFO_FIELDS = [
    "longName", "currentPrice", "previousClose", "marketCap", "enterpriseValue",
    "beta", "sharesOutstanding", "floatShares", "trailingEps", "forwardEps",
    "trailingPE", "forwardPE", "priceToBook", "enterpriseToRevenue",
    "enterpriseToEbitda", "totalRevenue", "ebitda", "totalDebt",
    "bookValue", "revenueGrowth", "earningsGrowth", "sector", "industry",
    "fiftyTwoWeekHigh", "fiftyTwoWeekLow", "currency",
]


def pull_financials(ticker: str) -> dict:
    for attempt in (1, 2):
        try:
            t = yf.Ticker(ticker)
            info = {k: t.info.get(k) for k in INFO_FIELDS}
            frames = {
                "income_annual": t.income_stmt,
                "balance_annual": t.balance_sheet,
                "cashflow_annual": t.cashflow,
                "income_quarterly": t.quarterly_income_stmt,
                "balance_quarterly": t.quarterly_balance_sheet,
            }
            out = {}
            for name, df in frames.items():
                if df is None or df.empty:
                    print(f"WARN: {ticker} {name} empty")
                    out[name] = None
                    continue
                fn = OUT / f"{ticker}_{name}_{AS_OF}.csv"
                df.to_csv(fn)
                out[name] = str(fn)
            fn = OUT / f"{ticker}_info_{AS_OF}.json"
            with open(fn, "w") as f:
                json.dump({"as_of": AS_OF, "ticker": ticker, "info": info}, f, indent=2, default=str)
            out["info"] = str(fn)
            return out
        except Exception as e:  # noqa: BLE001
            print(f"ERROR: {ticker} attempt {attempt}: {e}")
            time.sleep(5)
    raise RuntimeError(f"Failed to pull {ticker} after 2 attempts")


def main():
    manifest = {"as_of": AS_OF, "ticker": TICKER, "peers": PEERS, "tickers": {}}
    manifest["tickers"][TICKER] = pull_financials(TICKER)
    for i, p in enumerate(PEERS):
        if i:
            time.sleep(2)  # rate-limit courtesy
        manifest["tickers"][p] = pull_financials(p)
    with open(OUT / f"manifest_{AS_OF}.json", "w") as f:
        json.dump(manifest, f, indent=2, default=str)
    print("Wrote manifest:", OUT / f"manifest_{AS_OF}.json")


if __name__ == "__main__":
    main()
