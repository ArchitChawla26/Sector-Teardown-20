"""Raw data pull for the GS teardown (Diversified Financials). As-of: 2026-10-19.

Rerun:  python3 pull_data.py   (run from ~/workspace/sector-teardown-20/companies/GS/)
Outputs raw CSVs + info JSON into ./data/ with the as-of stamp, for GS
plus its sector comps.
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

TICKER = "GS"
PEERS = ['JPM', 'MS', 'BLK']

INFO_FIELDS = [
    "longName", "currentPrice", "previousClose", "marketCap", "enterpriseValue",
    "beta", "sharesOutstanding", "floatShares", "trailingEps", "forwardEps",
    "trailingPE", "forwardPE", "priceToBook", "bookValue", "totalRevenue",
    "ebitda", "totalDebt", "revenueGrowth", "earningsGrowth", "sector",
    "industry", "fiftyTwoWeekHigh", "fiftyTwoWeekLow", "currency",
    "dividendYield", "dividendRate", "payoutRatio", "trailingAnnualDividendRate",
    "trailingAnnualDividendYield", "lastDividendValue", "lastDividendDate",
]


def pull_ticker(ticker: str) -> None:
    for attempt in (1, 2):
        try:
            t = yf.Ticker(ticker)
            info = {k: t.info.get(k) for k in INFO_FIELDS}
            frames = {
                "income_annual": t.income_stmt,
                "balance_annual": t.balance_sheet,
                "cashflow_annual": t.cashflow,
                "income_quarterly": t.quarterly_income_stmt,
            }
            for name, df in frames.items():
                if df is None or df.empty:
                    print(f"WARN: {ticker} {name} empty (attempt {attempt})")
                    continue
                df.to_csv(OUT / f"{ticker}_{name}_{AS_OF}.csv")
            hist = t.history(start="2026-10-13", end="2026-10-20", auto_adjust=False)
            if hist is None or hist.empty:
                print(f"WARN: {ticker} history empty (attempt {attempt})")
            else:
                hist.to_csv(OUT / f"{ticker}_history_{AS_OF}.csv")
            with open(OUT / f"{ticker}_info_{AS_OF}.json", "w") as f:
                json.dump({"as_of": AS_OF, "ticker": ticker, "info": info},
                          f, indent=2, default=str)
            print(f"OK: {ticker} pulled")
            return
        except Exception as e:  # noqa: BLE001
            print(f"ERROR: {ticker} attempt {attempt}: {e}")
            time.sleep(5)
    raise RuntimeError(f"Failed to pull {ticker} after 2 attempts")


def main():
    for i, ticker in enumerate([TICKER] + PEERS):
        if i:
            time.sleep(2)  # rate-limit courtesy
        pull_ticker(ticker)
    manifest = {
        "as_of": AS_OF,
        "ticker": TICKER,
        "peers": PEERS,
        "files": sorted(p.name for p in OUT.iterdir()),
    }
    with open(OUT / f"manifest_{AS_OF}.json", "w") as f:
        json.dump(manifest, f, indent=2)
    print("Wrote manifest:", OUT / f"manifest_{AS_OF}.json")
    # Bank cross-check: revenue, net income, distributions (all annual cols)
    inc = pd.read_csv(OUT / f"{TICKER}_income_annual_{AS_OF}.csv", index_col=0)
    cf = pd.read_csv(OUT / f"{TICKER}_cashflow_annual_{AS_OF}.csv", index_col=0)
    for col in inc.columns:
        rev = inc.loc["Total Revenue", col] if "Total Revenue" in inc.index else None
        ni = inc.loc["Net Income Common Stockholders", col] \
            if "Net Income Common Stockholders" in inc.index else None
        div = -cf.loc["Cash Dividends Paid", col] \
            if "Cash Dividends Paid" in cf.index and col in cf.columns else None
        buy = -cf.loc["Repurchase Of Capital Stock", col] \
            if "Repurchase Of Capital Stock" in cf.index and col in cf.columns else None
        print(f"{col}: revenue={rev}, NI_common={ni}, div={div}, buyback={buy}")


if __name__ == "__main__":
    main()
