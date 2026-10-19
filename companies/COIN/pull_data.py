"""Raw data pull for the COIN (Coinbase) equity teardown. As-of: 2026-10-19.

Rerun:  python3 pull_data.py   (run from ~/workspace/sector-teardown-20/companies/COIN/)
Outputs raw CSVs + info JSONs into ./data/ with the as-of stamp, plus a
prices JSON with Oct 16, 2026 closes for the fintech sector comps set.
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

TICKER = "COIN"
SECTOR = ["HOOD", "COIN", "SOFI", "SCHW"]
PRICE_DATE = "2026-10-16"   # Friday; model market price is the Oct 16, 2026 close

INFO_FIELDS = [
    "longName", "currentPrice", "previousClose", "marketCap", "enterpriseValue",
    "beta", "sharesOutstanding", "floatShares", "trailingEps", "forwardEps",
    "trailingPE", "forwardPE", "priceToBook", "enterpriseToRevenue",
    "enterpriseToEbitda", "totalRevenue", "ebitda", "totalDebt",
    "bookValue", "revenueGrowth", "earningsGrowth", "sector", "industry",
    "fiftyTwoWeekHigh", "fiftyTwoWeekLow", "currency",
]


def pull_financials(ticker: str) -> dict:
    t = yf.Ticker(ticker)
    info = {k: t.info.get(k) for k in INFO_FIELDS}
    frames = {
        "income_annual": t.income_stmt,
        "balance_annual": t.balance_sheet,
        "cashflow_annual": t.cashflow,
        "income_quarterly": t.quarterly_income_stmt,
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


def pull_oct16_close(ticker: str):
    """Oct 16, 2026 close, retrying once; raises if unavailable."""
    last = None
    for attempt in (1, 2):
        try:
            t = yf.Ticker(ticker)
            h = t.history(start="2026-10-15", end="2026-10-20", auto_adjust=False)
            if not h.empty:
                row = h.loc[h.index.strftime("%Y-%m-%d") == PRICE_DATE]
                if not row.empty:
                    return float(row["Close"].iloc[0])
            last = RuntimeError(f"no {PRICE_DATE} bar for {ticker}")
        except Exception as e:  # noqa: BLE001
            last = e
        time.sleep(3)
    raise last


def main():
    manifest = {"as_of": AS_OF, "tickers": {}}
    manifest["tickers"][TICKER] = pull_financials(TICKER)
    time.sleep(2)
    for p in SECTOR:
        if p == TICKER:
            continue
        t = yf.Ticker(p)
        info = {k: t.info.get(k) for k in INFO_FIELDS}
        fn = OUT / f"{p}_info_{AS_OF}.json"
        with open(fn, "w") as f:
            json.dump({"as_of": AS_OF, "ticker": p, "info": info}, f, indent=2, default=str)
        manifest["tickers"][p] = {"info": str(fn)}
        time.sleep(2)
    closes = {}
    for p in SECTOR:
        closes[p] = pull_oct16_close(p)
        print(f"{p} Oct 16, 2026 close: {closes[p]:.2f}")
        time.sleep(2)
    fn = OUT / f"prices_oct16close_{AS_OF}.json"
    with open(fn, "w") as f:
        json.dump({"as_of": AS_OF, "price_date": PRICE_DATE, "closes": closes}, f, indent=2)
    manifest["tickers"]["prices_oct16close"] = str(fn)
    with open(OUT / f"manifest_{AS_OF}.json", "w") as f:
        json.dump(manifest, f, indent=2, default=str)
    print("Wrote manifest:", OUT / f"manifest_{AS_OF}.json")
    inc = pd.read_csv(OUT / f"{TICKER}_income_annual_{AS_OF}.csv", index_col=0)
    cf = pd.read_csv(OUT / f"{TICKER}_cashflow_annual_{AS_OF}.csv", index_col=0)
    print("\n=== COIN income statement rows (first 15) ===")
    print(inc.index.tolist()[:15])
    print("=== columns ===")
    print(inc.columns.tolist())


if __name__ == "__main__":
    main()
