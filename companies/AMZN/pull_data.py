"""Raw data pull for the <TICKER> (Consumer Tech) sector teardown. As-of: 2026-10-19.

TICKER is derived from the script's parent directory name, so this same file is
placed in each company's folder and runs for that company only.

Rerun:  python3 pull_data.py   (run from ~/workspace/sector-teardown-20/companies/<TICKER>/)
Outputs raw CSVs + info JSON + 1-mo price history into ./data/ with the as-of stamp.
Requires: pip install yfinance pandas
"""
import json
import time
from pathlib import Path

import pandas as pd
import yfinance as yf

AS_OF = "2026-10-19"
PRICE_ASOF = "2026-10-16"  # Oct 16, 2026 close is the market price (model price_asof)
TICKER = Path(__file__).resolve().parent.name
OUT = Path(__file__).resolve().parent / "data"
OUT.mkdir(parents=True, exist_ok=True)

INFO_FIELDS = [
    "longName", "currentPrice", "previousClose", "marketCap", "enterpriseValue",
    "beta", "sharesOutstanding", "floatShares", "trailingEps", "forwardEps",
    "trailingPE", "forwardPE", "priceToBook", "enterpriseToRevenue",
    "enterpriseToEbitda", "totalRevenue", "ebitda", "totalDebt",
    "bookValue", "revenueGrowth", "earningsGrowth", "sector", "industry",
    "fiftyTwoWeekHigh", "fiftyTwoWeekLow", "currency",
]


def pull_financials(ticker):
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
            print("WARN: %s %s empty" % (ticker, name))
            out[name] = None
            continue
        fn = OUT / ("%s_%s_%s.csv" % (ticker, name, AS_OF))
        df.to_csv(fn)
        out[name] = str(fn)
    # price history for the Oct 16 close
    px = {}
    try:
        hist = t.history(period="1mo")
        if hist is not None and not hist.empty:
            hist.index = pd.to_datetime(hist.index).tz_localize(None)
            fn = OUT / ("%s_history_1mo_%s.csv" % (ticker, AS_OF))
            hist.to_csv(fn)
            out["history_1mo"] = str(fn)
            mask = hist.index.date == pd.Timestamp(PRICE_ASOF).date()
            if mask.any():
                px["oct16_close"] = float(hist.loc[mask, "Close"].iloc[0])
            px["last_close"] = float(hist["Close"].iloc[-1])
            px["last_close_date"] = str(hist.index.date[-1])
    except Exception as e:
        print("WARN: %s history failed: %s" % (ticker, e))
    fn = OUT / ("%s_info_%s.json" % (ticker, AS_OF))
    with open(fn, "w") as f:
        json.dump({"as_of": AS_OF, "ticker": ticker, "info": info,
                   "price_asof": PRICE_ASOF, "price_check": px}, f, indent=2, default=str)
    out["info"] = str(fn)
    print("%s: oct16_close=%s, last=%s on %s" % (ticker, px.get("oct16_close"),
          px.get("last_close"), px.get("last_close_date")))
    return out


def main():
    manifest = {"as_of": AS_OF, "price_asof": PRICE_ASOF, "ticker": TICKER}
    try:
        manifest["data"] = pull_financials(TICKER)
    except Exception as e:
        print("ERROR: %s failed (%s); retrying once after 5s" % (TICKER, e))
        time.sleep(5)
        try:
            manifest["data"] = pull_financials(TICKER)
        except Exception as e2:
            print("FAILED: %s (%s)" % (TICKER, e2))
            manifest["error"] = str(e2)
    with open(OUT / ("manifest_%s.json" % AS_OF), "w") as f:
        json.dump(manifest, f, indent=2, default=str)
    print("Wrote manifest:", OUT / ("manifest_%s.json" % AS_OF))


if __name__ == "__main__":
    main()
