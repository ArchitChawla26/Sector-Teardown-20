#!/usr/bin/env python3
"""Daily refresh for the GitHub Actions cron.

Pulls the latest close for all 20 tickers via yfinance and updates
price / price_asof / upside_pct in site/data/companies.json.

Deliberately does NOT touch DCF values: fair values only change when the
per-company models are rebuilt on fresh fundamentals (quarterly). This keeps
the dashboard honest — prices move daily, valuations move on earnings.
"""
import json
import os
import sys
import time

try:
    import yfinance as yf
except ImportError:
    sys.exit("yfinance is not installed")

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "site", "data", "companies.json")


def latest_close(ticker):
    hist = yf.Ticker(ticker).history(period="5d")
    if hist.empty:
        return None, None
    px = round(float(hist["Close"].iloc[-1]), 2)
    asof = hist.index[-1].date().isoformat()
    return px, asof


def main():
    with open(DATA) as f:
        payload = json.load(f)
    companies = payload["companies"]
    updated = 0
    for c in companies:
        t = c["ticker"]
        try:
            px, asof = latest_close(t)
            time.sleep(1)  # be polite to the API
        except Exception as e:  # noqa - never let one ticker kill the run
            print(f"SKIP {t}: {e}")
            continue
        if px is None:
            print(f"SKIP {t}: no history returned")
            continue
        c["price"] = px
        c["price_asof"] = asof
        c["upside_pct"] = round((c["dcf_base"] - px) / px * 100, 1)
        updated += 1
        print(f"OK {t}: {px} as of {asof}")
    payload["meta"]["price_asof"] = max(c["price_asof"] for c in companies)
    with open(DATA, "w") as f:
        json.dump(payload, f, indent=2)
    print(f"Updated {updated}/{len(companies)} prices -> {DATA}")


if __name__ == "__main__":
    main()
