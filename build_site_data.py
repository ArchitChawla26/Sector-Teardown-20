#!/usr/bin/env python3
"""Aggregate per-company summary.json files into site/data/companies.json.

Run after all 20 company builds are complete:
    python3 build_site_data.py
Recomputes upside_pct from (dcf_base - price) / price as a guard against
stale arithmetic. Refuses to silently ship fewer than 20 companies.
"""
import json
import os

ROOT = os.path.dirname(os.path.abspath(__file__))
COMPANIES_DIR = os.path.join(ROOT, "companies")
OUT = os.path.join(ROOT, "site", "data", "companies.json")

META = {
    "universe_asof": "2026-10-19",
    "price_asof": "2026-10-16",
    "methodology_version": "1.0",
    "risk_free": 0.051,
    "erp": 0.045,
    "terminal_growth": 0.03,
    "author": "Archit Chawla",
    "disclaimer": "Independent project for skill demonstration; not investment advice.",
}

REQUIRED = ["ticker", "name", "sector", "price", "price_asof",
            "dcf_base", "dcf_bull", "dcf_bear",
            "model_asof", "upside_pct", "recommendation"]


def main():
    companies = []
    for ticker in sorted(os.listdir(COMPANIES_DIR)):
        sp = os.path.join(COMPANIES_DIR, ticker, "summary.json")
        if not os.path.isfile(sp):
            print(f"SKIP {ticker}: no summary.json")
            continue
        with open(sp) as f:
            s = json.load(f)
        missing = [k for k in REQUIRED if k not in s]
        if missing:
            raise SystemExit(f"{ticker} summary.json missing fields: {missing}")
        if s["recommendation"] not in ("BUY", "HOLD", "PASS"):
            raise SystemExit(f"{ticker}: bad recommendation {s['recommendation']!r}")
        # Recompute upside as a guard; the stored value must match.
        recomputed = round((s["dcf_base"] - s["price"]) / s["price"] * 100, 1)
        if abs(recomputed - s["upside_pct"]) > 0.15:
            print(f"NOTE {ticker}: stored upside {s['upside_pct']} -> recomputed {recomputed}")
        s["upside_pct"] = recomputed
        companies.append(s)
    if len(companies) != 20:
        print(f"WARNING: expected 20 companies, found {len(companies)}")
    payload = {"meta": META, "companies": companies}
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w") as f:
        json.dump(payload, f, indent=2)
    print(f"WROTE {OUT} with {len(companies)} companies")


if __name__ == "__main__":
    main()
