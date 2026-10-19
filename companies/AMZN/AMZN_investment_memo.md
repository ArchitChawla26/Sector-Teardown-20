# Amazon.com, Inc. (NASDAQ: AMZN) — Investment Memo

**By Archit Chawla — October 19, 2026** | Price: $250.00 (Oct 16 close) | **Recommendation: PASS**

## Business snapshot

Amazon is two companies in a trench coat. FY2025 revenue was $716.9B (+12.4% YoY) with $80.0B of operating income: AWS did $128.7B of sales (+20%) at a 35.4% operating margin — 18% of revenue but 57% of operating income; North America retail did $426.3B at a 6.9% margin; International did $161.9B at 2.9%. Advertising ($69B, +22%) is the quiet third engine. Q2 2026 sales hit $200.6B (+20% YoY) with AWS contributing $16.6B of the $27.5B in segment operating income. Q3 2026 earnings are due October 29 — this model is pre-earnings.

## Investment thesis

- **AWS is a superb business; the quote prices all of Amazon like AWS.** At 23.9x forward earnings on the consolidated entity, the market applies a cloud multiple to a company where 82% of revenue is retail and ads at structurally lower margins.
- **The AI build is eating the cash.** 2025 capex was $131.8B (18.4% of revenue); free cash flow collapsed to $7.7B from $32.9B in 2024 despite $139.5B of operating cash flow. In Q1 2026 Amazon added $53.4B of long-term debt in a single quarter to fund the build.
- **The most interesting tension in this name: multiples vs. cash.** Median peer forward P/E implies $232.60 — essentially the market price; the multiple lens says fairly valued. My DCF says $79.68. I trust the cash: you cannot pay 24x earnings on profits that the capex cycle keeps converting into data centers.
- **The price implies ~8.8% perpetual FCF growth** (my arithmetic) — the highest hurdle of the four Consumer Tech names, on a company whose FCF was negative in 2022 and $7.7B last year.
- **Even my bull case ($128.28) trails the quote by ~49%.** AWS sustaining 20%+ with margin expansion still can't justify $250 at an 11.1% WACC (beta 1.46 — the highest in the group).

## Valuation summary

| Method | Implied value / share |
|---|---|
| DCF — bear case | $19.05 |
| DCF — base case | $79.68 |
| DCF — bull case | $128.28 |
| Comps — median peer forward P/E (22.3x × $10.44 EPS) | $232.60 |
| **Current market price** | **$250.00** |

Base case: 5-year revenue CAGR of ~11.2% (14% → 9% taper), EBITDA margin 23.1% → 25.5% on AWS/ads mix, WACC 11.1% (5.1% risk-free + 1.46 beta × 4.5% ERP), 3.0% terminal growth, capex held at 17% of revenue. Segment discussion: the model stays consolidated, but the thesis is explicitly two-part — AWS deserves a premium multiple, retail does not, and the quote blends them at the AWS end.

## Key risks

- **AWS competition intensifies.** Azure and GCP are both re-accelerating; a price war compresses the 35% margin that carries the whole valuation.
- **The capex cycle overbuilds.** $130B+ annual capex earning cloud returns is the bull case; earning retail returns is the bear case ($19.05).
- **Consumer slowdown** hits the 82% of revenue that is retail, where margins are already thin.
- **Q3 earnings (Oct 29) re-prices the AWS growth rate** — the single number that matters is AWS YoY, not the consolidated beat.

## Recommendation: PASS

Great assets, wrong price. The DCF math is unambiguous — even the bull case is roughly half the quote — and the comps-implied "fair value" only works if you believe retail deserves a cloud multiple.

## Methodology & verification notes

I pulled filing-derived financials and market data via Yahoo Finance (frozen in `data/`, 2026-10-19) and verified every figure by hand, catching the following: (1) the Oct 16 close printed as exactly $250.00 — verified real via the 1-mo history OHLC (O 255.10 / H 255.54 / L 248.60 / C 250.00) and Yahoo's previousClose of 250.00; (2) Yahoo's quarterly balance sheet for AMZN stops at Q1 2026, so net debt ($66,799M) and book debt ($209,888M) use March 2026 — the $53.4B QoQ long-term-debt jump is confirmed in the quarterly detail (capital leases $90.8B); (3) beta 1.46 verified across three sources (Yahoo 1.456, stockanalysis.com 1.46, Finnhub 1.51); (4) FY2025 segment figures (AWS $128.7B sales / $45.6B op income; ads $69B) are from Amazon's 2025 10-K; Q2 2026 figures ($200.6B sales, AWS $16.6B of $27.5B op income, ads +26%) from the Q2 2026 release. All judgment inputs are tagged [JUDGMENT] with a one-line rationale on the Assumptions sheet.

*Independent project for skill demonstration; not investment advice.*
