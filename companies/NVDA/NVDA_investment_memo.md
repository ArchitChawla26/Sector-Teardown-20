# NVIDIA Corporation (NASDAQ: NVDA) — Investment Memo

**By Archit Chawla — October 19, 2026** | Price: $214.13 (Oct 16 close) | **Recommendation: PASS**

## Business snapshot

NVIDIA is the defining company of the AI build-out: the dominant supplier of datacenter GPUs and AI networking, with a full-stack platform (CUDA, rack-scale systems) that has made it the toll collector on global AI capex. FY2026 revenue was $215.94B (+65.4% YoY) with $120.07B of net income — a 55.6% net margin and 67.0% EBITDA margin. Q2 FY2027 revenue hit $96.22B, up 105.9% YoY. The board authorized a $150B buyback on September 28.

## Investment thesis

- **The AI infrastructure king.** CUDA's moat, rack-scale platforms, and a ~2-year product cadence keep NVDA a generation ahead; FY2023–2026 revenue went from $27B to $216B — an 8x in three years.
- **Economics are absurd.** 67% EBITDA margins and a 96% return on equity (per MarketBeat) on a fabless model — this is the best business in the group by a wide margin.
- **But the cycle is the cycle.** Semis always digest. My bear case puts a real −10% down year in FY2030, when three years of triple-digit growth run into datacenter digestion — that is what happened to every prior semi supercycle.
- **The price demands a flawless decade.** At a 15.0% WACC (beta 2.22 × 5.1% risk-free), even my bull case — 37% five-year revenue CAGR to a $1T revenue run-rate, margins pinned at ~70% — sits ~26% below the quote.
- **Rate regime dominates.** The sensitivity runs both ways: if the 10Y normalizes toward 4%, my bull case roughly clears the price. At 5.1%, there is no margin of safety.

## Valuation summary

| Method | Implied value / share |
|---|---|
| DCF — bear case | $47.48 |
| DCF — base case | $103.36 |
| DCF — bull case | $157.35 |
| Comps — EV/Revenue at median peer 20.7x | $258.83 |
| **Current market price** | **$214.13** |

Base case: FY2027–2031 revenue CAGR of ~29% (tapering from 65.4%), EBITDA margin 68% → 65%, WACC 15.0% (5.1% risk-free + 2.22 beta × 4.5% ERP), 3.0% terminal growth. Peer EV/EBITDA is distorted by AMD/ALAB growth-phase multiples, so EV/Revenue is the operative comps lens — and even it sits above the DCF range, which tells me the whole group is priced for perfection.

## Key risks

- **Customer concentration and capex cyclicality.** A handful of hyperscalers drive the P&L; any AI capex pause hits revenue immediately and the 2.22 beta amplifies the drawdown.
- **Competition and custom silicon.** Hyperscaler in-house chips and AMD's Instinct line attack the highest-margin sockets over time.
- **China/export controls.** A recurring source of demand volatility outside the company's control.

## Recommendation: PASS

The best business in the sector at a price that needs the bull case just to break even — and my bull case is already heroic. I would need the stock ~30% lower, or the 10Y meaningfully below 5%, to engage.

## Methodology & verification notes

Every figure in the model was verified against the frozen data pulls, and I caught and resolved the following: (1) the FY2022 column in the Yahoo Finance annual pull was entirely blank — excluded, model uses FY2023–FY2026; (2) the Oct 16 close ($214.13) was verified three ways — `previousClose` in the pull, daily history showing 2026-10-16 close = $214.13, and MarketBeat's "opened at $214.13" on Oct 19; (3) Yahoo's market cap ($5.28T) vs price × sharesOutstanding ($5.17T) differ by ~2% — I used the info-field market cap for WACC weights and flagged the delta; (4) total debt jumped from $12.3B (Q1 FY2027) to $38.4B (Q2 FY2027) — verified against the info-field totalDebt ($38.9B), consistent with new issuance, not a data artifact; (5) beta 2.22 cross-checked across Yahoo (2.216) and MarketBeat (2.22); (6) FY2026 effective tax was ~15% — I used the 21% statutory rate forward, which is conservative; (7) peer EV/EBITDA medians are distorted by AMD (89x) and ALAB (170x) growth-phase multiples — EV/Revenue is the operative comps lens, tagged [JUDGMENT]. The bear case includes an actual −10% revenue down year (FY2030), per the semi-cyclicality requirement. All judgment inputs are tagged [JUDGMENT] with a one-line rationale on the Assumptions sheet.

*Independent project for skill demonstration; not investment advice.*
