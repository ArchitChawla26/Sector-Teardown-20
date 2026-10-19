# Advanced Micro Devices, Inc. (NASDAQ: AMD) — Investment Memo

**By Archit Chawla — October 19, 2026** | Price: $528.75 (Oct 16 close) | **Recommendation: PASS**

## Business snapshot

AMD is the credible challenger across the compute stack: EPYC server CPUs, Ryzen client, and Instinct AI accelerators attacking NVIDIA's datacenter monopoly. FY2025 revenue was $34.64B (+34.3% YoY) with EBITDA of $7.28B — a 21.0% margin, expanding fast on datacenter mix. Q2 2026 revenue hit $11.54B, up 50.1% YoY. The balance sheet is clean: $810M net cash.

## Investment thesis

- **The challenger ramp is real.** +50% YoY quarterly growth, datacenter mix lifting margins from 21% toward 30% — the operating-leverage story is working, and EPYC's share gains are structural.
- **But 141x trailing earnings is not a valuation, it is a verdict.** At $528.75 the market has already priced in the entire challenger arc — flawless Instinct execution against a generational incumbent.
- **My bull case is still 69% light.** Even with a 34% five-year revenue CAGR and 34% EBITDA margins, a 16.1% WACC (beta 2.45 × 5.1% risk-free) discounts the ramp to $163.81. The market needs roughly double my bull-case growth or a much lower rate regime.
- **The challenger gets hit first in a downturn.** My bear case has FY2029 at −15% revenue: when AI capex pauses, customers consolidate orders with the incumbent, not the challenger.
- **Comps confirm the stretch.** Median peer EV/EBITDA (30.7x) implies $180.38 — less than half the quote. AMD's own 89x EV/EBITDA is a growth-phase multiple, not a comp.

## Valuation summary

| Method | Implied value / share |
|---|---|
| DCF — bear case | $32.25 |
| DCF — base case | $101.50 |
| DCF — bull case | $163.81 |
| Comps — EV/EBITDA at median peer 30.7x | $180.38 |
| **Current market price** | **$528.75** |

Base case: FY2026–2030 revenue CAGR of ~27% (tapering from 34.3%), EBITDA margin 22% → 30%, WACC 16.1% (5.1% risk-free + 2.45 beta × 4.5% ERP), 3.0% terminal growth. EV/EBITDA is the operative comps lens — AMD's margin expansion makes EV/Revenue understate the leverage, and its own 141x trailing P/E is not usable.

## Key risks

- **Instinct execution against CUDA.** The moat is software, not silicon; share gains can stall even with competitive hardware.
- **High beta in a high-rate world.** At 2.45 beta, any growth scare reprices the stock violently — the Oct 14–16 drawdown (−12.7% in three sessions) is a preview.
- **Xilinx amortization overhang.** $3.0B of annual D&A (8.7% of revenue) keeps GAAP earnings depressed and complicates the P/E narrative.

## Recommendation: PASS

I like the business and the balance sheet, but every lens I built — DCF base, bull, and peer comps — sits 66–94% below $528.75. This is a story stock at a story price. Revisit on a major drawdown or after the rate regime normalizes.

## Methodology & verification notes

Every figure in the model was verified against the frozen data pulls, and I caught and resolved the following: (1) the FY2021 column in the Yahoo Finance annual pull was entirely blank — excluded, model uses FY2022–FY2025; (2) the Oct 16 close ($528.75) was verified three ways — `previousClose` in the pull, daily history showing 2026-10-16 close = $528.75, and MarketBeat/Finnhub "opened at $528.75"/"Previous Close Price: 528.75" on Oct 19; (3) 2022/2023/2025 tax provisions were benefits (valuation-allowance moves), so I used the 21% statutory rate forward rather than distorted effective rates; (4) beta 2.45 cross-checked across Yahoo (2.454), MarketBeat (2.45), and Finnhub (2.54); (5) Yahoo's market cap ($902.0B) vs price × sharesOutstanding ($863.2B) differ by ~4.5% — I used the info-field market cap for WACC weights and flagged the delta; (6) AMD's own trailing P/E (141x) and EV/EBITDA (89.3x) are growth-phase multiples — the operative comps lens is median peer EV/EBITDA, tagged [JUDGMENT]. The bear case includes an actual −15% revenue down year (FY2029), per the semi-cyclicality requirement. All judgment inputs are tagged [JUDGMENT] with a one-line rationale on the Assumptions sheet.

*Independent project for skill demonstration; not investment advice.*
