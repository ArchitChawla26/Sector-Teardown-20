# AI-Augmented Sector Teardown: 20-Company Outlook

**By Archit Chawla — October 19, 2026**

*Prices as of October 16, 2026 close · Models as of October 19, 2026 · US 10Y at 5.1%*

---

## Executive summary

I ran the same teardown framework over 20 companies in 5 sectors — a 5-year DCF with bull and bear cases, sector comps, and a BUY/HOLD/PASS call on each — so the ranking below is apples to apples. The result is unambiguous:

- **1 BUY, 2 HOLDs, 17 PASSes.** This is not a cheap market. At a 5.1% risk-free rate, the math says most of what I cover is priced for perfection.
- **Charles Schwab (SCHW) is the single BUY** (+75.2% base-case upside): a 0.78 beta against demonstrated 48% through-cycle pretax margins. The market is pricing the rate-cut risk; my framework says the compensation is more than adequate.
- **JPMorgan (−7.7%) and Salesforce (−7.1%) are HOLDs** — fairly priced inside my bands, with identifiable drags (net interest normalization for JPM, a $33.4B net-debt load for CRM).
- **Ranked by median base-case upside, Diversified Financials screens best (−18.1%), then Fintech (−32.9%), Enterprise SaaS (−41.7%), Consumer Tech (−54.5%), and AI Infra & Semis last (−66.2%).** The ranking is relative — every sector median is negative. Read it as "least overvalued," not "cheap."

The through-line: elevated rates punish long-duration growth cash flows mechanically. AI hardware and hyperscalers carry the highest betas and the heaviest capex, so they screen worst under this framework. That is a statement about price versus cash math, not about the quality of the businesses — several of these are superb companies I would happily own 30–40% lower.

## Methodology

**One framework, twenty companies.** Each name gets: 5-year explicit DCF (FY2026–2030, 3-year granular detail), bull and bear cases, and a comps cross-check against its three sector peers. WACC via CAPM on identical constants: 5.1% US 10-year risk-free (October 19, 2026), 4.5% equity risk premium, per-company beta cross-checked across multiple sources, 21% tax. Terminal growth 3.0% base (3.5% bull, 2.0% bear); bull/bear discount rates shift ∓100 bps. Every computed cell in every workbook is a live Excel formula; every judgment input is tagged with its rationale.

**The one sanctioned adaptation: banks.** You cannot run an enterprise DCF on a bank. For JPM, GS, MS I discount distributable cash flow (dividends + buybacks) at the cost of equity — the standard approach. SOFI and SCHW use an operating-proxy DCF (pretax income + D&A, with only the equity slice of balance-sheet reinvestment treated as a shareholder cash cost), since loan-book flows distort their cash-flow statements. Both adaptations are documented per company; the comparison across the two is directionally consistent but not identical, and I say so plainly.

**AI-augmented workflow.** I use LLM tooling to accelerate data extraction and drafting, then verify every figure against primary sources myself. Each memo's Methodology & Verification Notes logs what I caught: stale share counts (HOOD, COIN, META, GOOGL), blank history columns in data pulls, distorted tax provisions normalized to statutory, fiscal-year mappings (NVDA, CRM, SNOW report on January fiscal years). Nothing is invented; where data was missing it is marked N/A with the workaround stated.

**Recommendation bands:** BUY needs 20%+ base-case upside with a margin of safety; HOLD is −10% to +20%; PASS is anything worse. The bands are applied mechanically — the calls below follow the math even where they disagree with consensus.

## Sector leaderboard

Ranked by median base-case implied upside (DCF base vs. October 16 close):

| Rank | Sector | Median upside | Companies (upside / call) |
|---|---|---|---|
| 1 | Diversified Financials | −18.1% | JPM −7.7% HOLD · GS −12.0% PASS · MS −24.2% PASS · BLK −49.3% PASS |
| 2 | Fintech | −32.9% | SCHW +75.2% BUY · SOFI −20.1% PASS · COIN −45.7% PASS · HOOD −70.5% PASS |
| 3 | Enterprise SaaS | −41.7% | CRM −7.1% HOLD · NOW −28.2% PASS · SNOW −55.2% PASS · PLTR −82.9% PASS |
| 4 | Consumer Tech | −54.5% | META −42.2% PASS · GOOGL −51.6% PASS · AAPL −57.3% PASS · AMZN −68.1% PASS |
| 5 | AI Infra & Semis | −66.2% | AVGO −39.6% PASS · NVDA −51.7% PASS · AMD −80.8% PASS · ALAB −87.7% PASS |

## Sector notes

**1. Diversified Financials (−18.1% median).** The best of a fully-priced market. Low betas (0.78–1.2 range) mean the 5.1% rate regime hurts these names least in discount-rate terms, and distributions are tangible cash, not promises. JPM is fairly priced; the risk is net interest income normalizing as the Fed eventually eases — my bear cases size that. BLK screens worst in the sector (−49.3%): a superb franchise at a multiple that assumes perpetual double-digit AUM compounding. The sector's edge is arithmetic, not optimism.

**2. Fintech (−32.9% median).** A barbell. SCHW is the project's lone BUY — the market sees rate-cut risk to net interest income (49% of revenue) and I agree it's the key risk, but a 7.8% WACC against 48% through-cycle pretax margins leaves enormous room; my call sits above street consensus and the memo owns that gap explicitly. At the other end, HOOD (−70.5%) is priced as if elevated rates last forever: the quote implies ~40% five-year revenue CAGR or a ~10% WACC. COIN (−45.7%) and SOFI (−20.1%) sit between — real businesses, wrong prices.

**3. Enterprise SaaS (−41.7% median).** CRM (−7.1%, HOLD) is the closest call in the project: the $33.4B net-debt load from 2026 levering is the main drag on an otherwise clean franchise. NOW, SNOW, PLTR are all PASS by wide margins — the market is capitalizing AI-workflow narratives at multiples the cash math can't support. SNOW deserves a methodology footnote: never GAAP-profitable, so it runs on an explicit revenue→FCF-margin ramp (16.1% → 37.9%), stated as judgment. PLTR (−82.9%) is the most extreme quote-versus-cash disconnect in the entire universe.

**4. Consumer Tech (−54.5% median).** All four PASS. My reverse-DCF arithmetic says these quotes imply 6.4–8.8% *perpetual* FCF growth — that is the definition of priced for perfection. The honest tension: peer forward P/E multiples are kinder than my DCFs (META's comps-implied value even clears the quote), but multiples borrow the market's own optimism while the DCF counts cash. Capex is the swing factor — META at 30% and GOOGL at 20% of revenue against nine-figure 2026 guides is what breaks the models. If AI capex converts to revenue faster than I assume, these calls are wrong; the memos say so.

**5. AI Infra & Semis (−66.2% median).** Screens worst, and the bear cases here are real down-cycle years (−10% to −20% revenue), not just slower growth — semis are cyclical and the framework respects that. Even my bull cases (29–56% five-year revenue CAGRs at expanding margins) cannot clear these quotes under 11.5–21.75% WACCs. AVGO's bull case is the only one in the sector at the quote. ALAB (−87.7%) carries an additional flag: IPO'd March 2024, so the history is two annuals plus S-1-derived figures — the framework is identical but the confidence interval is widest here.

## Full scorecard

| Ticker | Company | Sector | Price | DCF base | DCF bull | DCF bear | Upside | Call |
|---|---|---|---|---|---|---|---|---|
| SCHW | Charles Schwab | Fintech | $91.96 | $161.09 | $269.82 | $77.41 | +75.2% | BUY |
| JPM | JPMorgan Chase | Div. Financials | $339.72 | $313.42 | $500.78 | $156.06 | −7.7% | HOLD |
| CRM | Salesforce | Ent. SaaS | $240.70 | $223.72 | $338.58 | $129.66 | −7.1% | HOLD |
| GS | Goldman Sachs | Div. Financials | $881.26 | $775.17 | $1,196.80 | $280.37 | −12.0% | PASS |
| SOFI | SoFi Technologies | Fintech | $14.55 | $11.62 | $17.34 | $5.06 | −20.1% | PASS |
| MS | Morgan Stanley | Div. Financials | $180.57 | $136.93 | $204.19 | $58.21 | −24.2% | PASS |
| NOW | ServiceNow | Ent. SaaS | $141.97 | $101.98 | $172.70 | $52.54 | −28.2% | PASS |
| AVGO | Broadcom | AI Infra & Semis | $328.75 | $198.72 | $334.49 | $83.29 | −39.6% | PASS |
| META | Meta Platforms | Consumer Tech | $709.33 | $409.83 | $652.74 | $181.59 | −42.2% | PASS |
| COIN | Coinbase | Fintech | $173.16 | $93.99 | $156.53 | $41.44 | −45.7% | PASS |
| BLK | BlackRock | Div. Financials | $1,134.38 | $575.01 | $896.74 | $337.01 | −49.3% | PASS |
| GOOGL | Alphabet | Consumer Tech | $340.59 | $164.95 | $243.51 | $87.64 | −51.6% | PASS |
| NVDA | Nvidia | AI Infra & Semis | $214.13 | $103.36 | $157.35 | $47.48 | −51.7% | PASS |
| SNOW | Snowflake | Ent. SaaS | $306.84 | $137.47 | $226.06 | $62.85 | −55.2% | PASS |
| AAPL | Apple | Consumer Tech | $343.46 | $146.75 | $217.80 | $81.97 | −57.3% | PASS |
| AMZN | Amazon | Consumer Tech | $250.00 | $79.68 | $128.28 | $19.05 | −68.1% | PASS |
| HOOD | Robinhood | Fintech | $98.87 | $29.15 | $50.38 | $9.70 | −70.5% | PASS |
| AMD | Adv. Micro Devices | AI Infra & Semis | $528.75 | $101.50 | $163.81 | $32.25 | −80.8% | PASS |
| PLTR | Palantir | Ent. SaaS | $186.03 | $31.79 | $50.69 | $14.98 | −82.9% | PASS |
| ALAB | Astera Labs | AI Infra & Semis | $285.82 | $35.27 | $57.96 | $9.32 | −87.7% | PASS |

## Forward outlook — what would change the calls

I hold these views conditionally; here is what flips them:

- **Rate normalization is the master variable.** If the 10Y moves from 5.1% toward 4%, every WACC in this project falls ~100–150 bps and DCF values rise mechanically — most PASSes narrow sharply and several flip to HOLD. Conversely, a move toward 6% makes even the HOLDs indefensible. Watch the Fed, not the tickers.
- **Hyperscaler capex conversion.** The consumer-tech PASSes rest heavily on elevated capex (META 30%, GOOGL 20% of revenue). If AI infrastructure spend converts to revenue faster than my ramps assume, AMZN/META/GOOGL rerate upward on cash flow, not narrative.
- **The semi cycle.** My bear cases assume a real down-cycle year for NVDA/AVGO/AMD/ALAB. If the AI hardware cycle extends without a digestion year, the bear cases are too punitive — but the base cases still don't clear the quotes.
- **SCHW is the idiosyncratic long.** The BUY does not need rates to fall; it needs Schwab's through-cycle margins to persist. The risk is a faster-than-expected NII compression — sized in the $77.41 bear case.

## Limitations

- One framework cannot be perfectly neutral: high-beta growth will always screen worse than low-beta compounders at high rates. That is a feature of valuation, but it means this report is structurally skeptical of exactly the stocks the market loves most right now.
- Two different bank adaptations (distribution DCF vs. operating-proxy DCF) are both defensible and both disclosed, but SOFI/SCHW are not perfectly comparable to JPM/GS/MS.
- ALAB's two-year history and SNOW's never-profitable P&L widen their confidence intervals beyond what the point estimates suggest.
- Comps are sector-relative (the other three names in each sector), not broad-market — deliberately, so the peer lens shares the framework's assumptions.

*Independent project for skill demonstration; not investment advice. All models, memos, raw data, and rerun scripts are published alongside this report.*
