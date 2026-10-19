# Astera Labs, Inc. (NASDAQ: ALAB) — Investment Memo

**By Archit Chawla — October 19, 2026** | Price: $285.82 (Oct 16 close) | **Recommendation: PASS**

## Business snapshot

**Short history — read the verification notes.** Astera Labs IPO'd in March 2024; only FY2024–FY2025 are full public-company annuals, and the public trading record is ~2 years. The company sells AI connectivity silicon — Aries PCIe/CXL retimers, Taurus Ethernet modules, Scorpio fabric switches, Leo CXL memory controllers — the plumbing inside every AI rack. FY2025 revenue was $853M (+115.4% YoY) with EBITDA of $180M (21.1% margin), the first profitable year. Q2 2026 revenue hit $392.4M, up 104.5% YoY, with Q3 2026 guided at 1.160–1.210 EPS.

## Investment thesis

- **The purest AI-connectivity play in the group.** Every incremental AI rack needs more retimers, more fabric, more CXL — Astera's content-per-rack grows with cluster scale. The product economics are elite: 75.6% gross margins.
- **Hyper-growth is real but the discount rate is brutal.** At a 21.75% WACC (beta 3.70 × 5.1% risk-free), even a ~45% five-year revenue CAGR with margins expanding to 34% only gets to $35.27. The math is unforgiving for 3.7-beta names in a 5.1% rate world.
- **My bull case still needs the quote to stand still.** A ~56% five-year CAGR to $7.8B revenue and 38% EBITDA margins lands at $57.96 — 80% below the market. The quote is pricing something beyond my bull case.
- **Short history + high beta = maximum fragility.** Two years of public data, a −14.5% five-day drawdown into Oct 16, and a single down-cycle year in my bear case (−20% revenue, FY2029) takes the equity to $9.32.
- **Comps agree.** Median peer EV/EBITDA (30.7x) implies $50.62. ALAB's own 169.5x EV/EBITDA is a hyper-growth multiple, not a comp.

## Valuation summary

| Method | Implied value / share |
|---|---|
| DCF — bear case | $9.32 |
| DCF — base case | $35.27 |
| DCF — bull case | $57.96 |
| Comps — EV/EBITDA at median peer 30.7x | $50.62 |
| **Current market price** | **$285.82** |

Base case: FY2026–2030 revenue CAGR of ~45% (tapering from 115.4%), EBITDA margin 24% → 34%, WACC 21.75% (all-equity; net cash), 3.0% terminal growth. EV/EBITDA is the operative comps lens — ALAB's 75%+ gross margins make EV/Revenue understate the earnings leverage.

## Key risks

- **Customer concentration in AI racks.** A handful of hyperscaler programs drive the revenue; any platform redesign or in-sourcing cuts the retimer content.
- **Cyclicality at 3.7 beta.** Connectivity orders are the first cut in a datacenter digestion year — the bear case shows how fast $285.82 becomes a $9 stock in that world.
- **Short public record.** Two years of audited public-company history means the through-cycle margin structure is still unproven.

## Recommendation: PASS

Explosive growth, elite gross margins, and a price that requires perfection I cannot underwrite with a two-year public record. Every lens — DCF base, bull, comps — sits 80–97% below the quote. Revisit after a real drawdown or a full cycle of public-company results.

## Methodology & verification notes

Every figure in the model was verified against the frozen data pulls, and I caught and resolved the following: (1) **ALAB IPO'd March 2024** — FY2022–FY2023 annuals are S-1/filing-derived and pre-date the public listing; only FY2024–FY2025 are full public-company annuals. The short history is flagged on the Assumptions sheet, the Historicals subtitle, and the Checks sheet; (2) the Yahoo quarterly balance sheet carries NaN for all ALAB debt rows — workaround: used the FY2025 annual (total debt $4M vs cash $168M = $164M net cash); immaterial either way, tagged N/A on exact quarterly debt; (3) pre-IPO diluted share counts (FY2022: 130M, FY2023: 153M) are S-1-derived — per-share history is not used in the valuation; (4) the Oct 16 close ($285.82) was verified three ways — `previousClose` in the pull, daily history showing 2026-10-16 close = $285.82, and MarketBeat/Finnhub "opened at $285.82"/"Previous Close Price: 285.82" on Oct 19; (5) FY2025 was the first profitable year (tax provision −$1M) — I used the 21% statutory rate forward; (6) beta 3.70 cross-checked across Yahoo (3.69), MarketBeat (3.68), Finnhub (3.73), and lf0's 5Y-monthly 3.78; (7) ALAB's own trailing P/E (144x) and EV/EBITDA (169.5x) are hyper-growth multiples — the operative comps lens is median peer EV/EBITDA, tagged [JUDGMENT]. The bear case includes an actual −20% revenue down year (FY2029), per the semi-cyclicality requirement. All judgment inputs are tagged [JUDGMENT] with a one-line rationale on the Assumptions sheet.

*Independent project for skill demonstration; not investment advice.*
