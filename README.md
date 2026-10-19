# AI-Augmented Sector Teardown — 20-Company Outlook

**Independent equity research project by Archit Chawla** · As of October 19, 2026

Twenty companies, five sectors, one identical DCF framework — so the comparison is honest.
**Result: 1 BUY, 2 HOLDs, 17 PASSes.** Ranked by median base-case upside: Diversified
Financials (−18.1%) > Fintech (−32.9%) > Enterprise SaaS (−41.7%) > Consumer Tech (−54.5%) >
AI Infra & Semis (−66.2%).

- **Start here:** [`SECTOR_REPORT.md`](SECTOR_REPORT.md) ([PDF](SECTOR_REPORT.pdf)) — methodology,
  sector leaderboard, per-sector notes, full 20-company scorecard, forward outlook
- **Live dashboard:** `site/` (deployed to GitHub Pages) — sector leaderboard, company
  valuations vs. live quotes, every figure dated
- **Per-company packages:** `companies/<TICKER>/` — live-formula DCF workbook
  (Cover, Assumptions, Historicals, DCF_Base/Bull/Bear, Comps, Sensitivity, Checks, Sources),
  one-page investment memo, frozen raw data, rerunnable `pull_data.py` + `build_model.py`

## Method in one paragraph

5-year explicit DCF per company (FY2026–2030, 3-year granular detail), bull/bear cases,
sector comps. WACC from CAPM: 5.1% US 10Y + 4.5% ERP × cross-checked beta; 3.0% terminal
growth. Banks valued on distributable cash flow at the cost of equity (the standard
adaptation — documented per company). "AI-augmented" is the workflow: LLM-assisted data
extraction and drafting, every figure manually verified against primary sources; each
memo logs what was caught. BUY needs 20%+ base-case upside; HOLD is −10% to +20%.

## Reproduce / refresh

```bash
python3 build_site_data.py   # rebuild site/data/companies.json from the 20 summary.json files
python3 refresh_pipeline.py  # daily price refresh (used by the GitHub Action)
```

Per-company: `cd companies/<TICKER> && python3 pull_data.py && python3 build_model.py`

See [`PUSH_TO_GITHUB.md`](PUSH_TO_GITHUB.md) for deploying the dashboard to GitHub Pages
and enabling the daily price refresh.

*Independent project for skill demonstration; not investment advice.*
