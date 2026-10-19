# AI-Augmented Sector Teardown — Palantir Technologies Inc. (PLTR) | Enterprise SaaS

**Independent equity research project by Archit Chawla** | As of October 19, 2026

A full sell-side-style teardown of Palantir Technologies Inc. (PLTR): a live-formula DCF valuation
(base / bull / bear), peer multiples against the other three Enterprise SaaS names
(CRM, NOW, SNOW, PLTR), and a one-page investment memo.

**Recommendation: PASS** — DCF base $31.79 / bull $50.69 / bear $14.98
vs the $186.03 market price (Oct 16 close), -82.9% base-case upside.

## File map

| File | Contents |
|---|---|
| `PLTR_valuation_model.xlsx` | The model. Sheets: Cover, Assumptions, Historicals, DCF_Base, DCF_Bull, DCF_Bear, Comps, Checks, Sources. Every computed cell is a live Excel formula. |
| `PLTR_investment_memo.md` | One-page investment memo: thesis, valuation summary, risks, recommendation, methodology notes. |
| `data/` | Raw pulls frozen at 2026-10-19: annual income statement, balance sheet, cash flow, quarterly statements, and `info` JSONs for PLTR + peers, plus a manifest. |
| `pull_data.py` | Re-pulls all raw data from Yahoo Finance into `data/`. |
| `build_model.py` | Rebuilds the workbook from the data in `data/` and prints the valuation summary + sanity checks. |
| `summary.json` | Machine-readable valuation summary (price, DCF cases, upside, recommendation). |

## How to rerun

```bash
cd ~/workspace/sector-teardown-20/companies/PLTR
python3 pull_data.py     # refresh raw data (as-of becomes the run date)
python3 build_model.py   # rebuild PLTR_valuation_model.xlsx from data/
```

Requires `pip install yfinance pandas openpyxl`. The shipped `data/` folder is frozen at the
2026-10-19 as-of date; re-running the pull refreshes it to the current date.

## Key assumptions

- **Shared framework:** 5.1% US 10Y risk-free (Oct 19, 2026), 4.5% ERP, 3.0% terminal growth
  (bull 3.5%, bear 2.0%), bull/bear discount rate ∓100 bps, 21% tax.
- **Stock-based comp is added back in the FCF bridge** for all four names and disclosed —
  SaaS SBC is material (SNOW: 34.2% of revenue).
- Every judgment input is tagged `[JUDGMENT]` with a one-line rationale on the Assumptions sheet.

## Verification summary

- Net income ties across income and cash-flow statements (max diffs: CRM $0.0M, NOW $0.0M, SNOW $3.6M, PLTR $9.6M).
- Beta cross-checked across 2+ sources per name; 10Y yield verified for the project at 5.1%.
- Workbook Checks sheet runs live formula tie-outs: FCF bridge and per-share arithmetic — all PASS.
- Comps market data as-of 2026-10-16; full URLs and access dates on the Sources sheet.

*Independent project for skill demonstration; not investment advice.*
