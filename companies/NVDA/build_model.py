"""Build NVDA_valuation_model.xlsx for the NVDA sector teardown (as of 2026-10-19).

Rerun:  python3 build_model.py   (from ~/workspace/sector-teardown-20/companies/NVDA/)
Requires: openpyxl (model_lib.py is shared under ~/workspace/sector-teardown-20/)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
import model_lib

BASE = Path(__file__).resolve().parent

CONFIG = {
    "ticker": "NVDA",
    "name": "NVIDIA Corporation",
    "exchange": "NASDAQ",
    "out_path": BASE / "NVDA_valuation_model.xlsx",
    "price": 214.13,
    "price_asof_label": "Oct 16, 2026",
    "recommendation": "PASS",
    "one_liner": ("The AI infrastructure king at a price that demands a flawless decade: even the bull case "
                  "sits ~26% below the quote in a 5.1% rate world."),
    # --- market / WACC inputs ---
    "rf": 0.051, "erp": 0.045, "tax": 0.21,
    "beta": 2.22,
    "beta_note": "Yahoo Finance 2.216 (2026-10-19); cross-checked MarketBeat 2.22",
    "cod_pre": 0.045,
    "cod_note": "AA-grade credit; through-cycle new-issue estimate (new debt issued in H1 FY2027)",
    "tax_note": "US statutory; FY2026 effective rate was ~15% (Singapore/FDII mix) \u2014 model is conservative",
    "mcap": 5278775.4,
    "mcap_note": "Yahoo Finance 2026-10-19 ($5.28T); price \u00d7 shares cross-check: $5.17T",
    "debt_book": 38351.0,
    "debt_note": "Q2 FY2027 balance sheet (2026-07-31); +$26B vs Q1 on new issuance \u2014 verified vs info totalDebt $38.9B",
    "net_debt": 15908.0,
    "net_debt_note": "Q2 FY2027 (2026-07-31): total debt $38,351M less cash & equivalents $22,443M",
    "shares": 24147.0,
    "shares_note": "Yahoo Finance sharesOutstanding, Oct 2026 (FY2026 diluted weighted avg: 24,514M)",
    # --- history ---
    "hist_years": [2023, 2024, 2025, 2026],
    "hist_label": "FY2023\u2013FY2026 (fiscal year ends January)",
    "hist_note": "FY2022 column blank in the Yahoo Finance pull \u2014 excluded.",
    "hist": {
        "revenue":      [26974.0, 60922.0, 130497.0, 215938.0],
        "gross_profit": [15356.0, 44301.0, 97858.0, 153463.0],
        "ebitda":       [5986.0, 35583.0, 86137.0, 144552.0],
        "op_income":    [5577.0, 32972.0, 81453.0, 130387.0],
        "net_income":   [4368.0, 29760.0, 72880.0, 120067.0],
        "dil_shares":   [25070.0, 24940.0, 24804.0, 24514.0],
        "ocf":          [5641.0, 28090.0, 64089.0, 102718.0],
        "capex":        [-1833.0, -1069.0, -3236.0, -6042.0],
        "fcf":          [3808.0, 27021.0, 60853.0, 96676.0],
        "cash":         [3389.0, 7280.0, 8589.0, 10605.0],
        "debt":         [12031.0, 11056.0, 9982.0, 11040.0],
    },
    # --- forecast ---
    "fcast_years": [2027, 2028, 2029, 2030, 2031],
    "fcast_label": "FY2027\u2013FY2031 (NVDA fiscal year ends in January)",
    "g_base_note": "Q2 FY2027 revenue +105.9% YoY ($96.2B); taper from FY2026 +65.4%",
    "g_bull_note": "Sustained AI datacenter supercycle; rack-scale platforms deepen attach",
    "g_bear_note": "Semi down-cycle in FY2030 (\u221210%): datacenter digestion after years of >100% growth",
    "m_base_note": "FY2026 EBITDA margin 67.0%; mild normalization as mix broadens",
    "m_bull_note": "Faster operating leverage on networking/software mix",
    "m_bear_note": "Pricing pressure and under-utilization in the down year",
    "cases": {
        "base": {"g": [0.45, 0.35, 0.28, 0.22, 0.18],
                 "m": [0.68, 0.68, 0.67, 0.66, 0.65],
                 "wacc_adj": 0.0, "tg": 0.030},
        "bull": {"g": [0.55, 0.42, 0.35, 0.30, 0.25],
                 "m": [0.70, 0.70, 0.69, 0.68, 0.67],
                 "wacc_adj": -0.01, "tg": 0.035},
        "bear": {"g": [0.35, 0.25, 0.12, -0.10, 0.15],
                 "m": [0.64, 0.62, 0.58, 0.55, 0.56],
                 "wacc_adj": 0.01, "tg": 0.020},
    },
    "da_pct": 0.015,
    "da_note": "FY2026: 1.3% of revenue",
    "sbc_pct": [0.025, 0.023, 0.020, 0.018, 0.015],
    "sbc_note": "FY2026: 3.0% of revenue; declining with scale",
    "capex_pct": 0.030,
    "capex_note": "FY2026: 2.8% of revenue (fabless model; test/infra capex)",
    "nwc_pct": 0.05,
    # --- comps ---
    "comps_order": ["NVDA", "AVGO", "AMD", "ALAB"],
    "peer_label": "AVGO, AMD, ALAB",
    "comps_data": {
        "NVDA": {"price": 214.13, "mcap": 5278775.4, "ev": 5136924.0, "rev": 302970.0,
                 "ebitda": 201266.0, "fwd_eps": 15.9146, "trail_pe": 27.64},
        "AVGO": {"price": 328.75, "mcap": 1620074.6, "ev": 1604774.8, "rev": 89104.0,
                 "ebitda": 52260.0, "fwd_eps": 19.3970, "trail_pe": 43.34},
        "AMD":  {"price": 528.75, "mcap": 902049.6, "ev": 854336.2, "rev": 41305.0,
                 "ebitda": 9562.0, "fwd_eps": 15.7621, "trail_pe": 140.96},
        "ALAB": {"price": 285.82, "mcap": 50829.4, "ev": 48377.0, "rev": 1201.9,
                 "ebitda": 285.4, "fwd_eps": 6.4031, "trail_pe": 144.33},
    },
    "evrev_note": "[JUDGMENT] Operative comps lens: EV/Revenue at median peer 20.7x \u2192 $258.83. Peer EV/EBITDA is distorted by AMD/ALAB growth-phase multiples.",
    "evolve_note": "Shown for reference; median peer EV/EBITDA (89.3x) is dominated by AMD/ALAB growth-phase multiples \u2014 not the operative lens.",
    "pe_note": "Forward P/E shown for reference; DCF is the primary lens.",
    # --- checks ---
    "check_ni": ("Cash-flow 'Net Income From Continuing Operations' = income-statement net income "
                 "for FY2023\u2013FY2026 (max diff $0.0M)"),
    "check_rev": ("FY2023\u2013FY2026 total revenue 26,974.0 / 60,922.0 / 130,497.0 / 215,938.0 ($M) from Yahoo Finance "
                  "annual income statement; FY2022 column blank in source \u2014 excluded"),
    "check_shares": ("24,147M outstanding (Yahoo Finance, Oct 2026); FY2026 diluted weighted avg 24,514M "
                     "(10-K via Yahoo). Delta <2%"),
    "check_netdebt": ("Q2 FY2027 balance sheet (2026-07-31): total debt $38,351M less cash & equivalents "
                      "$22,443M = net debt $15,908M"),
    "check_beta": "2.22 (Yahoo Finance 2.216, 2026-10-19); cross-checked MarketBeat 2.22",
    "extra_checks": [
        ("Debt jump verified",
         "Q2 FY2027 total debt $38,351M vs $12,348M in Q1 FY2027 \u2014 verified against info-field totalDebt "
         "$38.9B; consistent with new issuance in H1 FY2027, not a data artifact", "PASS"),
        ("Tax normalization",
         "FY2026 effective tax ~15%; model uses 21% statutory forward rate \u2014 conservative, flagged as judgment on Assumptions", "PASS"),
        ("Down-cycle in bear case",
         "Bear case includes an actual \u221210% revenue down year (FY2030), not just slower growth \u2014 semi cyclicality requirement", "PASS"),
        ("DCF vs market gap is explainable",
         "Base per-share trails $214.13 because a 15.0% WACC (beta 2.22 \u00d7 5.1% risk-free) discounts even 29% 5-yr "
         "revenue CAGR hard; the market price implies a ~10% WACC or flawless execution", "PASS"),
    ],
    # --- sources ---
    "sources": [
        ("Yahoo Finance \u2014 NVDA quote & financials", "https://finance.yahoo.com/quote/NVDA",
         "2026-10-19", "Price, beta, shares, annual/quarterly statements (frozen in data/)"),
        ("MarketBeat \u2014 NVDA stock analysis", "https://www.marketbeat.com/stocks/NASDAQ/NVDA/?Vote=1",
         "2026-10-19", "Beta (2.22) and shares cross-check"),
        ("Reuters \u2014 10Y yield 5.085% on Oct 19, 2026",
         "https://www.reuters.com/business/wall-st-futures-climb-ai-stocks-rebound-oil-yields-fall-2026-10-19/",
         "2026-10-19", "Risk-free rate (shared framework constant)"),
    ],
}

if __name__ == "__main__":
    model_lib.run(CONFIG)
