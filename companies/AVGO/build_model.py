"""Build AVGO_valuation_model.xlsx for the AVGO sector teardown (as of 2026-10-19).

Rerun:  python3 build_model.py   (from ~/workspace/sector-teardown-20/companies/AVGO/)
Requires: openpyxl (model_lib.py is shared under ~/workspace/sector-teardown-20/)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
import model_lib

BASE = Path(__file__).resolve().parent

CONFIG = {
    "ticker": "AVGO",
    "name": "Broadcom Inc.",
    "exchange": "NASDAQ",
    "out_path": BASE / "AVGO_valuation_model.xlsx",
    "price": 328.75,
    "price_asof_label": "Oct 16, 2026",
    "recommendation": "PASS",
    "one_liner": ("AI networking plus a software cash machine \u2014 but the quote prices the VMware flywheel as "
                  "fully derisked; my bull case is only fair value at today's rate regime."),
    # --- market / WACC inputs ---
    "rf": 0.051, "erp": 0.045, "tax": 0.21,
    "beta": 1.48,
    "beta_note": "Yahoo Finance 1.476 (2026-10-19); cross-checked MarketBeat 1.46, Finnhub 1.49",
    "cod_pre": 0.060,
    "cod_note": "BBB-grade, leveraged balance sheet; through-cycle new-issue estimate",
    "tax_note": "US statutory; FY2025 provision was a \u2013$397M benefit (valuation allowance) \u2014 normalized",
    "mcap": 1620074.6,
    "mcap_note": "Yahoo Finance 2026-10-19 ($1.62T); price \u00d7 shares cross-check: $1.57T",
    "debt_book": 59419.0,
    "debt_note": "Q3 FY2026 balance sheet (2026-07-31)",
    "net_debt": 35444.0,
    "net_debt_note": "Q3 FY2026 (2026-07-31): total debt $59,419M less cash & equivalents $23,975M",
    "shares": 4773.6,
    "shares_note": "Yahoo Finance sharesOutstanding, Oct 2026 (FY2025 diluted weighted avg: 4,853M)",
    # --- history ---
    "hist_years": [2022, 2023, 2024, 2025],
    "hist_label": "FY2022\u2013FY2025 (fiscal year ends October)",
    "hist_note": "FY2021 column blank in the Yahoo Finance pull \u2014 excluded.",
    "hist": {
        "revenue":      [33203.0, 35819.0, 51574.0, 63887.0],
        "gross_profit": [22095.0, 24690.0, 32509.0, 43294.0],
        "ebitda":       [19155.0, 20554.0, 23879.0, 34714.0],
        "op_income":    [14282.0, 16451.0, 14996.0, 26075.0],
        "net_income":   [11495.0, 14082.0, 5895.0, 23126.0],
        "dil_shares":   [4230.0, 4270.0, 4778.0, 4853.0],
        "ocf":          [16736.0, 18085.0, 19962.0, 27537.0],
        "capex":        [-424.0, -452.0, -548.0, -623.0],
        "fcf":          [16312.0, 17633.0, 19414.0, 26914.0],
        "cash":         [12416.0, 14189.0, 9348.0, 16178.0],
        "debt":         [39515.0, 39229.0, 67566.0, 65136.0],
    },
    # --- forecast ---
    "fcast_years": [2026, 2027, 2028, 2029, 2030],
    "fcast_label": "FY2026\u2013FY2030",
    "g_base_note": "Q3 FY2026 revenue +85.5% YoY ($29.59B); taper from FY2025 +23.9% (VMware comp annualizes)",
    "g_bull_note": "AI networking (XPU/Ethernet) plus VMware cross-sell sustains 30%+ growth",
    "g_bear_note": "Semi down-cycle in FY2029 (\u221212%): enterprise IT digestion after the AI build-out",
    "m_base_note": "FY2025 EBITDA margin 54.3%; expand toward 59% on software/VMware mix",
    "m_bull_note": "Faster operating leverage as VMware margins fully annualize",
    "m_bear_note": "Margin pressure from pricing competition and mix shift in the down year",
    "cases": {
        "base": {"g": [0.45, 0.30, 0.20, 0.15, 0.12],
                 "m": [0.55, 0.56, 0.57, 0.58, 0.59],
                 "wacc_adj": 0.0, "tg": 0.030},
        "bull": {"g": [0.55, 0.38, 0.28, 0.22, 0.18],
                 "m": [0.57, 0.59, 0.60, 0.61, 0.62],
                 "wacc_adj": -0.01, "tg": 0.035},
        "bear": {"g": [0.35, 0.18, 0.05, -0.12, 0.08],
                 "m": [0.52, 0.51, 0.50, 0.49, 0.50],
                 "wacc_adj": 0.01, "tg": 0.020},
    },
    "da_pct": 0.12,
    "da_note": "FY2025: 13.7% of revenue (VMware intangibles amortizing); declining",
    "sbc_pct": [0.105, 0.095, 0.085, 0.075, 0.070],
    "sbc_note": "FY2025: 11.8% of revenue; elevated post-VMware, normalizing",
    "capex_pct": 0.010,
    "capex_note": "FY2025: 1.0% of revenue (fabless + asset-light software)",
    "nwc_pct": 0.05,
    # --- comps ---
    "comps_order": ["AVGO", "NVDA", "AMD", "ALAB"],
    "peer_label": "NVDA, AMD, ALAB",
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
    "evrev_note": "[JUDGMENT] Operative comps lens: EV/Revenue at median peer 20.7x \u2192 $378.60. Peer EV/EBITDA is distorted by AMD/ALAB growth-phase multiples.",
    "evolve_note": "Shown for reference; median peer EV/EBITDA (89.3x) is dominated by AMD/ALAB growth-phase multiples \u2014 not the operative lens.",
    "pe_note": "Forward P/E shown for reference; DCF is the primary lens.",
    # --- checks ---
    "check_ni": ("Cash-flow 'Net Income From Continuing Operations' = income-statement net income "
                 "for FY2022\u2013FY2025 (max diff $0.0M)"),
    "check_rev": ("FY2022\u2013FY2025 total revenue 33,203.0 / 35,819.0 / 51,574.0 / 63,887.0 ($M) from Yahoo Finance "
                  "annual income statement; FY2021 column blank in source \u2014 excluded"),
    "check_shares": ("4,773.6M outstanding (Yahoo Finance, Oct 2026); FY2025 diluted weighted avg 4,853M "
                     "(10-K via Yahoo). Delta <2%"),
    "check_netdebt": ("Q3 FY2026 balance sheet (2026-07-31): total debt $59,419M less cash & equivalents "
                      "$23,975M = net debt $35,444M"),
    "check_beta": "1.48 (Yahoo Finance 1.476, 2026-10-19); cross-checked MarketBeat 1.46, Finnhub 1.49",
    "extra_checks": [
        ("VMware amortization distorts GAAP net income",
         "FY2024 net income $5,895M vs EBITDA $23,879M \u2014 VMware intangibles amortization; model values on "
         "EBITDA/FCF, not GAAP NI", "PASS"),
        ("Tax normalization",
         "FY2025 tax provision was \u2013$397M (benefit); model uses 21% statutory forward rate \u2014 flagged as judgment on Assumptions", "PASS"),
        ("Down-cycle in bear case",
         "Bear case includes an actual \u221212% revenue down year (FY2029), not just slower growth \u2014 semi cyclicality requirement", "PASS"),
        ("Dividend not modeled separately",
         "$0.65 quarterly dividend ($2.60 annualized, 0.8% yield) is cash returned within the enterprise framework \u2014 no double count", "PASS"),
        ("DCF vs market gap is explainable",
         "Base per-share trails $328.75 because an 11.5% WACC discounts even 24% 5-yr revenue CAGR hard; only the "
         "bull case (\u223C31% CAGR) reaches fair value \u2014 the market prices flawless execution", "PASS"),
    ],
    # --- sources ---
    "sources": [
        ("Yahoo Finance \u2014 AVGO quote & financials", "https://finance.yahoo.com/quote/AVGO",
         "2026-10-19", "Price, beta, shares, annual/quarterly statements (frozen in data/)"),
        ("MarketBeat \u2014 AVGO stock analysis",
         "https://www.marketbeat.com/stocks/NASDAQ/AVGO/competitors-and-alternatives/",
         "2026-10-19", "Beta (1.46) cross-check; price and earnings context"),
        ("Finnhub \u2014 AVGO financial market data", "https://finnhub.io/?q=%22AVGO%22",
         "2026-10-19", "Beta cross-check (1.49)"),
        ("Reuters \u2014 10Y yield 5.085% on Oct 19, 2026",
         "https://www.reuters.com/business/wall-st-futures-climb-ai-stocks-rebound-oil-yields-fall-2026-10-19/",
         "2026-10-19", "Risk-free rate (shared framework constant)"),
    ],
}

if __name__ == "__main__":
    model_lib.run(CONFIG)
