"""Build ALAB_valuation_model.xlsx for the ALAB sector teardown (as of 2026-10-19).

Rerun:  python3 build_model.py   (from ~/workspace/sector-teardown-20/companies/ALAB/)
Requires: openpyxl (model_lib.py is shared under ~/workspace/sector-teardown-20/)

NOTE: ALAB IPO'd March 2024. Annuals FY2022\u2013FY2025 are S-1/filing-derived; the
public trading history is short and is flagged prominently in the memo and Assumptions.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
import model_lib

BASE = Path(__file__).resolve().parent

CONFIG = {
    "ticker": "ALAB",
    "name": "Astera Labs, Inc.",
    "exchange": "NASDAQ",
    "out_path": BASE / "ALAB_valuation_model.xlsx",
    "price": 285.82,
    "price_asof_label": "Oct 16, 2026",
    "recommendation": "PASS",
    "one_liner": ("The purest AI-connectivity play in the group \u2014 explosive growth, but a 21.75% discount rate "
                  "and a two-year public record mean the quote needs the bull case just to stand still."),
    # --- market / WACC inputs ---
    "rf": 0.051, "erp": 0.045, "tax": 0.21,
    "beta": 3.70,
    "beta_note": "Yahoo Finance 3.69 (2026-10-19); cross-checked MarketBeat 3.68, Finnhub 3.73",
    "cod_pre": 0.060,
    "cod_note": "Speculative-grade proxy; immaterial \u2014 company is net cash (WACC \u2248 cost of equity)",
    "tax_note": "US statutory; company only turned profitable in FY2025 (tax provision \u2013$1M) \u2014 normalized",
    "mcap": 50829.4,
    "mcap_note": "Yahoo Finance 2026-10-19 ($50.8B); price \u00d7 shares cross-check: $49.6B",
    "debt_book": 4.0,
    "debt_note": "FY2025 annual balance sheet (quarterly debt field NaN in Yahoo pull; immaterial either way)",
    "net_debt": -164.0,
    "net_debt_note": "FY2025 annual: cash & equivalents $168M less total debt $4M = net cash $164M",
    "shares": 173.5,
    "shares_note": "Yahoo Finance sharesOutstanding, Oct 2026 (FY2025 diluted weighted avg: 180M)",
    # --- history ---
    "hist_years": [2022, 2023, 2024, 2025],
    "hist_label": "FY2022\u2013FY2025 (fiscal year ends December)",
    "hist_note": ("SHORT HISTORY \u2014 ALAB IPO'd March 2024; FY2022\u2013FY2023 annuals are S-1/filing-derived "
                  "and pre-date the public listing. Public trading history is ~2 years."),
    "hist": {
        "revenue":      [80.0, 116.0, 396.0, 853.0],
        "gross_profit": [59.0, 80.0, 303.0, 645.0],
        "ebitda":       [-59.0, -28.0, -113.0, 180.0],
        "op_income":    [-60.0, -29.0, -116.0, 173.0],
        "net_income":   [-58.0, -26.0, -83.0, 219.0],
        "dil_shares":   [130.0, 153.0, 131.0, 180.0],
        "ocf":          [-36.0, -13.0, 137.0, 319.0],
        "capex":        [-4.0, -3.0, -34.0, -38.0],
        "fcf":          [-40.0, -15.0, 102.0, 282.0],
        "cash":         [76.0, 45.0, 80.0, 168.0],
        "debt":         [2.0, 2.0, 1.0, 4.0],
    },
    # --- forecast ---
    "fcast_years": [2026, 2027, 2028, 2029, 2030],
    "fcast_label": "FY2026\u2013FY2030",
    "g_base_note": "Q2 2026 revenue +104.5% YoY ($392.4M); taper from FY2025 +115.4%",
    "g_bull_note": "Aries/Scorpio/Taurus ramps compound; CXL attach sustains >50% CAGR",
    "g_bear_note": "Semi down-cycle in FY2029 (\u221220%): connectivity orders cut as datacenter digestion hits",
    "m_base_note": "FY2025 EBITDA margin 21.1%; expand to 34% on scale",
    "m_bull_note": "38% EBITDA margin on software-rich connectivity mix",
    "m_bear_note": "Margin compression in the down year as volumes fall",
    "cases": {
        "base": {"g": [0.80, 0.55, 0.40, 0.30, 0.25],
                 "m": [0.24, 0.27, 0.30, 0.32, 0.34],
                 "wacc_adj": 0.0, "tg": 0.030},
        "bull": {"g": [1.00, 0.70, 0.50, 0.38, 0.30],
                 "m": [0.27, 0.31, 0.34, 0.36, 0.38],
                 "wacc_adj": -0.01, "tg": 0.035},
        "bear": {"g": [0.50, 0.30, 0.10, -0.20, 0.15],
                 "m": [0.20, 0.19, 0.17, 0.15, 0.16],
                 "wacc_adj": 0.01, "tg": 0.020},
    },
    "da_pct": 0.01,
    "da_note": "FY2025: 0.8% of revenue (fabless; minimal fixed assets)",
    "sbc_pct": [0.16, 0.14, 0.12, 0.10, 0.08],
    "sbc_note": "FY2025: 18.8% of revenue; declining as revenue scales past SBC",
    "capex_pct": 0.040,
    "capex_note": "FY2025: 4.4% of revenue",
    "nwc_pct": 0.05,
    # --- comps ---
    "comps_order": ["ALAB", "NVDA", "AVGO", "AMD"],
    "peer_label": "NVDA, AVGO, AMD",
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
    "evrev_note": "EV/Revenue (median peer 18.0x \u2192 $125.71) reflects peers' scale; ALAB's 40.3x is a hyper-growth multiple \u2014 not the operative lens.",
    "evolve_note": "[JUDGMENT] Operative comps lens: EV/EBITDA at median peer 30.7x \u2192 $50.62. ALAB's own 169.5x is a hyper-growth multiple, not a comp.",
    "pe_note": "Forward P/E shown for reference; DCF is the primary lens.",
    # --- checks ---
    "check_ni": ("Cash-flow 'Net Income From Continuing Operations' = income-statement net income "
                 "for FY2022\u2013FY2025 (max diff $0.0M)"),
    "check_rev": ("FY2022\u2013FY2025 total revenue 80.0 / 116.0 / 396.0 / 853.0 ($M) from Yahoo Finance "
                  "annual income statement; all years S-1/filing-derived, pre-IPO for FY2022\u2013FY2023"),
    "check_shares": ("173.5M outstanding (Yahoo Finance, Oct 2026); FY2025 diluted weighted avg 180M "
                     "(10-K via Yahoo). Delta <4%"),
    "check_netdebt": ("FY2025 annual: cash & equivalents $168M less total debt $4M = net cash $164M; "
                      "quarterly debt field NaN in Yahoo pull \u2014 immaterial either way"),
    "check_beta": "3.70 (Yahoo Finance 3.69, 2026-10-19); cross-checked MarketBeat 3.68, Finnhub 3.73",
    "extra_checks": [
        ("Short history flagged",
         "ALAB IPO'd March 2024; FY2022\u2013FY2023 annuals are S-1-derived and pre-date the public listing; "
         "only FY2024\u2013FY2025 are full public-company annuals \u2014 flagged on Assumptions and Historicals", "PASS"),
        ("Quarterly debt NaN workaround",
         "Yahoo quarterly balance sheet carries NaN for ALAB debt rows; used FY2025 annual debt $4M vs cash $168M "
         "\u2014 net cash either way; N/A on exact quarterly debt", "PASS"),
        ("Pre-IPO share counts",
         "FY2022\u2013FY2023 diluted share counts (130M/153M) are S-1-derived pre-IPO figures; per-share history "
         "not used in valuation", "PASS"),
        ("Down-cycle in bear case",
         "Bear case includes an actual \u221220% revenue down year (FY2029), not just slower growth \u2014 semi cyclicality requirement", "PASS"),
        ("DCF vs market gap is explainable",
         "Base per-share trails $285.82 because a 21.75% WACC (beta 3.70 \u00d7 5.1% risk-free) discounts even ~45% "
         "5-yr revenue CAGR hard; the market prices a flawless hyper-growth ramp", "PASS"),
    ],
    # --- sources ---
    "sources": [
        ("Yahoo Finance \u2014 ALAB quote & financials", "https://finance.yahoo.com/quote/ALAB",
         "2026-10-19", "Price, beta, shares, annual/quarterly statements (frozen in data/)"),
        ("MarketBeat \u2014 ALAB stock analysis",
         "https://www.marketbeat.com/instant-alerts/filing-astera-labs-inc-alab-shares-purchased-by-assenagon-asset-management-sa-2026-10-19/",
         "2026-10-19", "Beta (3.68) cross-check; price and earnings context"),
        ("Finnhub \u2014 ALAB financial market data", "https://finnhub.io/?q=%22ALAB%22",
         "2026-10-19", "Beta cross-check (3.73)"),
        ("lf0 \u2014 Astera Labs stock analysis", "https://lf0.com/research/astera-labs-stock-analysis-alab/",
         "2026-10-19", "Beta 5Y-monthly (3.78) cross-check; TTM margin context"),
        ("Reuters \u2014 10Y yield 5.085% on Oct 19, 2026",
         "https://www.reuters.com/business/wall-st-futures-climb-ai-stocks-rebound-oil-yields-fall-2026-10-19/",
         "2026-10-19", "Risk-free rate (shared framework constant)"),
    ],
}

if __name__ == "__main__":
    model_lib.run(CONFIG)
