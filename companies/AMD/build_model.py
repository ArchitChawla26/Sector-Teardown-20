"""Build AMD_valuation_model.xlsx for the AMD sector teardown (as of 2026-10-19).

Rerun:  python3 build_model.py   (from ~/workspace/sector-teardown-20/companies/AMD/)
Requires: openpyxl (model_lib.py is shared under ~/workspace/sector-teardown-20/)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
import model_lib

BASE = Path(__file__).resolve().parent

CONFIG = {
    "ticker": "AMD",
    "name": "Advanced Micro Devices, Inc.",
    "exchange": "NASDAQ",
    "out_path": BASE / "AMD_valuation_model.xlsx",
    "price": 528.75,
    "price_asof_label": "Oct 16, 2026",
    "recommendation": "PASS",
    "one_liner": ("The credible #2 in AI accelerators \u2014 but 141x trailing earnings leaves no room for a stumble, "
                  "and my bull case is still ~69% below the quote."),
    # --- market / WACC inputs ---
    "rf": 0.051, "erp": 0.045, "tax": 0.21,
    "beta": 2.45,
    "beta_note": "Yahoo Finance 2.454 (2026-10-19); cross-checked MarketBeat 2.45, Finnhub 2.54",
    "cod_pre": 0.050,
    "cod_note": "Investment-grade, lightly levered; through-cycle new-issue estimate",
    "tax_note": "US statutory; 2022/2023/2025 provisions were benefits (valuation-allowance moves) \u2014 normalized",
    "mcap": 902049.6,
    "mcap_note": "Yahoo Finance 2026-10-19 ($902.0B); price \u00d7 shares cross-check: $863.2B",
    "debt_book": 4276.0,
    "debt_note": "Q2 2026 balance sheet (2026-06-30)",
    "net_debt": -810.0,
    "net_debt_note": "Q2 2026 (2026-06-30): cash & equivalents $5,086M less total debt $4,276M = net cash $810M",
    "shares": 1632.5,
    "shares_note": "Yahoo Finance sharesOutstanding, Oct 2026 (FY2025 diluted weighted avg: 1,636M)",
    # --- history ---
    "hist_years": [2022, 2023, 2024, 2025],
    "hist_label": "FY2022\u2013FY2025 (fiscal year ends December)",
    "hist_note": "FY2021 column blank in the Yahoo Finance pull \u2014 excluded.",
    "hist": {
        "revenue":      [23601.0, 22680.0, 25785.0, 34639.0],
        "gross_profit": [10603.0, 10460.0, 12725.0, 17152.0],
        "ebitda":       [5534.0, 4051.0, 5145.0, 7275.0],
        "op_income":    [1264.0, 401.0, 2086.0, 3694.0],
        "net_income":   [1320.0, 854.0, 1641.0, 4335.0],
        "dil_shares":   [1571.0, 1625.0, 1637.0, 1636.0],
        "ocf":          [3565.0, 1667.0, 3041.0, 7709.0],
        "capex":        [-450.0, -546.0, -636.0, -974.0],
        "fcf":          [3115.0, 1121.0, 2405.0, 6735.0],
        "cash":         [4835.0, 3933.0, 3787.0, 5539.0],
        "debt":         [2863.0, 3003.0, 2212.0, 3847.0],
    },
    # --- forecast ---
    "fcast_years": [2026, 2027, 2028, 2029, 2030],
    "fcast_label": "FY2026\u2013FY2030",
    "g_base_note": "Q2 2026 revenue +50.1% YoY ($11.54B); taper from FY2025 +34.3%",
    "g_bull_note": "Instinct accelerators take durable datacenter share; EPYC/embedded compounding",
    "g_bear_note": "Semi down-cycle in FY2029 (\u221215%): AI capex pause hits the challenger first",
    "m_base_note": "FY2025 EBITDA margin 21.0%; expand to 30% on datacenter mix",
    "m_bull_note": "Accelerator-rich mix drives 34% EBITDA margin",
    "m_bear_note": "Pricing pressure vs NVDA and under-utilization in the down year",
    "cases": {
        "base": {"g": [0.45, 0.32, 0.25, 0.20, 0.15],
                 "m": [0.22, 0.24, 0.26, 0.28, 0.30],
                 "wacc_adj": 0.0, "tg": 0.030},
        "bull": {"g": [0.55, 0.40, 0.32, 0.26, 0.20],
                 "m": [0.24, 0.27, 0.30, 0.32, 0.34],
                 "wacc_adj": -0.01, "tg": 0.035},
        "bear": {"g": [0.30, 0.15, 0.05, -0.15, 0.10],
                 "m": [0.20, 0.19, 0.18, 0.17, 0.18],
                 "wacc_adj": 0.01, "tg": 0.020},
    },
    "da_pct": 0.08,
    "da_note": "FY2025: 8.7% of revenue (Xilinx intangibles amortizing); declining",
    "sbc_pct": [0.045, 0.043, 0.040, 0.038, 0.035],
    "sbc_note": "FY2025: 4.7% of revenue; declining with scale",
    "capex_pct": 0.030,
    "capex_note": "FY2025: 2.8% of revenue (fabless; lab/test capex)",
    "nwc_pct": 0.05,
    # --- comps ---
    "comps_order": ["AMD", "NVDA", "AVGO", "ALAB"],
    "peer_label": "NVDA, AVGO, ALAB",
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
    "evrev_note": "EV/Revenue (median peer 18.0x \u2192 $456.16) understates AMD's operating leverage as datacenter mix grows \u2014 not the operative lens.",
    "evolve_note": "[JUDGMENT] Operative comps lens: EV/EBITDA at median peer 30.7x \u2192 $180.38. AMD's own 89.3x multiple prices in the growth; peer median is the comp.",
    "pe_note": "Forward P/E shown for reference; DCF is the primary lens.",
    # --- checks ---
    "check_ni": ("Cash-flow 'Net Income From Continuing Operations' = income-statement net income "
                 "for FY2022\u2013FY2025 (max diff $0.0M)"),
    "check_rev": ("FY2022\u2013FY2025 total revenue 23,601.0 / 22,680.0 / 25,785.0 / 34,639.0 ($M) from Yahoo Finance "
                  "annual income statement; FY2021 column blank in source \u2014 excluded"),
    "check_shares": ("1,632.5M outstanding (Yahoo Finance, Oct 2026); FY2025 diluted weighted avg 1,636M "
                     "(10-K via Yahoo). Delta <1%"),
    "check_netdebt": ("Q2 2026 balance sheet (2026-06-30): cash & equivalents $5,086M less total debt "
                      "$4,276M = net cash $810M"),
    "check_beta": "2.45 (Yahoo Finance 2.454, 2026-10-19); cross-checked MarketBeat 2.45, Finnhub 2.54",
    "extra_checks": [
        ("Xilinx amortization distorts GAAP earnings",
         "FY2025 D&A $3,004M (8.7% of revenue) on Xilinx intangibles; model values on EBITDA/FCF, not GAAP NI", "PASS"),
        ("Tax normalization",
         "2022/2023/2025 tax provisions were benefits; model uses 21% statutory forward rate \u2014 flagged as judgment on Assumptions", "PASS"),
        ("Down-cycle in bear case",
         "Bear case includes an actual \u221215% revenue down year (FY2029), not just slower growth \u2014 semi cyclicality requirement", "PASS"),
        ("Comps distortion flagged",
         "AMD's own trailing P/E 141x and EV/EBITDA 89.3x are growth-phase multiples; operative lens is median peer EV/EBITDA", "PASS"),
        ("DCF vs market gap is explainable",
         "Base per-share trails $528.75 because a 16.1% WACC (beta 2.45 \u00d7 5.1% risk-free) discounts even 27% 5-yr "
         "revenue CAGR hard; the market prices a near-flawless challenger ramp", "PASS"),
    ],
    # --- sources ---
    "sources": [
        ("Yahoo Finance \u2014 AMD quote & financials", "https://finance.yahoo.com/quote/AMD",
         "2026-10-19", "Price, beta, shares, annual/quarterly statements (frozen in data/)"),
        ("MarketBeat \u2014 AMD stock analysis",
         "https://www.marketbeat.com/stocks/NASDAQ/AMD/competitors-and-alternatives/",
         "2026-10-19", "Beta (2.45) cross-check; price and earnings context"),
        ("Finnhub \u2014 AMD financial market data", "https://finnhub.io/?q=%22AMD%22",
         "2026-10-19", "Beta cross-check (2.54)"),
        ("Reuters \u2014 10Y yield 5.085% on Oct 19, 2026",
         "https://www.reuters.com/business/wall-st-futures-climb-ai-stocks-rebound-oil-yields-fall-2026-10-19/",
         "2026-10-19", "Risk-free rate (shared framework constant)"),
    ],
}

if __name__ == "__main__":
    model_lib.run(CONFIG)
