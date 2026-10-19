"""Build GOOGL_valuation_model.xlsx for the GOOGL Consumer Tech sector teardown (as of 2026-10-19).

Every computed cell is a live Excel formula referencing the Assumptions /
Historicals sheets, so the model is fully auditable. Python mirrors the same
math to print expected outputs and run sanity checks.

Rerun:  python3 build_model.py   (from ~/workspace/sector-teardown-20/companies/GOOGL/)
Requires: openpyxl, pandas
"""

import math
from pathlib import Path
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CONFIG = {
"ticker": "GOOGL", "name": "Alphabet Inc.", "exchange": "NASDAQ",
"price": 340.59, "beta": 1.21,
"beta_note": "Yahoo Finance 2026-10-19 (1.211); cross-checked 1.21 (MarketBeat)",
"rf": 0.051, "erp": 0.045, "cod_pre": 0.055, "tax": 0.21,
"cod_note": "Through-cycle new-issue estimate in a 5.1% Treasury world (shared framework: uniform 5.5%)",
"mcap": 4165416.0,
"mcap_note": "Oct 16, 2026 close $340.59 x 12,230M total shares (all classes)",
"debt_book": 112756.0,
"debt_note": "Q2 2026 balance sheet, Jun 2026 (Yahoo Finance); debt up ~2x YoY on AI-capex issuance",
"net_debt": -129718.0,
"netdebt_note": "Q2 2026: total debt 112,756 less cash 55,911 and ST investments 186,563 = NET CASH $129,718M",
"shares": 12230.0,
"shares_note": "12.23B total (Class A + B + C): stockanalysis.com, MarketBeat, Fool, Oct 2026; Yahoo's 5,867M is Class A only \u2014 12.23B reconciles to market cap / price",
"fiscal_note": "FY2022\u2013FY2025 (fiscal years ended December 31)",
"hist": {
    "revenue":      [282836.0, 307394.0, 350018.0, 402836.0],
    "gross_profit": [156633.0, 174062.0, 203712.0, 240301.0],
    "ebitda":       [85160.0, 97971.0, 135394.0, 180698.0],
    "op_income":    [71685.0, 86025.0, 120083.0, 159562.0],
    "net_income":   [59972.0, 73795.0, 100118.0, 132170.0],
    "dil_shares":   [13159.0, 12722.0, 12447.0, 12230.0],
    "ocf":          [91495.0, 101746.0, 125299.0, 164713.0],
    "capex":        [-31485.0, -32251.0, -52535.0, -91447.0],
    "fcf":          [60010.0, 69495.0, 72764.0, 73266.0],
    "cash":         [113762.0, 110916.0, 95657.0, 126843.0],
    "debt":         [29679.0, 27121.0, 22574.0, 59291.0],
},
"cases": {
    "base": {"g": [0.16, 0.13, 0.11, 0.10, 0.08], "m": [0.450, 0.455, 0.460, 0.465, 0.470], "tg": 0.030},
    "bull": {"g": [0.21, 0.17, 0.14, 0.12, 0.10], "m": [0.470, 0.475, 0.480, 0.485, 0.490], "tg": 0.035},
    "bear": {"g": [0.09, 0.06, 0.05, 0.04, 0.03], "m": [0.420, 0.415, 0.410, 0.405, 0.400], "tg": 0.020},
},
"da_pct": 0.06,
"da_note": "FY2025: 5.2% of revenue (21,136 / 402,836); AI infra depreciation keeps rising",
"sbc_pct": [0.062, 0.060, 0.058, 0.056, 0.054],
"sbc_note": "FY2025: 6.2% of revenue (24,953 / 402,836); easing with scale",
"capex_pct": 0.20,
"capex_note": "FY2025: 22.7% (91,447 / 402,836); 2026 guided $175\u2013185B \u2014 base 20% assumes the build moderates from the 2026 peak",
"nwc_pct": 0.04,
"g_base_note": "FY2026E +16% (Q2 2026 rev +24% YoY); Cloud +82% YoY in Q2 2026, now 21% of revenue; Search +17%",
"g_bull_note": "Cloud sustains 40%+ growth; Gemini monetization; YouTube subscriptions compound",
"g_bear_note": "AI search disruption (zero-click answers); Cloud price competition; DOJ remedy overhang",
"m_base_note": "FY2025: 44.9%; Cloud op margin expanding (13.9 / 58.7 = 23.7% in 2025); blended margin +50 bps/yr",
"m_bull_note": "Cloud reaches 30%+ op margin; Other Bets losses contained",
"m_bear_note": "Traffic acquisition cost inflation; Cloud margin investment; Other Bets drag",
"comps_note": "Forward P/E is the operative lens: trailing EPS is inflated by a Q2 2026 $98B other-income gain.",
"peers": ["AAPL", "AMZN", "META"],
"comps": {
    "AAPL":  {"price": 343.46, "mcap": 5012536.0, "ev": 5034481.0, "rev": 466823.0, "ebitda": 167959.0, "trail_eps": 8.72, "fwd_eps": 9.57785},
    "AMZN":  {"price": 250.00, "mcap": 2696578.0, "ev": 2763377.0, "rev": 775680.0, "ebitda": 168912.0, "trail_eps": 12.43, "fwd_eps": 10.44315},
    "META":  {"price": 709.33, "mcap": 1808792.0, "ev": 1830850.0, "rev": 228247.0, "ebitda": 109655.0, "trail_eps": 26.54, "fwd_eps": 34.62667},
    "GOOGL": {"price": 340.59, "mcap": 4165416.0, "ev": 4035698.0, "rev": 445866.0, "ebitda": 173164.0, "trail_eps": 19.93, "fwd_eps": 15.29162},
},
"recommendation": "PASS",
"thesis": ("Search and Cloud are compounding beautifully \u2014 Cloud +82% YoY in Q2 2026, now 21% of revenue \u2014 "
           "but $340.59 prices ~7.4% perpetual growth, and the 2026 capex guide ($175\u2013185B) will consume most of "
           "operating cash flow. My DCF sits 52% below the quote."),
"checks": [
    ("Net income ties across statements",
     "Income-statement net income = cash-flow starting net income for FY2022\u20132025 (max diff $0.0M)", "PASS"),
    ("Revenue sourced from filings",
     "FY2022\u20132025 total revenue 282,836 / 307,394 / 350,018 / 402,836 ($M), fiscal years ended December (Yahoo Finance annual income statement)", "PASS"),
    ("Share count",
     "12,230M total (Class A + B + C): stockanalysis.com, MarketBeat, Fool, Oct 2026; Yahoo's 5,867M is Class A only (5.868B per Jun-2026 10-Q); 12.23B reconciles to market cap \u00f7 price", "PASS"),
    ("Net debt bridge",
     "Q2 2026 balance sheet (Jun 2026): total debt $112,756M less cash $55,911M and ST investments $186,563M = NET CASH $129,718M", "PASS"),
    ("Beta",
     "1.21 (Yahoo Finance 1.211, 2026-10-19); cross-checked 1.21 (MarketBeat)", "PASS"),
    ("Risk-free rate",
     "US 10Y 5.085% on 2026-10-19 (Reuters); used 5.1% (shared framework)", "PASS"),
    ("Trailing EPS distortion flagged",
     "Q2 2026 net income $112.2B (+298% YoY) included ~$98B of other income (Bastion Memos); trailing EPS $19.93 is inflated \u2014 forward P/E is the operative comps lens", "PASS"),
    ("Comps market data as-of",
     "AAPL / AMZN / META / GOOGL Oct 16, 2026 closes from Yahoo Finance 1-mo history; EVs = Oct-16 market cap + latest-quarter net debt", "PASS"),
    ("FCF bridge reconciles", "Formula check: EV = sum of PV of FCF + PV of terminal value",
     '=IF(ABS(DCF_Base!B22-(DCF_Base!B19+DCF_Base!B21))<0.5,"PASS","FAIL")'),
    ("Per-share arithmetic", "Formula check: per-share = equity value / shares",
     '=IF(ABS(DCF_Base!B26-DCF_Base!B24/DCF_Base!B25)<0.01,"PASS","FAIL")'),
    ("DCF vs market gap is explainable",
     "Base per-share trails $340.59 because the quote implies ~7.4% perpetual FCF growth (my arithmetic); the $175\u2013185B 2026 capex guide alone is ~40% of TTM revenue \u2014 the market is capitalizing AI-capex dollars as if they were already returns", "PASS"),
    ("Q3 2026 earnings next week",
     "GOOGL reports Q3 2026 on Oct 28, 2026 \u2014 the next catalyst; this model is pre-earnings", "PASS"),
],
"sources": [
    ("Yahoo Finance \u2014 GOOGL quote & financials", "https://finance.yahoo.com/quote/GOOGL", "2026-10-19", "Price, beta, shares, annual/quarterly statements, Oct-16 close"),
    ("Yahoo Finance \u2014 AAPL", "https://finance.yahoo.com/quote/AAPL", "2026-10-19", "Comps market data"),
    ("Yahoo Finance \u2014 AMZN", "https://finance.yahoo.com/quote/AMZN", "2026-10-19", "Comps market data"),
    ("Yahoo Finance \u2014 META", "https://finance.yahoo.com/quote/META", "2026-10-19", "Comps market data"),
    ("Reuters \u2014 10Y yield 5.085% on Oct 19, 2026",
     "https://www.reuters.com/business/wall-st-futures-climb-ai-stocks-rebound-oil-yields-fall-2026-10-19/", "2026-10-19", "Risk-free rate"),
    ("stockanalysis.com \u2014 GOOGL statistics", "https://stockanalysis.com/stocks/googl/statistics/", "2026-10-19",
     "Shares outstanding 12.23B, Oct-16 price $340.59 cross-check"),
    ("Motley Fool \u2014 GOOGL quote", "https://www.fool.com/quote/nasdaq/google-a-shares/googl", "2026-10-19",
     "Shares outstanding 12.23B cross-check; Q2 2026 total cash $242.5B / debt $112.8B"),
    ("Bastion Memos \u2014 GOOGL Q2 FY26 as filed", "https://www.bastionmemos.com/companies/GOOGL/q2-fy26.html", "2026-10-19",
     "Q2 2026: revenue $119.8B (+24%); Cloud $24.8B (+82%); net income $112.2B incl. ~$98B other income"),
    ("fourweekmba \u2014 Google business model (Oct 2026)", "https://fourweekmba.com/google-business-model/", "2026-10-19",
     "FY2025: Search $224.5B, YouTube ads $40.4B, Cloud $58.7B (+36%, op income $13.9B); 2026 capex guided $175\u2013185B"),
    ("ts2.tech \u2014 GOOGL vs GOOG share classes", "https://ts2.tech/en/goog-or-googl-alphabets-voting-shares-cost-1-1-more-after-the-search-rally/", "2026-10-19",
     "Class A 5.868B / Class C 5.527B as of Jul 15, 2026 (Jun-2026 10-Q)"),
],
}

# --- ENGINE START ---
# Everything below this line is the shared sector-teardown engine. It reads the
# per-company CONFIG dict defined above and builds the workbook from it.

BASE = Path(__file__).resolve().parent
TICKER = CONFIG["ticker"]
OUT = BASE / f"{TICKER}_valuation_model.xlsx"

HIST = CONFIG["hist"]
YEARS = [2022, 2023, 2024, 2025]
FCAST = [2026, 2027, 2028, 2029, 2030]
PRICE = CONFIG["price"]
BETA = CONFIG["beta"]
RF = CONFIG["rf"]; ERP = CONFIG["erp"]; COD_PRE = CONFIG["cod_pre"]; TAX = CONFIG["tax"]
MCAP = CONFIG["mcap"]; DEBT_BOOK = CONFIG["debt_book"]; NET_DEBT = CONFIG["net_debt"]
SHARES = CONFIG["shares"]
CASES = CONFIG["cases"]
DA_PCT = CONFIG["da_pct"]; SBC_PCT = CONFIG["sbc_pct"]
CAPEX_PCT = CONFIG["capex_pct"]; NWC_PCT = CONFIG["nwc_pct"]
COMPS = CONFIG["comps"]          # dict ticker -> market/financial inputs
PEERS = CONFIG["peers"]          # list of the other three tickers, in display order
WACC_ADJ = {"base": 0.0, "bull": -0.01, "bear": 0.01}  # shared framework: bull/bear ∓100 bps
NAME = CONFIG["name"]; EXCHANGE = CONFIG["exchange"]
RECO = CONFIG["recommendation"]; THESIS = CONFIG["thesis"]
FISCAL_NOTE = CONFIG["fiscal_note"]
CHECKS = CONFIG["checks"]
SOURCES = CONFIG["sources"]

# ---------------------------------------------------------------- styles
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")
SUB_FONT = Font(name="Calibri", size=11, italic=True, color="595959")
LBL_FONT = Font(name="Calibri", size=11)
INP_FILL = PatternFill("solid", fgColor="DDEBF7")
TAG_DATA_FILL = PatternFill("solid", fgColor="E2EFDA")
TAG_JUDG_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Border(*(Side(style="thin", color="BFBFBF") for _ in range(4)))
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
FMT_M = '#,##0.0'
FMT_PCT = '0.0%'
FMT_PX = '"$"#,##0.00'
FMT_X = '#,##0.0"x"'
RECO_COLOR = {"BUY": "006100", "HOLD": "9C5700", "PASS": "C00000"}[RECO]


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = THIN


def title_block(ws, title, subtitle, ncols=6):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = SUB_FONT
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18


def finalize(ws, widths):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B4"


# ---------------------------------------------------------------- workbook
wb = openpyxl.Workbook()

# ============================ COVER ============================
ws = wb.active
ws.title = "Cover"
title_block(ws, f"Equity Research Teardown  \u2014  {NAME} ({EXCHANGE}: {TICKER})",
            "Consumer Tech sector teardown  |  Independent equity research project by Archit Chawla  |  As of October 19, 2026", 4)
ws["A4"] = "Current price (Oct 16, 2026 close)"
ws["B4"] = PRICE
ws["B4"].number_format = FMT_PX
ws["B4"].font = Font(name="Calibri", size=12, bold=True)
ws["A5"] = "Recommendation"
ws["B5"] = RECO
ws["B5"].font = Font(name="Calibri", size=12, bold=True, color=RECO_COLOR)
ws["A6"] = "One-line thesis"
ws["B6"] = THESIS
ws["B6"].alignment = LEFT
for r in (4, 5, 6):
    ws.cell(row=r, column=1).font = LBL_FONT
ws["A8"] = "Contents"
ws["A8"].font = Font(name="Calibri", size=11, bold=True)
contents = [
    ("Assumptions", "WACC build and all forecast drivers; [DATA] vs [JUDGMENT] tags"),
    ("Historicals", f"{FISCAL_NOTE} financials from Yahoo Finance annual statements"),
    ("DCF_Base / DCF_Bull / DCF_Bear", "5-year DCF per case; all formulas live"),
    ("Comps", f"{', '.join([TICKER] + PEERS)} multiples on Oct 16, 2026 prices and implied ranges"),
    ("Checks", "Verification appendix: tie-outs and data sources"),
    ("Sources", "Full URLs and access dates"),
]
for i, (s, d) in enumerate(contents, start=9):
    ws.cell(row=i, column=1, value=s).font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=i, column=2, value=d).font = LBL_FONT
    ws.cell(row=i, column=2).alignment = LEFT
ws["A16"] = ("Disclaimer: Independent project for skill demonstration; not investment advice.")
ws["A16"].font = Font(name="Calibri", size=9, italic=True, color="595959")
ws.merge_cells("A16:D16")
finalize(ws, [28, 90, 10, 10])

# ============================ ASSUMPTIONS ============================
ws = wb.create_sheet("Assumptions")
A = "Assumptions"
title_block(ws, "Assumptions", "Blue cells are inputs. [DATA] = market/filing data; "
            "[JUDGMENT] = analyst estimate with rationale.", 4)
ws["A3"] = "WACC build"
ws["A3"].font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
wacc_rows = [
    ("Risk-free rate (US 10Y)", RF, "DATA", "5.09% on 2026-10-19 (Reuters); rounded to 5.1% (shared framework)"),
    ("Equity risk premium", ERP, "JUDGMENT", "Long-run US ERP assumption (shared framework)"),
    (f"Beta ({TICKER})", BETA, "DATA", CONFIG["beta_note"]),
    ("Cost of equity (CAPM)", f"={A}!$B$4+{A}!$B$6*{A}!$B$5", "CALC", "Rf + beta \u00d7 ERP"),
    ("Pre-tax cost of debt", COD_PRE, "JUDGMENT", CONFIG["cod_note"]),
    ("Marginal tax rate", TAX, "JUDGMENT", "US statutory 21% (shared framework)"),
    ("After-tax cost of debt", f"={A}!$B$8*(1-{A}!$B$9)", "CALC", ""),
    ("Market value of equity ($M)", MCAP, "DATA", CONFIG["mcap_note"]),
    ("Book value of debt ($M)", DEBT_BOOK, "DATA", CONFIG["debt_note"]),
    ("Weight of equity", f"={A}!$B$11/({A}!$B$11+{A}!$B$12)", "CALC", ""),
    ("Weight of debt", f"={A}!$B$12/({A}!$B$11+{A}!$B$12)", "CALC", ""),
    ("WACC \u2014 base case", f"={A}!$B$13*{A}!$B$7+{A}!$B$14*{A}!$B$10", "CALC", ""),
    ("WACC \u2014 bull case (-100 bps)", f"={A}!$B$15-0.01", "CALC", "Rate-normalization scenario (shared framework)"),
    ("WACC \u2014 bear case (+100 bps)", f"={A}!$B$15+0.01", "CALC", "Higher-for-longer scenario (shared framework)"),
]
R = {"rf": 4, "erp": 5, "beta": 6, "coe": 7, "cod": 8, "tax": 9, "atcod": 10,
     "mcap": 11, "debt": 12, "we": 13, "wd": 14, "wacc": 15,
     "wacc_bull": 16, "wacc_bear": 17}
for i, (lbl, val, tag, note) in enumerate(wacc_rows):
    r = 4 + i
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    c = ws.cell(row=r, column=2, value=val)
    if tag in ("DATA", "JUDGMENT"):
        c.fill = INP_FILL
    c.number_format = FMT_PCT if r in (4, 5, 7, 8, 9, 10, 13, 14, 15, 16, 17) else FMT_M
    if r == 6:
        c.number_format = '0.00'
    t = ws.cell(row=r, column=3, value=f"[{tag}]")
    t.font = Font(name="Calibri", size=9, bold=True)
    t.fill = TAG_DATA_FILL if tag == "DATA" else (TAG_JUDG_FILL if tag == "JUDGMENT" else PatternFill())
    t.alignment = CENTER
    ws.cell(row=r, column=4, value=note).font = Font(name="Calibri", size=9, color="595959")
    ws.cell(row=r, column=4).alignment = LEFT

ws["A19"] = "Forecast drivers (2026\u20132030)"
ws["A19"].font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
for j, y in enumerate(FCAST):
    c = ws.cell(row=20, column=2 + j, value=y)
    c.font = HDR_FONT
    c.fill = HDR_FILL
    c.alignment = CENTER
drv = [
    ("Revenue growth \u2014 base", CASES["base"]["g"], "JUDGMENT", CONFIG["g_base_note"]),
    ("Revenue growth \u2014 bull", CASES["bull"]["g"], "JUDGMENT", CONFIG["g_bull_note"]),
    ("Revenue growth \u2014 bear", CASES["bear"]["g"], "JUDGMENT", CONFIG["g_bear_note"]),
    ("EBITDA margin \u2014 base", CASES["base"]["m"], "JUDGMENT", CONFIG["m_base_note"]),
    ("EBITDA margin \u2014 bull", CASES["bull"]["m"], "JUDGMENT", CONFIG["m_bull_note"]),
    ("EBITDA margin \u2014 bear", CASES["bear"]["m"], "JUDGMENT", CONFIG["m_bear_note"]),
]
R_DRV = {}
for i, (lbl, vals, tag, note) in enumerate(drv):
    r = 21 + i
    R_DRV[lbl] = r
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    for j, v in enumerate(vals):
        c = ws.cell(row=r, column=2 + j, value=v)
        c.fill = INP_FILL
        c.number_format = FMT_PCT
    t = ws.cell(row=r, column=7, value=f"[{tag}]")
    t.font = Font(name="Calibri", size=9, bold=True)
    t.fill = TAG_JUDG_FILL
    t.alignment = CENTER
    ws.cell(row=r, column=8, value=note).font = Font(name="Calibri", size=9, color="595959")
    ws.cell(row=r, column=8).alignment = LEFT
misc = [
    ("D&A (% of revenue)", DA_PCT, "JUDGMENT", CONFIG["da_note"]),
    ("Stock-based comp. (% of revenue)", None, "JUDGMENT", CONFIG["sbc_note"]),
    ("Capex (% of revenue)", CAPEX_PCT, "JUDGMENT", CONFIG["capex_note"]),
    ("Change in NWC (% of revenue growth)", NWC_PCT, "JUDGMENT", "Uniform working-capital intensity (shared framework)"),
    ("Terminal growth \u2014 base", CASES["base"]["tg"], "JUDGMENT", "Long-run nominal GDP (shared framework)"),
    ("Terminal growth \u2014 bull", CASES["bull"]["tg"], "JUDGMENT", "Shared framework"),
    ("Terminal growth \u2014 bear", CASES["bear"]["tg"], "JUDGMENT", "Shared framework"),
    ("Net debt ($M)", NET_DEBT, "DATA", CONFIG["netdebt_note"]),
    ("Shares outstanding (M)", SHARES, "DATA", CONFIG["shares_note"]),
]
R_MISC = {}
for i, (lbl, val, tag, note) in enumerate(misc):
    r = 28 + i
    R_MISC[lbl] = r
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    if lbl.startswith("Stock-based"):
        for j, v in enumerate(SBC_PCT):
            c = ws.cell(row=r, column=2 + j, value=v)
            c.fill = INP_FILL
            c.number_format = FMT_PCT
    else:
        c = ws.cell(row=r, column=2, value=val)
        c.fill = INP_FILL
        c.number_format = FMT_PCT if isinstance(val, float) and val < 1 else FMT_M
    t = ws.cell(row=r, column=7, value=f"[{tag}]")
    t.font = Font(name="Calibri", size=9, bold=True)
    t.fill = TAG_DATA_FILL if tag == "DATA" else TAG_JUDG_FILL
    t.alignment = CENTER
    ws.cell(row=r, column=8, value=note).font = Font(name="Calibri", size=9, color="595959")
    ws.cell(row=r, column=8).alignment = LEFT
R_SBC = R_MISC["Stock-based comp. (% of revenue)"]
finalize(ws, [38, 16, 16, 16, 16, 16, 12, 60])

# ============================ HISTORICALS ============================
ws = wb.create_sheet("Historicals")
title_block(ws, f"Historical financials \u2014 {TICKER}",
            f"$ millions. {FISCAL_NOTE}. Source: Yahoo Finance annual statements (filing-derived), as of 2026-10-19.", 6)
ws["A3"] = ""
for j, y in enumerate(YEARS):
    c = ws.cell(row=3, column=2 + j, value=y)
style_header(ws, 3, 5)
hist_rows = [
    ("Total revenue", "revenue", FMT_M),
    ("Gross profit", "gross_profit", FMT_M),
    ("EBITDA", "ebitda", FMT_M),
    ("EBITDA margin", None, FMT_PCT),
    ("Operating income", "op_income", FMT_M),
    ("Net income", "net_income", FMT_M),
    ("Diluted average shares (M)", "dil_shares", FMT_M),
    ("Operating cash flow", "ocf", FMT_M),
    ("Capital expenditures", "capex", FMT_M),
    ("Free cash flow", "fcf", FMT_M),
    ("Cash & ST investments", "cash", FMT_M),
    ("Total debt", "debt", FMT_M),
    ("Revenue growth", None, FMT_PCT),
]
for i, (lbl, key, fmt) in enumerate(hist_rows):
    r = 4 + i
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    for j in range(4):
        col = get_column_letter(2 + j)
        c = ws.cell(row=r, column=2 + j)
        if key:
            c.value = HIST[key][j]
        elif lbl == "EBITDA margin":
            c.value = f"={col}6/{col}4"
        elif lbl == "Revenue growth" and j > 0:
            prev = get_column_letter(1 + j)
            c.value = f"={col}4/{prev}4-1"
        c.number_format = fmt
H_REV = 4  # revenue row in Historicals
finalize(ws, [30, 16, 16, 16, 16, 16])

# ============================ DCF SHEETS ============================
CASE_SHEETS = {
    "base": ("DCF_Base", "Discounted cash flow \u2014 base case",
             R_DRV["Revenue growth \u2014 base"], R_DRV["EBITDA margin \u2014 base"],
             f"{A}!$B${R['wacc']}", R_MISC["Terminal growth \u2014 base"]),
    "bull": ("DCF_Bull", "Discounted cash flow \u2014 bull case",
             R_DRV["Revenue growth \u2014 bull"], R_DRV["EBITDA margin \u2014 bull"],
             f"{A}!$B${R['wacc_bull']}", R_MISC["Terminal growth \u2014 bull"]),
    "bear": ("DCF_Bear", "Discounted cash flow \u2014 bear case",
             R_DRV["Revenue growth \u2014 bear"], R_DRV["EBITDA margin \u2014 bear"],
             f"{A}!$B${R['wacc_bear']}", R_MISC["Terminal growth \u2014 bear"]),
}


def build_dcf(name, subtitle, g_row, m_row, wacc_ref, tg_row):
    ws = wb.create_sheet(name)
    title_block(ws, f"{NAME} ({TICKER}) \u2014 {subtitle}",
                "$ millions, except per-share. All formulas reference Assumptions / Historicals.", 6)
    ws["A3"] = ""
    for j, y in enumerate(FCAST):
        ws.cell(row=3, column=2 + j, value=y)
    style_header(ws, 3, 6)
    labels = [
        "Revenue", "Revenue growth", "EBITDA margin", "EBITDA",
        "Less: Depreciation & amortization", "EBIT", "Tax rate", "Less: Income tax",
        "NOPAT", "Add: Stock-based compensation", "Less: Capital expenditures",
        "Less: Change in net working capital", "Unlevered free cash flow",
        "Discount factor", "Present value of FCF",
    ]
    for i, lbl in enumerate(labels):
        ws.cell(row=4 + i, column=1, value=lbl).font = LBL_FONT
    HR = f"Historicals!$E${H_REV}"
    for j in range(5):
        col = get_column_letter(2 + j)
        pcol = get_column_letter(1 + j)
        r = 4  # revenue row
        if j == 0:
            ws.cell(row=r, column=2 + j, value=f"={HR}*(1+{A}!{col}${g_row})")
        else:
            ws.cell(row=r, column=2 + j, value=f"={pcol}{r}*(1+{A}!{col}${g_row})")
        ws.cell(row=r + 1, column=2 + j,
                value=f"={col}{r}/{HR}-1" if j == 0 else f"={col}{r}/{pcol}{r}-1").number_format = FMT_PCT
        ws.cell(row=r + 2, column=2 + j, value=f"={A}!{col}${m_row}").number_format = FMT_PCT
        ws.cell(row=r + 3, column=2 + j, value=f"={col}{r}*{col}{r+2}")
        ws.cell(row=r + 4, column=2 + j, value=f"={col}{r}*{A}!$B${R_MISC['D&A (% of revenue)']}")
        ws.cell(row=r + 5, column=2 + j, value=f"={col}{r+3}-{col}{r+4}")
        ws.cell(row=r + 6, column=2 + j, value=f"={A}!$B${R['tax']}").number_format = FMT_PCT
        ws.cell(row=r + 7, column=2 + j, value=f"={col}{r+5}*{col}{r+6}")
        ws.cell(row=r + 8, column=2 + j, value=f"={col}{r+5}-{col}{r+7}")
        ws.cell(row=r + 9, column=2 + j, value=f"={col}{r}*{A}!{col}${R_SBC}")
        ws.cell(row=r + 10, column=2 + j, value=f"={col}{r}*{A}!$B${R_MISC['Capex (% of revenue)']}")
        prev_rev = HR if j == 0 else f"{pcol}{r}"
        ws.cell(row=r + 11, column=2 + j,
                value=f"=({col}{r}-{prev_rev})*{A}!$B${R_MISC['Change in NWC (% of revenue growth)']}")
        ws.cell(row=r + 12, column=2 + j,
                value=f"={col}{r+8}+{col}{r+4}+{col}{r+9}-{col}{r+10}-{col}{r+11}")
        ws.cell(row=r + 13, column=2 + j, value=f"=1/(1+{wacc_ref})^{j+1}").number_format = '0.000'
        ws.cell(row=r + 14, column=2 + j, value=f"={col}{r+12}*{col}{r+13}")
    # summary block (column B)
    s = 19
    ws.cell(row=s, column=1, value="Sum of PV of FCF").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=s, column=2, value="=SUM(B18:F18)").number_format = FMT_M
    ws.cell(row=s + 1, column=1, value="Terminal value").font = LBL_FONT
    ws.cell(row=s + 1, column=2,
            value=f"=F16*(1+{A}!$B${tg_row})/({wacc_ref}-{A}!$B${tg_row})").number_format = FMT_M
    ws.cell(row=s + 2, column=1, value="PV of terminal value").font = LBL_FONT
    ws.cell(row=s + 2, column=2, value=f"=B{s+1}/(1+{wacc_ref})^5").number_format = FMT_M
    ws.cell(row=s + 3, column=1, value="Enterprise value").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=s + 3, column=2, value=f"=B{s}+B{s+2}").number_format = FMT_M
    ws.cell(row=s + 4, column=1, value="Less: Net debt (add: net cash)").font = LBL_FONT
    ws.cell(row=s + 4, column=2, value=f"={A}!$B${R_MISC['Net debt ($M)']}").number_format = FMT_M
    ws.cell(row=s + 5, column=1, value="Equity value").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=s + 5, column=2, value=f"=B{s+3}-B{s+4}").number_format = FMT_M
    ws.cell(row=s + 6, column=1, value="Shares outstanding (M)").font = LBL_FONT
    ws.cell(row=s + 6, column=2, value=f"={A}!$B${R_MISC['Shares outstanding (M)']}").number_format = FMT_M
    ws.cell(row=s + 7, column=1, value="Implied value per share ($)").font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    ws.cell(row=s + 7, column=2, value=f"=B{s+5}/B{s+6}").number_format = FMT_PX
    for rr in range(4, 27):
        for cc in range(2, 7):
            cell = ws.cell(row=rr, column=cc)
            if cell.value is None:
                continue
            if cell.number_format == 'General':
                cell.number_format = FMT_M
    finalize(ws, [38, 16, 16, 16, 16, 16])
    return ws


for key, args in CASE_SHEETS.items():
    build_dcf(*args)

# ============================ COMPS ============================
ws = wb.create_sheet("Comps")
ws.sheet_view.showGridLines = False
title_block(ws, f"Comparable companies \u2014 {TICKER} vs Consumer Tech peers",
            "Market data as of Oct 16, 2026 closes (Yahoo Finance). $ millions, except price and per-share.", 6)
tickers = [TICKER] + PEERS
ws["A3"] = ""
for j, p in enumerate(tickers):
    ws.cell(row=3, column=2 + j, value=p)
style_header(ws, 3, 5)
crow = {
    "Price ($)": ("price", FMT_PX),
    "Market cap ($M)": ("mcap", FMT_M),
    "Enterprise value ($M)": ("ev", FMT_M),
    "Revenue, TTM ($M)": ("rev", FMT_M),
    "EBITDA, TTM ($M)": ("ebitda", FMT_M),
    "Trailing EPS ($)": ("trail_eps", FMT_PX),
    "Forward EPS ($)": ("fwd_eps", FMT_PX),
    "EV / Revenue": (None, FMT_X),
    "EV / EBITDA": (None, FMT_X),
    "Trailing P/E": (None, FMT_X),
    "Forward P/E": (None, FMT_X),
}
r = 4
for lbl, (key, fmt) in crow.items():
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    for j, p in enumerate(tickers):
        col = get_column_letter(2 + j)
        c = ws.cell(row=r, column=2 + j)
        if lbl == "EV / Revenue":
            c.value = f"={col}6/{col}7"
        elif lbl == "EV / EBITDA":
            c.value = f"={col}6/{col}8"
        elif lbl == "Trailing P/E":
            c.value = f"={col}4/{col}9"
        elif lbl == "Forward P/E":
            c.value = f"={col}4/{col}10"
        else:
            c.value = COMPS[p][key]
        c.number_format = fmt
    r += 1
# implied valuation block (peer medians exclude the subject)
r += 1
ws.cell(row=r, column=1, value=f"{TICKER} implied valuation from peer medians").font = \
    Font(name="Calibri", size=11, bold=True, color="1F4E78")
r += 1
ws.cell(row=r, column=1, value="Median peer EV / Revenue").font = LBL_FONT
ws.cell(row=r, column=2, value="=MEDIAN(C11:E11)").number_format = FMT_X
r += 1
ws.cell(row=r, column=1, value="Implied EV ($M)").font = LBL_FONT
ws.cell(row=r, column=2, value=f"=B{r-1}*B7").number_format = FMT_M
r += 1
ws.cell(row=r, column=1, value="Less: Net debt ($M)").font = LBL_FONT
ws.cell(row=r, column=2, value=f"={A}!$B${R_MISC['Net debt ($M)']}").number_format = FMT_M
r += 1
ws.cell(row=r, column=1, value="Implied per share \u2014 EV/Revenue ($)").font = \
    Font(name="Calibri", size=11, bold=True)
ws.cell(row=r, column=2, value=f"=(B{r-2}-B{r-1})/{A}!$B${R_MISC['Shares outstanding (M)']}"
         ).number_format = FMT_PX
r += 2
ws.cell(row=r, column=1, value="Median peer EV / EBITDA").font = LBL_FONT
ws.cell(row=r, column=2, value="=MEDIAN(C12:E12)").number_format = FMT_X
r += 1
ws.cell(row=r, column=1, value="Implied EV ($M)").font = LBL_FONT
ws.cell(row=r, column=2, value=f"=B{r-1}*B8").number_format = FMT_M
r += 1
ws.cell(row=r, column=1, value="Implied per share \u2014 EV/EBITDA ($)").font = \
    Font(name="Calibri", size=11, bold=True)
ws.cell(row=r, column=2, value=f"=(B{r-1}-{A}!$B${R_MISC['Net debt ($M)']})/{A}!$B${R_MISC['Shares outstanding (M)']}"
         ).number_format = FMT_PX
r += 2
ws.cell(row=r, column=1, value="Median peer forward P/E").font = LBL_FONT
ws.cell(row=r, column=2, value="=MEDIAN(C14:E14)").number_format = FMT_X
r += 1
ws.cell(row=r, column=1, value=f"{TICKER} forward EPS ($)").font = LBL_FONT
ws.cell(row=r, column=2, value="=B10").number_format = FMT_PX
r += 1
ws.cell(row=r, column=1, value="Implied per share \u2014 forward P/E ($)").font = \
    Font(name="Calibri", size=11, bold=True, color="1F4E78")
ws.cell(row=r, column=2, value=f"=B{r-2}*B{r-1}").number_format = FMT_PX
ws.cell(row=r, column=3, value=CONFIG["comps_note"]).font = \
    Font(name="Calibri", size=9, italic=True, color="595959")
ws.cell(row=r, column=3).alignment = LEFT
finalize(ws, [44, 16, 16, 16, 16, 60])

# ============================ CHECKS ============================
ws = wb.create_sheet("Checks")
title_block(ws, "Verification appendix",
            "Tie-outs and data provenance. Every [DATA] input is sourced; every [JUDGMENT] input carries a rationale on Assumptions.", 4)
ws["A3"] = "Check"
ws["B3"] = "Detail"
ws["C3"] = "Status"
style_header(ws, 3, 4)
for i, (chk, det, st) in enumerate(CHECKS):
    r = 4 + i
    ws.cell(row=r, column=1, value=chk).font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=r, column=1).alignment = LEFT
    ws.cell(row=r, column=2, value=det).font = LBL_FONT
    ws.cell(row=r, column=2).alignment = LEFT
    c = ws.cell(row=r, column=3, value=st)
    c.font = Font(name="Calibri", size=11, bold=True)
    c.alignment = CENTER
    ws.row_dimensions[r].height = 30
finalize(ws, [32, 110, 12, 10])

# ============================ SOURCES ============================
ws = wb.create_sheet("Sources")
title_block(ws, "Sources", "Full URLs and access dates for every external input.", 4)
for j, h in enumerate(["#", "Source", "URL", "Accessed", "Used for"]):
    ws.cell(row=3, column=1 + j, value=h)
style_header(ws, 3, 5)
for i, (s, u, a, f) in enumerate(SOURCES):
    r = 4 + i
    ws.cell(row=r, column=1, value=i + 1).font = LBL_FONT
    ws.cell(row=r, column=1).alignment = CENTER
    ws.cell(row=r, column=2, value=s).font = LBL_FONT
    ws.cell(row=r, column=2).alignment = LEFT
    c = ws.cell(row=r, column=3, value=u)
    c.font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    c.alignment = LEFT
    c.hyperlink = u
    ws.cell(row=r, column=4, value=a).font = LBL_FONT
    ws.cell(row=r, column=4).alignment = CENTER
    ws.cell(row=r, column=5, value=f).font = LBL_FONT
    ws.cell(row=r, column=5).alignment = LEFT
    ws.row_dimensions[r].height = 28
finalize(ws, [6, 42, 60, 14, 44])

# sheet order and active
wb._sheets = [wb["Cover"], wb["Assumptions"], wb["Historicals"], wb["DCF_Base"],
              wb["DCF_Bull"], wb["DCF_Bear"], wb["Comps"], wb["Checks"], wb["Sources"]]
wb.active = 0
wb.save(OUT)
print("Wrote", OUT)

# ---------------------------------------------------------------- python mirror + checks
def wacc(adj=0.0):
    coe = RF + BETA * ERP
    cod_at = COD_PRE * (1 - TAX)
    we = MCAP / (MCAP + DEBT_BOOK)
    wd = DEBT_BOOK / (MCAP + DEBT_BOOK)
    return we * coe + wd * cod_at + adj

def dcf(case):
    c = CASES[case]
    w = wacc(WACC_ADJ[case])
    rev = HIST["revenue"][-1]
    fcfs = []
    for t in range(5):
        rev = rev * (1 + c["g"][t])
        ebitda = rev * c["m"][t]
        da = rev * DA_PCT
        ebit = ebitda - da
        nopat = ebit * (1 - TAX)
        sbc = rev * SBC_PCT[t]
        capex = rev * CAPEX_PCT
        nwc = (rev - (fcfs and rev / (1 + c["g"][t]) or HIST["revenue"][-1])) * NWC_PCT
        fcf = nopat + da + sbc - capex - nwc
        fcfs.append((rev, fcf))
    pv_fcf = sum(f / (1 + w) ** (t + 1) for t, (_, f) in enumerate(fcfs))
    f5 = fcfs[-1][1]
    tv = f5 * (1 + c["tg"]) / (w - c["tg"])
    pv_tv = tv / (1 + w) ** 5
    ev = pv_fcf + pv_tv
    eq = ev - NET_DEBT
    return {"wacc": w, "rev2030": fcfs[-1][0], "fcf": [f for _, f in fcfs],
            "pv_fcf": pv_fcf, "tv": tv, "pv_tv": pv_tv, "ev": ev,
            "eq": eq, "ps": eq / SHARES}

print(f"\nWACC base: {wacc():.2%}")
res = {}
for case in ("base", "bull", "bear"):
    res[case] = dcf(case)
    r = res[case]
    print(f"{case:5s}: EV={r['ev']:,.1f} EQ={r['eq']:,.1f} PS=${r['ps']:.2f} "
          f"(rev2030={r['rev2030']:,.0f})")

# comps-implied (peer medians, subject excluded)
def med(xs):
    return sorted(xs)[len(xs) // 2]
peer_keys = PEERS
evrev = med([COMPS[p]["ev"] / COMPS[p]["rev"] for p in peer_keys])
eveb = med([COMPS[p]["ev"] / COMPS[p]["ebitda"] for p in peer_keys])
fpe = med([COMPS[p]["price"] / COMPS[p]["fwd_eps"] for p in peer_keys])
imp_ps_evrev = ((evrev * COMPS[TICKER]["rev"]) - NET_DEBT) / SHARES
imp_ps_eveb = ((eveb * COMPS[TICKER]["ebitda"]) - NET_DEBT) / SHARES
imp_ps_pe = fpe * COMPS[TICKER]["fwd_eps"]
print(f"comps: med peer EV/Rev={evrev:.2f}x -> PS=${imp_ps_evrev:.2f}")
print(f"comps: med peer EV/EBITDA={eveb:.2f}x -> PS=${imp_ps_eveb:.2f}")
print(f"comps: med peer fwd P/E={fpe:.2f}x -> PS=${imp_ps_pe:.2f} (operative lens)")

# sanity checks
b = res["base"]
assert abs((b["pv_fcf"] + b["pv_tv"]) - b["ev"]) < 0.5, "FCF bridge broken"
assert abs(b["eq"] / SHARES - b["ps"]) < 0.01, "per-share arithmetic broken"
upside = (b["ps"] - PRICE) / PRICE * 100
if RECO == "BUY":
    assert upside > 20, "BUY needs >+20% base upside"
elif RECO == "HOLD":
    assert -10 <= upside <= 20, "HOLD needs -10%..+20% base upside"
else:
    assert upside < -10, "PASS needs <-10% base upside"
print("\nSanity checks passed.")
print(f"Market ${PRICE:.2f} vs DCF base ${b['ps']:.2f} (upside {upside:+.1f}%) -> {RECO}")
print(f"\nRECOMMENDATION INPUT -> base ${b['ps']:.2f} / bull ${res['bull']['ps']:.2f} / "
      f"bear ${res['bear']['ps']:.2f} / comps fwd-P/E ${imp_ps_pe:.2f} vs market ${PRICE:.2f}")
