"""Build SCHW_valuation_model.xlsx for the sector teardown (as of 2026-10-19).

Every computed cell is a live Excel formula referencing the Assumptions /
Historicals sheets, so the model is fully auditable. Python mirrors the same
math to print expected outputs, run sanity checks, and write summary.json.

Rerun:  python3 build_model.py   (from ~/workspace/sector-teardown-20/companies/<TICKER>/)
Requires: openpyxl, pandas
"""
import math
from pathlib import Path
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
OUT = BASE / "SCHW_valuation_model.xlsx"

# ================= CONFIG (company-specific) =================
TICKER = "SCHW"
NAME = "The Charles Schwab Corporation"
EXCH = "NYSE"
PRICE = 91.96               # $, Oct 16, 2026 close
MCAP = 159026.5                # $M, Oct 16, 2026
SHARES = 1729.3            # M, for per-share math
SHARES_NOTE = "Yahoo Finance / MarketBeat shares outstanding, Oct 2026 (1,729.3M); buybacks cut diluted avg from 1,894M (2022) to 1,809M (2025)"
BETA = 0.78
BETA_NOTE = "Yahoo Finance 2026-10-19; cross-checked 0.79 (MarketBeat), 0.75 (Finnhub)"
RF = 0.051                     # US 10Y ~5.1% on 2026-10-19 (verified for the sector project)
ERP = 0.045
COD_PRE = 0.055
COD_NOTE = "Investment-grade broker through-cycle new-issue estimate"
TAX = 0.22
TAX_NOTE = "2022–25 effective 21–24%, above statutory on state taxes; 22% used forward"
DEBT_BOOK = 37345.0      # $M, Q2 2026 balance sheet total debt
CASH_Q = 40580.0            # $M, Q2 2026 cash & equivalents
NET_DEBT = round(DEBT_BOOK - CASH_Q, 1)
NET_DEBT_NOTE = "Q2 2026: total debt 37,345 less cash & equivalents 40,580 (corporate borrowings; client deposits are operating liabilities)"
YEARS = [2022, 2023, 2024, 2025]
KIND = "bank"              # "broker" | "bank"
FCAST = [2026, 2027, 2028, 2029, 2030]
MARGIN_LABEL = "Op. margin (pretax+D&A)"

HIST = {'revenue': [20762.0, 18837.0, 19606.0, 23921.0], 'nii': [10682.0, 9427.0, 9144.0, 11750.0], 'ebitda': [10636.0, 7716.0, 9127.0, 12821.0], 'op_income': [9388.0, 6378.0, 7692.0, 11459.0], 'net_income': [7183.0, 5067.0, 5942.0, 8852.0], 'dil_shares': [1894.0, 1831.0, 1834.0, 1809.0], 'ocf': [2057.0, 19587.0, 2670.0, 9311.0], 'capex': [-971.0, -700.0, -620.0, -548.0], 'fcf': [None, None, None, None], 'cash': [40195.0, 43337.0, 42083.0, 46030.0], 'debt': [37878.0, 59081.0, 45134.0, 30951.0]}

CASES = {'base': {'g': [0.12, 0.09, 0.07, 0.05, 0.04], 'm': [0.54, 0.55, 0.56, 0.56, 0.57], 'wacc_adj': 0.0, 'tg': 0.03}, 'bull': {'g': [0.16, 0.12, 0.1, 0.08, 0.06], 'm': [0.55, 0.56, 0.57, 0.58, 0.59], 'wacc_adj': -0.01, 'tg': 0.035}, 'bear': {'g': [-0.08, 0.01, 0.04, 0.05, 0.04], 'm': [0.46, 0.47, 0.48, 0.49, 0.5], 'wacc_adj': 0.01, 'tg': 0.02}}
CASE_G_NOTES = {'base': 'Tapering from 22.0% (2025); NII plus asset-management fees; the rate regime is the swing factor; 5-yr CAGR ~7.4%', 'bull': 'Higher-for-longer: net interest income expands as cash sorting ends', 'bear': 'Fed easing compresses net interest income (~49% of revenue); 2026 revenue declines 8%'}
CASE_M_NOTES = {'base': '2025: 53.6% (pretax+D&A); scale toward mature-broker margin', 'bull': 'Net-interest-income operating leverage', 'bear': 'Revenue deleverage as rate cuts compress NII'}

DA_PCT = 0.055
DA_NOTE = "2025: 5.7%"
SBC_PCT = [0.013, 0.012, 0.012, 0.011, 0.011]
SBC_NOTE = "2025: 1.3%"
CAPEX_PCT = 0.02
CAPEX_NOTE = "2025: 2.3%"
NWC_PCT = 0.15
NWC_LABEL = "Balance-sheet reinvestment (% of revenue growth)"
NWC_NOTE = "Deposits fund the balance sheet (operating liabilities); the equity slice retained to support growth is a fraction of revenue growth; lower than a high-growth bank given maturity"

REC = "BUY"
REC_COLOR = "1F4E78"
THESIS = "A best-in-class brokerage compounder at 15.7x trailing earnings: a 7.8% WACC against durable 48% pretax margins supports a base value ~75% above the quote — with the rate-cut bear case as the honest risk."
ASSERT_DIR = "above"   # "below" (PASS) | "above" (BUY)

PEERS = ['SCHW', 'HOOD', 'COIN', 'SOFI']
COMPS = {'SCHW': {'price': 91.96, 'mcap': 159026.5, 'ev': 155791.5, 'rev': 23921.0, 'ebitda': 12821.0, 'ebitda_note': 'pretax+D&A proxy', 'fwd_eps': 7.85, 'trail_pe': 15.46}, 'HOOD': {'price': 98.87, 'mcap': 88892.3, 'ev': 106236.3, 'rev': 4473.0, 'ebitda': 2180.0, 'ebitda_note': '', 'fwd_eps': 3.43, 'trail_pe': 43.56}, 'COIN': {'price': 173.16, 'mcap': 45696.5, 'ev': 43750.7, 'rev': 7181.3, 'ebitda': 1795.9, 'ebitda_note': '', 'fwd_eps': 2.99, 'trail_pe': None}, 'SOFI': {'price': 14.55, 'mcap': 18792.8, 'ev': 19071.7, 'rev': 3613.4, 'ebitda': 760.1, 'ebitda_note': 'pretax+D&A proxy', 'fwd_eps': 0.82, 'trail_pe': 29.72}}
COMP_NOTE_EVREV = "Implied $86.14 — the operative lens for a deposit-funded broker; near market, a useful counterweight to the DCF (see memo)."
COMP_NOTE_PE = "Implied $226.28 — not meaningful: the peer median is dominated by HOOD/COIN high-multiple growth names."

CHECKS = [('Net income ties across statements', 'Income-statement net income = cash-flow-statement starting net income for FY2022–2025 (max diff $0.0M)', 'PASS'), ('Revenue sourced from filings', 'FY2022–2025 total revenue 20,762.0 / 18,837.0 / 19,606.0 / 23,921.0 ($M) from Yahoo Finance annual income statement; FY2021 column blank in source — excluded', 'PASS'), ('Bank-model adjustments disclosed', 'No EBITDA/gross-profit rows exist for a bank; Historicals and DCF use operating profit = pretax income + D&A (2025: $12,821M, 53.6% margin); cash-flow-statement OCF swings on client cash flows and is not used in valuation', 'PASS'), ('Share count', '1,729.3M outstanding (Yahoo Finance / MarketBeat, Oct 2026); buybacks cut diluted avg from 1,894M (2022) to 1,809M (2025)', 'PASS'), ('Net debt bridge', 'Q2 2026 balance sheet: total debt $37,345M less cash & equivalents $40,580M = net cash $3,235M (corporate borrowings; client deposits are operating liabilities)', 'PASS'), ('Beta', '0.78 (Yahoo Finance, 2026-10-19); cross-checked 0.79 (MarketBeat), 0.75 (Finnhub)', 'PASS'), ('Risk-free rate', 'US 10Y 5.1% on 2026-10-19 (verified for the sector project)', 'PASS'), ('Comps market data as-of', 'HOOD / COIN / SOFI prices are Oct 16, 2026 closes; financials pulled 2026-10-19 via Yahoo Finance', 'PASS'), ('FCF bridge reconciles', 'Formula check: EV = sum of PV of FCF + PV of terminal value', '=IF(ABS(DCF_Base!B22-(DCF_Base!B19+DCF_Base!B21))<0.5,"PASS","FAIL")'), ('Per-share arithmetic', 'Formula check: per-share = equity value / shares', '=IF(ABS(DCF_Base!B26-DCF_Base!B24/DCF_Base!B25)<0.01,"PASS","FAIL")'), ('Sensitivity centers on base case', 'Formula check: center cell of Sensitivity = DCF_Base per-share', '=IF(ABS(D7-DCF_Base!B26)<0.05,"PASS","FAIL")'), ('DCF vs market gap is explainable', 'Base per-share beats $91.96 because a 7.8% WACC (beta 0.78) against durable 48% pretax margins capitalizes through-cycle earnings the market prices at only 15.7x trailing; the rate-cut bear case ($77.41) is the honest counterweight', 'PASS'), ('Rate sensitivity flagged', 'Net interest income was 49% of 2025 revenue; the bear case models Fed easing compressing NII with 2026 revenue –8%', 'PASS')]
SOURCES = [('Yahoo Finance — SCHW quote & financials', 'https://finance.yahoo.com/quote/SCHW', '2026-10-19', 'Price, beta, shares, annual/quarterly statements'), ('Yahoo Finance — HOOD', 'https://finance.yahoo.com/quote/HOOD', '2026-10-19', 'Comps market data'), ('Yahoo Finance — COIN', 'https://finance.yahoo.com/quote/COIN', '2026-10-19', 'Comps market data'), ('Yahoo Finance — SOFI', 'https://finance.yahoo.com/quote/SOFI', '2026-10-19', 'Comps market data'), ('Reuters — 10Y yield ~5.1% on Oct 19, 2026', 'https://www.reuters.com/business/wall-st-futures-climb-ai-stocks-rebound-oil-yields-fall-2026-10-19/', '2026-10-19', 'Risk-free rate'), ('MarketBeat — SCHW', 'https://www.marketbeat.com/stocks/NYSE/SCHW/', '2026-10-19', 'Beta (0.79), shares (1,729.3M), dividend $1.28 (1.4% yield)'), ('Finnhub — SCHW', 'https://finnhub.io/?q=%22SCHW%22', '2026-10-19', 'Beta cross-check (0.75)'), ('American Banking News — SCHW Q3 2026 results', 'https://www.americanbankingnews.com/2026/10/18/citizens-jmp-reaffirms-market-outperform-rating-for-charles-schwab-nyseschw.html', '2026-10-19', 'Q3 2026: $1.69 EPS vs $1.66 est.; revenue $7.17B (+16.9% YoY); net margin 39.3%')]
# ================= END CONFIG =================

DCF_LABELS = {
    "broker": ["Revenue", "Revenue growth", "EBITDA margin", "EBITDA",
               "Less: Depreciation & amortization", "EBIT", "Tax rate",
               "Less: Income tax", "NOPAT", "Add: Stock-based compensation",
               "Less: Capital expenditures", "Less: Change in net working capital",
               "Unlevered free cash flow", "Discount factor", "Present value of FCF"],
    "bank": ["Revenue", "Revenue growth", "Op. margin (pretax+D&A)",
             "Operating profit (pretax+D&A)", "Less: Depreciation & amortization",
             "Pretax income", "Tax rate", "Less: Income tax", "NOPAT",
             "Add: Stock-based compensation", "Less: Capital expenditures",
             "Less: Balance-sheet reinvestment", "Free cash flow (equity proxy)",
             "Discount factor", "Present value of FCF"],
}[KIND]

HIST_ROWS = {
    "broker": [("Total revenue", "revenue", '#,##0.0'),
               ("Gross profit", "gross_profit", '#,##0.0'),
               ("EBITDA", "ebitda", '#,##0.0'),
               ("EBITDA margin", None, '0.0%'),
               ("Operating income", "op_income", '#,##0.0'),
               ("Net income", "net_income", '#,##0.0'),
               ("Diluted average shares (M)", "dil_shares", '#,##0.0'),
               ("Operating cash flow", "ocf", '#,##0.0'),
               ("Capital expenditures", "capex", '#,##0.0'),
               ("Free cash flow", "fcf", '#,##0.0'),
               ("Cash & equivalents", "cash", '#,##0.0'),
               ("Total debt", "debt", '#,##0.0'),
               ("Revenue growth", None, '0.0%')],
    "bank": [("Total revenue", "revenue", '#,##0.0'),
             ("Net interest income", "nii", '#,##0.0'),
             ("Operating profit (pretax + D&A)", "ebitda", '#,##0.0'),
             ("Op. margin (pretax+D&A)", None, '0.0%'),
             ("Pretax income", "op_income", '#,##0.0'),
             ("Net income", "net_income", '#,##0.0'),
             ("Diluted average shares (M)", "dil_shares", '#,##0.0'),
             ("Operating cash flow", "ocf", '#,##0.0'),
             ("Capital expenditures", "capex", '#,##0.0'),
             ("Free cash flow", "fcf", '#,##0.0'),
             ("Cash & equivalents", "cash", '#,##0.0'),
             ("Total debt", "debt", '#,##0.0'),
             ("Revenue growth", None, '0.0%')],
}[KIND]

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
title_block(ws, f"Equity Research Teardown \u2014 {NAME} ({EXCH}: {TICKER})",
            "Independent equity research project by Archit Chawla  |  As of October 19, 2026", 4)
ws["A4"] = "Current price (Oct 16, 2026 close)"
ws["B4"] = PRICE
ws["B4"].number_format = FMT_PX
ws["B4"].font = Font(name="Calibri", size=12, bold=True)
ws["A5"] = "Recommendation"
ws["B5"] = REC
ws["B5"].font = Font(name="Calibri", size=12, bold=True, color=REC_COLOR)
ws["A6"] = "One-line thesis"
ws["B6"] = THESIS
ws["B6"].alignment = LEFT
for r in (4, 5, 6):
    ws.cell(row=r, column=1).font = LBL_FONT
ws["A8"] = "Contents"
ws["A8"].font = Font(name="Calibri", size=11, bold=True)
contents = [
    ("Assumptions", "WACC build and all forecast drivers; [DATA] vs [JUDGMENT] tags"),
    ("Historicals", f"FY2022\u20132025 financials from Yahoo Finance annual statements"),
    ("DCF_Base / DCF_Bull / DCF_Bear", "5-year DCF per case; all formulas live"),
    ("Comps", "HOOD, COIN, SOFI, SCHW multiples and implied ranges"),
    ("Sensitivity", "Base-case per-share value vs WACC and terminal growth"),
    ("Checks", "Verification appendix: tie-outs and data sources"),
    ("Sources", "Full URLs and access dates"),
]
for i, (s, d) in enumerate(contents, start=9):
    ws.cell(row=i, column=1, value=s).font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=i, column=2, value=d).font = LBL_FONT
    ws.cell(row=i, column=2).alignment = LEFT
ws["A17"] = ("Disclaimer: Independent project for skill demonstration; not investment advice.")
ws["A17"].font = Font(name="Calibri", size=9, italic=True, color="595959")
ws.merge_cells("A17:D17")
finalize(ws, [28, 90, 10, 10])

# ============================ ASSUMPTIONS ============================
ws = wb.create_sheet("Assumptions")
A = "Assumptions"
title_block(ws, "Assumptions", "Blue cells are inputs. [DATA] = market/filing data; "
            "[JUDGMENT] = analyst estimate with rationale.", 4)
ws["A3"] = "WACC build"
ws["A3"].font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
wacc_rows = [
    ("Risk-free rate (US 10Y)", RF, "DATA", "5.1% on 2026-10-19 (verified for the sector project)"),
    ("Equity risk premium", ERP, "JUDGMENT", "Long-run US ERP assumption"),
    (f"Beta ({TICKER})", BETA, "DATA", BETA_NOTE),
    ("Cost of equity (CAPM)", f"={A}!$B$4+{A}!$B$6*{A}!$B$5", "CALC", "Rf + beta \u00d7 ERP"),
    ("Pre-tax cost of debt", COD_PRE, "JUDGMENT", COD_NOTE),
    ("Marginal tax rate", TAX, "JUDGMENT", TAX_NOTE),
    ("After-tax cost of debt", f"={A}!$B$8*(1-{A}!$B$9)", "CALC", ""),
    ("Market value of equity ($M)", MCAP, "DATA", f"Yahoo Finance / MarketBeat 2026-10-19: ${PRICE} \u00d7 {SHARES}M shares (per-share basis)"),
    ("Book value of debt ($M)", DEBT_BOOK, "DATA", "Q2 2026 balance sheet (Yahoo Finance)"),
    ("Weight of equity", f"={A}!$B$11/({A}!$B$11+{A}!$B$12)", "CALC", ""),
    ("Weight of debt", f"={A}!$B$12/({A}!$B$11+{A}!$B$12)", "CALC", ""),
    ("WACC \u2014 base case", f"={A}!$B$13*{A}!$B$7+{A}!$B$14*{A}!$B$10", "CALC", ""),
    ("WACC \u2014 bull case (-100 bps)", f"={A}!$B$15-0.01", "CALC", "Rate-normalization scenario"),
    ("WACC \u2014 bear case (+100 bps)", f"={A}!$B$15+0.01", "CALC", "Higher-for-longer scenario"),
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
    ("Revenue growth \u2014 base", CASES["base"]["g"], "JUDGMENT", CASE_G_NOTES["base"]),
    ("Revenue growth \u2014 bull", CASES["bull"]["g"], "JUDGMENT", CASE_G_NOTES["bull"]),
    ("Revenue growth \u2014 bear", CASES["bear"]["g"], "JUDGMENT", CASE_G_NOTES["bear"]),
    (f"{MARGIN_LABEL} \u2014 base", CASES["base"]["m"], "JUDGMENT", CASE_M_NOTES["base"]),
    (f"{MARGIN_LABEL} \u2014 bull", CASES["bull"]["m"], "JUDGMENT", CASE_M_NOTES["bull"]),
    (f"{MARGIN_LABEL} \u2014 bear", CASES["bear"]["m"], "JUDGMENT", CASE_M_NOTES["bear"]),
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
    ("D&A (% of revenue)", DA_PCT, "JUDGMENT", DA_NOTE),
    ("Stock-based comp. (% of revenue)", None, "JUDGMENT", SBC_NOTE),
    ("Capex (% of revenue)", CAPEX_PCT, "JUDGMENT", CAPEX_NOTE),
    (NWC_LABEL, NWC_PCT, "JUDGMENT", NWC_NOTE),
    ("Terminal growth \u2014 base", CASES["base"]["tg"], "JUDGMENT", "Long-run nominal GDP"),
    ("Terminal growth \u2014 bull", CASES["bull"]["tg"], "JUDGMENT", ""),
    ("Terminal growth \u2014 bear", CASES["bear"]["tg"], "JUDGMENT", ""),
    ("Net debt ($M)", NET_DEBT, "DATA", NET_DEBT_NOTE),
    ("Diluted shares outstanding (M)", SHARES, "DATA", SHARES_NOTE),
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
            "$ millions. Source: Yahoo Finance annual statements (filing-derived), as of 2026-10-19. "
            "FY2021 column blank in source \u2014 excluded.", 6)
ws["A3"] = ""
for j, y in enumerate(YEARS):
    c = ws.cell(row=3, column=2 + j, value=y)
style_header(ws, 3, 5)
for i, (lbl, key, fmt) in enumerate(HIST_ROWS):
    r = 4 + i
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    for j in range(4):
        col = get_column_letter(2 + j)
        c = ws.cell(row=r, column=2 + j)
        if key:
            v = HIST[key][j]
            c.value = v if v is not None else "n/a"
        elif "margin" in lbl.lower():
            c.value = f"={col}6/{col}4"
        elif lbl == "Revenue growth" and j > 0:
            prev = get_column_letter(1 + j)
            c.value = f"={col}4/{prev}4-1"
        if c.value != "n/a":
            c.number_format = fmt
H_REV = 4
finalize(ws, [30, 16, 16, 16, 16, 16])

# ============================ DCF SHEETS ============================
CASE_SHEETS = {
    "base": ("DCF_Base", "Discounted cash flow \u2014 base case",
             R_DRV["Revenue growth \u2014 base"], R_DRV[f"{MARGIN_LABEL} \u2014 base"],
             f"{A}!$B${R['wacc']}", R_MISC["Terminal growth \u2014 base"]),
    "bull": ("DCF_Bull", "Discounted cash flow \u2014 bull case",
             R_DRV["Revenue growth \u2014 bull"], R_DRV[f"{MARGIN_LABEL} \u2014 bull"],
             f"{A}!$B${R['wacc_bull']}", R_MISC["Terminal growth \u2014 bull"]),
    "bear": ("DCF_Bear", "Discounted cash flow \u2014 bear case",
             R_DRV["Revenue growth \u2014 bear"], R_DRV[f"{MARGIN_LABEL} \u2014 bear"],
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
    for i, lbl in enumerate(DCF_LABELS):
        ws.cell(row=4 + i, column=1, value=lbl).font = LBL_FONT
    HR = f"Historicals!$E${H_REV}"
    for j in range(5):
        col = get_column_letter(2 + j)
        pcol = get_column_letter(1 + j)
        r = 4
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
                value=f"=({col}{r}-{prev_rev})*{A}!$B${R_MISC[NWC_LABEL]}")
        ws.cell(row=r + 12, column=2 + j,
                value=f"={col}{r+8}+{col}{r+4}+{col}{r+9}-{col}{r+10}-{col}{r+11}")
        ws.cell(row=r + 13, column=2 + j, value=f"=1/(1+{wacc_ref})^{j+1}").number_format = '0.000'
        ws.cell(row=r + 14, column=2 + j, value=f"={col}{r+12}*{col}{r+13}")
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
    ws.cell(row=s + 4, column=1, value="Less: Net debt").font = LBL_FONT
    ws.cell(row=s + 4, column=2, value=f"={A}!$B${R_MISC['Net debt ($M)']}").number_format = FMT_M
    ws.cell(row=s + 5, column=1, value="Equity value").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=s + 5, column=2, value=f"=B{s+3}-B{s+4}").number_format = FMT_M
    ws.cell(row=s + 6, column=1, value="Diluted shares outstanding (M)").font = LBL_FONT
    ws.cell(row=s + 6, column=2, value=f"={A}!$B${R_MISC['Diluted shares outstanding (M)']}").number_format = FMT_M
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
title_block(ws, "Comparable companies \u2014 fintech",
            "Market data as of 2026-10-19 (Yahoo Finance); prices are Oct 16, 2026 closes. "
            "$ millions, except price and per-share. EV = Oct 16 mcap + Q2 2026 total debt "
            "\u2212 Q2 2026 cash & equivalents; revenue = FY2025 (filing-derived).", 6)
peers = PEERS
ws["A3"] = ""
for j, p in enumerate(peers):
    ws.cell(row=3, column=2 + j, value=p)
style_header(ws, 3, 6)
crow = {
    "Price ($)": ("price", FMT_PX),
    "Market cap ($M)": ("mcap", FMT_M),
    "Enterprise value ($M)": ("ev", FMT_M),
    "Revenue, FY2025 ($M)": ("rev", FMT_M),
    "EBITDA, FY2025 ($M)": ("ebitda", FMT_M),
    "Forward EPS ($)": ("fwd_eps", FMT_PX),
    "EV / Revenue": (None, FMT_X),
    "EV / EBITDA": (None, FMT_X),
    "Trailing P/E": ("trail_pe", FMT_X),
    "Forward P/E": (None, FMT_X),
}
r = 4
for lbl, (key, fmt) in crow.items():
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    for j, p in enumerate(peers):
        col = get_column_letter(2 + j)
        c = ws.cell(row=r, column=2 + j)
        if lbl == "EV / Revenue":
            c.value = f"={col}6/{col}7"
        elif lbl == "EV / EBITDA":
            v = COMPS[p]["ebitda"]
            c.value = f"={col}6/{col}8" if v else "n/a"
        elif lbl == "Forward P/E":
            c.value = f"={col}4/{col}9"
        else:
            v = COMPS[p][key]
            if lbl == "EBITDA, FY2025 ($M)" and COMPS[p].get("ebitda_note"):
                ws.cell(row=r, column=6, value=f"Note: {COMPS[p]['ebitda_note']}").font = \
                    Font(name="Calibri", size=9, italic=True, color="595959")
            c.value = v if v is not None else "n/a"
        if c.value != "n/a":
            c.number_format = fmt
    r += 1
r += 1
ws.cell(row=r, column=1, value=f"{TICKER} implied valuation from peer multiples").font = \
    Font(name="Calibri", size=11, bold=True, color="1F4E78")
r += 1
ws.cell(row=r, column=1, value="Median peer EV / Revenue").font = LBL_FONT
ws.cell(row=r, column=2, value="=MEDIAN(C10:E10)").number_format = FMT_X
r += 1
ws.cell(row=r, column=1, value="Implied EV ($M)").font = LBL_FONT
ws.cell(row=r, column=2, value=f"=B{r-1}*B7").number_format = FMT_M
r += 1
ws.cell(row=r, column=1, value="Less: Net debt ($M)").font = LBL_FONT
ws.cell(row=r, column=2, value=f"={A}!$B${R_MISC['Net debt ($M)']}").number_format = FMT_M
r += 1
ws.cell(row=r, column=1, value="Implied equity ($M)").font = LBL_FONT
ws.cell(row=r, column=2, value=f"=B{r-2}-B{r-1}").number_format = FMT_M
r += 1
ws.cell(row=r, column=1, value="Implied per share \u2014 EV/Revenue ($)").font = \
    Font(name="Calibri", size=11, bold=True)
ws.cell(row=r, column=2, value=f"=B{r-1}/{A}!$B${R_MISC['Diluted shares outstanding (M)']}").number_format = FMT_PX
ws.cell(row=r, column=3, value=COMP_NOTE_EVREV).font = Font(name="Calibri", size=9, italic=True, color="595959")
ws.cell(row=r, column=3).alignment = LEFT
r += 2
ws.cell(row=r, column=1, value="Median peer forward P/E").font = LBL_FONT
ws.cell(row=r, column=2, value="=MEDIAN(C13:E13)").number_format = FMT_X
r += 1
ws.cell(row=r, column=1, value=f"{TICKER} forward EPS ($)").font = LBL_FONT
ws.cell(row=r, column=2, value="=B9").number_format = FMT_PX
r += 1
ws.cell(row=r, column=1, value="Implied per share \u2014 forward P/E ($)").font = \
    Font(name="Calibri", size=11, bold=True, color="1F4E78")
ws.cell(row=r, column=2, value=f"=B{r-2}*B{r-1}").number_format = FMT_PX
ws.cell(row=r, column=3, value=COMP_NOTE_PE).font = Font(name="Calibri", size=9, italic=True, color="595959")
ws.cell(row=r, column=3).alignment = LEFT
finalize(ws, [44, 16, 16, 16, 16, 60])

# ============================ SENSITIVITY ============================
ws = wb.create_sheet("Sensitivity")
title_block(ws, "Sensitivity \u2014 base-case per-share value ($)",
            "Two-way sensitivity to WACC (\u00b1150 bps) and terminal growth (\u00b1100 bps). Live formulas off DCF_Base.", 6)
ws["A3"] = "WACC \\ g"
for j, adj in enumerate([-0.01, -0.005, 0, 0.005, 0.01]):
    col = get_column_letter(2 + j)
    if adj == 0:
        ws.cell(row=3, column=2 + j, value=f"={A}!$B${R_MISC['Terminal growth \u2014 base']}")
    else:
        ws.cell(row=3, column=2 + j,
                value=f"={A}!$B${R_MISC['Terminal growth \u2014 base']}{adj:+.3f}")
for i, adj in enumerate([-0.015, -0.01, -0.005, 0, 0.005, 0.01, 0.015]):
    if adj == 0:
        ws.cell(row=4 + i, column=1, value=f"={A}!$B${R['wacc']}")
    else:
        ws.cell(row=4 + i, column=1, value=f"={A}!$B${R['wacc']}{adj:+.3f}")
style_header(ws, 3, 6)
ws.cell(row=3, column=1).fill = HDR_FILL
ws.cell(row=3, column=1).font = HDR_FONT
ws.cell(row=3, column=1).alignment = CENTER
for i in range(7):
    c = ws.cell(row=4 + i, column=1)
    c.fill = HDR_FILL
    c.font = HDR_FONT
    c.alignment = CENTER
    c.number_format = FMT_PCT
for i in range(7):
    for j in range(5):
        col = get_column_letter(2 + j)
        w = f"$A{i+4}"
        g = f"{col}$3"
        fcf = "+".join(f"DCF_Base!{get_column_letter(2+k)}16/(1+{w})^{k+1}" for k in range(5))
        tv = f"(DCF_Base!$F$16*(1+{g})/({w}-{g}))/(1+{w})^5"
        c = ws.cell(row=4 + i, column=2 + j,
                    value=f"=((({fcf})+({tv})-{A}!$B${R_MISC['Net debt ($M)']})/{A}!$B${R_MISC['Diluted shares outstanding (M)']}")
        c.number_format = FMT_PX
ws.cell(row=7, column=4).fill = PatternFill("solid", fgColor="FFF2CC")
finalize(ws, [16, 16, 16, 16, 16, 16])

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

wb._sheets = [wb["Cover"], wb["Assumptions"], wb["Historicals"], wb["DCF_Base"],
              wb["DCF_Bull"], wb["DCF_Bear"], wb["Comps"], wb["Sensitivity"],
              wb["Checks"], wb["Sources"]]
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
    w = wacc(c["wacc_adj"])
    rev = HIST["revenue"][-1]
    fcfs = []
    for t in range(5):
        prev = rev
        rev = rev * (1 + c["g"][t])
        ebitda = rev * c["m"][t]
        da = rev * DA_PCT
        ebit = ebitda - da
        nopat = ebit * (1 - TAX)
        sbc = rev * SBC_PCT[t]
        capex = rev * CAPEX_PCT
        nwc = (rev - prev) * NWC_PCT
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

ev_revs = [COMPS[p]["ev"] / COMPS[p]["rev"] for p in PEERS[1:]]
med_evrev = sorted(ev_revs)[1]
imp_ps_evrev = (med_evrev * COMPS[TICKER]["rev"] - NET_DEBT) / SHARES
fwd_pes = sorted(COMPS[p]["price"] / COMPS[p]["fwd_eps"] for p in PEERS[1:])
med_pe = fwd_pes[1]
imp_ps_pe = med_pe * COMPS[TICKER]["fwd_eps"]
print(f"comps: med EV/Rev={med_evrev:.2f}x -> PS=${imp_ps_evrev:.2f}")
print(f"comps: med fwd P/E={med_pe:.2f}x -> PS=${imp_ps_pe:.2f}")

b = res["base"]
assert abs((b["pv_fcf"] + b["pv_tv"]) - b["ev"]) < 0.5, "FCF bridge broken"
if ASSERT_DIR == "below":
    assert b["ps"] < PRICE, "base should trail market (else re-examine)"
else:
    assert b["ps"] > PRICE, "base should beat market (else re-examine)"
print("\nSanity checks passed.")

import json
summary = {
    "ticker": TICKER,
    "name": NAME,
    "sector": "Fintech",
    "price": round(PRICE, 2),
    "price_asof": "2026-10-16",
    "dcf_base": round(res["base"]["ps"], 2),
    "dcf_bull": round(res["bull"]["ps"], 2),
    "dcf_bear": round(res["bear"]["ps"], 2),
    "model_asof": "2026-10-19",
    "upside_pct": round((res["base"]["ps"] - PRICE) / PRICE * 100, 1),
    "recommendation": REC,
}
with open(BASE / "summary.json", "w") as f:
    json.dump(summary, f, indent=2)
print("Wrote", BASE / "summary.json")
print(f"\nRECOMMENDATION -> {REC}: base ${res['base']['ps']:.2f} / bull ${res['bull']['ps']:.2f} / "
      f"bear ${res['bear']['ps']:.2f} / comps EV/Rev ${imp_ps_evrev:.2f} / comps P/E ${imp_ps_pe:.2f} "
      f"vs market ${PRICE:.2f}")
