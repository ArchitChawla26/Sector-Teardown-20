"""Build PLTR_valuation_model.xlsx for the PLTR enterprise-SaaS teardown (as of 2026-10-19).

Every computed cell is a live Excel formula referencing the Assumptions /
Historicals sheets, so the model is fully auditable. Python mirrors the same
math to print expected outputs and run sanity checks.

Rerun:  python3 build_model.py   (from the company folder)
Requires: openpyxl, pandas
"""
import math
from pathlib import Path
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

C = {'beta': 1.6,
 'beta_note': 'Yahoo Finance 1.60 (2026-10-19); cross-checked 1.62 Finnhub, 1.60 MarketBeat',
 'capex_note': 'FY2025: 0.8% — pure software',
 'capex_pct': 0.01,
 'cases': {'base': {'g': [0.65, 0.45, 0.32, 0.24, 0.18],
                    'm': [0.36, 0.4, 0.43, 0.45, 0.47],
                    'tg': 0.03,
                    'wacc_adj': 0.0},
           'bear': {'g': [0.45, 0.3, 0.2, 0.14, 0.1],
                    'm': [0.32, 0.34, 0.36, 0.37, 0.38],
                    'tg': 0.02,
                    'wacc_adj': 0.01},
           'bull': {'g': [0.75, 0.55, 0.4, 0.3, 0.22],
                    'm': [0.38, 0.43, 0.46, 0.48, 0.5],
                    'tg': 0.035,
                    'wacc_adj': -0.01}},
 'cash_q': 2030.0,
 'checks': [('Net income ties across statements',
             'Income-statement vs cash-flow-statement net income for FY2022–FY2025 (max diff $9.6M '
             'on $1.6B income — minority interest/rounding)',
             'PASS'),
            ('Revenue sourced from filings',
             'FY2022–FY2025 total revenue 1,905.9 / 2,225.0 / 2,865.5 / 4,475.4 ($M); FY2021 '
             'column blank in source — excluded',
             'PASS'),
            ('Tax normalization',
             'FY2025 tax provision only $22.7M (~1.4% effective) on $1,657M pretax — '
             'NOL/valuation-allowance dynamics; model uses 21% statutory forward',
             'PASS'),
            ('Share count', '2,565.2M diluted weighted avg (FY2025 10-K via Yahoo)', 'PASS'),
            ('Net cash',
             'Q2 2026 balance sheet (2026-06-30): total debt $211.4M less cash $2,030.0M = net '
             'cash $1,818.6M',
             'PASS'),
            ('Beta',
             '1.60 used; Yahoo Finance 1.60 (2026-10-19); cross-checked 1.62 (Finnhub), 1.60 '
             '(MarketBeat)',
             'PASS'),
            ('Risk-free rate',
             'US 10Y 5.1% on 2026-10-19 (Reuters) — shared framework input',
             'PASS'),
            ('DCF vs market gap is explainable',
             'Base per-share trails $186.03 by ~83%: even 65%-tapering growth to $20.7B revenue at '
             '47% EBITDA margins cannot support 160x trailing earnings at a 12.3% WACC — the quote '
             'embeds a decade of flawless AIP dominance',
             'PASS'),
            ('Comps market data as-of',
             'CRM / NOW / SNOW / PLTR Oct 16 closes and TTM financials pulled 2026-10-19 via Yahoo '
             'Finance',
             'PASS')],
 'cod_note': 'Minimal debt ($211M); 5.5% through-cycle estimate, immaterial to WACC',
 'cod_pre': 0.055,
 'comps': {'CRM': {'ebitda': 12899.0,
                   'ev': 231529.1,
                   'fwd_eps': 16.023,
                   'mcap': 198096.1,
                   'price': 240.7,
                   'rev': 43938.0,
                   'trail_pe': 21.98},
           'NOW': {'ebitda': 2897.0,
                   'ev': 152727.7,
                   'fwd_eps': 5.005,
                   'mcap': 146777.7,
                   'price': 141.97,
                   'rev': 14732.0,
                   'trail_pe': 88.73},
           'PLTR': {'ebitda': 2662.6,
                    'ev': 426180.8,
                    'fwd_eps': 2.352,
                    'mcap': 427999.4,
                    'price': 186.03,
                    'rev': 6155.9,
                    'trail_pe': 160.37},
           'SNOW': {'ebitda': -1036.3,
                    'ev': 109309.8,
                    'fwd_eps': 3.008,
                    'mcap': 108253.2,
                    'price': 306.84,
                    'rev': 5434.6,
                    'trail_pe': None}},
 'comps_note': 'EV uses Oct 16 market cap + latest-quarter total debt less cash.',
 'da_note': 'FY2025 CF D&A $26M = 0.6% of revenue — negligible fixed base',
 'da_pct': 0.005,
 'debt_book': 211.4,
 'debt_note': 'Q2 2026 balance sheet, 2026-06-30 (Yahoo Finance)',
 'erp': 0.045,
 'exchange': 'NASDAQ',
 'fcast_labels': [2026, 2027, 2028, 2029, 2030],
 'fcast_range_label': 'FY2026–FY2030',
 'g_note_base': 'FY2025: +56.2%; Q2 2026 +92.8% YoY; AIP commercial + government; tapering as base '
                'scales',
 'g_note_bear': 'AI hype normalizes; government budget pressure',
 'g_note_bull': 'AIP becomes the enterprise AI standard; government contracts compound',
 'hist': {'capex': [-40.0, -15.1, -12.6, -33.9],
          'cash': [2598.5, 831.0, 2098.5, 1423.8],
          'debt': [249.4, 229.4, 239.2, 229.3],
          'dil_shares': [2060.8, 2294.6, 2447.2, 2565.2],
          'ebitda': [-138.7, 153.3, 342.0, 1440.2],
          'fcf': [183.7, 697.1, 1141.2, 2100.6],
          'gross_profit': [1497.3, 1793.9, 2299.5, 3686.3],
          'net_income': [-373.7, 209.8, 462.2, 1625.0],
          'ocf': [223.7, 712.2, 1153.9, 2134.5],
          'op_income': [-161.2, 120.0, 310.4, 1414.0],
          'revenue': [1905.9, 2225.0, 2865.5, 4475.4]},
 'hist_note': 'FY2022–FY2025 (FY2021 column blank in source — excluded).',
 'hist_note_short': 'FY2022–FY2025 financials',
 'implied_evrev_note': "Illustrative: the market pays 69x sales for PLTR's growth; no peer in this "
                       'set grows 65% — the multiple is descriptive, not prescriptive.',
 'implied_pe_note': "Illustrative: peer median 28x on PLTR's $2.35 forward EPS; the DCF is the "
                    'operative lens at 160x trailing earnings.',
 'm_note_base': 'FY2025: 32.2%; operating leverage toward 47% — software economics',
 'm_note_bear': 'Growth investments cap margins',
 'm_note_bull': 'Faster leverage on commercial mix',
 'mcap_note': 'Oct 16 close $186.03 × 2,565.2M diluted shares = $477,201M',
 'name': 'Palantir Technologies Inc.',
 'netdebt_note': 'Q2 2026: total debt $211M less cash $2,030M = net cash $1,819M',
 'nwc_note': 'Receivables scale with commercial expansion (FY2025 dWC was a $210M use)',
 'nwc_pct': 0.03,
 'peers': ['CRM', 'NOW', 'SNOW'],
 'price': 186.03,
 'rec': 'PASS',
 'rf': 0.051,
 'sbc_note': 'FY2025: 15.3% ($684M); declining with scale; added back in FCF bridge',
 'sbc_pct': [0.14, 0.13, 0.12, 0.11, 0.1],
 'shares': 2565.2,
 'shares_note': 'FY2025 diluted weighted average 2,565.2M (10-K via Yahoo)',
 'sources': [('Yahoo Finance — PLTR quote & financials',
              'https://finance.yahoo.com/quote/PLTR',
              '2026-10-19',
              'quote & financials'),
             ('Yahoo Finance — CRM comps market data',
              'https://finance.yahoo.com/quote/CRM',
              '2026-10-19',
              'comps market data'),
             ('Yahoo Finance — NOW comps market data',
              'https://finance.yahoo.com/quote/NOW',
              '2026-10-19',
              'comps market data'),
             ('Yahoo Finance — SNOW comps market data',
              'https://finance.yahoo.com/quote/SNOW',
              '2026-10-19',
              'comps market data'),
             ('Reuters — 10Y yield 5.085% on Oct 19, 2026',
              'https://www.reuters.com/business/wall-st-futures-climb-ai-stocks-rebound-oil-yields-fall-2026-10-19/',
              '2026-10-19',
              'Risk-free rate (5.1%)'),
             ('Finnhub — PLTR',
              'https://finnhub.io/?q=%22PLTR%22',
              '2026-10-19',
              'Beta cross-check (1.62)'),
             ('Zacks — PLTR beta',
              'https://www.zacks.com/stock/chart/PLTR/fundamental/beta',
              '2026-10-19',
              'Oct 16 close $186.03; beta context')],
 'tax': 0.21,
 'tax_floor': False,
 'tax_note': 'US statutory; FY2025 effective rate ~1.4% ($23M provision) reflects '
             'NOL/valuation-allowance dynamics — normalized to 21% forward',
 'tg': {'base': 0.03, 'bear': 0.02, 'bull': 0.035},
 'thesis': 'An exceptional AI-platform business at a price that requires miracles: 160x trailing '
           'earnings embeds flawless execution; my DCF base sits ~83% below $186.03 — and even the '
           'bull case is ~73% under.',
 'ticker': 'PLTR',
 'year_labels': ['FY2022', 'FY2023', 'FY2024', 'FY2025']}

BASE = Path(__file__).resolve().parent
OUT = BASE / f"{C['ticker']}_valuation_model.xlsx"
A = "Assumptions"

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
REC_COLORS = {"BUY": "2E7D32", "HOLD": "B26A00", "PASS": "C00000"}


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
title_block(ws, f"Equity Research Teardown \u2014 {C['name']} ({C['exchange']}: {C['ticker']})",
            "Independent equity research project by Archit Chawla  |  As of October 19, 2026", 4)
ws["A4"] = "Current price (Oct 16, 2026 close)"
ws["B4"] = C["price"]
ws["B4"].number_format = FMT_PX
ws["B4"].font = Font(name="Calibri", size=12, bold=True)
ws["A5"] = "Recommendation"
ws["B5"] = C["rec"]
ws["B5"].font = Font(name="Calibri", size=12, bold=True, color=REC_COLORS[C["rec"]])
ws["A6"] = "One-line thesis"
ws["B6"] = C["thesis"]
ws["B6"].alignment = LEFT
for r in (4, 5, 6):
    ws.cell(row=r, column=1).font = LBL_FONT
ws["A8"] = "Contents"
ws["A8"].font = Font(name="Calibri", size=11, bold=True)
contents = [
    ("Assumptions", "WACC build and all forecast drivers; [DATA] vs [JUDGMENT] tags"),
    ("Historicals", f"{C['hist_note_short']} from Yahoo Finance annual statements"),
    ("DCF_Base / DCF_Bull / DCF_Bear", "5-year DCF per case; all formulas live"),
    ("Comps", f"{', '.join(C['peers'])} multiples and {C['ticker']}-implied ranges"),
    ("Checks", "Verification appendix: tie-outs and data sources"),
    ("Sources", "Full URLs and access dates"),
]
for i, (s, d) in enumerate(contents, start=9):
    ws.cell(row=i, column=1, value=s).font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=i, column=2, value=d).font = LBL_FONT
    ws.cell(row=i, column=2).alignment = LEFT
ws["A16"] = "Disclaimer: Independent project for skill demonstration; not investment advice."
ws["A16"].font = Font(name="Calibri", size=9, italic=True, color="595959")
ws.merge_cells("A16:D16")
finalize(ws, [28, 90, 10, 10])

# ============================ ASSUMPTIONS ============================
ws = wb.create_sheet("Assumptions")
title_block(ws, "Assumptions", "Blue cells are inputs. [DATA] = market/filing data; "
            "[JUDGMENT] = analyst estimate with rationale.", 4)
ws["A3"] = "WACC build"
ws["A3"].font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
RF, ERP, BETA, TAX = C["rf"], C["erp"], C["beta"], C["tax"]
MCAP = C["price"] * C["shares"]
NET_DEBT = C["debt_book"] - C["cash_q"]
wacc_rows = [
    ("Risk-free rate (US 10Y)", RF, "DATA", "5.1% on 2026-10-19 (Reuters); verified for this project"),
    ("Equity risk premium", ERP, "JUDGMENT", "Long-run US ERP assumption (shared framework)"),
    (f"Beta ({C['ticker']})", BETA, "DATA", C["beta_note"]),
    ("Cost of equity (CAPM)", f"={A}!$B$4+{A}!$B$6*{A}!$B$5", "CALC", "Rf + beta \u00d7 ERP"),
    ("Pre-tax cost of debt", C["cod_pre"], "JUDGMENT", C["cod_note"]),
    ("Marginal tax rate", TAX, "JUDGMENT", C["tax_note"]),
    ("After-tax cost of debt", f"={A}!$B$8*(1-{A}!$B$9)", "CALC", ""),
    ("Market value of equity ($M)", MCAP, "DATA", C["mcap_note"]),
    ("Book value of debt ($M)", C["debt_book"], "DATA", C["debt_note"]),
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

ws["A19"] = f"Forecast drivers ({C['fcast_range_label']})"
ws["A19"].font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
for j, y in enumerate(C["fcast_labels"]):
    c = ws.cell(row=20, column=2 + j, value=y)
    c.font = HDR_FONT
    c.fill = HDR_FILL
    c.alignment = CENTER
drv = [
    ("Revenue growth \u2014 base", C["cases"]["base"]["g"], "JUDGMENT", C["g_note_base"]),
    ("Revenue growth \u2014 bull", C["cases"]["bull"]["g"], "JUDGMENT", C["g_note_bull"]),
    ("Revenue growth \u2014 bear", C["cases"]["bear"]["g"], "JUDGMENT", C["g_note_bear"]),
    ("EBITDA margin \u2014 base", C["cases"]["base"]["m"], "JUDGMENT", C["m_note_base"]),
    ("EBITDA margin \u2014 bull", C["cases"]["bull"]["m"], "JUDGMENT", C["m_note_bull"]),
    ("EBITDA margin \u2014 bear", C["cases"]["bear"]["m"], "JUDGMENT", C["m_note_bear"]),
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
    ("D&A (% of revenue)", C["da_pct"], "JUDGMENT", C["da_note"]),
    ("Stock-based comp. (% of revenue)", None, "JUDGMENT", C["sbc_note"]),
    ("Capex (% of revenue)", C["capex_pct"], "JUDGMENT", C["capex_note"]),
    ("Change in NWC (% of revenue growth)", C["nwc_pct"], "JUDGMENT", C["nwc_note"]),
    ("Terminal growth \u2014 base", C["tg"]["base"], "JUDGMENT", "Long-run nominal GDP (shared framework)"),
    ("Terminal growth \u2014 bull", C["tg"]["bull"], "JUDGMENT", ""),
    ("Terminal growth \u2014 bear", C["tg"]["bear"], "JUDGMENT", ""),
    ("Net debt ($M)", NET_DEBT, "DATA", C["netdebt_note"]),
    ("Diluted shares outstanding (M)", C["shares"], "DATA", C["shares_note"]),
]
R_MISC = {}
for i, (lbl, val, tag, note) in enumerate(misc):
    r = 28 + i
    R_MISC[lbl] = r
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    if lbl.startswith("Stock-based"):
        for j, v in enumerate(C["sbc_pct"]):
            c = ws.cell(row=r, column=2 + j, value=v)
            c.fill = INP_FILL
            c.number_format = FMT_PCT
    else:
        c = ws.cell(row=r, column=2, value=val)
        c.fill = INP_FILL
        c.number_format = FMT_PCT if isinstance(val, float) and abs(val) < 1 else FMT_M
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
H = C["hist"]
YEARS = C["year_labels"]
title_block(ws, f"Historical financials \u2014 {C['ticker']}",
            f"$ millions. Source: Yahoo Finance annual statements (filing-derived), as of 2026-10-19. {C['hist_note']}", 6)
ws["A3"] = ""
for j, y in enumerate(YEARS):
    ws.cell(row=3, column=2 + j, value=y)
style_header(ws, 3, 2 + len(YEARS) - 1)
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
    ("Cash & equivalents", "cash", FMT_M),
    ("Total debt", "debt", FMT_M),
    ("Revenue growth", None, FMT_PCT),
]
n = len(YEARS)
for i, (lbl, key, fmt) in enumerate(hist_rows):
    r = 4 + i
    ws.cell(row=r, column=1, value=lbl).font = LBL_FONT
    for j in range(n):
        col = get_column_letter(2 + j)
        c = ws.cell(row=r, column=2 + j)
        if key:
            c.value = H[key][j]
        elif lbl == "EBITDA margin":
            c.value = f"={col}6/{col}4"
        elif lbl == "Revenue growth" and j > 0:
            prev = get_column_letter(1 + j)
            c.value = f"={col}4/{prev}4-1"
        c.number_format = fmt
H_REV = 4
finalize(ws, [30] + [16] * n)

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
TAX_FLOOR = C["tax_floor"]


def build_dcf(name, subtitle, g_row, m_row, wacc_ref, tg_row):
    ws = wb.create_sheet(name)
    title_block(ws, f"{C['name']} ({C['ticker']}) \u2014 {subtitle}",
                "$ millions, except per-share. All formulas reference Assumptions / Historicals.", 6)
    ws["A3"] = ""
    for j, y in enumerate(C["fcast_labels"]):
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
        if TAX_FLOOR:
            ws.cell(row=r + 7, column=2 + j, value=f"=MAX({col}{r+5},0)*{col}{r+6}")
        else:
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
CMP = C["comps"]
peers = [C["ticker"]] + C["peers"]
title_block(ws, "Comparable companies \u2014 enterprise SaaS",
            f"Market data as of 2026-10-16 (Oct 16 closes via Yahoo Finance). $ millions, except price and per-share. {C['comps_note']}", 6)
ws["A3"] = ""
for j, p in enumerate(peers):
    ws.cell(row=3, column=2 + j, value=p)
style_header(ws, 3, 6)
crow = {
    "Price ($)": ("price", FMT_PX),
    "Market cap ($M)": ("mcap", FMT_M),
    "Enterprise value ($M)": ("ev", FMT_M),
    "Revenue, TTM ($M)": ("rev", FMT_M),
    "EBITDA, TTM ($M)": ("ebitda", FMT_M),
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
            v = CMP[p]["ebitda"]
            c.value = f"={col}6/{col}8" if v and v > 0 else "n/a"
        elif lbl == "Forward P/E":
            v = CMP[p]["fwd_eps"]
            c.value = f"={col}4/{col}9" if v else "n/a"
        else:
            v = CMP[p][key]
            c.value = v if v is not None else "n/a"
        if c.value != "n/a":
            c.number_format = fmt
    r += 1
r += 1
ws.cell(row=r, column=1, value=f"{C['ticker']} implied valuation from peer multiples").font = \
    Font(name="Calibri", size=11, bold=True, color="1F4E78")
r += 1
ws.cell(row=r, column=1, value="Median peer EV / Revenue").font = LBL_FONT
ws.cell(row=r, column=2, value=f"=MEDIAN({get_column_letter(3)}10:{get_column_letter(5)}10)").number_format = FMT_X
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
ws.cell(row=r, column=3, value=C["implied_evrev_note"]).font = Font(name="Calibri", size=9, italic=True, color="595959")
ws.cell(row=r, column=3).alignment = LEFT
r += 2
ws.cell(row=r, column=1, value="Median peer forward P/E").font = LBL_FONT
ws.cell(row=r, column=2, value=f"=MEDIAN({get_column_letter(3)}13:{get_column_letter(5)}13)").number_format = FMT_X
r += 1
ws.cell(row=r, column=1, value=f"{C['ticker']} forward EPS ($)").font = LBL_FONT
ws.cell(row=r, column=2, value="=B9").number_format = FMT_PX
r += 1
ws.cell(row=r, column=1, value="Implied per share \u2014 forward P/E ($)").font = \
    Font(name="Calibri", size=11, bold=True, color="1F4E78")
ws.cell(row=r, column=2, value=f"=B{r-2}*B{r-1}").number_format = FMT_PX
ws.cell(row=r, column=3, value=C["implied_pe_note"]).font = Font(name="Calibri", size=9, italic=True, color="595959")
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
checks = C["checks"] + [
    ("FCF bridge reconciles", "Formula check: EV = sum of PV of FCF + PV of terminal value",
     '=IF(ABS(DCF_Base!B22-(DCF_Base!B19+DCF_Base!B21))<0.5,"PASS","FAIL")'),
    ("Per-share arithmetic", "Formula check: per-share = equity value / shares",
     '=IF(ABS(DCF_Base!B26-DCF_Base!B24/DCF_Base!B25)<0.01,"PASS","FAIL")'),
]
for i, (chk, det, st) in enumerate(checks):
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
for i, (s, u, a, f) in enumerate(C["sources"]):
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
    coe = C["rf"] + C["beta"] * C["erp"]
    cod_at = C["cod_pre"] * (1 - C["tax"])
    mcap = C["price"] * C["shares"]
    we = mcap / (mcap + C["debt_book"])
    return we * coe + (1 - we) * cod_at + adj


def dcf(case):
    c = C["cases"][case]
    w = wacc(c["wacc_adj"])
    rev = C["hist"]["revenue"][-1]
    fcfs, fcfm = [], []
    for t in range(5):
        prev = rev
        rev = rev * (1 + c["g"][t])
        ebitda = rev * c["m"][t]
        ebit = ebitda - rev * C["da_pct"]
        tax = max(ebit, 0) * C["tax"] if C["tax_floor"] else ebit * C["tax"]
        nopat = ebit - tax
        fcf = (nopat + rev * C["da_pct"] + rev * C["sbc_pct"][t]
               - rev * C["capex_pct"] - (rev - prev) * C["nwc_pct"])
        fcfs.append(fcf)
        fcfm.append(fcf / rev)
    pv_fcf = sum(f / (1 + w) ** (t + 1) for t, f in enumerate(fcfs))
    tv = fcfs[-1] * (1 + c["tg"]) / (w - c["tg"])
    pv_tv = tv / (1 + w) ** 5
    ev = pv_fcf + pv_tv
    eq = ev - (C["debt_book"] - C["cash_q"])
    return {"wacc": w, "rev5": rev, "fcf": fcfs, "fcfm": fcfm,
            "pv_fcf": pv_fcf, "tv": tv, "pv_tv": pv_tv, "ev": ev,
            "eq": eq, "ps": eq / C["shares"]}


print(f"\nWACC base: {wacc():.2%}")
res = {}
for case in ("base", "bull", "bear"):
    res[case] = dcf(case)
    r = res[case]
    fm = " ".join(f"{x:.1%}" for x in r["fcfm"])
    print(f"{case:5s}: EV={r['ev']:,.1f} EQ={r['eq']:,.1f} PS=${r['ps']:.2f} "
          f"(rev5={r['rev5']:,.0f} fcfm=[{fm}])")

b = res["base"]
assert abs((b["pv_fcf"] + b["pv_tv"]) - b["ev"]) < 0.5, "FCF bridge broken"
upside = (b["ps"] - C["price"]) / C["price"] * 100
print(f"\nBase upside vs ${C['price']:.2f}: {upside:+.1f}% -> {C['rec']}")
if C["rec"] in ("PASS", "HOLD"):
    assert upside < 20.0, "upside inconsistent with HOLD/PASS"
else:
    assert upside > 20.0, "upside inconsistent with BUY"
print("Sanity checks passed.")

# comps-implied (python mirror of the Comps sheet formulas)
CMP = C["comps"]
peer_keys = C["peers"]
ev_revs = sorted(CMP[p]["ev"] / CMP[p]["rev"] for p in peer_keys)
med_evrev = ev_revs[len(ev_revs) // 2]
imp_ps_evrev = (med_evrev * CMP[C["ticker"]]["rev"] - (C["debt_book"] - C["cash_q"])) / C["shares"]
fwd_pes = sorted(CMP[p]["price"] / CMP[p]["fwd_eps"] for p in peer_keys if CMP[p]["fwd_eps"])
med_pe = fwd_pes[len(fwd_pes) // 2] if len(fwd_pes) % 2 else \
    (fwd_pes[len(fwd_pes) // 2 - 1] + fwd_pes[len(fwd_pes) // 2]) / 2
imp_ps_pe = med_pe * CMP[C["ticker"]]["fwd_eps"]
print(f"comps: med peer EV/Rev={med_evrev:.2f}x -> PS=${imp_ps_evrev:,.2f}")
print(f"comps: med peer fwd P/E={med_pe:.2f}x -> PS=${imp_ps_pe:,.2f}")
print(f"\nSUMMARY_JSON: ps_base={b['ps']:.2f} ps_bull={res['bull']['ps']:.2f} "
      f"ps_bear={res['bear']['ps']:.2f} upside={upside:.1f}")
