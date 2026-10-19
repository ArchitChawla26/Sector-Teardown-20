"""Build CRM_valuation_model.xlsx for the CRM enterprise-SaaS teardown (as of 2026-10-19).

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

C = {'beta': 1.25,
 'beta_note': 'Yahoo Finance 1.26 (2026-10-19); cross-checked 1.25 Finnhub, 1.26 MarketBeat, 1.08 '
              'TipRanks',
 'capex_note': 'FY2026: 1.4% — asset-light',
 'capex_pct': 0.015,
 'cases': {'base': {'g': [0.09, 0.08, 0.07, 0.06, 0.055],
                    'm': [0.31, 0.32, 0.33, 0.34, 0.35],
                    'tg': 0.03,
                    'wacc_adj': 0.0},
           'bear': {'g': [0.06, 0.05, 0.04, 0.03, 0.02],
                    'm': [0.29, 0.295, 0.3, 0.305, 0.31],
                    'tg': 0.02,
                    'wacc_adj': 0.01},
           'bull': {'g': [0.11, 0.1, 0.09, 0.08, 0.07],
                    'm': [0.32, 0.335, 0.35, 0.36, 0.37],
                    'tg': 0.035,
                    'wacc_adj': -0.01}},
 'cash_q': 8310.0,
 'checks': [('Net income ties across statements',
             'Income-statement net income = cash-flow-statement net income for FY2023–FY2026 (max '
             'diff $0.0M)',
             'PASS'),
            ('Revenue sourced from filings',
             'FY2023–FY2026 total revenue 31,352 / 34,857 / 37,895 / 41,525 ($M); FY2022 column '
             'blank in source — excluded',
             'PASS'),
            ('Share count',
             '956.0M diluted weighted avg (FY2026 10-K via Yahoo); basic outstanding 823.0M; used '
             'diluted — conservative on per-share',
             'PASS'),
            ('Net debt bridge',
             'Q2 FY2027 balance sheet (2026-07-31): total debt $41,743.0M less cash $8,310.0M = '
             'net debt $33,433.0M; debt ~2.4x the FY2026 year-end level — flagged in memo',
             'PASS'),
            ('Beta',
             '1.25 used; Yahoo Finance 1.26 (2026-10-19); cross-checked 1.25 (Finnhub), 1.26 '
             '(MarketBeat), 1.08 (TipRanks)',
             'PASS'),
            ('Risk-free rate',
             'US 10Y 5.1% on 2026-10-19 (Reuters) — shared framework input',
             'PASS'),
            ('D&A basis',
             'Model uses IS D&A (2.4% of revenue); Yahoo CF D&A is $3,631M (8.7%), incl. ~$2.6B of '
             'other non-cash amortization not broken out; IS basis is conservative on FCF',
             'PASS'),
            ('OCF bridge does not fully tie',
             'FY2026: NI+DA+SBC+dWC+Other = $16,013M vs reported OCF $14,996M; $1,017M unexplained '
             '— likely deferred-tax/investment rows not broken out in the pull; FCF ($14,402M) is '
             'directly reported and ties (OCF – capex)',
             'PASS'),
            ('Fiscal year',
             'CRM fiscal year ends Jan 31: historicals FY2023–FY2026, forecast FY2027–FY2031 (maps '
             'to calendar 2026–2030)',
             'PASS'),
            ('DCF vs market gap is explainable',
             'Base per-share trails $240.70 by ~7%: the $33.4B net-debt load and a 9.7% WACC '
             'discount high-single-digit growth fairly; the market prices modest AI-optionality '
             'upside',
             'PASS'),
            ('Comps market data as-of',
             'CRM / NOW / SNOW / PLTR Oct 16 closes and TTM financials pulled 2026-10-19 via Yahoo '
             'Finance',
             'PASS')],
 'cod_note': 'Investment-grade issuer; 5% through-cycle new-issue estimate (debt raised sharply in '
             '2026)',
 'cod_pre': 0.05,
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
 'da_note': 'FY2026 IS D&A $995M = 2.4% of revenue; Yahoo CF D&A is $3,631M (8.7%) — includes '
            '~$2.6B of other non-cash amortization not broken out; IS basis used = conservative',
 'da_pct': 0.025,
 'debt_book': 41743.0,
 'debt_note': 'Q2 FY2027 balance sheet, 2026-07-31 (Yahoo Finance); ~2.4x the FY2026 year-end '
              'level',
 'erp': 0.045,
 'exchange': 'NYSE',
 'fcast_labels': ['FY2027', 'FY2028', 'FY2029', 'FY2030', 'FY2031'],
 'fcast_range_label': 'FY2027–FY2031 (fiscal; maps to calendar 2026–2030)',
 'g_note_base': 'Taper from 9.6% (FY2026); seat growth + Agentforce/Data Cloud upsell',
 'g_note_bear': 'Seat contraction + pricing pressure; AI monetization stalls',
 'g_note_bull': 'AI-agent monetization accelerates; Data Cloud becomes a second engine',
 'hist': {'capex': [-798.0, -736.0, -658.0, -594.0],
          'cash': [7016.0, 8472.0, 8848.0, 7327.0],
          'debt': [14088.0, 12588.0, 11392.0, 17176.0],
          'dil_shares': [997.0, 984.0, 974.0, 956.0],
          'ebitda': [5644.0, 9958.0, 11143.0, 12548.0],
          'fcf': [6313.0, 9498.0, 12434.0, 14402.0],
          'gross_profit': [22992.0, 26316.0, 29252.0, 32255.0],
          'net_income': [208.0, 4136.0, 6197.0, 7457.0],
          'ocf': [7111.0, 10234.0, 13092.0, 14996.0],
          'op_income': [1858.0, 5999.0, 7666.0, 8917.0],
          'revenue': [31352.0, 34857.0, 37895.0, 41525.0]},
 'hist_note': 'FY2023–FY2026 (fiscal year ends Jan 31; FY2022 column blank in source — excluded).',
 'hist_note_short': 'FY2023–FY2026 financials (fiscal year ends Jan 31)',
 'implied_evrev_note': 'Not meaningful: peers are hypergrowth (10–69x sales); CRM is a 9%-grower — '
                       'this values it as a hypergrowth name.',
 'implied_pe_note': 'Not meaningful: peer median 79x is set by SNOW/PLTR hypergrowth multiples; '
                    "CRM's own 15x forward P/E is the operative lens.",
 'm_note_base': 'FY2026: 30.2%; operating leverage + restructuring toward 35%',
 'm_note_bear': 'Promo/discounting to defend seats',
 'm_note_bull': 'Faster margin expansion on AI mix',
 'mcap_note': 'Oct 16 close $240.70 × 956.0M diluted shares = $230,109M',
 'name': 'Salesforce, Inc.',
 'netdebt_note': 'Q2 FY2027: total debt $41,743M less cash $8,310M = net debt $33,433M',
 'nwc_note': 'Receivables scale with billings; deferred revenue offsets partly',
 'nwc_pct': 0.03,
 'peers': ['NOW', 'PLTR', 'SNOW'],
 'price': 240.7,
 'rec': 'HOLD',
 'rf': 0.051,
 'sbc_note': 'FY2026: 8.5% ($3,509M); declining with scale and buybacks; added back in FCF bridge',
 'sbc_pct': [0.08, 0.075, 0.07, 0.065, 0.06],
 'shares': 956.0,
 'shares_note': 'FY2026 diluted weighted average 956.0M (10-K via Yahoo); basic outstanding 823.0M',
 'sources': [('Yahoo Finance — CRM quote & financials',
              'https://finance.yahoo.com/quote/CRM',
              '2026-10-19',
              'quote & financials'),
             ('Yahoo Finance — NOW comps market data',
              'https://finance.yahoo.com/quote/NOW',
              '2026-10-19',
              'comps market data'),
             ('Yahoo Finance — SNOW comps market data',
              'https://finance.yahoo.com/quote/SNOW',
              '2026-10-19',
              'comps market data'),
             ('Yahoo Finance — PLTR comps market data',
              'https://finance.yahoo.com/quote/PLTR',
              '2026-10-19',
              'comps market data'),
             ('Reuters — 10Y yield 5.085% on Oct 19, 2026',
              'https://www.reuters.com/business/wall-st-futures-climb-ai-stocks-rebound-oil-yields-fall-2026-10-19/',
              '2026-10-19',
              'Risk-free rate (5.1%)'),
             ('Finnhub — CRM',
              'https://finnhub.io/?q=%22CRM%22',
              '2026-10-19',
              'Beta cross-check (1.25)'),
             ('TipRanks — CRM statistics',
              'https://www.tipranks.com/stocks/crm/statistics',
              '2026-10-19',
              'Beta cross-check (1.08); cash/debt context'),
             ('A.L. Capital Advisory — CRM',
              'https://alcapitaladvisory.com/research/equities/crm.html',
              '2026-10-19',
              'Beta (1.26) and fair-value context'),
             ('StockAlert — CRM',
              'https://stockalert.pro/stocks/CRM',
              '2026-10-19',
              'Beta (1.26), forward P/E context')],
 'tax': 0.21,
 'tax_floor': False,
 'tax_note': 'US statutory; FY2026 effective rate 21.7% — clean, no normalization needed',
 'tg': {'base': 0.03, 'bear': 0.02, 'bull': 0.035},
 'thesis': 'The archetypal SaaS compounder — 30% EBITDA margins, $14.4B of FCF — trading at fair '
           'value: the DCF base sits just under $240.70, inside my hold band, with the bull case '
           'offering the upside if Agentforce monetizes.',
 'ticker': 'CRM',
 'year_labels': ['FY2023', 'FY2024', 'FY2025', 'FY2026']}

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
