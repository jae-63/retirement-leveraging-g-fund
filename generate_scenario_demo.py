"""
generate_scenario_demo.py

Builds 'household_portfolio_scenario_demo.xlsx' -- a standalone companion to
household_portfolio_engine.xlsx, sharing the same four-tab structure:

  1. Documentation       -- what this demo is, how to use the scenario
                             selector, a disclaimer, the color legend, and a
                             glossary.
  2. Variables_And_Log   -- annual COLA/hurdle inputs, a Jan-Dec monthly
                             TSP -> checking payment log, and three hardcoded
                             illustrative market-return paths (Bear / Bull /
                             Choppy) per year.
  3. Account_Tracker     -- generic starting balances (beginning of Year 1),
                             not a live snapshot.
  4. The_Refill_Engine   -- same Quarterly Refill Engine / Annual SRR Hurdle
                             Model / Emergency Liquidity Protocol as the live
                             tracker, but the Annual Hurdle Model now also
                             simulates actual household wealth forward under
                             whichever scenario is selected, so the GOOD/POOR
                             trigger is visibly driven by the market path
                             instead of a frozen wealth snapshot.

Unlike the live tracker, this workbook is meant to be opened and understood
immediately -- every input is pre-filled with generic, illustrative numbers.
All account balances, payments, and market returns are hypothetical and
generic on purpose: this file is meant to be shared with a financial advisor
and eventually posted publicly as a template, so it must never contain real
household data. See household_portfolio_engine.xlsx / generate_household_
portfolio_engine.py for the live, hand-maintained version of this workbook.

Run:  python3 generate_scenario_demo.py
"""

import math

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# --------------------------------------------------------------------------
# Corporate "Classic Navy" palette (matches household_portfolio_engine.xlsx)
# --------------------------------------------------------------------------
NAVY = "1F3864"
WHITE = "FFFFFF"
ICE_BLUE = "DCE6F1"
LIGHT_GRAY = "BFBFBF"
TITLE_GRAY = "44546A"
INPUT_YELLOW = "FFFDE7"

SOFT_GREEN_FILL = "E2EFDA"
DARK_GREEN_FONT = "375623"
SOFT_YELLOW_FILL = "FFF2CC"
DARK_BROWN_FONT = "7F6000"
SOFT_RED_FILL = "FCE4E4"
DARK_RED_FONT = "C00000"

FONT_NAME = "Segoe UI"

CURRENCY_FMT = '$#,##0'
PERCENT_FMT = '0.00%'

thin_border = Border(
    left=Side(style="thin", color=LIGHT_GRAY),
    right=Side(style="thin", color=LIGHT_GRAY),
    top=Side(style="thin", color=LIGHT_GRAY),
    bottom=Side(style="thin", color=LIGHT_GRAY),
)
year_boundary_border = Border(
    left=Side(style="thin", color=LIGHT_GRAY),
    right=Side(style="thin", color=LIGHT_GRAY),
    top=Side(style="medium", color=NAVY),
    bottom=Side(style="thin", color=LIGHT_GRAY),
)

header_font = Font(name=FONT_NAME, bold=True, color=WHITE, size=11)
header_fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
section_font = Font(name=FONT_NAME, bold=True, color=WHITE, size=11)
section_fill = PatternFill(start_color=TITLE_GRAY, end_color=TITLE_GRAY, fill_type="solid")
accent_fill = PatternFill(start_color=ICE_BLUE, end_color=ICE_BLUE, fill_type="solid")
input_fill = PatternFill(start_color=INPUT_YELLOW, end_color=INPUT_YELLOW, fill_type="solid")
body_font = Font(name=FONT_NAME, size=11)
bold_body_font = Font(name=FONT_NAME, size=11, bold=True)
label_font = Font(name=FONT_NAME, size=11, bold=True)
italic_note_font = Font(name=FONT_NAME, size=9, italic=True, color="595959")

green_fill = PatternFill(start_color=SOFT_GREEN_FILL, end_color=SOFT_GREEN_FILL, fill_type="solid")
yellow_fill = PatternFill(start_color=SOFT_YELLOW_FILL, end_color=SOFT_YELLOW_FILL, fill_type="solid")
red_fill = PatternFill(start_color=SOFT_RED_FILL, end_color=SOFT_RED_FILL, fill_type="solid")
green_font = Font(name=FONT_NAME, size=10, color=DARK_GREEN_FONT, bold=True)
yellow_font = Font(name=FONT_NAME, size=10, color=DARK_BROWN_FONT, bold=True)
red_font = Font(name=FONT_NAME, size=10, color=DARK_RED_FONT, bold=True)

LEGEND_TEXT = ("Legend: pale-yellow cells are yours to type into (pre-filled here with illustrative "
               "example numbers); white cells hold formulas or fixed scenario assumptions -- please "
               "don't overwrite them.")


def style_header_row(ws, row, first_col, last_col):
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_body_cell(ws, row, col, number_format=None, bold=False, fill=None, border=None):
    cell = ws.cell(row=row, column=col)
    cell.font = bold_body_font if bold else body_font
    cell.border = border or thin_border
    if number_format:
        cell.number_format = number_format
    if fill:
        cell.fill = fill
    return cell


def mark_input(cell):
    cell.fill = input_fill
    return cell


def autofit_columns(ws, min_width=10, max_width=60, padding=3, skip_rows=None):
    """Auto-size columns to their longest value. Merged-anchor cells that span
    multiple columns (section banners, wrapped paragraph cells) and rows in
    skip_rows are excluded so a single long banner/note doesn't force every
    row's column artificially wide -- those get explicit widths instead."""
    skip_rows = skip_rows or set()
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            if cell.row in skip_rows:
                continue
            is_multi_col_merge_anchor = any(
                rng.min_row <= cell.row <= rng.max_row
                and rng.min_col <= cell.column <= rng.max_col
                and (rng.max_col - rng.min_col) > 0
                and cell.coordinate == rng.coord.split(":")[0]
                for rng in ws.merged_cells.ranges
            )
            if is_multi_col_merge_anchor:
                continue
            length = len(str(cell.value))
            col_letter = cell.column_letter
            widths[col_letter] = max(widths.get(col_letter, 0), length)
    for col_letter, length in widths.items():
        ws.column_dimensions[col_letter].width = max(min_width, min(max_width, length + padding))


def enable_gridlines(ws):
    ws.sheet_view.showGridLines = True


def write_paragraph(ws, row, text, col=2, chars_per_line=80):
    """Write a wrapped prose paragraph into a single cell and size the row
    tall enough to show it without clipping (conservative estimate --
    Excel/Sheets don't auto-grow row height for wrapped text on import)."""
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = body_font
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    lines = max(1, math.ceil(len(text) / chars_per_line))
    ws.row_dimensions[row].height = lines * 15 + 8
    return cell


def write_section_banner(ws, row, text, end_col):
    ws.cell(row=row, column=1, value=text)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    for col in range(1, end_col + 1):
        c = ws.cell(row=row, column=col)
        c.font = section_font
        c.fill = section_fill
        c.border = thin_border
    ws.row_dimensions[row].height = 20


# ==========================================================================
YEARS = list(range(2026, 2036))  # 2026..2035 inclusive (10 years)
MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_COLS = [get_column_letter(6 + i) for i in range(12)]  # F..Q

# Hardcoded, illustrative-only annual market-return paths. These are hand-
# picked to make the Refill Engine's GOOD/POOR trigger visibly fire in both
# directions -- they are not historical data, forecasts, or investment advice.
BEAR_RETURNS = [-0.35, -0.08, 0.06, 0.09, 0.07, 0.07, 0.07, 0.06, 0.06, 0.06]
BULL_RETURNS = [0.10, 0.11, 0.09, 0.12, 0.10, 0.08, 0.11, 0.09, 0.10, 0.09]
CHOPPY_RETURNS = [0.09, -0.01, 0.09, -0.01, 0.09, -0.01, 0.09, -0.01, 0.09, -0.01]

# Generic illustrative monthly TSP -> checking payment, rising slightly year
# over year -- purely a round-number placeholder, not tied to any real budget.
ANNUAL_MONTHLY_PAYMENT = [4000, 4100, 4200, 4300, 4400, 4500, 4600, 4700, 4800, 4900]

wb = Workbook()

# --------------------------------------------------------------------------
# TAB 0: Documentation
# --------------------------------------------------------------------------
ws0 = wb.active
ws0.title = "Documentation"
DOC_CHARS_PER_LINE = 78  # matches column B's width below

ws0.cell(row=1, column=1, value="HOUSEHOLD PORTFOLIO ENGINE — SCENARIO DEMO")
ws0.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
for col in (1, 2):
    c = ws0.cell(row=1, column=col)
    c.font = header_font
    c.fill = header_fill
    c.border = thin_border
ws0.row_dimensions[1].height = 24

r = 3
write_section_banner(ws0, r, "WHAT THIS WORKBOOK IS", 2)
r += 1
for para in (
    "This is a demonstration copy of the Household Portfolio Engine -- a two-bucket, decumulation-phase "
    "retirement income strategy built to protect a household from Sequence-of-Returns Risk (SRR). It uses "
    "entirely generic example balances and three hardcoded, hypothetical market-return paths (Bear / Bull "
    "/ Choppy) so you can watch the engine's core mechanic -- the Quarterly Refill Engine's GOOD/POOR "
    "trigger -- actually fire in both directions, instead of having to imagine it.",

    "The live version of this workbook is a hand-maintained tracker: real account balances and monthly "
    "payments get typed in and the sheet only ever compares today's actual wealth to a target. It never "
    "simulates a market move on its own. This demo workbook adds that missing piece -- a simulated "
    "household wealth path that evolves year by year under whichever scenario you select -- purely to "
    "illustrate how the mechanism behaves under different market conditions.",
):
    write_paragraph(ws0, r, para, chars_per_line=DOC_CHARS_PER_LINE)
    r += 1
r += 1

write_section_banner(ws0, r, "HOW TO USE THE SCENARIO SELECTOR", 2)
r += 1
for para in (
    "On The_Refill_Engine tab, the top-left cell lets you choose Bear, Bull, or Choppy from a dropdown. "
    "Bear applies a front-loaded market crash in years 1-2 followed by a recovery -- watch the Quarterly "
    "Refill Engine freeze transfers and the G Fund shield draw down on its own. Bull applies steady, "
    "consistently strong growth -- the engine should recommend a transfer almost every quarter. Choppy "
    "applies small, oscillating returns near the hurdle rate -- GOOD and POOR should alternate "
    "unpredictably, which is the case the Safety Anchor is really built for.",

    "Every other input (account balances, COLA, monthly payments) is pre-filled with round, generic "
    "numbers so the workbook is fully self-contained -- nothing needs to be typed in to see it work. Feel "
    "free to change any pale-yellow cell to explore \"what if\" variations.",
):
    write_paragraph(ws0, r, para, chars_per_line=DOC_CHARS_PER_LINE)
    r += 1
r += 1

write_section_banner(ws0, r, "DISCLAIMER", 2)
r += 1
write_paragraph(
    ws0, r,
    "All balances, payments, and the Bear/Bull/Choppy return paths in this workbook are hypothetical and "
    "for illustration only. They are not historical market data, a forecast, or investment, tax, or legal "
    "advice. Use this workbook to understand the mechanics of the strategy, then work from your own real "
    "numbers in the live tracker.",
    chars_per_line=DOC_CHARS_PER_LINE,
)
r += 2

write_section_banner(ws0, r, "HOW THE FIVE TABS WORK TOGETHER", 2)
r += 1
for para in (
    "Documentation (this tab) -- What this demo shows, how to use the scenario selector, the disclaimer "
    "above, the color legend, and a glossary. Nothing here feeds any formula.",

    "Variables_And_Log -- The annual COLA/hurdle inputs, the monthly TSP -> checking payment log, and the "
    "three hardcoded Bear/Bull/Choppy annual return columns, plus a column that picks out whichever "
    "scenario is currently selected.",

    "Account_Tracker -- Generic starting balances for every account as of the beginning of Year 1. Unlike "
    "the live tracker, this is not an ongoing snapshot -- it's the starting point that The_Refill_Engine "
    "projects forward under the selected scenario.",

    "The_Refill_Engine -- The decision-making tab, plus the scenario selector. A block of configured "
    "inputs (Retirement Start Year, Ramp Length, the G Fund target percentages, etc.) each carry a "
    "Default, an optional Override, and the resulting Effective value everything else on the tab actually "
    "uses. The Quarterly Refill Engine table is the headline output. The Annual SRR Hurdle Model below it "
    "simulates actual household wealth forward year by year under the selected scenario and compares it "
    "to the hurdle target, which is what drives GOOD vs. POOR here.",

    "Graph -- A single chart, read-only, pulling straight from The_Refill_Engine's Annual SRR Hurdle "
    "Model: Simulated Actual Household Wealth vs. the hurdle target (P_track), by year, for whichever "
    "scenario is currently selected.",
):
    write_paragraph(ws0, r, para, chars_per_line=DOC_CHARS_PER_LINE)
    r += 1
r += 1

write_section_banner(ws0, r, "COLOR LEGEND", 2)
r += 1
legend_rows = [
    (input_fill, "Pale yellow -- an editable input, pre-filled with an illustrative example value."),
    (None, "White (no fill) -- a formula, or a fixed scenario assumption; please leave it alone."),
    (accent_fill, "Ice blue -- a computed summary or headline total."),
    (green_fill, "Green (Emergency Liquidity Protocol only) -- the TSP G Fund shield is active."),
    (yellow_fill, "Yellow (Emergency Liquidity Protocol only) -- the G Fund is exhausted; drawing from "
                  "the cash buffer."),
    (red_fill, "Red (Emergency Liquidity Protocol only) -- both cash shields are exhausted; the household "
               "is in an emergency-liquidation stage."),
]
for fill, text in legend_rows:
    swatch = ws0.cell(row=r, column=1)
    swatch.border = thin_border
    if fill:
        swatch.fill = fill
    write_paragraph(ws0, r, text, chars_per_line=DOC_CHARS_PER_LINE)
    r += 1
r += 1

write_section_banner(ws0, r, "GLOSSARY", 2)
r += 1
ws0.cell(row=r, column=1, value="Term")
ws0.cell(row=r, column=2, value="Definition")
style_header_row(ws0, r, 1, 2)
r += 1
GLOSSARY = [
    ("COLA", "The federal Cost-of-Living Adjustment, applied annually to Social Security/federal "
             "retirement benefits."),
    ("Hurdle Rate", "The minimum real (inflation-adjusted) growth rate the portfolio needs to clear to be "
                     "considered \"on track\" -- COLA layered on a fixed real-growth baseline."),
    ("W_t / W_q", "Total dollar outflows (withdrawals) for a given year (t) or quarter (q)."),
    ("P_track", "The simulated target portfolio value (the hurdle trajectory) that simulated actual "
                "wealth is measured against each year."),
    ("Scenario", "One of three hardcoded, hypothetical annual market-return paths -- Bear, Bull, or "
                 "Choppy -- selected on The_Refill_Engine tab."),
    ("Simulated Actual Household Wealth", "This demo's stand-in for real net worth: last year's simulated "
                                           "wealth, grown by the selected scenario's return for the year, "
                                           "less that year's outflows. Compared against P_track to decide "
                                           "GOOD vs. POOR."),
    ("SRR", "Sequence-of-Returns Risk -- the risk that a downturn early in retirement forces permanent "
            "sale of depreciated assets."),
    ("TSP G Fund", "The federal Thrift Savings Plan's government-securities fund: principal-protected, "
                    "instantly liquid, and the household's operating cash vault."),
    ("Safety Anchor", "The G Fund cash buffer this engine tries to maintain. Ramps linearly (on "
                       "The_Refill_Engine) from a Starting % down to an Ending % over the Ramp Length, "
                       "then holds flat -- largest during the early years when SRR risk is heavily "
                       "front-loaded, easing as that window passes. Applied each year to that year's "
                       "P_track rather than a frozen snapshot."),
    ("Ramp Length", "How many years the G Fund target percentage takes to slide from its Starting % to "
                     "its Ending %, before holding flat. A configured input on The_Refill_Engine."),
    ("Default / Override / Effective", "The three-column pattern used for every configured input on "
                                        "The_Refill_Engine. Default is the built-in value; Override is "
                                        "blank unless you want to change it; Effective is whichever one "
                                        "is actually in force, and is what every other formula on the tab "
                                        "references."),
    ("Conduit", "A Traditional IRA used as a pass-through for pre-tax rollovers back into the TSP."),
    ("Allocation Mix", "Each account's share of total household net worth, shown on Account_Tracker."),
]
for term, definition in GLOSSARY:
    term_cell = ws0.cell(row=r, column=1, value=term)
    term_cell.font = bold_body_font
    term_cell.border = thin_border
    term_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    def_cell = ws0.cell(row=r, column=2, value=definition)
    def_cell.font = body_font
    def_cell.border = thin_border
    def_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    lines = max(1, math.ceil(len(definition) / DOC_CHARS_PER_LINE))
    ws0.row_dimensions[r].height = lines * 15 + 8
    r += 1

ws0.column_dimensions["A"].width = 22
ws0.column_dimensions["B"].width = 90
enable_gridlines(ws0)
ws0.freeze_panes = "A2"

# --------------------------------------------------------------------------
# TAB 1: Variables_And_Log
# --------------------------------------------------------------------------
ws1 = wb.create_sheet("Variables_And_Log")

BANNER_ROW_1 = 1
ws1.cell(row=BANNER_ROW_1, column=1,
          value="HOUSEHOLD VARIABLE LOG — ANNUAL HURDLE INPUTS, MARKET SCENARIOS & MONTHLY "
                "TSP → CHECKING PAYMENT LOG (2026-2035)")
LAST_COL_1 = 6 + 12 + 4  # 5 lead cols + 12 months + 4 scenario cols
ws1.merge_cells(start_row=BANNER_ROW_1, start_column=1, end_row=BANNER_ROW_1, end_column=LAST_COL_1)
for col in range(1, LAST_COL_1 + 1):
    c = ws1.cell(row=BANNER_ROW_1, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border
ws1.row_dimensions[BANNER_ROW_1].height = 20

HEADER_ROW_1 = 2
ws1_headers = [
    "Date / Year",
    "Actual Federal COLA (C_i)",
    "Rule of Four Hurdle Baseline",
    "Total Annual Hurdle Rate",
    "Total Annual Outflows (W_t)",
] + MONTH_NAMES + [
    "Bear Scenario Annual Return (Illustrative)",
    "Bull Scenario Annual Return (Illustrative)",
    "Choppy Scenario Annual Return (Illustrative)",
    "Selected Scenario Annual Return (This Year)",
]
for idx, h in enumerate(ws1_headers, start=1):
    ws1.cell(row=HEADER_ROW_1, column=idx, value=h)
style_header_row(ws1, HEADER_ROW_1, 1, len(ws1_headers))
ws1.row_dimensions[HEADER_ROW_1].height = 32

SCENARIO_COL_BEAR = 6 + 12      # R
SCENARIO_COL_BULL = SCENARIO_COL_BEAR + 1   # S
SCENARIO_COL_CHOPPY = SCENARIO_COL_BEAR + 2  # T
SCENARIO_COL_SELECTED = SCENARIO_COL_BEAR + 3  # U
SCENARIO_SELECTOR_CELL = "'The_Refill_Engine'!$B$1"

# Illustrative COLA path -- a plain example series, not a forecast.
COLA_PATH = [0.026, 0.025, 0.024, 0.025, 0.026, 0.025, 0.024, 0.025, 0.026, 0.025]

DATA_START_ROW_1 = HEADER_ROW_1 + 1  # row 3
for i, year in enumerate(YEARS):
    r = DATA_START_ROW_1 + i
    style_body_cell(ws1, r, 1, number_format="0").value = year
    mark_input(style_body_cell(ws1, r, 2, number_format=PERCENT_FMT)).value = COLA_PATH[i]
    style_body_cell(ws1, r, 3, number_format=PERCENT_FMT).value = 0.020
    style_body_cell(ws1, r, 4, number_format=PERCENT_FMT).value = f"=(1+B{r})*(1+C{r})"
    month_first, month_last = MONTH_COLS[0], MONTH_COLS[-1]
    style_body_cell(ws1, r, 5, number_format=CURRENCY_FMT).value = f"=SUM({month_first}{r}:{month_last}{r})"
    for m, col_letter in enumerate(MONTH_COLS):
        mark_input(style_body_cell(ws1, r, 6 + m, number_format=CURRENCY_FMT)).value = ANNUAL_MONTHLY_PAYMENT[i]

    style_body_cell(ws1, r, SCENARIO_COL_BEAR, number_format=PERCENT_FMT).value = BEAR_RETURNS[i]
    style_body_cell(ws1, r, SCENARIO_COL_BULL, number_format=PERCENT_FMT).value = BULL_RETURNS[i]
    style_body_cell(ws1, r, SCENARIO_COL_CHOPPY, number_format=PERCENT_FMT).value = CHOPPY_RETURNS[i]
    bear_letter = get_column_letter(SCENARIO_COL_BEAR)
    bull_letter = get_column_letter(SCENARIO_COL_BULL)
    choppy_letter = get_column_letter(SCENARIO_COL_CHOPPY)
    selected_formula = (
        f'=IF({SCENARIO_SELECTOR_CELL}="Bear",{bear_letter}{r},'
        f'IF({SCENARIO_SELECTOR_CELL}="Bull",{bull_letter}{r},{choppy_letter}{r}))'
    )
    style_body_cell(ws1, r, SCENARIO_COL_SELECTED, number_format=PERCENT_FMT, fill=accent_fill).value = selected_formula

LAST_DATA_ROW_1 = DATA_START_ROW_1 + len(YEARS) - 1  # row 12 (2035)

note_row_1 = LAST_DATA_ROW_1 + 2
ws1.cell(row=note_row_1, column=1,
          value="Notes: COLA, monthly payments, and the account balances on Account_Tracker are all "
                "generic illustrative numbers -- change any pale-yellow cell to explore a variation. The "
                "Bear/Bull/Choppy columns are fixed scenario definitions (hypothetical, not historical); "
                "the selector that picks among them lives on The_Refill_Engine tab.")
ws1.cell(row=note_row_1, column=1).font = italic_note_font

legend_row_1 = note_row_1 + 1
ws1.cell(row=legend_row_1, column=1, value=LEGEND_TEXT)
ws1.cell(row=legend_row_1, column=1).font = italic_note_font

enable_gridlines(ws1)
ws1.freeze_panes = f"A{DATA_START_ROW_1}"

# --------------------------------------------------------------------------
# TAB 2: Account_Tracker
# --------------------------------------------------------------------------
ws2 = wb.create_sheet("Account_Tracker")

ws2_headers = ["Asset Pool", "Custodian / Platform", "Starting Balance (Beginning of Year 1)", "Allocation Mix"]
for idx, h in enumerate(ws2_headers, start=1):
    ws2.cell(row=1, column=idx, value=h)
style_header_row(ws2, 1, 1, 4)
ws2.row_dimensions[1].height = 24

# Generic PLACEHOLDER balances only -- illustrative, round numbers, not any
# real account figures. This file is meant to be shared publicly.
ACCOUNTS = [
    ("Your Pre-Tax Account", "TSP G Fund", 100_000),
    ("Your Conversion Conduit", "Private Traditional IRA", 100_000),
    ("Your Tax-Free Growth", "Private Roth IRA", 100_000),
    ("Spouse's Tax-Free Growth", "Private Roth IRA", 100_000),
    ("Spouse's Safe Pre-Tax Account", "Current Employer 403b (Active + Reverse Rollover)", 100_000),
    ("Spouse's Conversion Conduit", "Private Traditional IRA", 100_000),
    ("Short-Term Liquidity", "Merrill Edge Cash Buffer", 100_000),
]

DATA_START_ROW_2 = 2
for i, (pool, custodian, balance) in enumerate(ACCOUNTS):
    r = DATA_START_ROW_2 + i
    style_body_cell(ws2, r, 1).value = pool
    style_body_cell(ws2, r, 2).value = custodian
    mark_input(style_body_cell(ws2, r, 3, number_format=CURRENCY_FMT)).value = balance
    style_body_cell(ws2, r, 4, number_format=PERCENT_FMT)  # formula added after total row known

LAST_DATA_ROW_2 = DATA_START_ROW_2 + len(ACCOUNTS) - 1  # row 8
TOTAL_ROW_2 = LAST_DATA_ROW_2 + 1  # row 9

ws2.cell(row=TOTAL_ROW_2, column=1, value="Total Household Net Worth")
ws2.cell(row=TOTAL_ROW_2, column=3, value=f"=SUM(C{DATA_START_ROW_2}:C{LAST_DATA_ROW_2})")
ws2.cell(row=TOTAL_ROW_2, column=4, value=f"=C{TOTAL_ROW_2}/C{TOTAL_ROW_2}")
for col in range(1, 5):
    cell = ws2.cell(row=TOTAL_ROW_2, column=col)
    cell.font = bold_body_font
    cell.fill = accent_fill
    cell.border = thin_border
ws2.cell(row=TOTAL_ROW_2, column=3).number_format = CURRENCY_FMT
ws2.cell(row=TOTAL_ROW_2, column=4).number_format = PERCENT_FMT

TOTAL_NET_WORTH_REF = f"'Account_Tracker'!$C${TOTAL_ROW_2}"
for i in range(len(ACCOUNTS)):
    r = DATA_START_ROW_2 + i
    ws2.cell(row=r, column=4, value=f"=C{r}/$C${TOTAL_ROW_2}")

note_row_2 = TOTAL_ROW_2 + 1
ws2.cell(row=note_row_2, column=1,
          value="Note: unlike the live tracker, these are not an ongoing snapshot -- they're the "
                "starting point (beginning of Year 1) that The_Refill_Engine projects forward under "
                "whichever scenario is selected.")
ws2.cell(row=note_row_2, column=1).font = italic_note_font

legend_row_2 = note_row_2 + 1
ws2.cell(row=legend_row_2, column=1, value=LEGEND_TEXT)
ws2.cell(row=legend_row_2, column=1).font = italic_note_font

G_FUND_BALANCE_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2}"
PRIVATE_TRAD_IRA_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 1}"
SPOUSE_403B_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 4}"
MERRILL_CASH_BUFFER_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 6}"

enable_gridlines(ws2)
ws2.freeze_panes = "A2"

# --------------------------------------------------------------------------
# TAB 3: The_Refill_Engine
# --------------------------------------------------------------------------
ws3 = wb.create_sheet("The_Refill_Engine")


def write_kv_row(row, label, value, number_format, is_input=False, is_summary=False):
    label_cell = ws3.cell(row=row, column=1, value=label)
    label_cell.font = label_font
    label_cell.border = thin_border
    val_cell = ws3.cell(row=row, column=2, value=value)
    val_cell.number_format = number_format
    val_cell.font = bold_body_font
    val_cell.border = thin_border
    if is_summary:
        label_cell.fill = accent_fill
        val_cell.fill = accent_fill
    if is_input:
        mark_input(val_cell)
    return val_cell


def write_config_row(row, label, default_value, number_format):
    """A configured input laid out as Default (col B) / Override (col C,
    blank = use Default) / Effective (col D, the absolute ref every other
    formula on this sheet should point at). Matches household_portfolio_
    engine.xlsx's The_Refill_Engine layout."""
    label_cell = ws3.cell(row=row, column=1, value=label)
    label_cell.font = label_font
    label_cell.border = thin_border

    default_cell = ws3.cell(row=row, column=2, value=default_value)
    default_cell.number_format = number_format
    default_cell.font = body_font
    default_cell.border = thin_border

    override_cell = ws3.cell(row=row, column=3)
    override_cell.number_format = number_format
    override_cell.font = body_font
    override_cell.border = thin_border
    mark_input(override_cell)

    effective_cell = ws3.cell(row=row, column=4, value=f'=IF(C{row}="",B{row},C{row})')
    effective_cell.number_format = number_format
    effective_cell.font = bold_body_font
    effective_cell.fill = accent_fill
    effective_cell.border = thin_border
    return f"$D${row}"


SCENARIO_ROW = 1
write_kv_row(SCENARIO_ROW, "Scenario Selector", "Bear", "@", is_input=True)
scenario_dv = DataValidation(type="list", formula1='"Bear,Bull,Choppy"', allow_blank=False)
ws3.add_data_validation(scenario_dv)
scenario_dv.add(ws3.cell(row=SCENARIO_ROW, column=2))
ws3.cell(row=SCENARIO_ROW, column=3,
          value="← choose Bear / Bull / Choppy to see the engine react to each hardcoded market path")
ws3.cell(row=SCENARIO_ROW, column=3).font = italic_note_font
ws3.merge_cells(start_row=SCENARIO_ROW, start_column=3, end_row=SCENARIO_ROW, end_column=9)

r = 3
write_kv_row(r, "Starting Household Wealth (Year 1 Opening)", f"={TOTAL_NET_WORTH_REF}", CURRENCY_FMT, is_summary=True)
STARTING_WEALTH_ROW = r
r += 2  # blank spacer row

write_section_banner(ws3, r, "CONFIGURED INPUTS — DEFAULT / OVERRIDE / EFFECTIVE", 4)
r += 1
config_header_row = r
for idx, h in enumerate(["Input", "Default", "Override (blank = use Default)", "Effective (used below)"], start=1):
    ws3.cell(row=config_header_row, column=idx, value=h)
style_header_row(ws3, config_header_row, 1, 4)
ws3.row_dimensions[config_header_row].height = 28
r += 1

RETIREMENT_START_CELL = write_config_row(r, "Retirement Start Year", 2026, "0")
r += 1
AGE_AT_START_CELL = write_config_row(r, "Age at Retirement Start", 63, "0")
r += 1
RAMP_LENGTH_CELL = write_config_row(r, "Ramp Length (Years)", 10, "0")
r += 1
STARTING_TARGET_CELL = write_config_row(
    r, "Starting G Fund Target % (elevated, early retirement)", 0.30, PERCENT_FMT)
r += 1
ENDING_TARGET_CELL = write_config_row(r, "Ending G Fund Target % (steady state)", 0.20, PERCENT_FMT)
r += 1
r += 1  # blank spacer row

write_kv_row(r, "Current G Fund Target % (Today)",
             f"=IF(YEAR(TODAY())-{RETIREMENT_START_CELL}>={RAMP_LENGTH_CELL},{ENDING_TARGET_CELL},"
             f"{STARTING_TARGET_CELL}+({ENDING_TARGET_CELL}-{STARTING_TARGET_CELL})*"
             f"(YEAR(TODAY())-{RETIREMENT_START_CELL})/{RAMP_LENGTH_CELL})",
             PERCENT_FMT, is_summary=True)
CURRENT_TARGET_PCT_ROW = r
r += 1
write_kv_row(r, "Target Safety Anchor Balance (Year 1 Opening)",
             f"=B{STARTING_WEALTH_ROW}*B{CURRENT_TARGET_PCT_ROW}", CURRENCY_FMT, is_summary=True)
r += 1

ws3.cell(row=r, column=1,
          value="Sequence-of-Returns-Risk (SRR) Shield: the G Fund target ramps down linearly from the "
                "Starting % to the Ending % over the Ramp Length, then holds flat -- SRR is concentrated in "
                "the first years of retirement, so the shield is largest when it matters most and eases as "
                "that window passes. Each year's actual target is computed per-year in the Annual SRR Hurdle "
                "Model below and drives the Quarterly Refill Engine's 'Target Safety Anchor Balance (This "
                "Year)' column, scaled against that year's simulated wealth path, not a value frozen at "
                "Year 1. The row above is only a Year-1 snapshot for quick reference. In this demo, GOOD/"
                "POOR below is driven by the Scenario Selector above, not a live account snapshot.")
ws3.cell(row=r, column=1).font = italic_note_font
CONFIG_NOTE_ROW = r
r += 2  # blank spacer row

# ---- THE QUARTERLY REFILL ENGINE (primary output) ----
QUARTERLY_BANNER_ROW = r
ws3.cell(row=QUARTERLY_BANNER_ROW, column=1,
          value="THE QUARTERLY REFILL ENGINE — PRIMARY OUTPUT: RECOMMENDED TSP TOP-UP TRANSFER")
ws3.merge_cells(start_row=QUARTERLY_BANNER_ROW, start_column=1, end_row=QUARTERLY_BANNER_ROW, end_column=9)
for col in range(1, 10):
    c = ws3.cell(row=QUARTERLY_BANNER_ROW, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border
ws3.row_dimensions[QUARTERLY_BANNER_ROW].height = 20

QUARTERLY_HEADER_ROW = QUARTERLY_BANNER_ROW + 1
quarterly_headers = [
    "Year",
    "Quarter",
    "Quarterly Outflow (W_q)",
    "Annual Performance Signal",
    "G Fund Balance (Start of Qtr)",
    "Recommended Transfer This Quarter",
    "Actual Transferred (Log It Here)",
    "G Fund Balance (End of Qtr, Projected)",
    "Target Safety Anchor Balance (This Year)",
]
for idx, h in enumerate(quarterly_headers, start=1):
    ws3.cell(row=QUARTERLY_HEADER_ROW, column=idx, value=h)
style_header_row(ws3, QUARTERLY_HEADER_ROW, 1, len(quarterly_headers))
ws3.row_dimensions[QUARTERLY_HEADER_ROW].height = 40

QUARTER_LABELS = ["Q1", "Q2", "Q3", "Q4"]
QUARTER_MONTH_SLICES = [(0, 3), (3, 6), (6, 9), (9, 12)]

QUARTERLY_DATA_START = QUARTERLY_HEADER_ROW + 1
ANNUAL_BANNER_ROW = QUARTERLY_DATA_START + 4 * len(YEARS) - 1 + 3
ANNUAL_HEADER_ROW = ANNUAL_BANNER_ROW + 1
ANNUAL_DATA_START = ANNUAL_HEADER_ROW + 1

for i, year in enumerate(YEARS):
    var_log_row = DATA_START_ROW_1 + i
    annual_row = ANNUAL_DATA_START + i
    for q, qlabel in enumerate(QUARTER_LABELS):
        r = QUARTERLY_DATA_START + 4 * i + q
        border = year_boundary_border if q == 0 else thin_border

        style_body_cell(ws3, r, 1, number_format="0", border=border).value = year
        style_body_cell(ws3, r, 2, border=border).value = qlabel

        m0, m1 = QUARTER_MONTH_SLICES[q]
        col_first, col_last = MONTH_COLS[m0], MONTH_COLS[m1 - 1]
        style_body_cell(ws3, r, 3, number_format=CURRENCY_FMT, border=border).value = (
            f"=SUM('Variables_And_Log'!{col_first}{var_log_row}:{col_last}{var_log_row})"
        )

        d_cell = style_body_cell(ws3, r, 4, border=border)
        d_cell.value = f"=G{annual_row}"
        d_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        e_cell = style_body_cell(ws3, r, 5, number_format=CURRENCY_FMT, border=border)
        e_cell.value = f"={G_FUND_BALANCE_CELL}" if r == QUARTERLY_DATA_START else f"=H{r - 1}"

        i_cell = style_body_cell(ws3, r, 9, number_format=CURRENCY_FMT, border=border)
        i_cell.value = f"=H{annual_row}*D{annual_row}"

        f_cell = style_body_cell(ws3, r, 6, number_format=CURRENCY_FMT, bold=True,
                                   fill=accent_fill, border=border)
        f_cell.value = f'=IF(ISNUMBER(SEARCH("GOOD",D{r})),MAX(0,I{r}-E{r}),0)'

        g_cell = mark_input(style_body_cell(ws3, r, 7, number_format=CURRENCY_FMT, border=border))

        h_cell = style_body_cell(ws3, r, 8, number_format=CURRENCY_FMT, border=border)
        h_cell.value = f'=MAX(0,E{r}-C{r}+IF(G{r}="",F{r},G{r}))'

        ws3.row_dimensions[r].height = 28

QUARTERLY_DATA_END = QUARTERLY_DATA_START + 4 * len(YEARS) - 1
assert ANNUAL_BANNER_ROW == QUARTERLY_DATA_END + 3, "ANNUAL_BANNER_ROW constant must match actual layout"

# ---- ANNUAL SRR HURDLE MODEL (gates the quarterly engine above) ----
ws3.cell(row=ANNUAL_BANNER_ROW, column=1,
          value="ANNUAL SRR HURDLE MODEL — SIMULATED WEALTH VS. HURDLE (GATES THE QUARTERLY REFILL ABOVE)")
ws3.merge_cells(start_row=ANNUAL_BANNER_ROW, start_column=1, end_row=ANNUAL_BANNER_ROW, end_column=8)
for col in range(1, 9):
    c = ws3.cell(row=ANNUAL_BANNER_ROW, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border
ws3.row_dimensions[ANNUAL_BANNER_ROW].height = 20

annual_headers = [
    "Year",
    "Total Annual Hurdle Rate",
    "Annual Outflows (W_t)",
    "Simulated Target Portfolio (P_track)",
    "Scenario Return (Selected, This Year)",
    "Simulated Actual Household Wealth (This Year)",
    "Performance Logic Trigger",
    "Safety Anchor % (This Year)",
]
for idx, h in enumerate(annual_headers, start=1):
    ws3.cell(row=ANNUAL_HEADER_ROW, column=idx, value=h)
style_header_row(ws3, ANNUAL_HEADER_ROW, 1, len(annual_headers))
ws3.row_dimensions[ANNUAL_HEADER_ROW].height = 40

for i, year in enumerate(YEARS):
    r = ANNUAL_DATA_START + i
    var_log_row = DATA_START_ROW_1 + i

    style_body_cell(ws3, r, 1, number_format="0").value = year
    style_body_cell(ws3, r, 2, number_format=PERCENT_FMT).value = f"='Variables_And_Log'!D{var_log_row}"
    style_body_cell(ws3, r, 3, number_format=CURRENCY_FMT).value = f"='Variables_And_Log'!E{var_log_row}"

    if i == 0:
        target_formula = f"=ROUND(($B$3*B{r})-C{r},2)"
    else:
        target_formula = f"=ROUND((D{r - 1}*B{r})-C{r},2)"
    style_body_cell(ws3, r, 4, number_format=CURRENCY_FMT).value = target_formula

    style_body_cell(ws3, r, 5, number_format=PERCENT_FMT).value = (
        f"='Variables_And_Log'!{get_column_letter(SCENARIO_COL_SELECTED)}{var_log_row}"
    )

    if i == 0:
        actual_formula = f"=ROUND(($B$3*(1+E{r}))-C{r},2)"
    else:
        actual_formula = f"=ROUND((F{r - 1}*(1+E{r}))-C{r},2)"
    style_body_cell(ws3, r, 6, number_format=CURRENCY_FMT, fill=accent_fill, bold=True).value = actual_formula

    trigger_formula = (
        f'=IF(F{r}>=D{r},'
        f'"✅ GOOD: Execute T2065 Rebalance & Rollover to TSP",'
        f'"⚠️ POOR: Freeze Equities. Draw Down G Fund")'
    )
    trig_cell = style_body_cell(ws3, r, 7)
    trig_cell.value = trigger_formula
    trig_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Linear ramp: Starting Target % at Retirement Start Year, sliding to Ending Target %
    # over Ramp Length years, then flat at Ending Target % forever after.
    style_body_cell(ws3, r, 8, number_format=PERCENT_FMT).value = (
        f"=IF(A{r}-{RETIREMENT_START_CELL}>={RAMP_LENGTH_CELL},{ENDING_TARGET_CELL},"
        f"{STARTING_TARGET_CELL}+({ENDING_TARGET_CELL}-{STARTING_TARGET_CELL})*"
        f"(A{r}-{RETIREMENT_START_CELL})/{RAMP_LENGTH_CELL})"
    )

    ws3.row_dimensions[r].height = 28

ANNUAL_DATA_END = ANNUAL_DATA_START + len(YEARS) - 1

# ---- EXTRA CAPACITY INDICATORS ----
KPI_HEADER_ROW = ANNUAL_DATA_END + 2
ws3.cell(row=KPI_HEADER_ROW, column=1, value="EXTRA CAPACITY INDICATORS")
ws3.merge_cells(start_row=KPI_HEADER_ROW, start_column=1, end_row=KPI_HEADER_ROW, end_column=3)
for col in range(1, 4):
    c = ws3.cell(row=KPI_HEADER_ROW, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border

CONDUIT_ROW = KPI_HEADER_ROW + 1
ws3.cell(row=CONDUIT_ROW, column=1, value="Your Pre-Tax Rollover Conduit Capacity")
ws3.cell(row=CONDUIT_ROW, column=1).font = label_font
ws3.cell(row=CONDUIT_ROW, column=1).border = thin_border
conduit_cell = ws3.cell(row=CONDUIT_ROW, column=2, value=f"={PRIVATE_TRAD_IRA_CELL}")
conduit_cell.number_format = CURRENCY_FMT
conduit_cell.font = body_font
conduit_cell.border = thin_border
conduit_status_cell = ws3.cell(row=CONDUIT_ROW, column=3,
                                 value=f'=IF(B{CONDUIT_ROW}>0,"Conduit Open","Conduit Closed")')
conduit_status_cell.font = bold_body_font
conduit_status_cell.fill = accent_fill
conduit_status_cell.border = thin_border

RUNWAY_ROW = CONDUIT_ROW + 2
ws3.cell(row=RUNWAY_ROW, column=1, value="TSP Cash Buffer Runway (Months)")
ws3.cell(row=RUNWAY_ROW, column=1).font = label_font
ws3.cell(row=RUNWAY_ROW, column=1).border = thin_border
current_year_months_range = f"'Variables_And_Log'!${MONTH_COLS[0]}${DATA_START_ROW_1}:${MONTH_COLS[-1]}${DATA_START_ROW_1}"
runway_cell = ws3.cell(
    row=RUNWAY_ROW, column=2,
    value=f'=IFERROR({G_FUND_BALANCE_CELL}/AVERAGE({current_year_months_range}),"Log Monthly Payments to Calculate")'
)
runway_cell.number_format = "0.0"
runway_cell.font = body_font
runway_cell.border = thin_border
runway_status_cell = ws3.cell(
    row=RUNWAY_ROW, column=3,
    value=f'=IF(ISNUMBER(B{RUNWAY_ROW}),IF(B{RUNWAY_ROW}>36,"Safe: 3+ Year Volatility Shield",'
          f'"Caution: Shield Shrinking"),"Awaiting Payment Data")'
)
runway_status_cell.font = bold_body_font
runway_status_cell.fill = accent_fill
runway_status_cell.border = thin_border

note_row_3 = RUNWAY_ROW + 1
ws3.cell(row=note_row_3, column=1,
          value="(Runway uses the average of 2026's monthly TSP payments, and the G Fund's Year-1 "
                "starting balance -- it does not reflect the quarter-by-quarter projection above.)")
ws3.cell(row=note_row_3, column=1).font = italic_note_font

# ---- EMERGENCY LIQUIDITY PROTOCOL ----
PROTOCOL_HEADER_ROW = note_row_3 + 2
ws3.cell(row=PROTOCOL_HEADER_ROW, column=1, value="EMERGENCY LIQUIDITY PROTOCOL — NEXT-DOLLAR PRIORITY ENGINE")
ws3.merge_cells(start_row=PROTOCOL_HEADER_ROW, start_column=1, end_row=PROTOCOL_HEADER_ROW, end_column=5)
for col in range(1, 6):
    c = ws3.cell(row=PROTOCOL_HEADER_ROW, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border

LABEL_ROW = PROTOCOL_HEADER_ROW + 1
ws3.cell(row=LABEL_ROW, column=1, value="Next-Dollar Liquidity Priority")
ws3.cell(row=LABEL_ROW, column=1).font = label_font
ws3.cell(row=LABEL_ROW, column=1).border = thin_border

OUTPUT_ROW_START = LABEL_ROW
OUTPUT_ROW_END = LABEL_ROW + 2
ws3.merge_cells(start_row=OUTPUT_ROW_START, start_column=2, end_row=OUTPUT_ROW_END, end_column=5)
protocol_cell = ws3.cell(row=OUTPUT_ROW_START, column=2)

PRIORITY_1_TEXT = (
    "\U0001F6E1️ SHIELD ACTIVE: Continue automated monthly draws directly from the TSP G Fund. "
    "Your equity engines are fully insulated from sequence-of-returns risk."
)
PRIORITY_2_TEXT = (
    "⚠️ SHIELD BROKEN: TSP G Fund is exhausted. IMMEDIATELY divert monthly checking draws to "
    "your Merrill Edge Cash Buffer. This buys next-day liquidity via Bank of America without forced "
    "retirement asset sales."
)
PRIORITY_3_TEXT = (
    "\U0001F6A8 EMERGENCY LIQUIDATION - STAGE 1: G Fund and Cash Buffers are exhausted. Turn on monthly "
    "distributions from Spouse's Current Employer 403b. Their pre-tax employer plan qualifies for the Maryland "
    "Pension Exclusion at age 65, making it more tax-efficient than touching your private traditional IRA."
)
PRIORITY_4_TEXT = (
    "\U0001F4A5 NUCLEAR PROTOCOL - STAGE 2: Absolute liquidity crisis. You are forced to break all shields. "
    "Begin selling shares of the 2065 Target Date Funds within your Private Traditional IRAs to fund monthly "
    "expenses. Maximize tax withholding at the time of distribution to shield your remaining capital from "
    "underpayment penalties."
)

protocol_formula = (
    f'=IF({G_FUND_BALANCE_CELL}>0,"{PRIORITY_1_TEXT}",'
    f'IF(AND({G_FUND_BALANCE_CELL}=0,{MERRILL_CASH_BUFFER_CELL}>0),"{PRIORITY_2_TEXT}",'
    f'IF(AND({G_FUND_BALANCE_CELL}=0,{MERRILL_CASH_BUFFER_CELL}=0,{SPOUSE_403B_CELL}>0),"{PRIORITY_3_TEXT}",'
    f'"{PRIORITY_4_TEXT}")))'
)
protocol_cell.value = protocol_formula
protocol_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
protocol_cell.font = Font(name=FONT_NAME, size=10)
protocol_cell.border = thin_border
for r in range(OUTPUT_ROW_START, OUTPUT_ROW_END + 1):
    for c in range(2, 6):
        ws3.cell(row=r, column=c).border = thin_border
ws3.row_dimensions[OUTPUT_ROW_START].height = 60
ws3.row_dimensions[OUTPUT_ROW_START + 1].height = 20
ws3.row_dimensions[OUTPUT_ROW_END].height = 20

legend_row_3 = OUTPUT_ROW_END + 2
ws3.cell(row=legend_row_3, column=1, value=LEGEND_TEXT)
ws3.cell(row=legend_row_3, column=1).font = italic_note_font

cf_range = f"B{OUTPUT_ROW_START}:B{OUTPUT_ROW_END}"

ws3.conditional_formatting.add(
    cf_range,
    FormulaRule(formula=[f"{G_FUND_BALANCE_CELL}>0"], fill=green_fill, font=green_font),
)
ws3.conditional_formatting.add(
    cf_range,
    FormulaRule(
        formula=[f"AND({G_FUND_BALANCE_CELL}=0,{MERRILL_CASH_BUFFER_CELL}>0)"],
        fill=yellow_fill, font=yellow_font,
    ),
)
ws3.conditional_formatting.add(
    cf_range,
    FormulaRule(
        formula=[f"AND({G_FUND_BALANCE_CELL}=0,{MERRILL_CASH_BUFFER_CELL}=0)"],
        fill=red_fill, font=red_font,
    ),
)

enable_gridlines(ws3)
ws3.freeze_panes = f"A{QUARTERLY_DATA_START}"

# --------------------------------------------------------------------------
# Autofit + explicit widths + final gridline pass for every tab
# --------------------------------------------------------------------------
autofit_columns(ws1, skip_rows={note_row_1, legend_row_1})
autofit_columns(ws2, skip_rows={note_row_2, legend_row_2})
autofit_columns(ws3, skip_rows={CONFIG_NOTE_ROW, note_row_3, legend_row_3})
for ws in (ws1, ws2, ws3):
    ws.sheet_view.showGridLines = True

ws1.column_dimensions["A"].width = 14
for col_letter in MONTH_COLS:
    ws1.column_dimensions[col_letter].width = 12
for col_idx in (SCENARIO_COL_BEAR, SCENARIO_COL_BULL, SCENARIO_COL_CHOPPY, SCENARIO_COL_SELECTED):
    ws1.column_dimensions[get_column_letter(col_idx)].width = 16

ws3.column_dimensions["A"].width = 42
ws3.column_dimensions["B"].width = 14
ws3.column_dimensions["C"].width = 22
ws3.column_dimensions["D"].width = 30
ws3.column_dimensions["E"].width = 22
ws3.column_dimensions["F"].width = 30
ws3.column_dimensions["G"].width = 22
ws3.column_dimensions["H"].width = 22
ws3.column_dimensions["I"].width = 22

# --------------------------------------------------------------------------
# TAB 4: Graph
# --------------------------------------------------------------------------
ws4 = wb.create_sheet("Graph")

ws4.cell(row=1, column=1, value="SIMULATED WEALTH VS. HURDLE — SELECTED SCENARIO")
ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
for col in range(1, 7):
    c = ws4.cell(row=1, column=col)
    c.font = header_font
    c.fill = header_fill
    c.border = thin_border
ws4.row_dimensions[1].height = 24

write_paragraph(
    ws4, 3,
    "This chart reads directly from the Annual SRR Hurdle Model table on The_Refill_Engine -- there is "
    "nothing to edit here. It redraws automatically for whichever scenario is selected there: watch "
    "Simulated Actual Household Wealth stay under the hurdle (P_track) for the whole Bear run, clear it "
    "every year under Bull, and cross back and forth under Choppy.",
    col=1, chars_per_line=100,
)

chart = BarChart()
chart.type = "col"
chart.grouping = "clustered"
chart.overlap = -10
chart.title = "Simulated Actual Household Wealth vs. Hurdle Target (P_track)"
chart.y_axis.title = "Household Wealth ($)"
chart.y_axis.numFmt = CURRENCY_FMT
chart.x_axis.title = "Year"
chart.height = 12
chart.width = 26

cats = Reference(ws3, min_col=1, min_row=ANNUAL_DATA_START, max_row=ANNUAL_DATA_END)
data_target = Reference(ws3, min_col=4, min_row=ANNUAL_HEADER_ROW, max_row=ANNUAL_DATA_END)
data_simulated = Reference(ws3, min_col=6, min_row=ANNUAL_HEADER_ROW, max_row=ANNUAL_DATA_END)

chart.add_data(data_target, titles_from_data=True)
chart.add_data(data_simulated, titles_from_data=True)
chart.set_categories(cats)

for series, color in zip(chart.series, (TITLE_GRAY, NAVY)):
    series.graphicalProperties.solidFill = color

ws4.add_chart(chart, "A6")

enable_gridlines(ws4)

wb.save("household_portfolio_scenario_demo.xlsx")
print("Wrote household_portfolio_scenario_demo.xlsx")
