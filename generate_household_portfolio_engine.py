"""
generate_household_portfolio_engine.py

Builds 'household_portfolio_engine.xlsx' -- a five-tab household retirement
income engine implementing a Sequence-of-Returns-Risk (SRR) shield strategy:

  1. Documentation       -- motivation, how the tabs fit together, an
                             annual checklist, the color legend, and a
                             glossary. Meant to be read first.
  2. Variables_And_Log   -- annual COLA/hurdle inputs, a Jan-Dec monthly
                             TSP -> checking payment log, an annual Roth
                             conversion tax-cost log, and an annual Actual
                             Balance log (2026-2035).
  3. Account_Tracker     -- unified household net worth across all accounts
  4. The_Refill_Engine   -- a wealth-adjustment block (isolates the tax cost
                             of Roth conversions from market performance), a
                             Default/Override/Effective block of configured
                             inputs, the Annual Refill Engine (the sheet's
                             primary output: how much to transfer between a
                             non-TSP account and the TSP G Fund each year),
                             an annual SRR hurdle model that gates
                             it (a linear G Fund target ramp, not a step
                             function), capacity KPIs, and the Emergency
                             Liquidity Protocol.
  5. Graph                -- a read-only chart (Original vs. Current vs.
                             Actual household wealth by year), kept off the
                             data-entry tabs.

This script is meant to be run ONCE to scaffold the workbook. The generated
.xlsx is then imported into Google Sheets and becomes the long-lived, hand-
maintained document -- real account balances and payment amounts get typed
in there, not re-generated here. Account_Tracker balances below are generic
placeholders on purpose; do not put real dollar figures into this script.

Run:        python3 generate_household_portfolio_engine.py
Verify:     make test  (see scripts/verify_workbook.sh and
            scripts/verify_household_engine.sh)
"""

import math

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

# --------------------------------------------------------------------------
# Corporate "Classic Navy" palette
# --------------------------------------------------------------------------
NAVY = "1F3864"          # dark navy header fill
WHITE = "FFFFFF"         # header text
ICE_BLUE = "DCE6F1"      # soft ice-blue accent / summary fill
LIGHT_GRAY = "BFBFBF"    # thin grid border
TITLE_GRAY = "44546A"    # secondary section-title fill
INPUT_YELLOW = "FFFDE7"  # pale-yellow fill marking user-editable input cells

SOFT_GREEN_FILL = "E2EFDA"
DARK_GREEN_FONT = "375623"
SOFT_YELLOW_FILL = "FFF2CC"
DARK_BROWN_FONT = "7F6000"
SOFT_RED_FILL = "FCE4E4"
DARK_RED_FONT = "C00000"

FONT_NAME = "Segoe UI"

CURRENCY_FMT = '$#,##0'
CURRENCY_FMT_SIGNED = '$#,##0;($#,##0)'
PERCENT_FMT = '0.00%'

thin_border = Border(
    left=Side(style="thin", color=LIGHT_GRAY),
    right=Side(style="thin", color=LIGHT_GRAY),
    top=Side(style="thin", color=LIGHT_GRAY),
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

LEGEND_TEXT = "Legend: pale-yellow cells are yours to type into; white cells hold formulas -- please don't overwrite them."


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

wb = Workbook()

# --------------------------------------------------------------------------
# TAB 0: Documentation
# --------------------------------------------------------------------------
ws0 = wb.active
ws0.title = "Documentation"
DOC_CHARS_PER_LINE = 78  # matches column B's width below

ws0.cell(row=1, column=1, value="HOUSEHOLD PORTFOLIO ENGINE — DOCUMENTATION & USER GUIDE")
ws0.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
for col in (1, 2):
    c = ws0.cell(row=1, column=col)
    c.font = header_font
    c.fill = header_fill
    c.border = thin_border
ws0.row_dimensions[1].height = 24

r = 3
write_section_banner(ws0, r, "WHY THIS SPREADSHEET EXISTS", 2)
r += 1
for para in (
    "This workbook manages a two-bucket, decumulation-phase retirement income strategy built to protect "
    "the household from Sequence-of-Returns Risk (SRR) -- the danger that a market downturn early in "
    "retirement forces the sale of depreciated stocks to cover monthly living expenses, permanently "
    "damaging the portfolio's ability to recover. The strategy rests on three ideas, described below.",

    "1. The TSP G Fund as a \"free lunch\" cash vault. By law, the G Fund earns the yield of intermediate- "
    "to long-term Treasuries while never losing principal value and staying instantly liquid. It serves as "
    "the household's operating cash account -- monthly bills are paid straight out of it, so equities are "
    "never sold to cover routine expenses.",

    "2. A dynamic Refill Engine, not a fixed withdrawal rate. Instead of a static 4%-rule, this workbook "
    "tracks a real-growth hurdle (COLA layered on a fixed real-growth baseline) year over year. When actual "
    "household wealth clears that hurdle, it's safe to harvest equity gains and roll them back into the "
    "G Fund to top the shield back up. When wealth falls short, all equity sales freeze and the G Fund alone "
    "absorbs spending, giving stocks time to recover.",

    "3. Asset location by account. Safe, fixed-income assets are deliberately concentrated in accounts "
    "built for stability (the TSP G Fund, a spouse's workplace plan), while private IRAs hold the aggressive "
    "growth positions and serve as Roth-conversion conduits -- keeping each spouse's accounts doing the job "
    "they're best suited for.",

    "4. An optional longevity bucket (QLAC). A Qualified Longevity Annuity Contract bought with pre-tax "
    "money insures the far tail -- income that starts late (up to age 85) and lasts for life -- while the G "
    "Fund protects the near term. Both are \"safe\" money but they do opposite jobs, so the QLAC is tracked "
    "as its own illiquid asset: its premium leaves the plan's spendable, tracked wealth (subtracted from "
    "P_track the same way a Roth conversion is), it is never part of the shield's runway or the Emergency "
    "Liquidity Protocol, and it can optionally be counted toward the G Fund target's safe share -- at cost "
    "or at present value -- so the G Fund only has to hold the rest.",
):
    write_paragraph(ws0, r, para, chars_per_line=DOC_CHARS_PER_LINE)
    r += 1
r += 1

write_section_banner(ws0, r, "HOW THE FIVE TABS WORK TOGETHER", 2)
r += 1
for para in (
    "Documentation (this tab) -- Motivation, how-to, an annual checklist, the color legend, and a "
    "glossary. Nothing here feeds any formula; it's for the humans using the sheet.",

    "Variables_And_Log -- Where you log the things that change constantly: the actual Social "
    "Security/Federal COLA each year, every month's actual TSP -> checking payment, the total cost of any "
    "Roth conversion (funded separately, from non-TSP accounts -- converted principal plus whatever share "
    "of the taxes came out of IRA/retirement funds rather than out-of-pocket, combined into one number), "
    "any new money continued work (yours or your spouse's) adds to a tracked non-Roth account -- a payroll "
    "contribution, an IRA contribution -- and once a year, your actual total household balance. Everything "
    "else on this tab is a formula built from those inputs.",

    "Account_Tracker -- A snapshot of every account's live balance, updated by hand whenever you check "
    "statements (at least once a year, right before the annual refill review). This is the single source "
    "of truth for household net worth "
    "that every other tab's formulas reference.",

    "The_Refill_Engine -- The decision-making tab. A block of configured inputs (Retirement Start Year, "
    "Ramp Length, the G Fund target percentages, the minimum transfer threshold, etc.) each carry a "
    "Default, an optional Override, and the resulting Effective value everything else on the tab actually "
    "uses. The Annual Refill Engine table is the headline output -- once a year it tells you how much "
    "(if anything) to move between a private pre-tax account and the TSP G Fund, in either direction. "
    "Below that, the Annual SRR Hurdle Model is the strategic gauge: since Roth funds are entirely out of "
    "this tool's scope, a Roth conversion's real cost to the plan -- the converted principal, plus "
    "whatever share of the taxes came out of IRA/retirement funds rather than out-of-pocket -- "
    "is treated as one real, permanent reduction to the plan's tracked wealth -- and the mirror image "
    "applies to new money continued work adds to a tracked non-Roth account: it's a real, permanent "
    "addition to that same tracked wealth, not investment performance. Both feed the same trajectory that "
    "drives the GOOD/POOR trigger and every real dollar amount (the Safety Anchor balance, recommended "
    "transfers), so there's no separate 'adjusted' track to keep in sync, and neither a conversion's cost "
    "nor a fresh contribution can be mistaken for a market move. Because a Roth conversion's cost typically "
    "comes out of non-TSP accounts, it can lower the G Fund's dollar target below its actual balance -- in "
    "that case the engine recommends pulling money back out of the G Fund, the mirror image of its usual "
    "top-up job. The Emergency Liquidity Protocol tells you, in a full-blown crisis, exactly which account "
    "to draw from next.",

    "Graph -- A single chart, read-only, pulling straight from The_Refill_Engine's Annual SRR Hurdle "
    "Model: Original Predicted Wealth vs. Current Predicted Wealth vs. Actual Balance, by year. Kept on "
    "its own tab so it never crowds the data-entry tabs -- open it only when you want the picture.",
):
    write_paragraph(ws0, r, para, chars_per_line=DOC_CHARS_PER_LINE)
    r += 1
r += 1

write_section_banner(ws0, r, "YOUR ANNUAL CHECKLIST", 2)
r += 1
checklist = [
    "Update Account_Tracker with current balances from your statements.",
    "Confirm Variables_And_Log has the year's monthly TSP payments logged (they can change any month; "
    "the review itself happens once a year).",
    "Open The_Refill_Engine and find this year's row in the Annual Refill Engine table.",
    "Check the \"Recommended Transfer This Year\" cell. If it's $0, do nothing -- the gap between "
    "target and actual is too small to bother with, or the market hasn't cleared its hurdle yet.",
    "If it's a positive number, execute a pre-tax rollover of that amount from your Conversion Conduit "
    "(Traditional IRA) into the TSP G Fund. If it's negative, transfer that amount the other way -- from "
    "the TSP G Fund back to your Conversion Conduit -- since the target has dropped below what's already "
    "sitting in the G Fund.",
    "Log whatever you actually transferred in the \"Actual Transferred\" column, even if it matches the "
    "recommendation exactly -- leaving it blank makes the sheet assume the recommendation happened. Use "
    "the same sign the recommendation used: positive for money moved into the G Fund, negative for money "
    "moved out.",
    "If you executed a Roth conversion this year, log its cost to the plan -- the converted principal, "
    "plus any taxes you paid from IRA/retirement funds -- as one combined number in Variables_And_Log. "
    "Skip any portion of the taxes you paid out-of-pocket; that cash was never part of this tool's tracked "
    "wealth, so it's not a reduction to it. Roth funds themselves are out of this tool's scope entirely, "
    "so what you do log is a real, permanent reduction to the plan's tracked wealth, not a market downturn "
    "to explain away.",
    "If continued work (yours or your spouse's) added new money to a tracked non-Roth account this year -- "
    "a TSP/403b payroll contribution, an IRA contribution -- log the total in the New Retirement "
    "Contributions column in Variables_And_Log. This keeps fresh savings from being mistaken for market "
    "outperformance in the GOOD/POOR trigger, while still growing the Safety Anchor's dollar target to "
    "match the larger tracked pool.",
    "Log the real total household balance at year-end in the \"Actual Balance (Logged, Year-End)\" "
    "column of the Annual Refill Engine table, so the Graph tab has something to plot against the "
    "projections.",
    "The first time your real starting balances and assumptions are in Google Sheets, select The_Refill_"
    "Engine's Original Predicted Wealth column and Paste Special -> Values Only over itself, once, to "
    "freeze it as your baseline -- never touch it again after that.",
    "If the TSP G Fund is ever fully drawn down, check the Emergency Liquidity Protocol block for the "
    "next-dollar priority order.",
    "If you bought a QLAC tranche this year, log its premium and the guaranteed annual income it buys in "
    "Variables_And_Log (year of purchase), and confirm the QLAC inputs on The_Refill_Engine (birth year, "
    "income start age, count mode). If the annuitant dies, set \"Annuitant Living?\" to 0.",
]
for i, item in enumerate(checklist, start=1):
    num_cell = ws0.cell(row=r, column=1, value=f"{i}.")
    num_cell.font = bold_body_font
    num_cell.alignment = Alignment(horizontal="right", vertical="top")
    write_paragraph(ws0, r, item, chars_per_line=DOC_CHARS_PER_LINE)
    r += 1
r += 1

write_section_banner(ws0, r, "COLOR LEGEND", 2)
r += 1
legend_rows = [
    (input_fill, "Pale yellow -- type your own numbers here."),
    (None, "White (no fill) -- a formula; please leave it alone."),
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
GLOSSARY_HEADER_ROW = r
ws0.cell(row=r, column=1, value="Term")
ws0.cell(row=r, column=2, value="Definition")
style_header_row(ws0, r, 1, 2)
r += 1
GLOSSARY = [
    ("COLA", "The federal Cost-of-Living Adjustment, applied annually to Social Security/federal "
             "retirement benefits."),
    ("Hurdle Rate", "The minimum real (inflation-adjusted) growth rate the portfolio needs to clear to be "
                     "considered \"on track\" -- COLA layered on a fixed real-growth baseline."),
    ("W_t", "Total dollar outflows (withdrawals) for a given year (t) -- the sum of that year's monthly "
            "TSP -> checking payments."),
    ("P_track", "The simulated target portfolio value that the household's actual wealth is measured "
                "against each year. Reduced by that year's logged Roth Conversion + Tax Cost and increased "
                "by that year's logged New Retirement Contributions, since neither is a market event -- "
                "both are real, permanent changes to the plan's tracked wealth that would otherwise distort "
                "the GOOD/POOR trigger."),
    ("SRR", "Sequence-of-Returns Risk -- the risk that a downturn early in retirement forces permanent "
            "sale of depreciated assets."),
    ("TSP G Fund", "The federal Thrift Savings Plan's government-securities fund: principal-protected, "
                    "instantly liquid, and the household's operating cash vault."),
    ("Safety Anchor", "The G Fund cash buffer this engine tries to maintain. Ramps linearly (on "
                       "The_Refill_Engine) from a Starting % down to an Ending % over the Ramp Length, then "
                       "holds flat -- largest during the early years when SRR risk is heavily front-loaded, "
                       "easing as that window passes. Applied each year to that year's projected wealth "
                       "(P_track), not to today's balance, so the dollar target keeps moving with the plan "
                       "instead of staying pinned to today's net worth."),
    ("Ramp Length", "How many years the G Fund target percentage takes to slide from its Starting % to its "
                     "Ending %, before holding flat. A configured input on The_Refill_Engine."),
    ("Default / Override / Effective", "The three-column pattern used for every configured input on "
                                        "The_Refill_Engine. Default is the built-in value; Override is blank "
                                        "unless you want to change it; Effective is whichever one is actually "
                                        "in force, and is what every other formula on the tab references."),
    ("Roth Conversion + Tax Cost", "The one number you log per year for a Roth conversion: the converted "
                                    "principal, plus any taxes paid from IRA/retirement funds -- never taxes "
                                    "paid out-of-pocket, since that cash was never part of this tool's "
                                    "tracked wealth. Since Roth funds are entirely out of this tool's scope, "
                                    "this whole amount is treated as a "
                                    "real, permanent reduction to Current Predicted Wealth (P_track) -- the "
                                    "same reduction feeds the GOOD/POOR trigger and every real dollar amount "
                                    "the plan sizes, so there's no separate adjusted track to reconcile."),
    ("New Retirement Contributions", "The one number you log per year for new money continued work (yours "
                                      "or your spouse's) adds to a tracked non-Roth account -- a payroll "
                                      "contribution, an IRA contribution. The mirror image of Roth "
                                      "Conversion + Tax Cost: added to Current Predicted Wealth (P_track) "
                                      "instead of subtracted, so fresh savings raise the plan's tracked "
                                      "wealth (and the Safety Anchor's dollar target with it) without ever "
                                      "looking like market outperformance to the GOOD/POOR trigger."),
    ("Minimum Transfer Threshold", "A configured input (Default/Override/Effective, on The_Refill_Engine) "
                                    "setting how large the gap between the Target Safety Anchor Balance and "
                                    "the G Fund's actual balance must be, as a percent of that year's P_track, "
                                    "before the engine bothers recommending a transfer. Applies in either "
                                    "direction, so a gap too small to matter produces a $0 recommendation "
                                    "instead of a token transfer."),
    ("Original Predicted Wealth", "A column on the Annual SRR Hurdle Model that mirrors Current Predicted "
                                   "Wealth (P_track) -- the real, unadjusted trajectory -- until you freeze "
                                   "it by hand (Paste Special -> Values Only, once, after your real numbers "
                                   "are in). After that it stays fixed as your baseline while Current keeps "
                                   "moving with your latest assumptions. Feeds the Graph tab."),
    ("Actual Balance", "Your real total household balance, logged once a year (at year-end) directly on "
                        "the Annual Refill Engine table. Compared on the Graph tab against Current and "
                        "Original Predicted Wealth for the same year."),
    ("Conduit", "A Traditional IRA used as a pass-through for pre-tax rollovers back into the TSP."),
    ("Rollover", "A pre-tax transfer of funds between retirement accounts (e.g. Traditional IRA to TSP) "
                 "that doesn't trigger a taxable distribution."),
    ("Allocation Mix", "Each account's share of total household net worth, shown on Account_Tracker."),
    ("QLAC", "Qualified Longevity Annuity Contract: a deferred income annuity bought inside a pre-tax "
             "account, excluded from RMDs until its income starts (no later than age 85). Modeled here with "
             "no death benefit: illiquid, mortality-contingent, worth nothing once the annuitant dies."),
    ("QLAC Premium Paid", "The one number you log per purchase year for a QLAC. It leaves the plan's liquid, "
                           "tracked wealth, so the Annual SRR Hurdle Model subtracts it from P_track exactly "
                           "the way it subtracts a Roth Conversion + Tax Cost -- the GOOD/POOR trigger is "
                           "never fooled by the drop, and the dollar targets shrink to match the smaller "
                           "liquid pool."),
    ("QLAC Present Value", "The expected, survival-weighted, discounted value of the QLAC's remaining "
                            "income stream at a given age. Starts below the premium, rises toward the "
                            "income start age, then falls as payments are received. Shown on "
                            "Account_Tracker below the liquid total, and projected per year on the Annual "
                            "SRR Hurdle Model."),
    ("Gompertz Mortality", "A two-parameter survival curve used to weight QLAC payments: the modal age "
                            "(most common age at death) and the dispersion (spread, in years). The defaults "
                            "approximate a woman in the top income quintile; adjust them, don't treat them "
                            "as precise."),
    ("QLAC Count Mode", "How the QLAC counts toward the G Fund target's safe share: COST (premiums paid, "
                         "the conservative default -- doesn't let an appreciating, mortality-contingent "
                         "value crowd out liquid G Fund), PV (present value), or NONE (outside the target). "
                         "In COST mode, a purchase lowers the G Fund's own dollar target by exactly the "
                         "premium, matching a purchase funded out of the safe share."),
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
GLOSSARY_END_ROW = r - 1

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
          value="HOUSEHOLD VARIABLE LOG — ANNUAL HURDLE INPUTS, MONTHLY TSP → CHECKING PAYMENT LOG, "
                "ROTH CONVERSION COST, NEW CONTRIBUTIONS & QLAC PURCHASES (2026-2035)")
LAST_COL_1 = 5 + 12 + 2 + 2  # 5 lead cols + 12 months + Roth cost + new contributions + 2 QLAC cols
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
    "Roth Conversion + Tax Cost (This Year)",
    "New Retirement Contributions (This Year)",
    "QLAC Premium Paid (This Year)",
    "QLAC Guaranteed Annual Income Purchased (This Year)",
]
for idx, h in enumerate(ws1_headers, start=1):
    ws1.cell(row=HEADER_ROW_1, column=idx, value=h)
style_header_row(ws1, HEADER_ROW_1, 1, len(ws1_headers))
ws1.row_dimensions[HEADER_ROW_1].height = 32

ROTH_CONVERSION_COL = 6 + 12       # R -- IRA-funded Roth conversion + tax total cost, logged as it happens
ROTH_CONVERSION_COL_LETTER = get_column_letter(ROTH_CONVERSION_COL)
CONTRIBUTIONS_COL = ROTH_CONVERSION_COL + 1   # S -- new retirement contributions, logged as they happen
CONTRIBUTIONS_COL_LETTER = get_column_letter(CONTRIBUTIONS_COL)
QLAC_PREMIUM_COL = CONTRIBUTIONS_COL + 1   # T -- premium moved from a tracked pre-tax account into a QLAC
QLAC_PREMIUM_COL_LETTER = get_column_letter(QLAC_PREMIUM_COL)
QLAC_INCOME_COL = CONTRIBUTIONS_COL + 2    # U -- guaranteed annual income that tranche buys
QLAC_INCOME_COL_LETTER = get_column_letter(QLAC_INCOME_COL)

DATA_START_ROW_1 = HEADER_ROW_1 + 1  # row 3
for i, year in enumerate(YEARS):
    r = DATA_START_ROW_1 + i
    style_body_cell(ws1, r, 1, number_format="0").value = year
    cola_cell = mark_input(style_body_cell(ws1, r, 2, number_format=PERCENT_FMT))
    if year == 2026:
        cola_cell.value = 0.026
    # else: left blank for user input
    style_body_cell(ws1, r, 3, number_format=PERCENT_FMT).value = 0.020
    style_body_cell(ws1, r, 4, number_format=PERCENT_FMT).value = f"=(1+B{r})*(1+C{r})"
    month_first, month_last = MONTH_COLS[0], MONTH_COLS[-1]
    style_body_cell(ws1, r, 5, number_format=CURRENCY_FMT).value = f"=SUM({month_first}{r}:{month_last}{r})"
    for m, col_letter in enumerate(MONTH_COLS):
        mark_input(style_body_cell(ws1, r, 6 + m, number_format=CURRENCY_FMT))
    mark_input(style_body_cell(ws1, r, ROTH_CONVERSION_COL, number_format=CURRENCY_FMT)).value = 0
    mark_input(style_body_cell(ws1, r, CONTRIBUTIONS_COL, number_format=CURRENCY_FMT)).value = 0
    mark_input(style_body_cell(ws1, r, QLAC_PREMIUM_COL, number_format=CURRENCY_FMT)).value = 0
    mark_input(style_body_cell(ws1, r, QLAC_INCOME_COL, number_format=CURRENCY_FMT)).value = 0

LAST_DATA_ROW_1 = DATA_START_ROW_1 + len(YEARS) - 1  # row 12 (2035)

note_row_1 = LAST_DATA_ROW_1 + 2  # row 14
ws1.cell(row=note_row_1, column=1,
          value="Notes: Seed 2026 COLA is set; enter 2027-2035 COLA as it is announced each fall. "
                "Log each month's actual TSP -> checking payment in columns F:Q as it is set or changes -- "
                "Total Annual Outflows (E) sums automatically. If a payment amount changes mid-year, just "
                "overwrite that month's cell going forward. Roth conversions are funded entirely from non-TSP "
                "accounts -- log the total cost of each conversion (the converted principal, plus any taxes "
                "paid from IRA/retirement funds -- skip taxes paid out-of-pocket, that cash was never part "
                "of this tool's tracked wealth) in the Roth Conversion + Tax Cost column. Roth funds "
                "are entirely out of this tool's scope, so The_Refill_Engine treats that whole amount as a "
                "real, permanent reduction to the plan's tracked wealth. If continued work (yours or your "
                "spouse's) adds new money to a tracked non-Roth account -- a TSP/403b payroll contribution, "
                "an IRA contribution -- log that total in the New Retirement Contributions column; "
                "The_Refill_Engine adds it to the plan's tracked wealth the same way it subtracts a Roth "
                "conversion, so new savings can't be mistaken for market outperformance in the GOOD/POOR "
                "trigger, while the Safety Anchor's dollar target still grows to match the larger pool. "
                "If you buy a QLAC (longevity annuity), log the premium and the guaranteed annual income "
                "that tranche buys in the year of purchase -- one row per purchase year, so a laddered "
                "purchase is just several logged years. The premium is treated like a Roth conversion: "
                "it leaves the plan's liquid, tracked wealth, so P_track drops by it and the GOOD/POOR "
                "trigger isn't fooled. "
                "Actual Balance is logged once a year instead, directly in the Annual Refill Engine table "
                "on The_Refill_Engine, alongside everything else you already update there.")
ws1.cell(row=note_row_1, column=1).font = italic_note_font

legend_row_1 = note_row_1 + 1  # row 15
ws1.cell(row=legend_row_1, column=1, value=LEGEND_TEXT)
ws1.cell(row=legend_row_1, column=1).font = italic_note_font

enable_gridlines(ws1)
ws1.freeze_panes = f"A{DATA_START_ROW_1}"

# --------------------------------------------------------------------------
# TAB 2: Account_Tracker
# --------------------------------------------------------------------------
ws2 = wb.create_sheet("Account_Tracker")

ws2_headers = ["Asset Pool", "Custodian / Platform", "Live Balance", "Allocation Mix"]
for idx, h in enumerate(ws2_headers, start=1):
    ws2.cell(row=1, column=idx, value=h)
style_header_row(ws2, 1, 1, 4)
ws2.row_dimensions[1].height = 24

# NOTE: These are illustrative PLACEHOLDER balances only -- deliberately
# generic round numbers, not any real account figures. Type your actual
# balances into the generated .xlsx / Google Sheet directly; do not put
# real dollar amounts back into this script or commit them to source control.
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

# Backfill Allocation Mix formulas referencing the total cell
TOTAL_NET_WORTH_REF = f"'Account_Tracker'!$C${TOTAL_ROW_2}"
for i in range(len(ACCOUNTS)):
    r = DATA_START_ROW_2 + i
    ws2.cell(row=r, column=4, value=f"=C{r}/$C${TOTAL_ROW_2}")

# Illiquid longevity asset, deliberately kept OUT of the liquid total above
# (which every hurdle/refill formula reads). Its value is a formula filled in
# once The_Refill_Engine's QLAC valuation cell exists, further down.
QLAC_ROW_2 = TOTAL_ROW_2 + 1
TOTAL_INCL_QLAC_ROW_2 = TOTAL_ROW_2 + 2
style_body_cell(ws2, QLAC_ROW_2, 1).value = "Longevity Annuity (QLAC) -- illiquid, not in total above"
style_body_cell(ws2, QLAC_ROW_2, 2).value = "Insurer QLAC (present value, computed -- don't type here)"
style_body_cell(ws2, QLAC_ROW_2, 3, number_format=CURRENCY_FMT)
style_body_cell(ws2, QLAC_ROW_2, 4, number_format=PERCENT_FMT).value = (
    f"=IFERROR(C{QLAC_ROW_2}/C{TOTAL_INCL_QLAC_ROW_2},0)")
ws2.cell(row=TOTAL_INCL_QLAC_ROW_2, column=1, value="Household Net Worth incl. QLAC")
ws2.cell(row=TOTAL_INCL_QLAC_ROW_2, column=3, value=f"=C{TOTAL_ROW_2}+C{QLAC_ROW_2}")
for col in range(1, 5):
    cell = ws2.cell(row=TOTAL_INCL_QLAC_ROW_2, column=col)
    cell.font = bold_body_font
    cell.fill = accent_fill
    cell.border = thin_border
ws2.cell(row=TOTAL_INCL_QLAC_ROW_2, column=3).number_format = CURRENCY_FMT

note_row_2 = TOTAL_INCL_QLAC_ROW_2 + 1
ws2.cell(row=note_row_2, column=1,
          value="Note: Live Balance figures are a manually-updated snapshot -- refresh them whenever you "
                "check statements, and at least once a year before the annual refill review, so Tab 3's "
                "projections start from an "
                "accurate baseline. The QLAC row is computed from The_Refill_Engine's QLAC inputs and the "
                "purchases logged on Variables_And_Log; it sits below the liquid total on purpose, so the "
                "hurdle model and refill engine only ever see money that can actually be spent.")
ws2.cell(row=note_row_2, column=1).font = italic_note_font

legend_row_2 = note_row_2 + 1
ws2.cell(row=legend_row_2, column=1, value=LEGEND_TEXT)
ws2.cell(row=legend_row_2, column=1).font = italic_note_font

# Named/explicit cell references used elsewhere in the workbook
G_FUND_BALANCE_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2}"                    # row 2
PRIVATE_TRAD_IRA_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 1}"              # row 3
SPOUSE_403B_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 4}"                     # row 6
MERRILL_CASH_BUFFER_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 6}"           # row 8

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
    formula on this sheet should point at)."""
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


r = 1
write_kv_row(r, "Total Current Actual Wealth (Live Account_Tracker Snapshot)",
             f"={TOTAL_NET_WORTH_REF}", CURRENCY_FMT, is_summary=True)
ACTUAL_WEALTH_SNAPSHOT_ROW = r
ACTUAL_WEALTH_CELL = f"$B${ACTUAL_WEALTH_SNAPSHOT_ROW}"
r += 1
write_kv_row(r, "Cumulative Roth Conversion + Tax Cost (Logged To Date, Informational)",
             f"=SUM('Variables_And_Log'!{ROTH_CONVERSION_COL_LETTER}{DATA_START_ROW_1}:"
             f"{ROTH_CONVERSION_COL_LETTER}{LAST_DATA_ROW_1})", CURRENCY_FMT, is_summary=True)
r += 1
write_kv_row(r, "Cumulative QLAC Premiums Paid (Logged To Date, Informational)",
             f"=SUM('Variables_And_Log'!{QLAC_PREMIUM_COL_LETTER}{DATA_START_ROW_1}:"
             f"{QLAC_PREMIUM_COL_LETTER}{LAST_DATA_ROW_1})", CURRENCY_FMT, is_summary=True)
r += 1
ROTH_NOTE_ROW = r
ws3.cell(row=r, column=1,
          value="Roth conversions are funded from non-TSP accounts. Since Roth funds are entirely out of "
                "this tool's scope, the converted principal is NOT a net-zero move as far as this plan is "
                "concerned -- it, plus any taxes paid from IRA/retirement funds (never taxes paid "
                "out-of-pocket, since that cash was never part of this tool's tracked wealth), are a real, "
                "permanent reduction to the plan's tracked wealth. Log the combined total per year on "
                "Variables_And_Log; the Annual SRR Hurdle Model below subtracts it from Current Predicted "
                "Wealth (P_track) the same way it subtracts Annual Outflows, and that same reduced P_track "
                "feeds both the GOOD/POOR trigger and every real dollar amount (Safety Anchor balance, "
                "recommended transfers) -- deliberately no longer two separate tracks. A QLAC premium is "
                "handled the same way: it moves money out of the plan's liquid, tracked accounts into an "
                "illiquid annuity, so it is subtracted from P_track in its purchase year.")
ws3.cell(row=r, column=1).font = italic_note_font
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
MIN_TRANSFER_PCT_CELL = write_config_row(
    r, "Minimum Transfer Threshold (% of That Year's P_track)", 0.001, PERCENT_FMT)
r += 1
# QLAC (longevity annuity) assumptions. Birth year is a generic placeholder --
# type the real one into the Google Sheets copy, never back into this script.
QLAC_BIRTH_YEAR_CELL = write_config_row(r, "QLAC: Annuitant Birth Year", 1965, "0")
r += 1
QLAC_START_AGE_CELL = write_config_row(r, "QLAC: Income Start Age", 85, "0")
r += 1
QLAC_ALIVE_CELL = write_config_row(r, "QLAC: Annuitant Living? (1 = yes, 0 = no)", 1, "0")
r += 1
QLAC_DISCOUNT_CELL = write_config_row(r, "QLAC: Valuation Discount Rate (~ G Fund yield)", 0.0425, PERCENT_FMT)
r += 1
QLAC_GOMPERTZ_M_CELL = write_config_row(r, "QLAC: Mortality -- Gompertz Modal Age", 93, "0.0")
r += 1
QLAC_GOMPERTZ_B_CELL = write_config_row(r, "QLAC: Mortality -- Gompertz Dispersion (Years)", 9.5, "0.0")
r += 1
QLAC_COUNT_MODE_ROW = r
QLAC_COUNT_MODE_CELL = write_config_row(r, "QLAC: Count Toward G Fund Target As (COST / PV / NONE)",
                                        "COST", "@")
qlac_mode_dv = DataValidation(type="list", formula1='"COST,PV,NONE"', allow_blank=True)
ws3.add_data_validation(qlac_mode_dv)
qlac_mode_dv.add(ws3.cell(row=QLAC_COUNT_MODE_ROW, column=3))
r += 1
r += 1  # blank spacer row

# Present value of the QLAC's remaining income stream, valued at annuitant age
# `age` with `income` dollars/year of guaranteed income in force. Payments are
# modeled annually, mid-year, starting at the Income Start Age (or now, if
# already past it), for up to 45 years, each weighted by Gompertz survival from
# `age` and discounted at the configured rate. Zero if no income is in force
# or the annuitant has died (no death benefit is modeled). The array constant
# keeps this a single-cell formula that Excel, LibreOffice and Google Sheets
# all evaluate inside SUMPRODUCT -- no helper table needed.
QLAC_PAYMENT_OFFSETS = "{" + ",".join(f"{k + 0.5}" for k in range(45)) + "}"


def qlac_pv_formula(age, income):
    pay_age = f"(MAX({QLAC_START_AGE_CELL},{age})+{QLAC_PAYMENT_OFFSETS})"
    survival = (f"EXP(-(EXP(({pay_age}-{QLAC_GOMPERTZ_M_CELL})/{QLAC_GOMPERTZ_B_CELL})"
                f"-EXP((({age})-{QLAC_GOMPERTZ_M_CELL})/{QLAC_GOMPERTZ_B_CELL})))")
    discount = f"(1+{QLAC_DISCOUNT_CELL})^(-({pay_age}-({age})))"
    return (f"IF(({income})=0,0,{QLAC_ALIVE_CELL}*({income})*"
            f"SUMPRODUCT({survival}*{discount}))")


def qlac_counted_formula(pv, premiums):
    """How much of the QLAC counts toward the G Fund (safe-asset) target, per the
    configured mode: COST = premiums paid (while the annuitant lives), PV = the
    present value above, NONE = it sits outside the target entirely."""
    return (f'IF({QLAC_COUNT_MODE_CELL}="PV",{pv},'
            f'IF({QLAC_COUNT_MODE_CELL}="COST",{QLAC_ALIVE_CELL}*({premiums}),0))')


QLAC_LOG_YEARS = f"'Variables_And_Log'!$A${DATA_START_ROW_1}:$A${LAST_DATA_ROW_1}"
QLAC_LOG_PREMIUMS = (f"'Variables_And_Log'!${QLAC_PREMIUM_COL_LETTER}${DATA_START_ROW_1}:"
                     f"${QLAC_PREMIUM_COL_LETTER}${LAST_DATA_ROW_1}")
QLAC_LOG_INCOME = (f"'Variables_And_Log'!${QLAC_INCOME_COL_LETTER}${DATA_START_ROW_1}:"
                   f"${QLAC_INCOME_COL_LETTER}${LAST_DATA_ROW_1}")

write_kv_row(r, "Current G Fund Target % (Today)",
             f"=IF(YEAR(TODAY())-{RETIREMENT_START_CELL}>={RAMP_LENGTH_CELL},{ENDING_TARGET_CELL},"
             f"{STARTING_TARGET_CELL}+({ENDING_TARGET_CELL}-{STARTING_TARGET_CELL})*"
             f"(YEAR(TODAY())-{RETIREMENT_START_CELL})/{RAMP_LENGTH_CELL})",
             PERCENT_FMT, is_summary=True)
CURRENT_TARGET_PCT_ROW = r
r += 1
write_kv_row(r, "QLAC Present Value (Today)",
             "=" + qlac_pv_formula(f"(YEAR(TODAY())-{QLAC_BIRTH_YEAR_CELL})", f"SUM({QLAC_LOG_INCOME})"),
             CURRENCY_FMT, is_summary=True)
QLAC_PV_TODAY_ROW = r
QLAC_PV_TODAY_CELL = f"'The_Refill_Engine'!$B${QLAC_PV_TODAY_ROW}"
r += 1
write_kv_row(r, "QLAC Value Counted Toward G Fund Target (Today)",
             "=" + qlac_counted_formula(f"B{QLAC_PV_TODAY_ROW}", f"SUM({QLAC_LOG_PREMIUMS})"),
             CURRENCY_FMT, is_summary=True)
QLAC_COUNTED_TODAY_ROW = r
r += 1
# The G Fund target % applies to liquid wealth PLUS whatever portion of the QLAC
# is counted as a safe asset; the counted QLAC value then fills part of that
# safe share, and the G Fund only has to hold the rest (floored at zero).
write_kv_row(r, "Target Safety Anchor Balance (Today)",
             f"=MAX(0,(B{ACTUAL_WEALTH_SNAPSHOT_ROW}+B{QLAC_COUNTED_TODAY_ROW})*B{CURRENT_TARGET_PCT_ROW}"
             f"-B{QLAC_COUNTED_TODAY_ROW})", CURRENCY_FMT, is_summary=True)
r += 1
write_kv_row(r, "Household Net Worth incl. QLAC (Present Value)",
             f"=B{ACTUAL_WEALTH_SNAPSHOT_ROW}+B{QLAC_PV_TODAY_ROW}", CURRENCY_FMT, is_summary=True)
r += 1

ws3.cell(row=r, column=1,
          value="Sequence-of-Returns-Risk (SRR) Shield: the G Fund target ramps down linearly from the "
                "Starting % to the Ending % over the Ramp Length, then holds flat -- SRR is concentrated in "
                "the first years of retirement, so the shield is largest when it matters most and eases as "
                "that window passes. Each year's actual target is computed per-year in the Annual SRR Hurdle "
                "Model below and drives the Annual Refill Engine's 'Target Safety Anchor Balance (This "
                "Year)' column, scaled against that year's Current Predicted Wealth (P_track), not a value "
                "frozen at Year 1. The rows above are only a live snapshot of today's applicable target for "
                "quick reference. A QLAC can be counted toward the safe share (COST or PV mode), in which "
                "case it fills part of the target and the G Fund holds the rest -- but it can never pay a "
                "bill before its income starts, so the Runway KPI below always uses the real G Fund balance. The Minimum Transfer Threshold keeps the recommended-transfer cell from "
                "firing on a trivially small gap between that target and the G Fund's actual balance -- in "
                "either direction.")
ws3.cell(row=r, column=1).font = italic_note_font
CONFIG_NOTE_ROW = r
r += 2  # blank spacer row

# ---- THE ANNUAL REFILL ENGINE (primary output) ----
# One row per year. Transfers between the TSP G Fund and the non-TSP conduit
# are reviewed and executed once a year (quarterly transfers proved too
# cumbersome); monthly TSP -> checking payments still happen every month and
# are summed into that year's outflow. Every column here is either a per-year
# input or reads straight off the Annual SRR Hurdle Model below, so there is
# no interpolation between periods anymore.
REFILL_BANNER_ROW = r
ws3.cell(row=REFILL_BANNER_ROW, column=1,
          value="THE ANNUAL REFILL ENGINE — PRIMARY OUTPUT: RECOMMENDED TSP TRANSFER (ONCE A YEAR)")
ws3.merge_cells(start_row=REFILL_BANNER_ROW, start_column=1, end_row=REFILL_BANNER_ROW, end_column=9)
for col in range(1, 10):
    c = ws3.cell(row=REFILL_BANNER_ROW, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border
ws3.row_dimensions[REFILL_BANNER_ROW].height = 20

REFILL_HEADER_ROW = REFILL_BANNER_ROW + 1
refill_headers = [
    "Year",                                                   # A
    "Annual Outflow (W_t)",                                   # B
    "Annual Performance Signal",                              # C
    "G Fund Balance (Start of Year)",                         # D
    "Recommended Transfer This Year (+ Into TSP / − Out Of TSP)",  # E
    "Actual Transferred (Log It Here, Same Sign)",            # F
    "G Fund Balance (End of Year, Projected)",                # G
    "Target Safety Anchor Balance (This Year)",               # H
    "Actual Balance (Logged, Year-End)",                      # I
]
for idx, h in enumerate(refill_headers, start=1):
    ws3.cell(row=REFILL_HEADER_ROW, column=idx, value=h)
style_header_row(ws3, REFILL_HEADER_ROW, 1, len(refill_headers))
ws3.row_dimensions[REFILL_HEADER_ROW].height = 40

REFILL_DATA_START = REFILL_HEADER_ROW + 1
REFILL_DATA_END = REFILL_DATA_START + len(YEARS) - 1
ANNUAL_BANNER_ROW = REFILL_DATA_END + 3
ANNUAL_HEADER_ROW = ANNUAL_BANNER_ROW + 1
ANNUAL_DATA_START = ANNUAL_HEADER_ROW + 1

for i, year in enumerate(YEARS):
    var_log_row = DATA_START_ROW_1 + i
    annual_row = ANNUAL_DATA_START + i
    r = REFILL_DATA_START + i

    style_body_cell(ws3, r, 1, number_format="0").value = year

    # The whole year's logged monthly TSP -> checking payments.
    style_body_cell(ws3, r, 2, number_format=CURRENCY_FMT).value = f"='Variables_And_Log'!E{var_log_row}"

    c_cell = style_body_cell(ws3, r, 3)
    c_cell.value = f"=G{annual_row}"
    c_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    d_cell = style_body_cell(ws3, r, 4, number_format=CURRENCY_FMT)
    d_cell.value = f"={G_FUND_BALANCE_CELL}" if i == 0 else f"=G{r - 1}"

    h_cell = style_body_cell(ws3, r, 8, number_format=CURRENCY_FMT)
    h_cell.value = f"=N{annual_row}"

    # Bidirectional: pushes into the G Fund only in a GOOD year, but pulls money back out
    # whenever the target has dropped below the actual balance by more than the threshold --
    # e.g. after a Roth conversion+tax cost or a QLAC premium shrinks P_track -- regardless of
    # that year's GOOD/POOR signal, since that's a rebalance-to-target, not a performance call.
    # Either direction is suppressed if the gap is smaller than the Minimum Transfer Threshold
    # (a percent of that year's P_track), so a trivial gap doesn't trigger a token transfer.
    threshold_expr = f"({MIN_TRANSFER_PCT_CELL}*D{annual_row})"
    e_cell = style_body_cell(ws3, r, 5, number_format=CURRENCY_FMT_SIGNED, bold=True, fill=accent_fill)
    e_cell.value = (
        f'=IF(H{r}-D{r}<-{threshold_expr},H{r}-D{r},'
        f'IF(AND(ISNUMBER(SEARCH("GOOD",C{r})),H{r}-D{r}>{threshold_expr}),H{r}-D{r},0))'
    )

    mark_input(style_body_cell(ws3, r, 6, number_format=CURRENCY_FMT_SIGNED))

    # Projected year-end G Fund balance: start of year, plus the transfer (actual if logged,
    # otherwise the recommendation), minus the whole year's outflow -- floored at zero.
    g_cell = style_body_cell(ws3, r, 7, number_format=CURRENCY_FMT)
    g_cell.value = f'=MAX(0,D{r}-B{r}+IF(F{r}="",E{r},F{r}))'

    # Logged once a year, at year-end -- feeds the Annual SRR Hurdle Model's Actual
    # Balance column (which applies the NA()-if-blank gap logic the Graph tab needs),
    # so this cell stays a plain, empty input until you type into it.
    mark_input(style_body_cell(ws3, r, 9, number_format=CURRENCY_FMT))

    ws3.row_dimensions[r].height = 28

# ---- ANNUAL SRR HURDLE MODEL (gates the annual refill engine above) ----
ws3.cell(row=ANNUAL_BANNER_ROW, column=1,
          value="ANNUAL SRR HURDLE MODEL — STRATEGIC PACING CHECK (GATES THE ANNUAL REFILL ABOVE)")
ws3.merge_cells(start_row=ANNUAL_BANNER_ROW, start_column=1, end_row=ANNUAL_BANNER_ROW, end_column=14)
for col in range(1, 15):
    c = ws3.cell(row=ANNUAL_BANNER_ROW, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border
ws3.row_dimensions[ANNUAL_BANNER_ROW].height = 20

annual_headers = [
    "Year",
    "Total Annual Hurdle Rate",
    "Annual Outflows (W_t)",
    "Current Predicted Wealth (P_track)",
    "Original Predicted Wealth (Freeze By Hand Once)",
    "Actual Balance (Logged)",
    "Performance Logic Trigger",
    "Safety Anchor % (This Year)",
    "Roth Conversion + Tax Cost (This Year)",
    "New Retirement Contributions (This Year)",
    "QLAC Premium Paid (This Year)",
    "QLAC Present Value (Year-End, Projected)",
    "QLAC Value Counted Toward G Fund Target",
    "Target Safety Anchor Balance ($, This Year)",
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

    # Seeded from the raw Account_Tracker snapshot, then rolled forward each year net of
    # Annual Outflows, that year's logged Roth Conversion + Tax Cost (column I, subtracted --
    # Roth funds are entirely out of this tool's scope, so that cost is a real, permanent
    # exit from the plan, not a market event), and that year's logged New Retirement
    # Contributions (column J, added -- new savings from continued work are real money
    # entering the plan, not market performance, so they belong here rather than showing up
    # as apparent outperformance). Both feed the same trajectory everything else (the
    # GOOD/POOR trigger, the Safety Anchor's dollar sizing) reads from -- there is
    # deliberately no separate adjusted track for either.
    if i == 0:
        target_formula = f"=ROUND((B{ACTUAL_WEALTH_SNAPSHOT_ROW}*B{r})-C{r}-I{r}+J{r}-K{r},2)"
    else:
        target_formula = f"=ROUND((D{r - 1}*B{r})-C{r}-I{r}+J{r}-K{r},2)"
    style_body_cell(ws3, r, 4, number_format=CURRENCY_FMT).value = target_formula

    # Mirrors Current (D) until you freeze it: once your real starting balances and
    # assumptions are in Google Sheets, select this column and Paste Special -> Values
    # Only over itself, one time, so it stops moving and becomes your fixed baseline.
    style_body_cell(ws3, r, 5, number_format=CURRENCY_FMT).value = f"=D{r}"

    # Year-end snapshot, pulled from this year's row on the Annual Refill Engine
    # table above (where Actual Balance is actually logged, once a year).
    # NA() rather than blank/0 so an unlogged year shows as a true gap.
    refill_row = REFILL_DATA_START + i
    style_body_cell(ws3, r, 6, number_format=CURRENCY_FMT).value = (
        f'=IF(I{refill_row}="",NA(),I{refill_row})'
    )

    trigger_formula = (
        f'=IF({ACTUAL_WEALTH_CELL}>=D{r},'
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

    # Pulled straight from this year's logged total on Variables_And_Log -- the converted
    # principal plus any taxes paid from IRA/retirement funds (never out-of-pocket taxes,
    # since that cash was never part of this tool's tracked wealth). Feeds column D's
    # recursive formula above as another category of outflow, same as Annual Outflows.
    style_body_cell(ws3, r, 9, number_format=CURRENCY_FMT).value = (
        f"='Variables_And_Log'!{ROTH_CONVERSION_COL_LETTER}{var_log_row}"
    )

    # Pulled straight from this year's logged total on Variables_And_Log -- new money added
    # to a tracked non-Roth account from continued work (a payroll contribution, an IRA
    # contribution). Feeds column D's recursive formula above as an addition, the mirror
    # image of the Roth conversion+tax cost column, so new savings raise the plan's tracked
    # wealth without ever looking like market outperformance to the GOOD/POOR trigger.
    style_body_cell(ws3, r, 10, number_format=CURRENCY_FMT).value = (
        f"='Variables_And_Log'!{CONTRIBUTIONS_COL_LETTER}{var_log_row}"
    )

    # QLAC premium paid this year -- subtracted in column D's recursive formula, exactly
    # like a Roth conversion: the money leaves the plan's liquid, tracked accounts for an
    # illiquid annuity, so it is a real exit from P_track, not a market event.
    style_body_cell(ws3, r, 11, number_format=CURRENCY_FMT).value = (
        f"='Variables_And_Log'!{QLAC_PREMIUM_COL_LETTER}{var_log_row}"
    )

    # QLAC value at this year's end: annuitant age = year - birth year, with only the
    # tranches logged in this year or earlier counted as in force.
    income_in_force = f'SUMIF({QLAC_LOG_YEARS},"<="&A{r},{QLAC_LOG_INCOME})'
    premiums_to_date = f'SUMIF({QLAC_LOG_YEARS},"<="&A{r},{QLAC_LOG_PREMIUMS})'
    style_body_cell(ws3, r, 12, number_format=CURRENCY_FMT).value = (
        "=" + qlac_pv_formula(f"(A{r}-{QLAC_BIRTH_YEAR_CELL})", income_in_force)
    )
    style_body_cell(ws3, r, 13, number_format=CURRENCY_FMT).value = (
        "=" + qlac_counted_formula(f"L{r}", premiums_to_date)
    )
    # Dollar target the G Fund itself must hold: the safe share of (P_track + counted
    # QLAC value), minus the counted QLAC value, floored at zero. With no QLAC (or mode
    # NONE) this is exactly P_track x Safety Anchor %, as before. In COST mode a purchase
    # lowers this target by exactly the premium -- the purchase came out of the safe share.
    style_body_cell(ws3, r, 14, number_format=CURRENCY_FMT).value = (
        f"=MAX(0,(D{r}+M{r})*H{r}-M{r})"
    )

    ws3.row_dimensions[r].height = 28

ANNUAL_DATA_END = ANNUAL_DATA_START + len(YEARS) - 1

# ---- EXTRA CAPACITY INDICATORS ----
KPI_HEADER_ROW = ANNUAL_DATA_END + 2  # row 63
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
          value="(Runway uses the average of whichever 2026 monthly TSP payments you've logged so far.)")
ws3.cell(row=note_row_3, column=1).font = italic_note_font

# ---- EMERGENCY LIQUIDITY PROTOCOL ----
PROTOCOL_HEADER_ROW = note_row_3 + 2  # row 69
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

protocol_note_row = OUTPUT_ROW_END + 1
ws3.cell(row=protocol_note_row, column=1,
          value="A QLAC never appears in this priority order: it has no cash value and can't be surrendered "
                "or drawn early, so it is never a source of emergency liquidity -- only of income once its "
                "start age arrives.")
ws3.cell(row=protocol_note_row, column=1).font = italic_note_font
legend_row_3 = OUTPUT_ROW_END + 2
ws3.cell(row=legend_row_3, column=1, value=LEGEND_TEXT)
ws3.cell(row=legend_row_3, column=1).font = italic_note_font

# Conditional formatting: green / yellow / red priority states
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
ws3.freeze_panes = f"A{REFILL_DATA_START}"

# Now that the QLAC valuation cell exists, wire it into Account_Tracker.
ws2.cell(row=QLAC_ROW_2, column=3).value = f"={QLAC_PV_TODAY_CELL}"

# --------------------------------------------------------------------------
# Autofit + explicit widths + final gridline pass for every tab
# --------------------------------------------------------------------------
autofit_columns(ws1, skip_rows={note_row_1, legend_row_1})
autofit_columns(ws2, skip_rows={note_row_2, legend_row_2})
autofit_columns(ws3, skip_rows={ROTH_NOTE_ROW, CONFIG_NOTE_ROW, note_row_3, legend_row_3,
                                protocol_note_row})
for ws in (ws1, ws2, ws3):
    ws.sheet_view.showGridLines = True

ws1.column_dimensions["A"].width = 14
for col_letter in MONTH_COLS:
    ws1.column_dimensions[col_letter].width = 12
ws1.column_dimensions[ROTH_CONVERSION_COL_LETTER].width = 18
ws1.column_dimensions[QLAC_PREMIUM_COL_LETTER].width = 18
ws1.column_dimensions[QLAC_INCOME_COL_LETTER].width = 22

ws3.column_dimensions["A"].width = 42
ws3.column_dimensions["B"].width = 14
ws3.column_dimensions["C"].width = 22
ws3.column_dimensions["D"].width = 44
ws3.column_dimensions["E"].width = 22
ws3.column_dimensions["F"].width = 22
ws3.column_dimensions["G"].width = 22
ws3.column_dimensions["H"].width = 22
ws3.column_dimensions["I"].width = 22
ws3.column_dimensions["J"].width = 22
ws3.column_dimensions["K"].width = 24
ws3.column_dimensions["L"].width = 24
ws3.column_dimensions["M"].width = 24
ws3.column_dimensions["N"].width = 24

# --------------------------------------------------------------------------
# TAB 4: Graph
# --------------------------------------------------------------------------
ws4 = wb.create_sheet("Graph")

ws4.cell(row=1, column=1, value="HOUSEHOLD WEALTH — ORIGINAL VS. CURRENT PROJECTION VS. ACTUAL")
ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
for col in range(1, 7):
    c = ws4.cell(row=1, column=col)
    c.font = header_font
    c.fill = header_fill
    c.border = thin_border
ws4.row_dimensions[1].height = 24

write_paragraph(
    ws4, 3,
    "This chart reads directly from the Annual SRR Hurdle Model on The_Refill_Engine -- there is "
    "nothing to edit here. Actual Balance is whatever you've logged at year-end in the Annual Refill "
    "Engine table; a year you haven't logged shows as a gap, not a zero. Original only means something "
    "once you've frozen it there (see the note on that column).",
    col=1, chars_per_line=100,
)

chart = BarChart()
chart.type = "col"
chart.grouping = "clustered"
chart.overlap = -10
chart.title = "Household Wealth: Original vs. Current Projection vs. Actual, By Year"
chart.y_axis.title = "Household Wealth ($)"
chart.y_axis.numFmt = CURRENCY_FMT
chart.x_axis.title = "Year"
chart.height = 12
chart.width = 32

cats = Reference(ws3, min_col=1, min_row=ANNUAL_DATA_START, max_row=ANNUAL_DATA_END)
data_original = Reference(ws3, min_col=5, min_row=ANNUAL_HEADER_ROW, max_row=ANNUAL_DATA_END)
data_current = Reference(ws3, min_col=4, min_row=ANNUAL_HEADER_ROW, max_row=ANNUAL_DATA_END)
data_actual = Reference(ws3, min_col=6, min_row=ANNUAL_HEADER_ROW, max_row=ANNUAL_DATA_END)

chart.add_data(data_original, titles_from_data=True)
chart.add_data(data_current, titles_from_data=True)
chart.add_data(data_actual, titles_from_data=True)
chart.set_categories(cats)

for series, color in zip(chart.series, (TITLE_GRAY, NAVY, "E8A33D")):
    series.graphicalProperties.solidFill = color

ws4.add_chart(chart, "A6")

enable_gridlines(ws4)

wb.save("household_portfolio_engine.xlsx")
print("Wrote household_portfolio_engine.xlsx")
