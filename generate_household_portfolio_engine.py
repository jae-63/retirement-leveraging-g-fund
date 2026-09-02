"""
generate_household_portfolio_engine.py

Builds 'household_portfolio_engine.xlsx' -- a three-tab household retirement
income engine implementing a Sequence-of-Returns-Risk (SRR) shield strategy:

  1. Variables_And_Log   -- annual COLA / hurdle-rate / outflow log (2026-2035)
  2. Account_Tracker     -- unified household net worth across all accounts
  3. The_Refill_Engine   -- SRR "Refill Engine" hurdle-tracking model, capacity
                            KPIs, and the Emergency Liquidity Protocol.

Run:  python3 generate_household_portfolio_engine.py
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Corporate "Classic Navy" palette
# --------------------------------------------------------------------------
NAVY = "1F3864"          # dark navy header fill
WHITE = "FFFFFF"         # header text
ICE_BLUE = "DCE6F1"      # soft ice-blue accent / summary fill
LIGHT_GRAY = "BFBFBF"    # thin grid border
TITLE_GRAY = "44546A"    # secondary section-title fill

SOFT_GREEN_FILL = "E2EFDA"
DARK_GREEN_FONT = "375623"
SOFT_YELLOW_FILL = "FFF2CC"
DARK_BROWN_FONT = "7F6000"
SOFT_RED_FILL = "FCE4E4"
DARK_RED_FONT = "C00000"

FONT_NAME = "Segoe UI"

CURRENCY_FMT = '$#,##0'
CURRENCY_FMT_2DP = '$#,##0.00'
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


def style_header_row(ws, row, first_col, last_col):
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_body_cell(ws, row, col, number_format=None, bold=False, fill=None):
    cell = ws.cell(row=row, column=col)
    cell.font = bold_body_font if bold else body_font
    cell.border = thin_border
    if number_format:
        cell.number_format = number_format
    if fill:
        cell.fill = fill
    return cell


def autofit_columns(ws, min_width=10, max_width=60, padding=3, skip_rows=None):
    """Auto-size columns to their longest value. Merged-anchor cells that span
    multiple columns (section banners, wrapped paragraph cells) and rows in
    skip_rows are excluded so a single long banner/note doesn't force every
    row's column artificially wide -- those get explicit widths instead."""
    skip_rows = skip_rows or set()
    merged_anchor_cells = {
        rng.min_row: rng for rng in ws.merged_cells.ranges
    }
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            if cell.row in skip_rows:
                continue
            # Skip the anchor cell of any multi-column merge -- its full
            # text length shouldn't dictate a single column's width.
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


# ==========================================================================
YEARS = list(range(2026, 2036))  # 2026..2035 inclusive (10 years)

wb = Workbook()

# --------------------------------------------------------------------------
# TAB 1: Variables_And_Log
# --------------------------------------------------------------------------
ws1 = wb.active
ws1.title = "Variables_And_Log"

ws1_headers = [
    "Date / Year",
    "Actual Federal COLA (C_i)",
    "Rule of Four Hurdle Baseline",
    "Total Annual Hurdle Rate",
    "Total Annual Outflows (W_t)",
]
for idx, h in enumerate(ws1_headers, start=1):
    ws1.cell(row=1, column=idx, value=h)
style_header_row(ws1, 1, 1, 5)
ws1.row_dimensions[1].height = 32

DATA_START_ROW_1 = 2
for i, year in enumerate(YEARS):
    r = DATA_START_ROW_1 + i
    style_body_cell(ws1, r, 1, number_format="0").value = year
    cola_cell = style_body_cell(ws1, r, 2, number_format=PERCENT_FMT)
    if year == 2026:
        cola_cell.value = 0.026
    # else: left blank for user input
    style_body_cell(ws1, r, 3, number_format=PERCENT_FMT).value = 0.020
    style_body_cell(ws1, r, 4, number_format=PERCENT_FMT).value = f"=(1+B{r})*(1+C{r})"
    style_body_cell(ws1, r, 5, number_format=CURRENCY_FMT)  # blank -- user logs annual outflow total

LAST_DATA_ROW_1 = DATA_START_ROW_1 + len(YEARS) - 1  # row 11 (2035)

note_row = LAST_DATA_ROW_1 + 2
ws1.cell(row=note_row, column=1,
          value="Notes: Seed 2026 COLA is set; enter 2027-2035 COLA as it is announced each fall. "
                "Log actual annual $6k-$20k/mo withdrawal totals in column E as they occur.")
ws1.cell(row=note_row, column=1).font = italic_note_font

enable_gridlines(ws1)
ws1.freeze_panes = "A2"

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
# generic round numbers, not any real account figures. Overwrite them in the
# generated .xlsx with your actual balances after opening it; do not put
# real dollar amounts back into this script or commit them to source control.
ACCOUNTS = [
    ("Your Pre-Tax Account", "TSP G Fund", 100_000),
    ("Your Conversion Conduit", "Private Traditional IRA", 100_000),
    ("Your Tax-Free Growth", "Private Roth IRA", 100_000),
    ("Wife's Tax-Free Growth", "Private Roth IRA", 100_000),
    ("Wife's Safe Pre-Tax Account", "Current Employer 403b (Active + Reverse Rollover)", 100_000),
    ("Wife's Conversion Conduit", "Private Traditional IRA", 100_000),
    ("Short-Term Liquidity", "Merrill Edge Cash Buffer", 100_000),
]

DATA_START_ROW_2 = 2
for i, (pool, custodian, balance) in enumerate(ACCOUNTS):
    r = DATA_START_ROW_2 + i
    style_body_cell(ws2, r, 1).value = pool
    style_body_cell(ws2, r, 2).value = custodian
    style_body_cell(ws2, r, 3, number_format=CURRENCY_FMT).value = balance
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

# Now backfill Allocation Mix formulas referencing the total cell
TOTAL_NET_WORTH_REF = f"'Account_Tracker'!$C${TOTAL_ROW_2}"
for i in range(len(ACCOUNTS)):
    r = DATA_START_ROW_2 + i
    ws2.cell(row=r, column=4, value=f"=C{r}/$C${TOTAL_ROW_2}")

# Named/explicit cell references used elsewhere in the workbook
G_FUND_BALANCE_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2}"                    # row 2
PRIVATE_TRAD_IRA_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 1}"              # row 3
WIFE_403B_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 4}"                     # row 6
MERRILL_CASH_BUFFER_CELL = f"'Account_Tracker'!$C${DATA_START_ROW_2 + 6}"           # row 8

enable_gridlines(ws2)
ws2.freeze_panes = "A2"

# --------------------------------------------------------------------------
# TAB 3: The_Refill_Engine
# --------------------------------------------------------------------------
ws3 = wb.create_sheet("The_Refill_Engine")

ws3.cell(row=1, column=1, value="Total Current Actual Wealth")
ws3.cell(row=1, column=1).font = label_font
b1 = ws3.cell(row=1, column=2, value=f"={TOTAL_NET_WORTH_REF}")
b1.number_format = CURRENCY_FMT
b1.font = bold_body_font
b1.fill = accent_fill
b1.border = thin_border
ws3.cell(row=1, column=1).border = thin_border
ws3.cell(row=1, column=1).fill = accent_fill

ws3.cell(row=2, column=1,
          value="Sequence-of-Returns-Risk (SRR) Shield: dynamic hurdle-tracking Refill Engine")
ws3.cell(row=2, column=1).font = italic_note_font

# ---- Hurdle-tracking table (rows 4-14) ----
TABLE_HEADER_ROW = 4
ws3_headers = [
    "Year",
    "Total Annual Hurdle Rate",
    "Annual Outflows (W_t)",
    "Simulated Target Portfolio (P_track)",
    "Performance Logic Trigger",
]
for idx, h in enumerate(ws3_headers, start=1):
    ws3.cell(row=TABLE_HEADER_ROW, column=idx, value=h)
style_header_row(ws3, TABLE_HEADER_ROW, 1, 5)
ws3.row_dimensions[TABLE_HEADER_ROW].height = 32

DATA_START_ROW_3 = TABLE_HEADER_ROW + 1  # row 5
# Row-1 baseline is the LIVE total household net worth (B1, linked from
# Account_Tracker) rather than a hardcoded dollar figure -- keeps this
# formula free of any actual balance and correct as balances change.

for i, year in enumerate(YEARS):
    r = DATA_START_ROW_3 + i
    var_log_row = DATA_START_ROW_1 + i  # matching row in Variables_And_Log

    style_body_cell(ws3, r, 1, number_format="0").value = year
    style_body_cell(ws3, r, 2, number_format=PERCENT_FMT).value = f"='Variables_And_Log'!D{var_log_row}"
    style_body_cell(ws3, r, 3, number_format=CURRENCY_FMT).value = f"='Variables_And_Log'!E{var_log_row}"

    if i == 0:
        target_formula = f"=ROUND(($B$1*B{r})-C{r},2)"
    else:
        target_formula = f"=ROUND((D{r - 1}*B{r})-C{r},2)"
    style_body_cell(ws3, r, 4, number_format=CURRENCY_FMT).value = target_formula

    trigger_formula = (
        f'=IF(B$1>=D{r},'
        f'"✅ GOOD: Execute T2065 Rebalance & Rollover to TSP",'
        f'"⚠️ POOR: Freeze Equities. Draw Down G Fund")'
    )
    trig_cell = style_body_cell(ws3, r, 5)
    trig_cell.value = trigger_formula
    trig_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

LAST_DATA_ROW_3 = DATA_START_ROW_3 + len(YEARS) - 1  # row 14

# ---- KPI block (rows 16-20) ----
KPI_HEADER_ROW = 16
ws3.cell(row=KPI_HEADER_ROW, column=1, value="EXTRA CAPACITY INDICATORS")
for col in range(1, 4):
    c = ws3.cell(row=KPI_HEADER_ROW, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border
ws3.merge_cells(start_row=KPI_HEADER_ROW, start_column=1, end_row=KPI_HEADER_ROW, end_column=3)

# Row 17: Conduit capacity
ws3.cell(row=17, column=1, value="Your Pre-Tax Rollover Conduit Capacity")
ws3.cell(row=17, column=1).font = label_font
ws3.cell(row=17, column=1).border = thin_border
conduit_cell = ws3.cell(row=17, column=2, value=f"={PRIVATE_TRAD_IRA_CELL}")
conduit_cell.number_format = CURRENCY_FMT
conduit_cell.font = body_font
conduit_cell.border = thin_border
conduit_status_cell = ws3.cell(row=17, column=3, value='=IF(B17>0,"Conduit Open","Conduit Closed")')
conduit_status_cell.font = bold_body_font
conduit_status_cell.fill = accent_fill
conduit_status_cell.border = thin_border

# Row 19: TSP Cash Buffer Runway
CURRENT_YEAR_OUTFLOW_CELL = f"'Variables_And_Log'!$E${DATA_START_ROW_1}"  # 2026 row = current year
ws3.cell(row=19, column=1, value="TSP Cash Buffer Runway (Months)")
ws3.cell(row=19, column=1).font = label_font
ws3.cell(row=19, column=1).border = thin_border
runway_cell = ws3.cell(
    row=19, column=2,
    value=f'=IFERROR({G_FUND_BALANCE_CELL}/({CURRENT_YEAR_OUTFLOW_CELL}/12),"Log Outflows to Calculate")'
)
runway_cell.number_format = "0.0"
runway_cell.font = body_font
runway_cell.border = thin_border
runway_status_cell = ws3.cell(
    row=19, column=3,
    value='=IF(ISNUMBER(B19),IF(B19>36,"Safe: 3+ Year Volatility Shield","Caution: Shield Shrinking"),'
          '"Awaiting Outflow Data")'
)
runway_status_cell.font = bold_body_font
runway_status_cell.fill = accent_fill
runway_status_cell.border = thin_border

ws3.cell(row=20, column=1,
          value="(Runway assumes the current year's logged outflow rate persists steadily.)")
ws3.cell(row=20, column=1).font = italic_note_font

# ---- Emergency Liquidity Protocol (rows 22-25) ----
PROTOCOL_HEADER_ROW = 22
ws3.cell(row=PROTOCOL_HEADER_ROW, column=1, value="EMERGENCY LIQUIDITY PROTOCOL — NEXT-DOLLAR PRIORITY ENGINE")
for col in range(1, 6):
    c = ws3.cell(row=PROTOCOL_HEADER_ROW, column=col)
    c.font = section_font
    c.fill = section_fill
    c.border = thin_border
ws3.merge_cells(start_row=PROTOCOL_HEADER_ROW, start_column=1, end_row=PROTOCOL_HEADER_ROW, end_column=5)

LABEL_ROW = PROTOCOL_HEADER_ROW + 1  # row 23
ws3.cell(row=LABEL_ROW, column=1, value="Next-Dollar Liquidity Priority")
ws3.cell(row=LABEL_ROW, column=1).font = label_font
ws3.cell(row=LABEL_ROW, column=1).border = thin_border
ws3.merge_cells(start_row=LABEL_ROW, start_column=1, end_row=LABEL_ROW, end_column=1)

OUTPUT_ROW_START = LABEL_ROW  # row 23
OUTPUT_ROW_END = 25
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
    "distributions from Wife's Current Employer 403b. Her pre-tax employer plan qualifies for the Maryland "
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
    f'IF(AND({G_FUND_BALANCE_CELL}=0,{MERRILL_CASH_BUFFER_CELL}=0,{WIFE_403B_CELL}>0),"{PRIORITY_3_TEXT}",'
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

# Conditional formatting: green / yellow / red priority states
cf_range = f"B{OUTPUT_ROW_START}:B{OUTPUT_ROW_END}"

# Priority 1 -- G Fund has a balance
ws3.conditional_formatting.add(
    cf_range,
    FormulaRule(formula=[f"{G_FUND_BALANCE_CELL}>0"], fill=green_fill, font=green_font),
)
# Priority 2 -- G Fund exhausted, Merrill Edge still has cash
ws3.conditional_formatting.add(
    cf_range,
    FormulaRule(
        formula=[f"AND({G_FUND_BALANCE_CELL}=0,{MERRILL_CASH_BUFFER_CELL}>0)"],
        fill=yellow_fill, font=yellow_font,
    ),
)
# Priority 3/4 -- G Fund and Merrill Edge both exhausted (Stage 1 or Stage 2 duress)
ws3.conditional_formatting.add(
    cf_range,
    FormulaRule(
        formula=[f"AND({G_FUND_BALANCE_CELL}=0,{MERRILL_CASH_BUFFER_CELL}=0)"],
        fill=red_fill, font=red_font,
    ),
)

enable_gridlines(ws3)
ws3.freeze_panes = "A5"

# --------------------------------------------------------------------------
# Autofit + final gridline pass (defensive, per spec) for every tab
# --------------------------------------------------------------------------
autofit_columns(ws1, skip_rows={note_row})
autofit_columns(ws2)
autofit_columns(ws3, skip_rows={2, 20})
for ws in (ws1, ws2, ws3):
    ws.sheet_view.showGridLines = True

# Explicit widths for label/banner columns that autofit deliberately skipped
ws1.column_dimensions["A"].width = 14
ws3.column_dimensions["A"].width = 42
# Widen the protocol output column on Tab 3 for wrapped paragraph text
ws3.column_dimensions["B"].width = 60
ws3.column_dimensions["C"].width = 18
ws3.column_dimensions["D"].width = 18
ws3.column_dimensions["E"].width = 18

wb.save("household_portfolio_engine.xlsx")
print("Wrote household_portfolio_engine.xlsx")
