#!/bin/bash
# Functional/behavioral regression checks for household_portfolio_engine.xlsx --
# not unit tests of the Python generator, but assertions on the RECALCULATED
# workbook's actual numbers. The bugs this project has actually hit (the G
# Fund floor going negative, a Roth-conversion accounting change leaking into
# formulas it shouldn't touch) live in how formulas interact with each other,
# not in any single Python function -- so that's what these check, by
# mutating a few cells, recalculating for real, and asserting on the result.
# Complements (does not replace) verify_workbook.sh's generic formula-error
# scan.
#
# Every check reverts its test input(s) before moving to the next check. The
# whole check sequence runs inside a try block so the workbook is always
# closed (saving no) whether the checks pass or a check fails partway through
# -- a failed check that left the file open was itself caught and fixed while
# building this script (confirmed via a deliberate negative-control run: a
# reintroduced bug was caught with the exact expected before/after numbers,
# but left Excel open for the next invocation until this try/close wrapper
# was added). Nothing here is ever written back to disk, and this never
# touches any other workbook already open in Excel.
#
# Usage: scripts/verify_household_engine.sh <path-to-xlsx>

set -euo pipefail

if [ $# -lt 1 ]; then
  echo "Usage: $0 <path-to-xlsx>" >&2
  exit 1
fi

XLSX_PATH="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
XLSX_NAME="$(basename "$1")"

osascript <<EOF
-- Scans column 1 of worksheet w, rows 1..maxRow, for a cell whose text
-- contains searchText, and returns its row number. One bulk read, not a
-- round trip per row.
on findRow(w, searchText, maxRow)
    tell application "Microsoft Excel"
        set colVals to value of (range ("A1:A" & maxRow) of w)
    end tell
    set i to 0
    repeat with rowVals in colVals
        set i to i + 1
        repeat with v in rowVals
            if class of v is string then
                if v contains searchText then return i
            end if
        end repeat
    end repeat
    error "verify_household_engine.sh: could not find a row containing \"" & searchText & "\""
end findRow

on roughlyEqual(a, b, tolerance)
    return (a - b < tolerance) and (b - a < tolerance)
end roughlyEqual

tell application "Microsoft Excel"
    open (POSIX file "$XLSX_PATH")
    calculate full
    set wsRefill to worksheet "The_Refill_Engine" of active workbook
    set wsLog to worksheet "Variables_And_Log" of active workbook
    set wsTracker to worksheet "Account_Tracker" of active workbook
end tell

try
    -- Shared row lookups, used by several checks below.
    set quarterlyBannerRow to findRow(wsRefill, "PRIMARY OUTPUT", 200)
    set quarterlyDataStart to quarterlyBannerRow + 2
    set quarterlyDataEnd to quarterlyDataStart + 39 -- 10 years x 4 quarters
    set annualBannerRow to findRow(wsRefill, "STRATEGIC PACING CHECK", 200)
    set annualDataStart to annualBannerRow + 2

    -- ---- Check 1: a logged Roth conversion + tax cost is a real, permanent
    -- ---- reduction to Current Predicted Wealth (P_track) for that year --
    -- ---- and that SAME reduction flows through to both the Target Safety
    -- ---- Anchor Balance (This Year) and (structurally) the GOOD/POOR
    -- ---- trigger, since both now read off the one unadjusted P_track
    -- ---- trajectory. There is deliberately no isolated "trigger-only"
    -- ---- track anymore -- Roth funds are entirely out of this tool's
    -- ---- scope, so the logged cost is real money leaving the plan, not a
    -- ---- market event to net back out. ----
    tell application "Microsoft Excel"
        set pTrackBefore to value of cell annualDataStart of column 4 of wsRefill
        set safetyAnchorThisYearBefore to value of cell quarterlyDataStart of column 9 of wsRefill

        set value of range "R3" of wsLog to 50000
        calculate full

        set rothColAfter to value of cell annualDataStart of column 9 of wsRefill
        set pTrackAfter to value of cell annualDataStart of column 4 of wsRefill
        set safetyAnchorThisYearAfter to value of cell quarterlyDataStart of column 9 of wsRefill
        set safetyPct to value of cell annualDataStart of column 8 of wsRefill

        set value of range "R3" of wsLog to 0
        calculate full
    end tell

    if not roughlyEqual(rothColAfter, 50000, 1) then
        error "Roth cost pull-through: Annual table's Roth Conversion + Tax Cost column should read the logged 50000, got " & rothColAfter
    end if
    if not roughlyEqual(pTrackAfter, pTrackBefore - 50000, 1) then
        error "Roth unification: Current Predicted Wealth (P_track) should drop by exactly the logged Roth conversion+tax cost -- expected " & (pTrackBefore - 50000) & ", got " & pTrackAfter
    end if
    if not roughlyEqual(safetyAnchorThisYearAfter, safetyAnchorThisYearBefore - (50000 * safetyPct), 1) then
        error "Roth unification: Target Safety Anchor Balance (This Year) should fall by exactly the logged cost times that year's Safety Anchor % (proving it shares the same reduced basis as the trigger) -- expected drop of " & (50000 * safetyPct) & ", observed " & (safetyAnchorThisYearBefore - safetyAnchorThisYearAfter)
    end if

    -- ---- Check 2: the G Fund balance floors at zero under a sustained draw;
    -- ---- it must never go negative, in any quarter, regardless of whether
    -- ---- the draw is large enough to also flip a given year's GOOD/POOR
    -- ---- trigger (the floor must hold either way -- this doesn't rely on
    -- ---- staying POOR, just on a real, nontrivial draw over the horizon). ----
    tell application "Microsoft Excel"
        repeat with rowIdx from 3 to 12 -- Variables_And_Log data rows (2026-2035)
            repeat with col in {"F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"}
                set value of range (col & rowIdx) of wsLog to 2000
            end repeat
        end repeat
        calculate full
        set gFundEndVals to value of (range (("H" & quarterlyDataStart) & ":H" & quarterlyDataEnd) of wsRefill)
        repeat with rowIdx from 3 to 12
            repeat with col in {"F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"}
                set value of range (col & rowIdx) of wsLog to ""
            end repeat
        end repeat
        calculate full
    end tell

    set minGFundEnd to 999999999
    repeat with rowVals in gFundEndVals
        repeat with v in rowVals
            if v < minGFundEnd then set minGFundEnd to v
        end repeat
    end repeat
    if minGFundEnd < 0 then
        error "G Fund floor: end-of-quarter balance went negative (minimum seen: " & minGFundEnd & ") under a sustained \$2,000/month draw -- MAX(0, ...) floor appears broken"
    end if

    -- ---- Check 3: the Safety Anchor % ramp is linear between Starting and
    -- ---- Ending, using the default 30% -> 20% over 10 years. ----
    tell application "Microsoft Excel"
        set year1Pct to value of cell annualDataStart of column 8 of wsRefill
        set year4Pct to value of cell (annualDataStart + 3) of column 8 of wsRefill
    end tell

    if not roughlyEqual(year1Pct, 0.30, 0.001) then
        error "Ramp linearity: expected ~30% in the first year with default inputs, got " & year1Pct
    end if
    if not roughlyEqual(year4Pct, 0.27, 0.001) then
        error "Ramp linearity: expected ~27% three years in (30% ramping to 20% over 10 years), got " & year4Pct
    end if

    -- ---- Check 4: Default / Override / Effective precedence on a
    -- ---- configured input (Ramp Length). ----
    set rampRow to findRow(wsRefill, "Ramp Length", 20)

    tell application "Microsoft Excel"
        set rampDefault to value of cell rampRow of column 2 of wsRefill
        set rampEffectiveBefore to value of cell rampRow of column 4 of wsRefill

        set value of range ("C" & rampRow) of wsRefill to 5
        calculate full
        set rampEffectiveOverridden to value of cell rampRow of column 4 of wsRefill

        set value of range ("C" & rampRow) of wsRefill to ""
        calculate full
        set rampEffectiveReverted to value of cell rampRow of column 4 of wsRefill
    end tell

    if not roughlyEqual(rampEffectiveBefore, rampDefault, 0.001) then
        error "Default/Override: with no override set, Effective should equal Default (" & rampDefault & "), got " & rampEffectiveBefore
    end if
    if not roughlyEqual(rampEffectiveOverridden, 5, 0.001) then
        error "Default/Override: setting Override to 5 should make Effective 5, got " & rampEffectiveOverridden
    end if
    if not roughlyEqual(rampEffectiveReverted, rampDefault, 0.001) then
        error "Default/Override: clearing Override should revert Effective to Default (" & rampDefault & "), got " & rampEffectiveReverted
    end if

    -- ---- Check 5: an unlogged Actual Balance year shows as a true gap
    -- ---- (NA()), not a phantom zero, so the Graph tab doesn't plot a fake
    -- ---- data point. ----
    tell application "Microsoft Excel"
        set actualBalanceCell to value of cell annualDataStart of column 6 of wsRefill
    end tell

    if actualBalanceCell is not missing value then
        error "NA() gap: an unlogged Actual Balance year should read as a gap (missing value), got " & actualBalanceCell
    end if

    -- ---- Check 6: quarterly Actual Balance logging -- the interpolated
    -- ---- Current Predicted Wealth lands at the expected point between the
    -- ---- starting wealth and Year 1's P_track, and logging Q4 of a year
    -- ---- flows through to that year's row on the Annual table. ----
    tell application "Microsoft Excel"
        set startingWealth to value of cell 1 of column 2 of wsRefill
        set year1PTrack to value of cell annualDataStart of column 4 of wsRefill
        set q1InterpCurrent to value of cell quarterlyDataStart of column 11 of wsRefill
    end tell

    set expectedQ1 to startingWealth + (year1PTrack - startingWealth) * 0.25
    if not roughlyEqual(q1InterpCurrent, expectedQ1, 1) then
        error "Quarterly interpolation: Q1's interpolated Current Predicted Wealth should be 25% of the way from starting wealth to Year 1's P_track -- expected " & expectedQ1 & ", got " & q1InterpCurrent
    end if

    set q4Row to quarterlyDataStart + 3
    tell application "Microsoft Excel"
        set annualBalanceBefore to value of cell annualDataStart of column 6 of wsRefill
        set value of range (("J" & q4Row)) of wsRefill to 740000
        calculate full
        set annualBalanceAfter to value of cell annualDataStart of column 6 of wsRefill
        set value of range (("J" & q4Row)) of wsRefill to ""
        calculate full
    end tell

    if annualBalanceBefore is not missing value then
        error "Quarterly Actual Balance: expected the Annual table's Actual Balance to start as a gap before Q4 is logged, got " & annualBalanceBefore
    end if
    if not roughlyEqual(annualBalanceAfter, 740000, 1) then
        error "Quarterly Actual Balance: logging Q4 should flow through to the Annual table's year-end figure -- expected 740000, got " & annualBalanceAfter
    end if

    -- ---- Check 7: the Recommended Transfer is bidirectional and respects
    -- ---- the Minimum Transfer Threshold in either direction. A gap smaller
    -- ---- than the threshold (as a % of that year's P_track) recommends $0;
    -- ---- a gap well beyond it, with the G Fund OVER-funded relative to
    -- ---- target, recommends pulling money back OUT of the G Fund (a
    -- ---- negative transfer) -- and does so regardless of that year's
    -- ---- GOOD/POOR trigger, since an over-funded reserve is a rebalance,
    -- ---- not a performance call. ----
    set minTransferRow to findRow(wsRefill, "Minimum Transfer Threshold", 30)

    tell application "Microsoft Excel"
        set gFundOriginal to value of cell 2 of column 3 of wsTracker
        set conduitOriginal to value of cell 3 of column 3 of wsTracker
        set targetThisYear to value of cell quarterlyDataStart of column 9 of wsRefill
        set thresholdPct to value of cell minTransferRow of column 4 of wsRefill
        set pTrackYear1 to value of cell annualDataStart of column 4 of wsRefill
    end tell

    set thresholdDollar to thresholdPct * pTrackYear1

    -- Both sub-checks below shift dollars from the Conversion Conduit into the
    -- G Fund (and back), rather than just inflating the G Fund balance alone --
    -- the G Fund is one of the accounts summed into Total Household Net Worth,
    -- so moving money INTO it without an offsetting move would also inflate
    -- P_track and the target itself, defeating the point of the test. Shifting
    -- between two tracked accounts keeps total wealth (and the target) fixed
    -- while making the G Fund's own balance the only thing that moves.

    -- 7a: G Fund balance sits ABOVE target, but by less than the threshold --
    -- should still recommend $0 (the gap is too small to bother with).
    set desiredGFundSmall to targetThisYear + (thresholdDollar * 0.5)
    set deltaSmall to desiredGFundSmall - gFundOriginal
    tell application "Microsoft Excel"
        set value of range "C2" of wsTracker to gFundOriginal + deltaSmall
        set value of range "C3" of wsTracker to conduitOriginal - deltaSmall
        calculate full
        set recommendedSmallGap to value of cell quarterlyDataStart of column 6 of wsRefill
        set value of range "C2" of wsTracker to gFundOriginal
        set value of range "C3" of wsTracker to conduitOriginal
        calculate full
    end tell

    if not roughlyEqual(recommendedSmallGap, 0, 0.01) then
        error "Minimum Transfer Threshold: a gap smaller than the threshold should recommend \$0, got " & recommendedSmallGap
    end if

    -- 7b: G Fund balance sits well ABOVE target -- should recommend pulling
    -- the exact excess back out (a negative transfer), and the projected
    -- End of Qtr balance should reflect that pull, not just the raw balance.
    set desiredGFundLarge to targetThisYear + (thresholdDollar * 3)
    set deltaLarge to desiredGFundLarge - gFundOriginal
    tell application "Microsoft Excel"
        set value of range "C2" of wsTracker to gFundOriginal + deltaLarge
        set value of range "C3" of wsTracker to conduitOriginal - deltaLarge
        calculate full
        set recommendedPullOut to value of cell quarterlyDataStart of column 6 of wsRefill
        set endOfQtrProjected to value of cell quarterlyDataStart of column 8 of wsRefill
        set quarterlyOutflow to value of cell quarterlyDataStart of column 3 of wsRefill

        set value of range "C2" of wsTracker to gFundOriginal
        set value of range "C3" of wsTracker to conduitOriginal
        calculate full
    end tell

    if not roughlyEqual(recommendedPullOut, -(thresholdDollar * 3), 1) then
        error "Bidirectional transfer: an over-funded G Fund should recommend pulling the exact excess back out -- expected " & -(thresholdDollar * 3) & ", got " & recommendedPullOut
    end if
    if not roughlyEqual(endOfQtrProjected, targetThisYear - quarterlyOutflow, 1) then
        error "Bidirectional transfer: G Fund Balance (End of Qtr, Projected) should reflect the negative transfer pulling the balance back down to (roughly) target -- expected " & (targetThisYear - quarterlyOutflow) & ", got " & endOfQtrProjected
    end if

    tell application "Microsoft Excel"
        close workbook "$XLSX_NAME" saving no
    end tell
    log "OK: $XLSX_NAME -- Roth conversion unification, G Fund floor, ramp linearity, Default/Override precedence, the NA() gap, quarterly Actual Balance logging, and bidirectional transfer + deadband all behave as designed"
on error errMsg number errNum
    try
        tell application "Microsoft Excel"
            close workbook "$XLSX_NAME" saving no
        end tell
    end try
    error errMsg number errNum
end try
EOF
