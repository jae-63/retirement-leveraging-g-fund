#!/bin/bash
# Regression check for household_portfolio_scenario_demo.xlsx's core
# mechanic: cycles the Scenario Selector (The_Refill_Engine!B1) through
# Bear / Bull / Choppy and confirms the GOOD/POOR trigger behaves as
# designed for each one -- Bear stays POOR throughout, Bull stays GOOD
# throughout, Choppy shows both. Also runs the generic formula-error scan
# (same check as verify_workbook.sh, including its known limitation -- see
# that script's header -- that a real error surfaces via AppleScript as
# `missing value`, not text, so this only catches the rare case it comes
# back as a literal string) while the file is open.
#
# Counts "GOOD"/"POOR" substrings across The_Refill_Engine tab's used range
# (skipping column 1, which holds prose labels/notes -- including one that
# explains the GOOD/POOR mechanic by name) rather than fixed row numbers, so
# it survives layout changes as long as the trigger text itself still
# contains those words. Leaves the selector set back to "Bear" (the shipped
# default) and never touches any other workbook already open in Excel.
#
# Usage: scripts/verify_scenario_demo.sh <path-to-xlsx>

set -euo pipefail

if [ $# -lt 1 ]; then
  echo "Usage: $0 <path-to-xlsx>" >&2
  exit 1
fi

XLSX_PATH="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
XLSX_NAME="$(basename "$1")"

osascript <<EOF
tell application "Microsoft Excel"
    open (POSIX file "$XLSX_PATH")
    -- Excel's open can return before the workbook is actually loaded (seen in
    -- practice: "active workbook" was still missing value right after open),
    -- so wait for it by name -- and address it by name below rather than via
    -- "active workbook", which is only as reliable as whatever has focus.
    set openedOK to false
    repeat with attempt from 1 to 120
        try
            get name of workbook "$XLSX_NAME"
            set openedOK to true
            exit repeat
        on error
            delay 0.5
        end try
    end repeat
    if not openedOK then error "$XLSX_NAME did not open within 60 seconds -- check Excel for a dialog (file repair, already open, etc.)" number 1
    calculate full

    -- generic formula-error scan across all five tabs
    set errCount to 0
    repeat with wsName in {"Documentation", "Variables_And_Log", "Account_Tracker", "The_Refill_Engine", "Graph"}
        set w to worksheet wsName of workbook "$XLSX_NAME"
        set allVals to value of (used range of w)
        repeat with rowVals in allVals
            repeat with v in rowVals
                if class of v is string then
                    if v is in {"#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "#SPILL!", "#CALC!"} then
                        set errCount to errCount + 1
                        log "ERROR in " & wsName & ": " & v
                    end if
                end if
            end repeat
        end repeat
    end repeat
    if errCount > 0 then
        close workbook "$XLSX_NAME" saving no
        error "$XLSX_NAME: found " & errCount & " formula error(s) -- see log above" number 1
    end if

    set ws to worksheet "The_Refill_Engine" of workbook "$XLSX_NAME"

    set value of range "B1" of ws to "Bear"
    calculate full
    set bearGood to 0
    set bearPoor to 0
    set bearVals to value of (used range of ws)
    repeat with rowVals in bearVals
        repeat with v in (items 2 thru -1 of rowVals)
            if class of v is string then
                if v contains "GOOD" then set bearGood to bearGood + 1
                if v contains "POOR" then set bearPoor to bearPoor + 1
            end if
        end repeat
    end repeat

    set value of range "B1" of ws to "Bull"
    calculate full
    set bullGood to 0
    set bullPoor to 0
    set bullVals to value of (used range of ws)
    repeat with rowVals in bullVals
        repeat with v in (items 2 thru -1 of rowVals)
            if class of v is string then
                if v contains "GOOD" then set bullGood to bullGood + 1
                if v contains "POOR" then set bullPoor to bullPoor + 1
            end if
        end repeat
    end repeat

    set value of range "B1" of ws to "Choppy"
    calculate full
    set choppyGood to 0
    set choppyPoor to 0
    set choppyVals to value of (used range of ws)
    repeat with rowVals in choppyVals
        repeat with v in (items 2 thru -1 of rowVals)
            if class of v is string then
                if v contains "GOOD" then set choppyGood to choppyGood + 1
                if v contains "POOR" then set choppyPoor to choppyPoor + 1
            end if
        end repeat
    end repeat

    -- restore the shipped default before closing
    set value of range "B1" of ws to "Bear"
    calculate full
    close workbook "$XLSX_NAME" saving no

    if bearPoor is 0 or bearGood is not 0 then
        error "Bear scenario: expected only POOR triggers, got " & bearGood & " GOOD / " & bearPoor & " POOR" number 1
    end if
    if bullGood is 0 or bullPoor is not 0 then
        error "Bull scenario: expected only GOOD triggers, got " & bullGood & " GOOD / " & bullPoor & " POOR" number 1
    end if
    if choppyGood is 0 or choppyPoor is 0 then
        error "Choppy scenario: expected a mix of GOOD and POOR, got " & choppyGood & " GOOD / " & choppyPoor & " POOR" number 1
    end if

    log "OK: $XLSX_NAME -- Bear all-POOR, Bull all-GOOD, Choppy mixed, no formula errors"
end tell
EOF
