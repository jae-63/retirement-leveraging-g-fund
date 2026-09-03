#!/bin/bash
# Opens an .xlsx in Microsoft Excel, forces a full recalculation, and scans
# every sheet's used range for formula errors that surface as literal text
# (#REF!, #VALUE!, #DIV/0!, etc). openpyxl never evaluates formulas, so this
# is a cheap sanity net before committing.
#
# KNOWN LIMITATION: Excel's AppleScript bridge returns a genuine formula
# error from `value of range` as the literal AppleScript `missing value`,
# indistinguishable from a truly blank cell -- confirmed by deliberately
# introducing a #DIV/0! and reading it back. There is no reliable, safe way
# found so far to recover the error token generically (a `save workbook as`
# + re-read via openpyxl(data_only=True) round trip was attempted; probing
# further AppleScript-side, e.g. bulk `formula of` over a multi-cell range
# after a bad cross-sheet reference, crashed Excel outright). So this scan
# only catches an error in the rare case it surfaces as text -- it is not a
# substitute for spot-checking key computed cells by hand after a change,
# which remains the reliable method.
#
# #N/A is deliberately NOT treated as an error -- the household-portfolio
# workbooks use NA() on purpose (e.g. an unlogged year's Actual Balance) so
# a chart shows a true gap instead of a phantom zero.
#
# Never touches any other workbook already open in Excel -- it opens the
# given file, closes exactly that file (by name, without saving), and
# leaves everything else alone.
#
# Usage: scripts/verify_workbook.sh <path-to-xlsx> [sheet1 sheet2 ...]
# With no sheet names given, checks the four standard household-portfolio
# tabs: Documentation, Variables_And_Log, Account_Tracker, The_Refill_Engine.

set -euo pipefail

if [ $# -lt 1 ]; then
  echo "Usage: $0 <path-to-xlsx> [sheet1 sheet2 ...]" >&2
  exit 1
fi

XLSX_PATH="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
XLSX_NAME="$(basename "$1")"
shift

SHEETS=("$@")
if [ ${#SHEETS[@]} -eq 0 ]; then
  SHEETS=(Documentation Variables_And_Log Account_Tracker The_Refill_Engine)
fi

SHEET_LIST=""
for s in "${SHEETS[@]}"; do
  SHEET_LIST="${SHEET_LIST}\"${s}\", "
done
SHEET_LIST="{${SHEET_LIST%, }}"

osascript <<EOF
tell application "Microsoft Excel"
    open (POSIX file "$XLSX_PATH")
    calculate full
    set errCount to 0
    repeat with wsName in $SHEET_LIST
        set w to worksheet wsName of active workbook
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
    close workbook "$XLSX_NAME" saving no
    if errCount > 0 then
        error "$XLSX_NAME: found " & errCount & " formula error(s) -- see log above" number 1
    end if
    log "OK: $XLSX_NAME -- no formula errors across ${#SHEETS[@]} sheet(s)"
end tell
EOF
