#!/bin/bash
# Opens an .xlsx in Microsoft Excel, forces a full recalculation, and scans
# every sheet's used range for formula errors (#REF!, #VALUE!, #DIV/0!,
# etc). openpyxl never evaluates formulas, so this is the only reliable way
# to catch a broken formula before it's committed.
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
                    if v starts with "#" then
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
