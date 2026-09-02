# Household Portfolio Engine workbooks.
#
#   make build   regenerate both .xlsx files from their Python generators
#                (only actually runs a generator if its .py is newer than
#                the .xlsx it produces).
#   make test    open each generated workbook in Microsoft Excel, force a
#                full recalculation, and check the formulas actually
#                evaluate correctly -- openpyxl never evaluates formulas,
#                so this is the only reliable way to catch a broken one.
#                Requires macOS + Microsoft Excel. Never touches any other
#                workbook you already have open.
#   make all     build then test (the default).
#   make clean   remove Python build artifacts (never touches the .xlsx
#                files -- they're committed, hand-maintained documents).
#
# The generated .xlsx files are meant to be imported into Google Sheets and
# hand-maintained there; re-running `make build` only matters when a
# generator script itself changes.

PYTHON := python3

TRACKER_SCRIPT := generate_household_portfolio_engine.py
TRACKER_XLSX   := household_portfolio_engine.xlsx
DEMO_SCRIPT    := generate_scenario_demo.py
DEMO_XLSX      := household_portfolio_scenario_demo.xlsx

.PHONY: all build test verify-tracker verify-demo clean

all: build test

build: $(TRACKER_XLSX) $(DEMO_XLSX)

$(TRACKER_XLSX): $(TRACKER_SCRIPT)
	$(PYTHON) $(TRACKER_SCRIPT)

$(DEMO_XLSX): $(DEMO_SCRIPT)
	$(PYTHON) $(DEMO_SCRIPT)

test: verify-tracker verify-demo

verify-tracker: $(TRACKER_XLSX)
	scripts/verify_workbook.sh $(TRACKER_XLSX)

verify-demo: $(DEMO_XLSX)
	scripts/verify_scenario_demo.sh $(DEMO_XLSX)

clean:
	rm -rf __pycache__
