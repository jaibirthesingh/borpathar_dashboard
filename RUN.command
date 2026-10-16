#!/bin/bash
# BORPATHAR BLF — refresh the data pipe. Double-click to update data.xlsx.
cd "$(dirname "$0")"
echo "=============================================="
echo "  BORPATHAR BLF — refreshing data.xlsx"
echo "=============================================="
if ! command -v python3 >/dev/null 2>&1; then
  echo; echo "  Python 3 isn't installed. Run:  xcode-select --install"; echo
  read -p "Press Enter to close."; exit 1
fi
if [ ! -d ".venv" ]; then
  echo "  First run — setting up (one minute, one time)..."
  python3 -m venv .venv
  ./.venv/bin/pip install --quiet --upgrade pip
  ./.venv/bin/pip install --quiet pdfplumber xlrd openpyxl
fi
echo "  Reading your reports..."
./.venv/bin/python scripts/build_data.py
STATUS=$?
echo
if [ $STATUS -eq 0 ]; then
  echo "  Done. data.xlsx refreshed."
  echo "  Now open your master dashboard and hit Update (or it refreshes on open)."
else
  echo "  Something went wrong above. Screenshot it and send to Claude."
fi
echo
read -p "Press Enter to close."
