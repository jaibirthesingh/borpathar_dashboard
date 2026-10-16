#!/bin/bash
# BORPATHAR BLF dashboard — double-click to rebuild. Self-bootstrapping.
cd "$(dirname "$0")"
echo "=============================================="
echo "  BORPATHAR BLF — rebuilding dashboard"
echo "=============================================="

if ! command -v python3 >/dev/null 2>&1; then
  echo
  echo "  Python 3 isn't installed on this Mac."
  echo "  One-time fix: open Terminal and run:  xcode-select --install"
  echo "  (or install from https://www.python.org/downloads/macos/)"
  echo
  read -p "Press Enter to close."; exit 1
fi

# first-run: create a local venv and install the libraries
if [ ! -d ".venv" ]; then
  echo "  First run — setting up (one minute, one time)..."
  python3 -m venv .venv
  ./.venv/bin/pip install --quiet --upgrade pip
  ./.venv/bin/pip install --quiet pdfplumber xlrd openpyxl
fi

echo "  Reading your reports..."
./.venv/bin/python scripts/build_dashboard.py
STATUS=$?
echo
if [ $STATUS -eq 0 ]; then
  echo "  Done. Opening dashboard.xlsx ..."
  open dashboard.xlsx
else
  echo "  Something went wrong above. Screenshot it and send to Claude."
fi
echo
read -p "Press Enter to close."
