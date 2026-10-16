BORPATHAR BLF — MASTER DASHBOARD
================================

DAILY USE (the whole manual):
  1. Drop new reports into the matching folder:
       daily MFG .xls         -> mfg_reports/
       weekly Parcon .pdf     -> sale_advices/
       fortnightly leaf price -> green_leaf_prices/
       monthly Tally .pdf     -> tally_reports/
  2. Double-click  RUN.command
  3. Open  dashboard.xlsx

Filenames don't matter — dates are read from inside each file.
Never hand-edit dashboard.xlsx; it is rebuilt from scratch every run.

GREEN LEAF FORMAT (green_leaf_prices/):
  CSV with columns: period_from,period_to,kg,rate,amount
  (send Claude one real fortnightly file to switch this to your actual format)

STATUS: Tally feed not wired yet — conversion cost sits on the Rs 45
assumption until a sample Tally PDF is provided. Everything else is live.
