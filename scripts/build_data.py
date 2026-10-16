"""
BORPATHAR BLF — DATA PIPE.
Reads the four feed folders and writes ../data.xlsx: clean tidy tables only,
one per dataset, as Excel Tables. NO charts, NO layout — that lives in your
own master dashboard, which pulls from this file.
Run via RUN.command. Safe to run any time.
"""
import os, glob, sys
from datetime import datetime
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import mfg_parser, advice_parser, green_leaf_parser, tally_parser, mandays_parser

RECOVERY = 0.22
MULT = 4.545
CONV_COST = 45.0
SEASON_SOLD_KG = 149583.2  # all-channel sold to date (sales ledger, manual)

# 2024 reference for YoY (historical; 1-Apr-24 to 31-Aug-24, made tea 113,770 kg)
YOY_2024_MADE = 113770.0
YOY_2024 = {
    "Power & Fuel Expenses": 3299060.00, "Bank Interest and Charges": 1392867.36,
    "Employees Benefit Expenses": 1191322.97, "Packing Expenses": 300861.00,
    "Selling & Distribution Expenses": 277715.67, "Insurance Charges": 178591.00,
    "Freight, Travelling, Carriage, Loading": 147000.00, "Repairing & Maintenance": 120694.18,
    "GOVT RATES, FEES AND TAXES": 52352.00, "Miscellaneous Expenses": 24643.00,
}

# ---------- load feeds ----------
mfg = []
for f in sorted(glob.glob(os.path.join(ROOT, "mfg_reports", "*.xls*"))):
    try: mfg.append(mfg_parser.parse_mfg(f))
    except Exception as e: print("  ! mfg skip", f, e)
mfg = [m for m in mfg if m.get("date")]; mfg.sort(key=lambda m: m["date"])

advices = []
for f in sorted(glob.glob(os.path.join(ROOT, "sale_advices", "*.pdf"))):
    try: advices.append(advice_parser.parse_advice(f))
    except Exception as e: print("  ! advice skip", f, e)
advices = [a for a in advices if a.get("sale_no")]; advices.sort(key=lambda a: a["sale_no"])

green = []
for f in glob.glob(os.path.join(ROOT, "green_leaf_prices", "*.csv")) + glob.glob(os.path.join(ROOT, "green_leaf_prices", "*.xls*")):
    try: green += green_leaf_parser.parse_green_leaf(f)
    except Exception as e: print("  ! green skip", f, e)
green.sort(key=lambda r: r["period_from"])

# tally: latest cumulative file wins
import re as _re, datetime as _dt
def _tally_enddate(path):
    try:
        if path.lower().endswith(".pdf"):
            import pdfplumber
            with pdfplumber.open(path) as _p:
                t = "\n".join(pg.extract_text() or "" for pg in _p.pages)
            m = _re.search(r"to\s+(\d{1,2})-([A-Za-z]{3})-(\d{2,4})", t)
            if m:
                mon = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}[m.group(2).lower()]
                yr = int(m.group(3)); yr = yr+2000 if yr < 100 else yr
                return _dt.date(yr, mon, int(m.group(1))).isoformat()
    except Exception: pass
    mm = _re.findall(r"(\d{4}-\d{2}-\d{2})", os.path.basename(path))
    if mm: return max(mm)
    return _dt.date.fromtimestamp(os.path.getmtime(path)).isoformat()
_tc = []
for f in glob.glob(os.path.join(ROOT, "tally_reports", "*.csv")) + glob.glob(os.path.join(ROOT, "tally_reports", "*.pdf")):
    try:
        rows = tally_parser.parse_tally(f)
        if rows: _tc.append((_tally_enddate(f), rows))
    except Exception as e: print("  ! tally skip", f, e)
tally = max(_tc, key=lambda x: x[0])[1] if _tc else []

# mandays: one TOTAL per fortnightly wages file, summed to date
mandays = []
for f in glob.glob(os.path.join(ROOT, "mandays", "*.xls*")):
    if os.path.basename(f).startswith("~$"): continue
    try:
        m = mandays_parser.parse_mandays(f)
        if m and m.get("total_mandays") is not None: mandays.append(m)
    except Exception as e: print("  ! mandays skip", f, e)
mandays.sort(key=lambda m: (m.get("start") or ""))
mandays_total = sum(m["total_mandays"] for m in mandays)

print(f"loaded: {len(mfg)} mfg | {len(advices)} advices | {len(green)} green | {len(tally)} tally heads | {len(mandays)} mandays periods")

# ---------- derive ----------
gl_kg = sum(r["kg"] for r in green); gl_spend = sum(r["amount"] for r in green)
gl_rate = gl_spend/gl_kg if gl_kg else 0
tea_todate = mfg[-1]["tea_made_todate_kg"] if mfg else gl_kg*RECOVERY
gl_todate = mfg[-1]["gl_todate_kg"] if mfg else gl_kg
recovery_td = (mfg[-1]["recovery_todate_pct"] if mfg else RECOVERY*100)
mandays_per_kg = (mandays_total/tea_todate) if tea_todate else 0
kg_per_manday = (tea_todate/mandays_total) if mandays_total else 0

tally_total = sum(t["amount"] for t in tally)
tvar = sum(t["amount"] for t in tally if t["cls"]=="variable")
tfix = sum(t["amount"] for t in tally if t["cls"]=="fixed")
tsemi = sum(t["amount"] for t in tally if t["cls"]=="semi")
tloaded = tally_total > 0
conv_used = (tally_total/tea_todate) if (tloaded and tea_todate) else CONV_COST
marg_conv = ((tvar + 0.5*tsemi)/tea_todate) if (tloaded and tea_todate) else 0
made_cost = gl_rate*MULT + conv_used
marginal_floor = gl_rate*MULT + marg_conv

from collections import defaultdict, Counter
inv = defaultdict(lambda: False)
for a in advices:
    for lot in a["lots"]:
        if lot["sold"]: inv[lot["inv"]] = True
        elif lot["inv"] not in inv: inv[lot["inv"]] = False
uniq = len(inv); ever = sum(1 for v in inv.values() if v); clearance = ever/uniq if uniq else 0
auc_sold_kg = sum(a["sold"] or 0 for a in advices)
auc_sold_val = sum((a["sold"] or 0)*(a["avg"] or 0) for a in advices)
auc_real = auc_sold_val/auc_sold_kg if auc_sold_kg else 0
stock = tea_todate - SEASON_SOLD_KG

# ---------- write workbook ----------
wb = openpyxl.Workbook(); wb.remove(wb.active)
TSTYLE = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)

def sheet(name, headers, rows):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for r in rows: ws.append(r)
    nrow = len(rows) + 1; ncol = len(headers)
    ref = f"A1:{get_column_letter(ncol)}{max(nrow,2)}"
    t = Table(displayName=name, ref=ref); t.tableStyleInfo = TSTYLE
    ws.add_table(t)
    for i in range(1, ncol+1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, len(str(headers[i-1]))+3)
    return ws

# KPIs
sheet("KPIs", ["metric", "value", "unit"], [
    ["green_leaf_taken", round(gl_kg,1), "kg"],
    ["green_spend", round(gl_spend,0), "INR"],
    ["avg_leaf_rate", round(gl_rate,2), "INR/kg"],
    ["tea_made_todate", round(tea_todate,0), "kg"],
    ["recovery_todate", round(recovery_td,2), "%"],
    ["conversion_cost", round(conv_used,2), "INR/kg"],
    ["made_tea_cost", round(made_cost,2), "INR/kg"],
    ["auction_realization", round(auc_real,2), "INR/kg"],
    ["full_cost_spread", round(auc_real-made_cost,2), "INR/kg"],
    ["marginal_floor", round(marginal_floor,2), "INR/kg"],
    ["headroom_above_floor", round(auc_real-marginal_floor,2), "INR/kg"],
    ["invoice_clearance", round(clearance,4), "fraction"],
    ["closing_stock", round(stock,0), "kg"],
    ["mandays_todate", round(mandays_total,1), "mandays"],
    ["mandays_per_kg", round(mandays_per_kg,5), "mandays/kg"],
    ["kg_per_manday", round(kg_per_manday,1), "kg/manday"],
    ["generated", datetime.now().strftime("%Y-%m-%d %H:%M"), "timestamp"],
])
# Mandays — per-period totals (trend)
sheet("Mandays", ["period_start","period_end","total_mandays"],
    [[m.get("start"), m.get("end"), m["total_mandays"]] for m in mandays])
# GreenLeaf
sheet("GreenLeaf", ["period_from","period_to","gl_kg","rate_inr_kg","amount_inr","made_equiv_kg","made_cost_inr_kg"],
    [[g["period_from"], g["period_to"], round(g["kg"],1), round(g["rate"],2), round(g["amount"],0),
      round(g["kg"]*RECOVERY,1), round(g["rate"]*MULT+conv_used,2)] for g in green])
# Auction
sheet("Auction", ["sale_no","date","offered_kg","sold_kg","avg_inr_kg","unsold_pct"],
    [[a["sale_no"], a["date"], round(a["offered"] or 0,1), round(a["sold"] or 0,1), round(a["avg"] or 0,2),
      round((1-(a["sold"] or 0)/a["offered"]) if a["offered"] else 0,4)] for a in advices])
# Production
gcols = ["BOP","SM","BP","PF","PD","DUST","SEC"]
sheet("Production", ["date","gl_consumed_kg","gl_received_kg","tea_made_kg","recovery_today_pct","fine_pct","labour_inr_kg","coal_kg_kg"]+gcols,
    [[m["date"], m.get("gl_consumed_kg"), m.get("gl_received_today_kg"), m.get("tea_made_today_kg"),
      m.get("recovery_today_pct"), (m.get("avg_gl_fine_pct") or 0)*100, round(m.get("labour_cost_perkg_today") or 0,2),
      round(m.get("coal_perkg_today") or 0,3)] + [(m.get("sorting_grades_kg") or {}).get(g) for g in gcols] for m in mfg])
# Expenses
if tloaded:
    sheet("Expenses", ["head","amount_inr","class","pct_of_total","per_kg_made"],
        [[t["head"], round(t["amount"],0), t["cls"], round(t["amount"]/tally_total,4), round(t["amount"]/tea_todate,2)]
         for t in sorted(tally, key=lambda x:-x["amount"])])
else:
    sheet("Expenses", ["head","amount_inr","class","pct_of_total","per_kg_made"], [])
# Economics
sheet("Economics", ["item","value","unit"], [
    ["green_leaf_kg", round(gl_kg,1), "kg"],
    ["green_spend_inr", round(gl_spend,0), "INR"],
    ["avg_leaf_rate", round(gl_rate,2), "INR/kg"],
    ["tally_total_inr", round(tally_total,0), "INR"],
    ["variable_inr", round(tvar,0), "INR"],
    ["semi_inr", round(tsemi,0), "INR"],
    ["fixed_inr", round(tfix,0), "INR"],
    ["conversion_inr_kg", round(conv_used,2), "INR/kg"],
    ["marginal_conversion_inr_kg", round(marg_conv,2), "INR/kg"],
    ["tea_made_kg", round(tea_todate,0), "kg"],
    ["made_cost_inr_kg", round(made_cost,2), "INR/kg"],
    ["auction_sold_kg", round(auc_sold_kg,1), "kg"],
    ["auction_realization_inr_kg", round(auc_real,2), "INR/kg"],
    ["all_channel_sold_kg", SEASON_SOLD_KG, "kg"],
    ["full_cost_spread_inr_kg", round(auc_real-made_cost,2), "INR/kg"],
    ["marginal_floor_inr_kg", round(marginal_floor,2), "INR/kg"],
    ["closing_stock_kg", round(stock,0), "kg"],
    ["mandays_todate", round(mandays_total,1), "mandays"],
    ["mandays_per_kg", round(mandays_per_kg,5), "mandays/kg"],
    ["kg_per_manday", round(kg_per_manday,1), "kg/manday"],
])
# YoY Expenses (2024 vs 2026, per kg)
yrows = []
head_map = {h: t["amount"] for t in tally for h in [t["head"]]}
# align 2024 heads to 2026 by loose name match
def find26(h24):
    key = h24.lower().split()[0]
    for t in tally:
        if key in t["head"].lower(): return t["amount"]
    return None
for h24, amt24 in YOY_2024.items():
    amt26 = find26(h24)
    pk24 = amt24/YOY_2024_MADE
    pk26 = (amt26/tea_todate) if amt26 else None
    yrows.append([h24, round(amt24,0), round(pk24,2),
                  round(amt26,0) if amt26 else None,
                  round(pk26,2) if pk26 else None,
                  round(pk26-pk24,2) if pk26 else None])
yrows.append(["TOTAL (all heads)", round(sum(YOY_2024.values()),0), round(sum(YOY_2024.values())/YOY_2024_MADE,2),
              round(tally_total,0), round(tally_total/tea_todate,2),
              round(tally_total/tea_todate - sum(YOY_2024.values())/YOY_2024_MADE,2)])
sheet("YoY_Expenses", ["head","amount_2024","per_kg_2024","amount_2026","per_kg_2026","delta_per_kg"], yrows)

out = os.path.join(ROOT, "data.xlsx")
wb.save(out)
print("wrote", out, "with tabs:", wb.sheetnames)
