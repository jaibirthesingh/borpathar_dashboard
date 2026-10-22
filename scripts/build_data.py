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
import mfg_parser, advice_parser, green_leaf_parser, tally_parser, mandays_parser, electricity_parser, private_advice_parser, grade_gap_parser, leaf_quality_parser

RECOVERY = 0.22
MULT = 4.545
CONV_COST = 45.0
# (removed hardcoded SEASON_SOLD_KG — stock is now computed live from auction sold every run)

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

# grade-gap benchmark: one Parcon Range-Grade-Report .xlsx per sale (GUW+BLF market avg/grade)
gradegap = []
for f in sorted(glob.glob(os.path.join(ROOT, "grade_gap", "*.xls*"))):
    if os.path.basename(f).startswith("~$"): continue
    try: gradegap.append(grade_gap_parser.parse_grade_gap(f))
    except Exception as e: print("  ! grade_gap skip", f, e)
gradegap = [g for g in gradegap if g.get("sale_no")]; gradegap.sort(key=lambda g: g["sale_no"])

# leaf-quality weekly (GL_KG / AV.RATE export): separate feed for the fine%/premium work.
# Does NOT touch the green-leaf price feed above.
leafq = []
for f in (glob.glob(os.path.join(ROOT, "leaf_quality", "*.pdf")) +
          glob.glob(os.path.join(ROOT, "leaf_quality", "*.xls*")) +
          glob.glob(os.path.join(ROOT, "leaf_quality", "*.csv"))):
    if os.path.basename(f).startswith("~$"): continue
    try: leafq += leaf_quality_parser.parse_leaf_quality(f)
    except Exception as e: print("  ! leaf_quality skip", f, e)
_lq = {}
for r in leafq: _lq[r["period_from"]] = r
leafq = sorted(_lq.values(), key=lambda r: r["period_from"])
# weighted-avg green-leaf fine % to date (kg-weighted, from the leaf-quality feed)
_flk = sum(r["gl_kg"] for r in leafq if r.get("fine_pct") is not None and r.get("gl_kg"))
_flw = sum(r["gl_kg"]*r["fine_pct"] for r in leafq if r.get("fine_pct") is not None and r.get("gl_kg"))
avg_gl_fine_pct = (_flw/_flk) if _flk else 0

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

# electricity: one bill per month -> kWh + cost, summed
elec = []
for f in glob.glob(os.path.join(ROOT, "electricity_bills", "*.pdf")):
    try:
        b = electricity_parser.parse_bill(f)
        if b and b.get("units_kwh") is not None: elec.append(b)
    except Exception as e: print("  ! elec skip", f, e)
elec.sort(key=lambda b: (b.get("start") or ""))
elec_kwh = sum(b["units_kwh"] for b in elec)
elec_cost = sum(b.get("bill_amount") or 0 for b in elec)
elec_maxdemand = max((b.get("max_demand_kva") or 0 for b in elec), default=0)
elec_contracted = elec[-1].get("contracted_kva") if elec else None
elec_pf = elec[-1].get("power_factor") if elec else None

# private sales: per-invoice GST .xlsx (HUL, Tata, direct) — NON-auction, grade-wise.
# Adds to the weekly auction figure; no overlap. One invoice = one file (subfolders ok).
priv_invoices = []
for f in glob.glob(os.path.join(ROOT, "private_sale_advices", "**", "*.xls*"), recursive=True):
    if os.path.basename(f).startswith("~$"): continue
    try:
        inv = private_advice_parser.parse_private_invoice(f)
        if inv and inv.get("lines"): priv_invoices.append(inv)
    except Exception as e: print("  ! private skip", f, e)
priv_lines = [l for inv in priv_invoices for l in inv["lines"]]
priv_sold_kg = sum(l["qty"] for l in priv_lines)
priv_sold_val = sum(l["amount"] or 0 for l in priv_lines)

print(f"loaded: {len(mfg)} mfg | {len(advices)} advices | {len(green)} green | {len(tally)} tally heads | {len(mandays)} mandays periods")

# ---------- derive ----------
gl_kg = sum(r["kg"] for r in green); gl_spend = sum(r["amount"] for r in green)
gl_rate = gl_spend/gl_kg if gl_kg else 0
tea_todate = mfg[-1]["tea_made_todate_kg"] if mfg else gl_kg*RECOVERY
gl_todate = mfg[-1]["gl_todate_kg"] if mfg else gl_kg
recovery_td = (mfg[-1]["recovery_todate_pct"] if mfg else RECOVERY*100)
mandays_per_kg = (mandays_total/tea_todate) if tea_todate else 0
kg_per_manday = (tea_todate/mandays_total) if mandays_total else 0
kwh_per_kg = (elec_kwh/tea_todate) if tea_todate else 0
electricity_inr_kg = (elec_cost/tea_todate) if tea_todate else 0

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
# primary vs secondary realization: second-grade teas end in "1" (BP1, D1, BOPSM1)
_pv=_pk=_sv=_sk=0
for _a in advices:
    for _lot in _a["lots"]:
        if _lot["sold"] and _lot["price"]:
            if _re.search(r"1$", _lot["grade"]): _sv+=_lot["price"]*_lot["net_wt"]; _sk+=_lot["net_wt"]
            else: _pv+=_lot["price"]*_lot["net_wt"]; _pk+=_lot["net_wt"]
# fold private-sale invoice lines in -> primary/secondary become ALL-channel
for _l in priv_lines:
    if _l["amount"] and _l["qty"]:
        if _re.search(r"1$", _l["grade"]): _sv+=_l["amount"]; _sk+=_l["qty"]
        else: _pv+=_l["amount"]; _pk+=_l["qty"]
avg_primary_rate = _pv/_pk if _pk else 0
avg_secondary_rate = _sv/_sk if _sk else 0
# total sold = auction (weekly advices) + private (invoices). They add; no overlap.
all_channel_sold = auc_sold_kg + priv_sold_kg
stock = tea_todate - all_channel_sold

# --- NEW METRICS (Oct 2026): channel split, grade prices, primary/secondary made % ---
# Channel split by kg (share of total tonnage sold)
auction_sale_pct = (auc_sold_kg/all_channel_sold*100) if all_channel_sold else 0
private_sale_pct = (priv_sold_kg/all_channel_sold*100) if all_channel_sold else 0
# Grade-wise avg sold price, ALL-CHANNEL (auction lots + private invoice lines), weighted by kg
_gp = defaultdict(lambda: [0.0, 0.0])  # grade -> [value_inr, kg]
for _a in advices:
    for _lot in _a["lots"]:
        if _lot["sold"] and _lot["price"]:
            _gp[_lot["grade"]][0] += _lot["price"]*_lot["net_wt"]
            _gp[_lot["grade"]][1] += _lot["net_wt"]
for _l in priv_lines:
    if _l["amount"] and _l["qty"]:
        _gp[_l["grade"]][0] += _l["amount"]
        _gp[_l["grade"]][1] += _l["qty"]
grade_prices = sorted(([g, round(v/k, 2), round(k, 1)] for g, (v, k) in _gp.items() if k > 0),
                      key=lambda r: -r[2])  # sorted by sold kg, biggest first
# Primary vs Secondary made tea %, from MFG daily sorting.
# SEC column = secondary ("1"-suffix marks) bucket; everything else (incl. DUST) = primary.
_mp = _ms = 0.0
for _m in mfg:
    for _g, _kg in (_m.get("sorting_grades_kg") or {}).items():
        if _g == "TOTAL" or _kg is None: continue
        if _g == "SEC": _ms += _kg
        else: _mp += _kg
_made_sorted = _mp + _ms
primary_made_pct = (_mp/_made_sorted*100) if _made_sorted else 0
secondary_made_pct = (_ms/_made_sorted*100) if _made_sorted else 0

# --- PREMIUM OVER MARKET (primary grades; benchmark = GUW+BLF grade gap) ---
# Secondaries ("1"-suffix) have no benchmark in the report -> primary-only by design.
_PRIM = {"BOP","BOPSM","BP","PF","PD","D"}
_mkt = {}  # (sale, grade) -> market avg price (current season)
for _g in gradegap:
    for _grd, _px in _g["grades"].items():
        _mkt[(_g["sale_no"], _grd)] = _px
# my auction price per grade per sale, kg-weighted over sold lots
_myg = defaultdict(lambda: [0.0, 0.0])  # (sale, grade) -> [value, kg]
for _a in advices:
    for _lot in _a["lots"]:
        if _lot["sold"] and _lot["price"] and _lot["grade"] in _PRIM:
            _myg[(_a["sale_no"], _lot["grade"])][0] += _lot["price"]*_lot["net_wt"]
            _myg[(_a["sale_no"], _lot["grade"])][1] += _lot["net_wt"]
_bench_rows = sorted(([g["sale_no"], grd, round(px,2)] for g in gradegap for grd,px in g["grades"].items()),
                     key=lambda r: (r[0], r[1]))
_prem_rows = []
_bysale = defaultdict(lambda: [0.0, 0.0, 0.0])  # sale -> [sum(kg*prem), sum(kg), sum(kg*my)]
for (_sale, _grd), (_val, _kg) in sorted(_myg.items()):
    if _kg <= 0: continue
    _m = _mkt.get((_sale, _grd))
    if _m is None: continue
    _my = _val/_kg; _p = _my - _m
    _prem_rows.append([_sale, _grd, round(_my,2), round(_m,2), round(_p,1), round(_kg,1)])
    _bysale[_sale][0] += _kg*_p; _bysale[_sale][1] += _kg; _bysale[_sale][2] += _kg*_my
_adate = {a["sale_no"]: a["date"] for a in advices}
_prembysale_rows = []
for _sale in sorted(_bysale):
    _wp, _w, _wmy = _bysale[_sale]
    if _w <= 0: continue
    _prembysale_rows.append([_sale, _adate.get(_sale), round(_wp/_w,1), round(_w,1),
                             round(_wmy/_w,2), round(_wmy/_w - _wp/_w,2)])

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
    ["avg_primary_rate", round(avg_primary_rate,2), "INR/kg"],
    ["avg_secondary_rate", round(avg_secondary_rate,2), "INR/kg"],
    ["full_cost_spread", round(auc_real-made_cost,2), "INR/kg"],
    ["marginal_floor", round(marginal_floor,2), "INR/kg"],
    ["headroom_above_floor", round(auc_real-marginal_floor,2), "INR/kg"],
    ["invoice_clearance", round(clearance,4), "fraction"],
    ["closing_stock", round(stock,0), "kg"],
    ["mandays_todate", round(mandays_total,1), "mandays"],
    ["kg_per_manday", round(kg_per_manday,1), "kg/manday"],
    ["kwh_per_kg", round(kwh_per_kg,3), "kWh/kg"],
    ["electricity_inr_kg", round(electricity_inr_kg,2), "INR/kg"],
    ["max_demand_kva", round(elec_maxdemand,1) if elec_maxdemand else None, "KVA"],
    ["contracted_kva", elec_contracted, "KVA"],
    ["power_factor", elec_pf, "%"],
    ["Auction Sale %", round(auction_sale_pct, 2), "%"],
    ["Private Sale %", round(private_sale_pct, 2), "%"],
    ["Primary Made Tea %", round(primary_made_pct, 2), "%"],
    ["Secondary Made Tea %", round(secondary_made_pct, 2), "%"],
    ["Avg GL Fine %", round(avg_gl_fine_pct, 1), "%"],
    ["generated", datetime.now().strftime("%Y-%m-%d %H:%M"), "timestamp"],
])
# Mandays — per-period totals (trend)
sheet("Mandays", ["period_start","period_end","total_mandays"],
    [[m.get("start"), m.get("end"), m["total_mandays"]] for m in mandays])
# Electricity — per-bill (trend)
sheet("Electricity", ["period_start","period_end","units_kwh","bill_amount_inr","max_demand_kva","power_factor"],
    [[b.get("start"), b.get("end"), b.get("units_kwh"), b.get("bill_amount"),
      b.get("max_demand_kva"), b.get("power_factor")] for b in elec])
# GreenLeaf
sheet("GreenLeaf", ["period_from","period_to","gl_kg","rate_inr_kg","amount_inr","made_equiv_kg","made_cost_inr_kg"],
    [[g["period_from"], g["period_to"], round(g["kg"],1), round(g["rate"],2), round(g["amount"],0),
      round(g["kg"]*RECOVERY,1), round(g["rate"]*MULT+conv_used,2)] for g in green])
# Auction
sheet("Auction", ["sale_no","date","offered_kg","sold_kg","avg_inr_kg","unsold_pct"],
    [[a["sale_no"], a["date"], round(a["offered"] or 0,1), round(a["sold"] or 0,1), round(a["avg"] or 0,2),
      round((1-(a["sold"] or 0)/a["offered"]) if a["offered"] else 0,4)] for a in advices])
# GradePrices — all-channel (auction + private) avg sold price per grade, kg-weighted
sheet("GradePrices", ["Grade","Avg Sold Price","Sold Kg"], grade_prices)
# Benchmark — GUW+BLF market avg price, per sale per grade (Parcon grade gap)
sheet("Benchmark", ["sale_no","grade","market_inr_kg"], _bench_rows)
# Premium — my auction price minus market, per sale per primary grade
sheet("Premium", ["sale_no","grade","my_inr_kg","market_inr_kg","premium_inr_kg","sold_kg"], _prem_rows)
# PremiumBySale — headline: volume-weighted premium per sale (your price vs market)
sheet("PremiumBySale", ["sale_no","date","vol_wtd_premium","sold_kg","my_avg_inr_kg","market_avg_inr_kg"], _prembysale_rows)
# LeafQuality — weekly GL_KG/AV.RATE feed: fine% + volume for the quality analysis (separate feed)
sheet("LeafQuality", ["period_from","period_to","gl_kg","rate_inr_kg","amount_inr","fine_pct"],
    [[r["period_from"], r["period_to"], round(r["gl_kg"],1), round((r["rate"] or 0),2),
      round((r["amount"] or 0),0), r["fine_pct"]] for r in leafq])
# LeafVsAuction — LAGGED: each sale paired with the leaf (rate+fine%) from ~3wk prior that fed it.
# For the headline quality chart: purchased GL rate + fine% vs auction realization.
def _iso2d(s):
    try: return _dt.date.fromisoformat(str(s))
    except Exception: return None
_mkt_by_sale = {r[0]: r[5] for r in _prembysale_rows}  # sale_no -> market_avg_inr_kg
# Keyed by LEAF WEEK so leaf extends back to 01-Apr. Each week gets the auction number it feeds,
# back-calculated from the weekly cadence (anchor: sale 20 = 12-May-2026) and the 3-week leaf->auction
# lag. Auction/market fill only where a real sale exists (20-42); earlier leaf sits at sale<20,
# and the most recent leaf (not yet auctioned) at sale>42. One contiguous table => Mac-safe chart.
_SALE20_NO, _SALE20_DT = 20, _dt.date(2026, 5, 12)
_sale_data = {a["sale_no"]: (round(a["avg"], 2) if a.get("avg") else None, _mkt_by_sale.get(a["sale_no"]))
              for a in advices if a.get("sale_no")}
_lv = []
for w in leafq:
    _pt = _iso2d(w["period_to"])
    if not _pt: continue
    _fed = _SALE20_NO + round(((_pt + _dt.timedelta(days=21)) - _SALE20_DT).days / 7)
    _auc, _mkt = _sale_data.get(_fed, (None, None))
    _lv.append([_fed, w["period_from"], round((w.get("rate") or 0), 2),
                w.get("fine_pct"), _auc, _mkt])
_lv.sort(key=lambda r: r[0])
sheet("LeafVsAuction", ["sale_no","leaf_week_from","gl_rate_inr_kg","fine_pct","auction_avg_inr_kg","market_avg_inr_kg"], _lv)
# Production
gcols = ["BOP","SM","BP","PF","PD","DUST","SEC"]
sheet("Production", ["date","gl_consumed_kg","gl_received_kg","tea_made_kg","recovery_today_pct","fine_pct","labour_inr_kg","coal_kg_kg"]+gcols,
    [[m["date"], m.get("gl_consumed_kg"), m.get("gl_received_today_kg"), m.get("tea_made_today_kg"),
      m.get("recovery_today_pct"), (m.get("avg_gl_fine_pct") or 0)*100, round(m.get("labour_cost_perkg_today") or 0,2),
      round(m.get("coal_perkg_today") or 0,3)] + [(m.get("sorting_grades_kg") or {}).get(g) for g in gcols] for m in mfg])
# Expenses — line-item labels cleaned to your dashboard names at source (survives refresh).
# Raw Tally ledger head -> clean label. Edit here to add/adjust; class & YoY matching are
# unaffected (they read the raw tally list, not these display labels).
_EXP_LABEL_SRC = {
    "Power & Fuel Expenses": "Power & Fuel",
    "Bank Interest and Charges": "Bank Interest & Charges",
    "Employees Benefit Expenses": "Salary & Labour",
    "Packing Expenses": "Packaging",
    "Freight Travelling Carriage Loading": "Transport",
    "Selling & Distribution Expenses": "Selling & Distribution",
    "Repairing & Maintenance": "Repairing & Maintenance",
    "Insurance Charges": "Insurance Charges",
    "INTEREST (PARCON INDIA)": "Interest-Parcon",
    "Miscellaneous Expenses": "Miscellaneous",
    "GOVT RATES FEES AND TAXES": "Govt Taxes",
    "Professional-Fee": "Professional-Fee",
    "Office Rental Expenses": "HO Rental",
}
_norm = lambda s: _re.sub(r"\s+", " ", str(s)).strip().upper()
_EXP_LABEL = {_norm(k): v for k, v in _EXP_LABEL_SRC.items()}
def _exp_label(raw): return _EXP_LABEL.get(_norm(raw), raw)
for _t in tally:  # loudly flag any new ledger head with no clean label yet (shows raw)
    if _norm(_t["head"]) not in _EXP_LABEL:
        print(f"  ! expense head has no dashboard label (showing raw): {_t['head']!r}")
if tloaded:
    sheet("Expenses", ["Parameter","Amount","class","pct_of_total","per_kg_made"],
        [[_exp_label(t["head"]), round(t["amount"],0), t["cls"], round(t["amount"]/tally_total,4), round(t["amount"]/tea_todate,2)]
         for t in sorted(tally, key=lambda x:-x["amount"])])
else:
    sheet("Expenses", ["Parameter","Amount","class","pct_of_total","per_kg_made"], [])
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
    ["all_channel_sold_kg", round(all_channel_sold,1), "kg"],
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

# Timeline — cost vs realization on one weekly grid (auction dates).
# Leaf rate carried forward: the rate in force on each auction date.
def _leaf_rate_on(dstr):
    r = None
    for g in green:  # green sorted by period_from
        if g["period_from"] <= dstr: r = g["rate"]
    return r
_tl = []
for a in advices:
    if not a.get("date"): continue
    lr = _leaf_rate_on(a["date"])
    cost = round(lr*MULT + conv_used, 2) if lr else None
    real = a["avg"] if a.get("avg") else None   # skip 100%-unsold weeks
    _tl.append([a["date"], real, cost])
sheet("Timeline", ["date","realization","made_cost"], _tl)

out = os.path.join(ROOT, "data.xlsx")
wb.save(out)
print("wrote", out, "with tabs:", wb.sheetnames)
