"""
BORPATHAR BLF — master dashboard builder.
Reads the four feed folders, parses everything, writes ../dashboard.xlsx.
Run via ../RUN.command (self-bootstrapping). Safe to run any time.
"""
import os, glob, sys
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import mfg_parser, advice_parser, green_leaf_parser, tally_parser

# ---------- assumptions (single source of truth) ----------
RECOVERY = 0.22          # made tea / green leaf
MULT = 4.545             # green kg per made-tea kg (= 1/recovery)
CONV_COST = 45.0         # Rs/kg conversion — PENDING TALLY (placeholder)
SEASON_SOLD_KG = 149583.2  # ALL-channel sold to date, from Primary+Secondary sales ledger
                           # (manual cross-check; not a live feed yet)

# ---------- styles ----------
F = "Arial"
H1   = Font(name=F, size=16, bold=True, color="1F3A2E")
H2   = Font(name=F, size=12, bold=True, color="FFFFFF")
BOLD = Font(name=F, size=10, bold=True)
BODY = Font(name=F, size=10)
MUTE = Font(name=F, size=9, italic=True, color="7A7A7A")
BLUE = Font(name=F, size=10, color="0000FF")          # parsed/input values
HERO = Font(name=F, size=20, bold=True, color="1F3A2E")
HEADFILL = PatternFill("solid", fgColor="2E6B4F")
BANDFILL = PatternFill("solid", fgColor="EAF2EC")
WARNFILL = PatternFill("solid", fgColor="FFF4CC")
thin = Side(style="thin", color="D5D5D5")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header_row(ws, row, c1, c2):
    for c in range(c1, c2+1):
        cell = ws.cell(row=row, column=c)
        cell.font = H2; cell.fill = HEADFILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

def setw(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

# ---------- load feeds ----------
mfg = []
for f in sorted(glob.glob(os.path.join(ROOT, "mfg_reports", "*.xls*"))):
    try: mfg.append(mfg_parser.parse_mfg(f))
    except Exception as e: print("  ! mfg skip", f, e)
mfg = [m for m in mfg if m.get("date")]
mfg.sort(key=lambda m: m["date"])

advices = []
for f in sorted(glob.glob(os.path.join(ROOT, "sale_advices", "*.pdf"))):
    try: advices.append(advice_parser.parse_advice(f))
    except Exception as e: print("  ! advice skip", f, e)
advices = [a for a in advices if a.get("sale_no")]
advices.sort(key=lambda a: a["sale_no"])

green = []
for f in glob.glob(os.path.join(ROOT, "green_leaf_prices", "*.csv")) + \
         glob.glob(os.path.join(ROOT, "green_leaf_prices", "*.xls*")):
    try: green += green_leaf_parser.parse_green_leaf(f)
    except Exception as e: print("  ! green skip", f, e)
green.sort(key=lambda r: r["period_from"])

tally = []
for f in glob.glob(os.path.join(ROOT, "tally_reports", "*.csv")):
    try: tally += tally_parser.parse_tally(f)
    except Exception as e: print("  ! tally skip", f, e)

print(f"loaded: {len(mfg)} mfg | {len(advices)} advices | {len(green)} green rows | {len(tally)} tally rows")

# ---------- cross-advice clearance (dedupe by invoice) ----------
from collections import defaultdict, Counter
inv = defaultdict(lambda: {"sold": False, "appear": 0})
grade_sold = Counter()
for a in advices:
    for lot in a["lots"]:
        inv[lot["inv"]]["appear"] += 1
        if lot["sold"]:
            inv[lot["inv"]]["sold"] = True
            grade_sold[lot["grade"]] += lot["net_wt"]
uniq = len(inv); ever = sum(1 for v in inv.values() if v["sold"])
clearance = (ever/uniq) if uniq else 0

auc_sold_kg = sum(a["sold"] or 0 for a in advices)
auc_sold_val = sum((a["sold"] or 0)*(a["avg"] or 0) for a in advices)
auc_real = auc_sold_val/auc_sold_kg if auc_sold_kg else 0

gl_spend = sum(r["amount"] for r in green)
gl_kg = sum(r["kg"] for r in green)
gl_rate = gl_spend/gl_kg if gl_kg else 0
tea_todate = mfg[-1]["tea_made_todate_kg"] if mfg else gl_kg*RECOVERY
gl_todate = mfg[-1]["gl_todate_kg"] if mfg else gl_kg

# ---------- Tally expense aggregates ----------
tally_total = sum(t["amount"] for t in tally)
tally_var   = sum(t["amount"] for t in tally if t["cls"]=="variable")
tally_fix   = sum(t["amount"] for t in tally if t["cls"]=="fixed")
tally_semi  = sum(t["amount"] for t in tally if t["cls"]=="semi")
tally_loaded = tally_total > 0
conv_used = (tally_total/tea_todate) if (tally_loaded and tea_todate) else CONV_COST
marg_expense = tally_var + 0.5*tally_semi          # marginal basis: variable + half of semi
marg_conv_perkg = (marg_expense/tea_todate) if tea_todate else 0
marginal_floor = gl_rate*MULT + marg_conv_perkg    # below this realization: don't make

wb = openpyxl.Workbook()
GEN = datetime.now().strftime("%d %b %Y, %H:%M")

# =========================================================
# GUIDE
# =========================================================
ws = wb.active; ws.title = "Guide"; ws.sheet_view.showGridLines = False
setw(ws, {"A":2,"B":34,"C":52,"D":18})
ws["B2"]="BORPATHAR TEA CO. LLP"; ws["B2"].font=H1
ws["B3"]="Bought-Leaf Factory — Master Dashboard"; ws["B3"].font=Font(name=F,size=11,bold=True,color="2E6B4F")
ws["B4"]=f"Rebuilt: {GEN}"; ws["B4"].font=MUTE

r=6
ws.cell(r,2,"HOW TO OPERATE").font=BOLD
for i,step in enumerate([
    "1.  Drop new reports into the matching folder:",
    "         daily MFG .xls        →  /mfg_reports/",
    "         weekly Parcon .pdf    →  /sale_advices/",
    "         fortnightly leaf price →  /green_leaf_prices/  (CSV: period_from,period_to,kg,rate,amount)",
    "         monthly Tally .pdf    →  /tally_reports/",
    "2.  Double-click  RUN.command",
    "3.  Open  dashboard.xlsx  — fresh numbers.",
    "Filenames don't matter; dates are read from inside each file. Never hand-edit this workbook — it is rebuilt every run.",
]):
    ws.cell(r+1+i,2,step).font=BODY if not step.startswith("     ") else Font(name=F,size=9,color="444444")
r=16
ws.cell(r,2,"DATA LOADED").font=BOLD
status=[
    ("Daily MFG reports", f"{len(mfg)}", (mfg[0]['date']+' → '+mfg[-1]['date']) if mfg else "—"),
    ("Auction sale advices", f"{len(advices)}", (f"Sale {advices[0]['sale_no']}–{advices[-1]['sale_no']}") if advices else "—"),
    ("Green-leaf price rows", f"{len(green)}", (green[0]['period_from']+' → '+green[-1]['period_to']) if green else "—"),
    ("Tally expense heads", f"{len(tally)}", (f"loaded — ₹{tally_total/1e5:.1f}L  ·  actual conv ₹{conv_used:.1f}/kg" if tally_loaded else "PENDING — ₹45 assumption")),
]
ws.cell(r+1,2,"Feed").font=H2; ws.cell(r+1,3,"Range").font=H2; ws.cell(r+1,4,"Count").font=H2
style_header_row(ws,r+1,2,4)
for i,(a,b,c) in enumerate(status):
    rr=r+2+i
    ws.cell(rr,2,a).font=BODY; ws.cell(rr,3,c).font=BODY; ws.cell(rr,4,b).font=BLUE
    if "PENDING" in c:
        for cc in (2,3,4): ws.cell(rr,cc).fill=WARNFILL
r=23
ws.cell(r,2,"ASSUMPTIONS (change in build_dashboard.py)").font=BOLD
for i,(a,b) in enumerate([
    ("Recovery (made tea / green leaf)", f"{RECOVERY:.0%}"),
    ("Green→made multiplier (1/recovery)", f"{MULT}"),
    ("Conversion cost ₹/kg (from Tally)", (f"₹{conv_used:.1f}  (actual, all-in)" if tally_loaded else f"₹{CONV_COST:.0f}  ← PENDING TALLY")),
    ("Marginal floor ₹/kg (leaf + variable)", (f"₹{marginal_floor:.0f}" if tally_loaded else "—")),
]):
    ws.cell(r+1+i,2,a).font=BODY; ws.cell(r+1+i,3,b).font=BLUE
r=28
ws.cell(r,2,"KNOWN GAPS").font=BOLD
for i,g in enumerate([
    "• Tally wired from a transcribed, reconciled scan. Export Tally to a DIGITAL pdf for auto-parsing.",
    "• Verify: is factory labour fully inside 'Employees Benefit', or is there a separate Direct Expenses group?",
    "• Fixed/variable split is editable in tally_reports/*.csv (class column) — tune it to tighten the HOLD line.",
    "• Green-leaf parser reads a CSV seed; swap in your real fortnightly file format when ready.",
    "• Auction feed = Guwahati auction only; private/off-auction sales not captured here.",
]):
    ws.cell(r+1+i,2,g).font=Font(name=F,size=9,color="444444")

# =========================================================
# DASHBOARD
# =========================================================
ws = wb.create_sheet("Dashboard"); ws.sheet_view.showGridLines=False
setw(ws, {"A":2,"B":22,"C":16,"D":4,"E":22,"F":16,"G":4,"H":22,"I":16})
ws["B2"]="SEASON AT A GLANCE — 2026"; ws["B2"].font=H1
ws["B3"]=f"Rebuilt {GEN}  ·  provisional until Tally lands"; ws["B3"].font=MUTE
def kpi(anchor_col, row, label, value, sub=""):
    ws.cell(row,anchor_col,label).font=Font(name=F,size=9,bold=True,color="2E6B4F")
    ws.cell(row+1,anchor_col,value).font=HERO
    if sub: ws.cell(row+2,anchor_col,sub).font=MUTE
    for dr in range(0,3):
        ws.cell(row+dr,anchor_col).fill=BANDFILL
        ws.cell(row+dr,anchor_col+1).fill=BANDFILL
kpi(2,5,"GREEN LEAF TAKEN", f"{gl_kg:,.0f} kg", f"avg ₹{gl_rate:.2f}/kg")
kpi(5,5,"TEA MADE TO DATE", f"{tea_todate:,.0f} kg", f"recovery {mfg[-1]['recovery_todate_pct'] if mfg else RECOVERY*100:.1f}%")
kpi(8,5,"GREEN SPEND", f"₹{gl_spend/1e7:.2f} Cr", "input cost, season")
kpi(2,9,"AUCTION REALIZATION", f"₹{auc_real:.0f}/kg", f"on {auc_sold_kg:,.0f} kg sold")
kpi(5,9,"MADE-TEA COST", f"₹{gl_rate*MULT+conv_used:.0f}/kg", ("Tally actual ₹%.0f conv"%conv_used if tally_loaded else "₹45 assumed"))
kpi(8,9,"SPREAD", f"₹{auc_real-(gl_rate*MULT+conv_used):.0f}/kg", "realization − cost")
kpi(2,13,"INVOICE CLEARANCE", f"{clearance:.0%}", f"{ever}/{uniq} ever sold at auction")
kpi(5,13,"NON-CLEARANCE", f"{(1-clearance):.0%}", "stuck / carried / private")
kpi(8,13,"MARGINAL FLOOR", (f"₹{marginal_floor:.0f}/kg" if tally_loaded else "—"), "below this: don't make")

# data block for charts (hidden-ish, lower down)
dr0=22
ws.cell(dr0,2,"— chart data —").font=MUTE
# green leaf rate series
ws.cell(dr0+1,2,"Fortnight").font=BOLD; ws.cell(dr0+1,3,"Leaf ₹/kg").font=BOLD
for i,g in enumerate(green):
    ws.cell(dr0+2+i,2,g["period_from"]).font=BODY
    ws.cell(dr0+2+i,3,g["rate"]).font=BLUE
gln=len(green)
# auction series
ac0=dr0
ws.cell(ac0+1,5,"Sale").font=BOLD; ws.cell(ac0+1,6,"Avg ₹").font=BOLD; ws.cell(ac0+1,7,"Unsold %").font=BOLD
for i,a in enumerate(advices):
    ws.cell(ac0+2+i,5,f"S{a['sale_no']}").font=BODY
    ws.cell(ac0+2+i,6,a["avg"]).font=BLUE
    uns = (1-(a["sold"] or 0)/a["offered"]) if a["offered"] else 0
    ws.cell(ac0+2+i,7,round(uns,3)).font=BLUE
    ws.cell(ac0+2+i,7).number_format="0%"
an=len(advices)

# Chart 1: leaf rate over time
c1=LineChart(); c1.title="Green Leaf Rate (₹/kg) over the season"; c1.height=7; c1.width=15
c1.y_axis.title="₹/kg"; c1.legend=None
data=Reference(ws,min_col=3,min_row=dr0+1,max_row=dr0+1+gln)
cats=Reference(ws,min_col=2,min_row=dr0+2,max_row=dr0+1+gln)
c1.add_data(data,titles_from_data=True); c1.set_categories(cats)
c1.series[0].graphicalProperties.line.solidFill="2E6B4F"; c1.series[0].graphicalProperties.line.width=20000
ws.add_chart(c1,"B18" if False else "K5")

# Chart 2: auction realization
c2=LineChart(); c2.title="Auction Realization (₹/kg) by sale"; c2.height=7; c2.width=15
c2.y_axis.title="₹/kg"; c2.legend=None
d2=Reference(ws,min_col=6,min_row=ac0+1,max_row=ac0+1+an)
cat2=Reference(ws,min_col=5,min_row=ac0+2,max_row=ac0+1+an)
c2.add_data(d2,titles_from_data=True); c2.set_categories(cat2)
c2.series[0].graphicalProperties.line.solidFill="B5651D"; c2.series[0].graphicalProperties.line.width=20000
ws.add_chart(c2,"K20")

# Chart 3: unsold % by sale (bar)
c3=BarChart(); c3.title="Auction Unsold % by sale"; c3.height=7; c3.width=15
c3.y_axis.title="unsold %"; c3.legend=None
d3=Reference(ws,min_col=7,min_row=ac0+1,max_row=ac0+1+an)
c3.add_data(d3,titles_from_data=True); c3.set_categories(cat2)
c3.series[0].graphicalProperties.solidFill="A33B3B"
ws.add_chart(c3,"K35")

# =========================================================
# GREEN LEAF
# =========================================================
ws = wb.create_sheet("Green Leaf"); ws.sheet_view.showGridLines=False
setw(ws,{"A":2,"B":14,"C":14,"D":14,"E":12,"F":16,"G":18,"H":20})
ws["B2"]="GREEN LEAF — fortnightly input cost"; ws["B2"].font=H1
hdr=["Period From","Period To","GL (kg)","Rate ₹/kg","Amount ₹","Made-tea equiv (kg)","Made-tea cost ₹/kg"]
for i,h in enumerate(hdr): ws.cell(4,2+i,h)
style_header_row(ws,4,2,8)
r0=5
for i,g in enumerate(green):
    r=r0+i
    ws.cell(r,2,g["period_from"]).font=BODY
    ws.cell(r,3,g["period_to"]).font=BODY
    ws.cell(r,4,g["kg"]).font=BLUE; ws.cell(r,4).number_format="#,##0"
    ws.cell(r,5,g["rate"]).font=BLUE; ws.cell(r,5).number_format="0.00"
    ws.cell(r,6,g["amount"]).font=BLUE; ws.cell(r,6).number_format="#,##0"
    ws.cell(r,7,f"=D{r}*{RECOVERY}").font=BODY; ws.cell(r,7).number_format="#,##0"
    ws.cell(r,8,f"=E{r}*{MULT}+{conv_used:.4f}").font=BODY; ws.cell(r,8).number_format="#,##0.0"
    for c in range(2,9): ws.cell(r,c).border=BORDER
rtot=r0+len(green)
ws.cell(rtot,2,"TOTAL").font=BOLD
ws.cell(rtot,4,f"=SUM(D{r0}:D{rtot-1})").font=BOLD; ws.cell(rtot,4).number_format="#,##0"
ws.cell(rtot,6,f"=SUM(F{r0}:F{rtot-1})").font=BOLD; ws.cell(rtot,6).number_format="#,##0"
ws.cell(rtot,5,f"=F{rtot}/D{rtot}").font=BOLD; ws.cell(rtot,5).number_format="0.00"  # wavg rate
ws.cell(rtot,7,f"=SUM(G{r0}:G{rtot-1})").font=BOLD; ws.cell(rtot,7).number_format="#,##0"
for c in range(2,9): ws.cell(rtot,c).fill=BANDFILL

# =========================================================
# AUCTION SALES
# =========================================================
ws = wb.create_sheet("Auction Sales"); ws.sheet_view.showGridLines=False
setw(ws,{"A":2,"B":8,"C":14,"D":14,"E":14,"F":12,"G":12})
ws["B2"]="AUCTION SALES — weekly (Parcon, Guwahati)"; ws["B2"].font=H1
hdr=["Sale","Date","Offered (kg)","Sold (kg)","Avg ₹/kg","Unsold %"]
for i,h in enumerate(hdr): ws.cell(4,2+i,h)
style_header_row(ws,4,2,7)
r0=5
for i,a in enumerate(advices):
    r=r0+i
    ws.cell(r,2,a["sale_no"]).font=BLUE
    ws.cell(r,3,a["date"]).font=BODY
    ws.cell(r,4,a["offered"]).font=BLUE; ws.cell(r,4).number_format="#,##0"
    ws.cell(r,5,a["sold"]).font=BLUE; ws.cell(r,5).number_format="#,##0"
    ws.cell(r,6,a["avg"]).font=BLUE; ws.cell(r,6).number_format="0.00"
    ws.cell(r,7,f"=IF(D{r}=0,0,1-E{r}/D{r})").font=BODY; ws.cell(r,7).number_format="0%"
    for c in range(2,8): ws.cell(r,c).border=BORDER
rtot=r0+len(advices)
ws.cell(rtot,2,"TOTAL").font=BOLD
ws.cell(rtot,4,f"=SUM(D{r0}:D{rtot-1})").font=BOLD; ws.cell(rtot,4).number_format="#,##0"
ws.cell(rtot,5,f"=SUM(E{r0}:E{rtot-1})").font=BOLD; ws.cell(rtot,5).number_format="#,##0"
ws.cell(rtot,6,f"=SUMPRODUCT(E{r0}:E{rtot-1},F{r0}:F{rtot-1})/E{rtot}").font=BOLD; ws.cell(rtot,6).number_format="0.00"
ws.cell(rtot,7,f"=IF(D{rtot}=0,0,1-E{rtot}/D{rtot})").font=BOLD; ws.cell(rtot,7).number_format="0%"
for c in range(2,8): ws.cell(rtot,c).fill=BANDFILL
# clearance note
ws.cell(rtot+2,2,f"True invoice clearance (deduped across re-offers): {ever}/{uniq} unique invoices ever sold = {clearance:.0%}.").font=MUTE
ws.cell(rtot+3,2,"'Offered' double-counts re-catalogued lots; dedupe is the honest clearance number.").font=MUTE

# =========================================================
# PRODUCTION
# =========================================================
ws = wb.create_sheet("Production"); ws.sheet_view.showGridLines=False
setw(ws,{"A":2,"B":12,"C":14,"D":14,"E":12,"F":12,"G":10,"H":10,"I":9,"J":9,"K":9,"L":9,"M":9,"N":9})
ws["B2"]="PRODUCTION — daily manufacturing"; ws["B2"].font=H1
ws["B3"]="Grows as daily MFG reports are dropped in."; ws["B3"].font=MUTE
hdr=["Date","GL consumed","GL received","Tea made","Recovery %","Fine %","Labour ₹/kg","Coal kg/kg"]
gcols=["BOP","SM","BP","PF","PD","DUST","SEC"]
for i,h in enumerate(hdr+gcols): ws.cell(5,2+i,h)
style_header_row(ws,5,2,2+len(hdr)+len(gcols)-1)
r0=6
for i,m in enumerate(mfg):
    r=r0+i
    ws.cell(r,2,m["date"]).font=BODY
    ws.cell(r,3,m.get("gl_consumed_kg")).font=BLUE; ws.cell(r,3).number_format="#,##0"
    ws.cell(r,4,m.get("gl_received_today_kg")).font=BLUE; ws.cell(r,4).number_format="#,##0"
    ws.cell(r,5,m.get("tea_made_today_kg")).font=BLUE; ws.cell(r,5).number_format="#,##0"
    ws.cell(r,6,(m.get("recovery_today_pct") or 0)/100).font=BLUE; ws.cell(r,6).number_format="0.0%"
    ws.cell(r,7,m.get("avg_gl_fine_pct")).font=BLUE; ws.cell(r,7).number_format="0.0%"
    ws.cell(r,8,m.get("labour_cost_perkg_today")).font=BLUE; ws.cell(r,8).number_format="0.00"
    ws.cell(r,9,m.get("coal_perkg_today")).font=BLUE; ws.cell(r,9).number_format="0.00"
    for j,gcol in enumerate(gcols):
        ws.cell(r,10+j,(m.get("sorting_grades_kg") or {}).get(gcol)).font=BODY
        ws.cell(r,10+j).number_format="#,##0"
ws.cell(r0+len(mfg)+1,2,f"To date: {tea_todate:,.0f} kg made from {gl_todate:,.0f} kg leaf  ·  recovery {mfg[-1]['recovery_todate_pct']:.2f}%" if mfg else "no reports yet").font=BOLD

# =========================================================
# EXPENSES  (Tally breakdown + fixed/variable)
# =========================================================
ws = wb.create_sheet("Expenses"); ws.sheet_view.showGridLines=False
setw(ws,{"A":2,"B":42,"C":16,"D":10,"E":12,"F":12})
ws["B2"]="EXPENSES — Tally indirect, with fixed/variable split"; ws["B2"].font=H1
ws["B3"]="Edit the 'Class' in tally_reports/*.csv (variable / fixed / semi) to tune the marginal floor."; ws["B3"].font=MUTE
if tally_loaded:
    hdr=["Expense head","Amount ₹","% total","Class","₹/kg made"]
    for i,h in enumerate(hdr): ws.cell(5,2+i,h)
    style_header_row(ws,5,2,6)
    r0=6
    for i,t in enumerate(sorted(tally,key=lambda x:-x["amount"])):
        r=r0+i
        ws.cell(r,2,t["head"]).font=BODY
        ws.cell(r,3,t["amount"]).font=BLUE; ws.cell(r,3).number_format="#,##0"
        ws.cell(r,4,f"=C{r}/C{r0+len(tally)}").number_format="0.0%"
        cl=ws.cell(r,5,t["cls"]); cl.font=BODY; cl.alignment=Alignment(horizontal="center")
        if t["cls"]=="variable": cl.fill=PatternFill("solid",fgColor="FCE8D5")
        elif t["cls"]=="semi":   cl.fill=PatternFill("solid",fgColor="FFF4CC")
        ws.cell(r,6,f"=C{r}/{tea_todate}").number_format="0.00"
        for c in range(2,7): ws.cell(r,c).border=BORDER
    rt=r0+len(tally)
    ws.cell(rt,2,"TOTAL").font=BOLD
    ws.cell(rt,3,f"=SUM(C{r0}:C{rt-1})").font=BOLD; ws.cell(rt,3).number_format="#,##0"
    ws.cell(rt,6,f"=C{rt}/{tea_todate}").font=BOLD; ws.cell(rt,6).number_format="0.00"
    for c in range(2,7): ws.cell(rt,c).fill=BANDFILL
    # class subtotals
    sr=rt+2
    ws.cell(sr,2,"SPLIT").font=H2
    for c in (2,3,4): ws.cell(sr,c).fill=HEADFILL
    for i,(lab,amt) in enumerate([("Variable",tally_var),("Semi (split 50/50)",tally_semi),("Fixed",tally_fix)]):
        ws.cell(sr+1+i,2,lab).font=BODY
        ws.cell(sr+1+i,3,amt).font=BLUE; ws.cell(sr+1+i,3).number_format="#,##0"
        ws.cell(sr+1+i,6,f"=C{sr+1+i}/{tea_todate}").number_format="0.00"
    ws.cell(sr+5,2,"Marginal conversion ₹/kg (variable + ½ semi)").font=BOLD
    ws.cell(sr+5,6,f"=({tally_var}+0.5*{tally_semi})/{tea_todate}").font=BOLD; ws.cell(sr+5,6).number_format="0.00"
else:
    ws.cell(5,2,"No Tally file loaded — drop a tally_reports/*.csv to populate.").font=MUTE

# =========================================================
# ECONOMICS
# =========================================================
ws = wb.create_sheet("Economics"); ws.sheet_view.showGridLines=False
setw(ws,{"A":2,"B":42,"C":18,"D":34})
ws["B2"]="ECONOMICS — crush spread" + ("" if tally_loaded else " (provisional)"); ws["B2"].font=H1
sub = ("Conversion from Tally actuals. Sales = auction channel; all-channel sold is a manual cross-check."
       if tally_loaded else "PROVISIONAL: conversion ₹45 assumed until Tally lands.")
ws["B3"]=sub; ws["B3"].font=Font(name=F,size=9,italic=True,color=("2E6B4F" if tally_loaded else "A33B3B"))
rows=[
    ("INPUT — green leaf", None, None),
    ("Green leaf taken (kg)", gl_kg, "#,##0", "blue"),
    ("Green spend (₹)", gl_spend, "#,##0", "blue"),
    ("Weighted avg leaf rate (₹/kg)", None, "0.00", "f_rate"),
    ("CONVERSION — from Tally", None, None),
    ("Recovery (assumption)", RECOVERY, "0%", "yellow"),
    ("Total plant expenses ₹ (Tally)", tally_total, "#,##0", "blue"),
    ("Variable + ½·semi ₹ (marginal basis)", marg_expense, "#,##0", "blue"),
    ("Conversion ₹/kg (actual, all-in)", None, "0.0", "f_conv"),
    ("Marginal conversion ₹/kg (variable)", None, "0.0", "f_marg"),
    ("OUTPUT — production", None, None),
    ("Tea made to date (kg)", tea_todate, "#,##0", "blue"),
    ("Made-tea cost ₹/kg  (leaf + conv)", None, "0.0", "f_cost"),
    ("Total production cost (₹)", None, "#,##0", "f_prodcost"),
    ("SALES — auction channel", None, None),
    ("Auction sold (kg)", auc_sold_kg, "#,##0", "blue"),
    ("Auction realization ₹/kg", auc_real, "0.00", "blue"),
    ("Auction revenue (₹)", auc_sold_val, "#,##0", "blue"),
    ("Total sold — ALL channels (ledger, manual)", SEASON_SOLD_KG, "#,##0", "yellow"),
    ("RESULT", None, None),
    ("Full-cost spread ₹/kg (realization − cost)", None, "0.0", "f_spread"),
    ("Marginal floor ₹/kg (leaf + variable)", None, "0.0", "f_floor"),
    ("Headroom above floor ₹/kg", None, "0.0", "f_head"),
    ("Implied closing stock (kg)  made − all sold", None, "#,##0", "f_stock"),
]
r=5; rowmap={}
for label,val,fmt,*tag in rows:
    tg=tag[0] if tag else None
    if val is None and tg is None:
        ws.cell(r,2,label).font=H2
        for c in (2,3,4): ws.cell(r,c).fill=HEADFILL
        r+=1; continue
    ws.cell(r,2,label).font=BODY
    rowmap[label]=(r,val,fmt,tg); r+=1
def cell(label): return rowmap[label][0]
for label,(rr,val,fmt,tg) in rowmap.items():
    c=ws.cell(rr,3)
    if tg=="blue": c.value=val; c.font=BLUE
    elif tg=="yellow": c.value=val; c.font=Font(name=F,size=10,color="0000FF"); c.fill=WARNFILL
    if fmt: c.number_format=fmt
glkg_r=cell("Green leaf taken (kg)"); spend_r=cell("Green spend (₹)")
rate_r=cell("Weighted avg leaf rate (₹/kg)")
texp_r=cell("Total plant expenses ₹ (Tally)")
marg_r=cell("Variable + ½·semi ₹ (marginal basis)")
conv_r=cell("Conversion ₹/kg (actual, all-in)")
mconv_r=cell("Marginal conversion ₹/kg (variable)")
made_r=cell("Tea made to date (kg)")
cost_r=cell("Made-tea cost ₹/kg  (leaf + conv)")
prodcost_r=cell("Total production cost (₹)")
real_r=cell("Auction realization ₹/kg")
spread_r=cell("Full-cost spread ₹/kg (realization − cost)")
floor_r=cell("Marginal floor ₹/kg (leaf + variable)")
head_r=cell("Headroom above floor ₹/kg")
stock_r=cell("Implied closing stock (kg)  made − all sold")
allsold_r=cell("Total sold — ALL channels (ledger, manual)")
ws.cell(rate_r,3,f"=C{spend_r}/C{glkg_r}").number_format="0.00"
ws.cell(conv_r,3,f"=C{texp_r}/C{made_r}").number_format="0.0"
ws.cell(mconv_r,3,f"=C{marg_r}/C{made_r}").number_format="0.0"
ws.cell(cost_r,3,f"=C{rate_r}*{MULT}+C{conv_r}").number_format="0.0"
ws.cell(prodcost_r,3,f"=C{made_r}*C{cost_r}").number_format="#,##0"
ws.cell(spread_r,3,f"=C{real_r}-C{cost_r}").number_format="0.0"
ws.cell(floor_r,3,f"=C{rate_r}*{MULT}+C{mconv_r}").number_format="0.0"
ws.cell(head_r,3,f"=C{real_r}-C{floor_r}").number_format="0.0"
ws.cell(stock_r,3,f"=C{made_r}-C{allsold_r}").number_format="#,##0"
for rr in (rate_r,conv_r,mconv_r,cost_r,prodcost_r,spread_r,floor_r,head_r,stock_r):
    ws.cell(rr,3).font=Font(name=F,size=10,bold=True)
ws.cell(floor_r,4,"← below this realization, idle beats making: you'd lose cash on every kg").font=MUTE
ws.cell(spread_r,4,"← full-cost margin on what you sell").font=MUTE

# order tabs
wb.move_sheet("Guide", -wb.sheetnames.index("Guide"))
# clean print setup: fit each sheet to one page wide
from openpyxl.worksheet.properties import PageSetupProperties
for s in wb.worksheets:
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1; s.page_setup.fitToHeight = 0
    s.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    s.print_options.horizontalCentered = True
out=os.path.join(ROOT,"dashboard.xlsx")
wb.save(out)
print("wrote", out)
