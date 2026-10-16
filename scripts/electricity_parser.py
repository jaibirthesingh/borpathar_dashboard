"""APDCL electricity bill parser (digital PDF). Returns period, gross units
consumed (kWh, for efficiency), bill amount, and demand/power-factor extras."""
import pdfplumber, re
from datetime import date
_MON={'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
def _d(s):
    m=re.match(r'(\d{1,2})-([A-Za-z]+)-(\d{4})',s or '')
    return date(int(m.group(3)),_MON[m.group(2)[:3].lower()],int(m.group(1))).isoformat() if m else None
def _num(s):
    try: return float(str(s).replace(',',''))
    except: return None
def parse_bill(path):
    with pdfplumber.open(path) as pdf:
        t="\n".join(p.extract_text() or "" for p in pdf.pages)
    if not t.strip(): return None
    d={}
    m=re.search(r'Bill Period:\s*([\d]{1,2}-[A-Za-z]+-\d{4})\s*To\s*([\d]{1,2}-[A-Za-z]+-\d{4})',t)
    if m: d['start']=_d(m.group(1)); d['end']=_d(m.group(2))
    m=re.search(r'Bill Amount:\s*([\d,]+\.?\d*)',t); d['bill_amount']=_num(m.group(1)) if m else None
    # gross units consumed: first number after "Solar|/Peak|/Normal|" in the consumption block
    gross=re.findall(r'^(?:Solar|Peak|Normal)\s*\|\s*([\d,]+\.\d+)',t,flags=re.M)
    d['units_kwh']=round(sum(_num(x) for x in gross),2) if gross else None
    m=re.search(r'Maximum Demand \(in KVA\)\s*([\d.]+)',t); d['max_demand_kva']=_num(m.group(1)) if m else None
    m=re.search(r'Contracted Demand in KVA:\s*([\d.]+)',t); d['contracted_kva']=_num(m.group(1)) if m else None
    m=re.search(r'Average Power Factor\s*([\d.]+)',t); d['power_factor']=_num(m.group(1)) if m else None
    d['_source']=path.split('/')[-1]
    return d
if __name__=='__main__':
    import sys,json; print(json.dumps(parse_bill(sys.argv[1]),indent=1))
