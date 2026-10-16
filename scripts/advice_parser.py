"""Parcon auction Sale Advice PDF parser."""
import pdfplumber, re
from datetime import datetime

def _date(s):
    try: return datetime.strptime(s,'%d/%m/%Y').date().isoformat()
    except: return s

def parse_advice(path):
    with pdfplumber.open(path) as pdf:
        txt='\n'.join(p.extract_text() or '' for p in pdf.pages)
    sn=re.search(r'Sale No (\d+)',txt)
    dt=re.search(r'Sale Date: (\d{2}/\d{2}/\d{4})',txt)
    gt=re.search(r'Grand Total Offered Qty : ([\d.]+) Sold Qty : ([\d.]+) Sold at Average of Rs : ([\d.]+)',txt)
    lots=[]
    for line in txt.splitlines():
        m=re.match(r'(PAR\d+)\s+(\S+)\s+([A-Z0-9]+)\s+(\d+)\s+([\d.]+)\s+(UNSOLD|[\d.]+)',line)
        if m:
            lot,inv,grade,pk,wt,price=m.groups()
            lots.append(dict(lot=lot,inv=inv,grade=grade,pkgs=int(pk),net_wt=float(wt),
                             price=(None if price=='UNSOLD' else float(price)),sold=(price!='UNSOLD')))
    return dict(sale_no=int(sn.group(1)) if sn else None,
                date=_date(dt.group(1)) if dt else None,
                offered=float(gt.group(1)) if gt else None,
                sold=float(gt.group(2)) if gt else None,
                avg=float(gt.group(3)) if gt else None,
                lots=lots, _source=path.split('/')[-1])
