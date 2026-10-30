"""Parcon SOLD PRICE report parser (richer successor to the Sale Advice).

Per-lot line layout:
  LotNo  Invoice  Mark...  Grade  DealPkgs  DealQty  HBP  F/RP  Status  Buyer...
e.g.  PAR3532 D328 CLASSIC BORPATHAR D1 15 507.60 149 F Sold GIRNAR FOOD & BEVERAGES
Mark and Buyer are multi-word; the grade is the token right before the first number.

Output mirrors the advice interface (so the pipeline can use it unchanged) and adds buyer:
  dict(sale_no, date(None -> mapped externally), sold, offered(None; unsold weight absent),
       avg, lots=[dict(lot, inv, mark, grade, pkgs, net_wt, price, hbp, sold, status, f_rp, buyer)])
"""
import re
import pdfplumber

_SALE = re.compile(r'Sale\s*No\.?\s*(\d+)', re.I)
_AVG = re.compile(r'Avg\.?\s*Price\s*:\s*([\d.]+)', re.I)
# lot ... grade pkgs qty hbp f/rp status buyer   (grade = last token before the first int)
_LOT = re.compile(
    r'^(PAR\d+)\s+(\S+)\s+(.+?)\s+([A-Z0-9]+)\s+(\d+)\s+([\d.]+)\s+(\d+)\s+(RP-\d+|F)\s+(Sold|Unsold)\s*(.*)$',
    re.I)


def parse_sold_price(path):
    with pdfplumber.open(path) as pdf:
        txt = "\n".join(p.extract_text() or "" for p in pdf.pages)
    sn = _SALE.search(txt)
    av = _AVG.search(txt)
    lots = []
    for line in txt.splitlines():
        m = _LOT.match(line.strip())
        if not m:
            continue
        lot, inv, mark, grade, pkgs, qty, hbp, frp, status, buyer = m.groups()
        sold = status.strip().lower() == "sold"
        qtyf = float(qty)
        hbpf = float(hbp)
        lots.append(dict(
            lot=lot, inv=inv, mark=mark.strip(), grade=grade.strip().upper(),
            pkgs=int(pkgs), net_wt=qtyf,                 # deal (sold) qty; 0 for unsold
            price=(hbpf if sold else None),              # sold price (advice-compatible)
            hbp=hbpf,                                    # top bid, present even when unsold
            sold=sold, status=status.strip(),
            f_rp=frp.strip(), buyer=(buyer.strip() or None)))
    sold_kg = sum(l["net_wt"] for l in lots if l["sold"])
    return dict(sale_no=int(sn.group(1)) if sn else None,
                date=None,                               # not in file; map sale_no -> date outside
                offered=None,                            # unsold lots show 0 qty -> offered unknown
                sold=sold_kg,
                avg=float(av.group(1)) if av else None,
                lots=lots, _source=path.split('/')[-1])
