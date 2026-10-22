"""Leaf-quality weekly reader — SEPARATE from the green-leaf price feed.
Drop format = the GL_KG / AV.RATE export used for the fine%/premium analysis:
one row per week with a combined period cell and a fine% column.

Columns (header-matched): PERIOD FROM ('DD.MM.YY TO DD.MM.YY'), KG, RATE,
AMOUNT, GL AV.RATE (= weighted fine %, e.g. '26%').

Reads .pdf (pdfplumber), .xlsx/.xls (openpyxl), or .csv. Output:
period_from, period_to (ISO yyyy-mm-dd), gl_kg, rate, amount,
fine_pct (percent number, e.g. 26.0 or None). TOTAL / zero-kg weeks skipped.
This feed does NOT touch the green-leaf price feed or its GreenLeaf tab.
"""
import csv, re

_PERIOD = re.compile(r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s+TO\s+(\d{1,2})\.(\d{1,2})\.(\d{2,4})', re.I)


def _iso(dd, mm, yy):
    yy = int(yy); yy = yy + 2000 if yy < 100 else yy
    return f"{yy:04d}-{int(mm):02d}-{int(dd):02d}"


def _num(v):
    if v is None: return None
    try: return float(str(v).replace(',', '').strip())
    except Exception: return None


def _fine(v):
    """'26%' -> 26.0 ; 0.26 -> 26.0 ; '0' / 0 / '' -> None."""
    if v is None: return None
    s = str(v).replace('%', '').strip()
    if s == '' or s == '0': return None
    try: f = float(s)
    except Exception: return None
    if f == 0: return None
    return round(f * 100, 1) if f < 1 else round(f, 1)


def _row_from_cells(cells):
    if not cells: return None
    m = _PERIOD.search(str(cells[0]))
    if not m: return None
    kg = _num(cells[1]) if len(cells) > 1 else None
    if not kg: return None
    g = m.groups()
    return dict(period_from=_iso(g[0], g[1], g[2]), period_to=_iso(g[3], g[4], g[5]),
               gl_kg=kg, rate=_num(cells[2]) if len(cells) > 2 else None,
               amount=_num(cells[3]) if len(cells) > 3 else None,
               fine_pct=_fine(cells[4]) if len(cells) > 4 else None)


def parse_leaf_quality(path):
    ext = path.lower()
    rows = []
    if ext.endswith('.pdf'):
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            for pg in pdf.pages:
                for tb in (pg.extract_tables() or []):
                    for cells in tb:
                        r = _row_from_cells(cells)
                        if r: rows.append(r)
                if not rows:
                    for line in (pg.extract_text() or '').splitlines():
                        m = _PERIOD.search(line)
                        if not m: continue
                        r = _row_from_cells([line[:m.end()]] + line[m.end():].split())
                        if r: rows.append(r)
    elif ext.endswith('.csv'):
        with open(path, newline='') as f:
            for raw in csv.reader(f):
                r = _row_from_cells(raw)
                if r: rows.append(r)
    else:
        import openpyxl
        ws = openpyxl.load_workbook(path, data_only=True).active
        for row in ws.iter_rows(values_only=True):
            r = _row_from_cells(list(row))
            if r: rows.append(r)
    seen = {}
    for r in rows:
        seen[r['period_from']] = r
    return sorted(seen.values(), key=lambda r: r['period_from'])
