"""Parcon GRADE GAP (Range Grade Report) reader. One .xlsx per sale.

Layout: a FILTER block (SALE RANGE 'N - N', EST/BLF, CENTER, AREA), then a
SEASON header row (2026 2025 2024 ...), then one row per grade holding the
avg price for each season, interleaved with 'GRADE - GRADE' difference rows
(skipped). We take the current-season (max year) column.

Output: dict(sale_no, est_blf, center, area, year, grades={grade: price}).
Primary grades only exist in this report (no secondaries).
"""
import re
import openpyxl

GRADES = {"BOP", "BOPSM", "BP", "PF", "PD", "D"}


def parse_grade_gap(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]

    sale_no = est_blf = center = area = None
    year_col = None
    year = None
    grades = {}

    for r in rows:
        cells = [c for c in r]
        # locate the filter line: a cell 'N - N' with EST/BLF, CENTER, AREA to its right
        for i, c in enumerate(cells):
            if isinstance(c, str):
                m = re.match(r'^\s*(\d+)\s*-\s*(\d+)\s*$', c)
                if m and sale_no is None:
                    sale_no = int(m.group(1))
                    tail = [x for x in cells[i + 1:] if x not in (None, '')]
                    if len(tail) >= 1: est_blf = str(tail[0]).strip()
                    if len(tail) >= 2: center = str(tail[1]).strip()
                    if len(tail) >= 3: area = str(tail[2]).strip()
        # SEASON header -> pick the current (max) year column
        if cells and isinstance(cells[0], str) and cells[0].strip().upper() == "SEASON":
            years = [(idx, int(v)) for idx, v in enumerate(cells)
                     if isinstance(v, (int, float)) and 2000 <= v <= 2100]
            if years:
                year_col, year = max(years, key=lambda t: t[1])
        # grade rows (single grade token, numeric in the chosen year column)
        if year_col is not None and cells and isinstance(cells[0], str):
            name = cells[0].strip()
            if name in GRADES and " - " not in name:
                v = cells[year_col] if year_col < len(cells) else None
                if isinstance(v, (int, float)):
                    grades[name] = float(v)

    return dict(sale_no=sale_no, est_blf=est_blf, center=center, area=area,
                year=year, grades=grades, _source=path.split('/')[-1])
