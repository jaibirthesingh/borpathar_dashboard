"""Green-leaf fortnightly price reader. Expected drop format: CSV/XLSX with
columns period_from, period_to, kg, rate, amount."""
import csv, os

def parse_green_leaf(path):
    rows=[]
    if path.lower().endswith('.csv'):
        with open(path, newline='') as f:
            for r in csv.DictReader(f):
                rows.append(dict(period_from=r['period_from'],period_to=r['period_to'],
                    kg=float(r['kg']),rate=float(r['rate']),amount=float(r['amount'])))
    else:
        import openpyxl
        wb=openpyxl.load_workbook(path,data_only=True); ws=wb.active
        hdr=[str(c.value).strip().lower() if c.value else '' for c in ws[1]]
        idx={h:i for i,h in enumerate(hdr)}
        for row in ws.iter_rows(min_row=2,values_only=True):
            if row[idx['kg']] is None: continue
            rows.append(dict(period_from=str(row[idx['period_from']]),period_to=str(row[idx['period_to']]),
                kg=float(row[idx['kg']]),rate=float(row[idx['rate']]),amount=float(row[idx['amount']])))
    return rows
