"""Tally monthly plant-ledger reader.
Current input: a transcribed CSV (head,amount,class) because Tally was supplied
as a scanned image. For auto-parsing, export Tally to a DIGITAL pdf (text layer)
and this can be upgraded to read it directly. class in {variable,fixed,semi}."""
import csv, glob, os

def parse_tally(path):
    rows=[]
    if path.lower().endswith('.csv'):
        with open(path, newline='') as f:
            for r in csv.DictReader(f):
                try: amt=float(r['amount'])
                except: continue
                rows.append(dict(head=r['head'].strip(), amount=amt,
                                 cls=(r.get('class') or 'fixed').strip().lower()))
    return rows
