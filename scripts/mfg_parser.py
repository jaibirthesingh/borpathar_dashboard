"""Borpathar daily MFG-report parser. Label-anchored; template is fixed day-to-day."""
import xlrd
from datetime import datetime

def _num(v):
    try: return float(v)
    except: return None

def _date(v):
    if not v: return None
    for fmt in ('%d/%m/%Y','%d-%m-%Y','%d.%m.%Y','%d/%m/%y'):
        try: return datetime.strptime(str(v).strip(), fmt).date().isoformat()
        except: pass
    return str(v)

def parse_mfg(path):
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_name('MFG')
    g = [[sh.cell_value(r,c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    def rc(label, start=0):
        for r in range(start, sh.nrows):
            for c in range(sh.ncols):
                v=g[r][c]
                if isinstance(v,str) and label.lower() in v.lower().strip():
                    return r,c
        return None
    def right(label, start=0):
        hit=rc(label,start)
        if not hit: return None
        r,c=hit
        for cc in range(c+1,sh.ncols):
            if g[r][cc] not in ('',None): return g[r][cc]
        return None
    d={}
    d['date']=_date(right('Date'))
    d['gl_consumed_kg']=_num(right('G/L CONSUMED'))
    d['tea_made_today_kg']=_num(right('TEA MADE TODAY'))
    d['recovery_today_pct']=_num(right('RECOVERY % TODAY'))
    d['recovery_todate_pct']=_num(right('RECOVERY % TO DATE'))
    d['tea_made_todate_kg']=_num(right('TEA MADE TODATE'))
    d['working_days']=_num(right('WORKING DAYS'))
    d['gl_todate_kg']=_num(right('TO DATE G/L'))
    d['labour_cost_perkg_today']=_num(right('LABOUR COST'))
    d['coal_perkg_today']=_num(right('Coal Use'))
    # G/L received TODAY — scope search to AFTER the 'G/L REPORT' header
    glr=rc('G/L REPORT')
    if glr:
        d['gl_received_today_kg']=_num(right('TODAY', start=glr[0]))
    # fine %
    fh=rc('Avg G/L Fine')
    if fh:
        r,c=fh
        for rr in range(r,min(r+4,sh.nrows)):
            v=g[rr][c]
            if isinstance(v,(int,float)) and 0<v<1: d['avg_gl_fine_pct']=float(v); break
    # sorting grades
    hs=rc('BPS')
    grades={}
    if hs:
        r,c=hs; hdr=g[r]
        for rr in range(r+1,min(r+4,sh.nrows)):
            row=g[rr]
            if any(isinstance(x,(int,float)) and x>1 for x in row):
                for cc,h in enumerate(hdr):
                    if isinstance(h,str) and h.strip() in ('BOP','SM','BP','PF','PD','DUST','SEC','TOTAL'):
                        grades[h.strip()]=_num(row[cc])
                break
    d['sorting_grades_kg']=grades
    d['_source']=path.split('/')[-1]
    return d
