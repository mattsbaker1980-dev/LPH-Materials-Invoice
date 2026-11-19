"""Parse one C&T tech sheet read through the Microsoft 365 connector (tab-separated text) into day records.
Same rules as the tool's in-browser importer."""
import json, re, sys, datetime as dt, calendar
LAB={'DEMAND CALLS':'dc','DEMAND MONEY CALLS':'dm','DEMAND REVENUE':'dr','SPP CALLS':'sc','SPP MONEY CALLS':'sm','SPP REVENUE':'sr'}
def col_index(ref):
    n=0
    for ch in ref: n=n*26+ord(ch)-64
    return n-1
def load_text(path):
    raw=open(path).read()
    try:
        j=json.loads(raw); raw=''.join(x.get('text','') for x in j) if isinstance(j,list) else raw
    except Exception: pass
    return raw
def sheets(text):
    out=[]
    for m in re.finditer(r'## Sheet: (.+?) — \d+ rows × \d+ columns \(([A-Z]+)(\d+):[A-Z]+\d+\)\n(.*?)(?=\n## Sheet: |\Z)', text, re.S):
        name, c0, r0, body = m.group(1).strip(), col_index(m.group(2)), int(m.group(3)), m.group(4)
        body = body.split('\nFormulas')[0]
        rows=[]
        for line in body.split('\n'):
            e=re.fullmatch(r'\[(\d+) empty rows?\]', line.strip())
            if e: rows += [[] for _ in range(int(e.group(1)))]; continue
            rows.append(['']*c0 + line.split('\t'))
        out.append((name.replace(' [hidden]',''), [[]]*(r0-1) + rows))
    return out
def num(v):
    try: return float(str(v).replace(',','').replace('$',''))
    except: return None
def todate(v):
    s=str(v).strip()
    if re.fullmatch(r'\d{5}(\.0+)?', s): return dt.date(1899,12,30)+dt.timedelta(days=int(float(s)))
    m=re.match(r'(\d+)/(\d+)/(\d+)', s)
    if m:
        y=int(m.group(3)); y+=2000 if y<100 else 0
        return dt.date(y,int(m.group(1)),int(m.group(2)))
    return None
def prange(v):
    nums=re.findall(r'\d+',str(v or ''))
    if len(nums)>=6:
        a=[int(x) for x in nums[:6]]; y1=a[2]+2000 if a[2]<100 else a[2]; y2=a[5]+2000 if a[5]<100 else a[5]
        return dt.date(y1,a[0],a[1]), dt.date(y2,a[3],a[4])
    d=todate(v); return (d,d) if d else None
def parse_tech(name, rows, year, month):
    cell=lambda r,c: rows[r][c] if r<len(rows) and c<len(rows[r]) else ''
    if 'Individual Performance' not in ' '.join(rows[1] if len(rows)>1 else []): return None, []
    hdr=rows[0]
    starts=[i for i,v in enumerate(hdr) if re.match(r'Week \d',str(v))]
    c0=next((i for i,v in enumerate(hdr) if v=='Monthly Summary'),None)
    tc=next((i for i,v in enumerate(rows[3]) if v=='MONTHLY'),None) if len(rows)>3 else None
    mstart=dt.date(year,month,1); mend=dt.date(year,month,calendar.monthrange(year,month)[1])
    totals={}
    if c0 is not None and tc is not None:
        for r in range(min(45,len(rows))):
            lab=str(cell(r,c0)).strip()
            if lab in LAB: totals[LAB[lab]]=num(cell(r,tc)) or 0
    days={}; prev=mstart-dt.timedelta(days=1); warn=[]
    for off in starts:
        rg=prange(cell(2,off+1)) or (prev+dt.timedelta(days=1), mend)
        s=max(rg[0],mstart); e=min(max(rg[1],s),mend); prev=e
        sun=s-dt.timedelta(days=(s.weekday()+1)%7)
        for r in range(min(45,len(rows))):
            k=LAB.get(str(cell(r,off)).strip())
            if not k: continue
            tot=0
            for i in range(7):
                v=num(cell(r,off+1+i))
                if v:
                    d=min(max(sun+dt.timedelta(days=i),s),e); days.setdefault(d,{}); days[d][k]=days[d].get(k,0)+v; tot+=v
            wt=num(cell(r,off+8))
            if wt is not None and abs(wt-tot)>0.005: days.setdefault(e,{}); days[e][k]=days[e].get(k,0)+(wt-tot)
    sums={k:sum(x.get(k,0) for x in days.values()) for k in LAB.values()}
    for k in LAB.values():
        if abs(sums[k]-totals.get(k,0))>0.01: warn.append(f'{name}: {k} days add to {round(sums[k],2)} but monthly total is {totals.get(k)}')
    recs=[]
    for d,v in sorted(days.items()):
        v={k:round(x,2) for k,x in v.items() if abs(x)>0.001}
        if v: recs.append(dict(t=name,d=d.isoformat(),**v))
    return warn, recs
if __name__=='__main__':
    text=load_text(sys.argv[1]); y,m=int(sys.argv[2]),int(sys.argv[3])
    allr=[]; W=[]
    for name,rows in sheets(text):
        w,r=parse_tech(name,rows,y,m)
        if w is None: continue
        W+=w; allr+=r
    json.dump({'records':allr,'warnings':W},sys.stdout)
