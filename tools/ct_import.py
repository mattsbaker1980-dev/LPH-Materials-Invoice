"""Daily C&T auto-import for the Minimum Standards tool.

Reads C&T workbook text saved from the Microsoft 365 connector (read_resource output files),
parses every tech sheet with the same rules as the tool's importer, checks each tech's daily
numbers against the sheet's monthly totals, and replaces that department+month in
data/standards/data.json. Writes nothing if no tech data was found.

usage: python3 -I tools/ct_import.py --file-name "Plbg CT Mgr - Nov 2026.xlsx" --data data/standards/data.json SAVED_OUTPUT [SAVED_OUTPUT ...]
"""
import argparse, json, re, sys, os, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ct_parse import load_text, sheets, parse_tech
MON = {m: i for i, m in enumerate(['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'], 1)}
ap = argparse.ArgumentParser(); ap.add_argument('--file-name', required=True); ap.add_argument('--data', required=True); ap.add_argument('--dry-run', action='store_true'); ap.add_argument('paths', nargs='+')
a = ap.parse_args()
fn = a.file_name
dept = 'S' if re.search(r'sewer', fn, re.I) else 'P' if re.search(r'plbg|plumb', fn, re.I) else None
m = re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+(\d{4})', fn, re.I)
if not dept or not m: sys.exit(json.dumps({'ok': False, 'error': 'cannot tell department/month from file name: ' + fn}))
year, month = int(m.group(2)), MON[m.group(1).lower()[:3]]
ym = f'{year}-{month:02d}'
seen = {}; workbook_sheets = []
for p in a.paths:
    text = load_text(p)
    if not workbook_sheets:
        workbook_sheets = re.findall(r'^\d+\. (.+)$', text.split('This read covers')[0], re.M)
    for name, rows in sheets(text): seen[name] = rows
records, warnings, techs = [], [], []
for name, rows in seen.items():
    w, r = parse_tech(name, rows, year, month)
    if w is None: continue
    warnings += w
    if r: records += r; techs.append(name)
visible = [s for s in workbook_sheets if '[hidden]' not in s]
missing = [s for s in visible if s not in seen and s not in ('Month SPA by Week',)]
summ = lambda k: round(sum(r.get(k, 0) for r in records), 2)
res = {'ok': bool(records) and not missing, 'dept': dept, 'month': ym, 'techs': sorted(techs), 'day_rows': len(records),
       'demand_calls': summ('dc'), 'demand_money': summ('dm'), 'spp_calls': summ('sc'), 'spp_money': summ('sm'),
       'missing_sheets': missing, 'total_warnings': warnings, 'changed': False}
if not records: res['error'] = 'no tech data found'
if missing: res['error'] = 'these visible sheets were not read: ' + ', '.join(missing)
if res['ok']:
    d = json.load(open(a.data))
    old = sorted([{k: v for k, v in r.items() if k != 'p'} for r in d['records'] if r['p'] == dept and r['d'][:7] == ym], key=lambda r: (r['t'], r['d']))
    new = sorted(records, key=lambda r: (r['t'], r['d']))
    if old != new:
        res['changed'] = True
        if not a.dry_run:
            d['records'] = [r for r in d['records'] if not (r['p'] == dept and r['d'][:7] == ym)] + [dict(p=dept, **r) for r in new]
            dc, dm, sc, sm = res['demand_calls'], res['demand_money'], res['spp_calls'], res['spp_money']
            msg = (f'Auto-import {fn}: {len(techs)} techs · {dc:g} demand calls' + (f' ({dm/dc*100:.1f}%)' if dc else '') +
                   (f' · {sc:g} system checks ({sm/sc*100:.1f}%)' if sc else '') + (f' — {len(warnings)} total(s) did not match' if warnings else ''))
            d.setdefault('imports', []).insert(0, {'at': dt.datetime.now(dt.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'), 'msg': msg})
            d['imports'] = d['imports'][:50]
            json.dump(d, open(a.data, 'w'), separators=(',', ':'))
print(json.dumps(res))
