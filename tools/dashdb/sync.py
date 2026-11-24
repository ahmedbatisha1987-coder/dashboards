"""Read the dashboard Excel database and push it into the dashboard data.
usage: python3 sync.py <db.xlsx> <outdir> <page.html> [more pages...] [--check]
 - patches the embedded data (const D=...) and the data-date line of every standalone page given
 - writes, for each dataset that changed, the file the original Claude dashboard loads (csv or json) into <outdir>
 - writes <outdir>/changed.json and lists Arabic text that has no English translation yet"""
import sys, os, json, re, csv, io, datetime
from openpyxl import load_workbook
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import load_D

STATUS_AR = {'Open': 'مفتوحة', 'Closed': 'مغلقة', 'Cancelled': 'ملغي', '': 'غير محددة'}
NUM = re.compile(r'[-+]?\d[\d,.:\/%]*'); AR = re.compile('[؀-ۿ]')

def iso(v):
    if v is None or v == '': return None
    if isinstance(v, datetime.datetime): return v.date().isoformat()
    if isinstance(v, datetime.date): return v.isoformat()
    s = str(v).strip()
    m = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$', s)
    if m: return f'{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}'
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m: return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    return s

def numtxt(v):
    if isinstance(v, bool): return str(v)
    if isinstance(v, int): return str(v)
    if isinstance(v, float): return str(int(v)) if v.is_integer() and abs(v) < 1e15 else repr(v)
    return str(v).strip()

def conv(v, t):
    if isinstance(v, str) and v.strip() == '': v = None
    if v is None: return None
    if t == 'date': return iso(v)
    if t == 'numstr': return numtxt(v)
    if t == 'int':
        try: f = float(v); return int(f) if f.is_integer() else f
        except: return v
    if t == 'num':
        if isinstance(v, (int, float)) and not isinstance(v, bool): return int(v) if isinstance(v, float) and v.is_integer() and False else v
        try: f = float(str(v).replace(',', '')); return int(f) if f.is_integer() else f
        except: return v
    if t == 'bool': return v if isinstance(v, bool) else str(v).lower() in ('true', 'yes', '1')
    if t == 'str':
        if isinstance(v, (datetime.date, datetime.datetime)): return iso(v)
        return numtxt(v) if isinstance(v, (int, float)) else str(v)
    return v

def read_schema(wb):
    ws = wb['_schema']; parts = []
    r = 1
    while ws.cell(r, 1).value: parts.append(ws.cell(r, 1).value); r += 1
    return json.loads(''.join(parts))

def data_date(wb):
    g = wb[wb.sheetnames[0]]
    for row in g.iter_rows(min_row=1, max_row=20):
        for c in row:
            if c.value == 'تاريخ البيانات': return iso(g.cell(c.row, c.column + 1).value)
    return None

def read_sheet(wb, sh, dd, old=None):
    ws = wb[sh['sheet']]; t = sh['types']; edit = set(sh['edit'])
    if sh['kv']:
        r0 = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] in edit: r0[row[0]] = conv(row[2], t.get(row[0], 'str'))
        rows = [r0] if r0 else []
    else:
        keys = [c.value for c in ws[1]]
        rows = []
        for row in ws.iter_rows(min_row=3, values_only=True):
            if all(v is None or (isinstance(v, str) and not v.strip()) for v in row): continue
            r = {}
            for k, v in zip(keys, row):
                if k in edit: r[k] = conv(v, t.get(k, 'str'))
            rows.append(r)
    # derived columns
    ddate = datetime.date.fromisoformat(dd) if dd else datetime.date.today()
    d = lambda s: datetime.date.fromisoformat(s) if s else None
    for r in rows:
        for b, m in sh['maps'].items():
            if b in sh['derived']:
                a = r.get(m['from']); key = '' if a is None else str(a)
                if b == 'status_name': r[b] = m['map'].get(key) or STATUS_AR.get(key) or key
                elif b == 'theme': r[b] = m['map'].get(key) or 'other'
                else: r[b] = m['map'].get(key) or key
        if sh['ds'] in ('reg_sor_hse', 'reg_sor_q'):
            iss, clo = d(r.get('issued')), d(r.get('closed'))
            r['days_to_close'] = (clo - iss).days if iss and clo else None
            r['days_open'] = (ddate - iss).days if iss and r.get('status') == 'Open' else None
            r['month'] = r['issued'][:7] if r.get('issued') else ''
            r['complete'] = 'yes' if r.get('issued') and r.get('desc') else 'no'
        if sh['ds'] == 'sor_critical':
            m = re.search(r'(\d+)\s*$', r.get('ref') or ''); r['n'] = int(m.group(1)) if m else None
            dt = d(r.get('date')); r['days_open'] = (ddate - dt).days if dt and r.get('status') == 'Open' else None
    cols = sh['columns']; emp = sh.get('empty', {})
    out = [{c: (emp.get(c) if r.get(c) is None and c in emp else r.get(c)) for c in cols} for r in rows]
    # keep the previous text of numbers that did not change (e.g. '0.0' stays '0.0'), so only real edits count as changes
    for k, r in enumerate(out):
        o = old[k] if old and k < len(old) else None
        if not o: continue
        for c in cols:
            a, b = o.get(c), r.get(c)
            if a == b or a is None or b is None or t.get(c) not in ('numstr', 'num', 'int'): continue
            try:
                fa, fb = float(a), float(b)
                if abs(fa - fb) <= 1e-6 * max(1, abs(fa)): r[c] = a
            except: pass
    return out

def to_csv(rows, cols):
    s = io.StringIO(); w = csv.writer(s, lineterminator='\n'); w.writerow(cols)
    for r in rows: w.writerow(['' if r.get(c) is None else r.get(c) for c in cols])
    return s.getvalue()

def i18n_dict(h):
    a = h.find('const DICT = ');
    if a < 0: return None
    a += len('const DICT = '); b = h.find('};', a) + 1
    return json.loads(h[a:b])

def main():
    args = [a for a in sys.argv[1:] if a != '--check']; check = '--check' in sys.argv
    xlsx, outdir, pages = args[0], args[1], args[2:]
    wb = load_workbook(xlsx, data_only=True)
    sc = read_schema(wb); dd = data_date(wb)
    h0, i0, j0, D0 = load_D(pages[0])
    data = {sh['ds']: (sh, read_sheet(wb, sh, dd, (D0.get(sh['ds']) or {}).get('data'))) for sh in sc['sheets']}
    changed = [ds for ds, (sh, rows) in data.items() if (D0.get(ds) or {}).get('data') != rows]
    print('project', sc['proj'], 'data date', dd)
    for ds in data:
        old = (D0.get(ds) or {}).get('data') or []; new = data[ds][1]
        if ds in changed:
            nd = sum(1 for a, b in zip(old, new) if a != b) + abs(len(old) - len(new))
            print(f'  changed {ds}: {len(old)} -> {len(new)} rows, {nd} rows differ')
    if check: return
    os.makedirs(outdir, exist_ok=True)
    for p in pages:
        h, i, j, D = load_D(p)
        for ds in changed:
            sh, rows = data[ds]
            D[ds] = {'status': 'ok', 'data': rows, 'columns': sh['columns']}
        h = h[:i] + json.dumps(D, ensure_ascii=False).replace('</script', '<\\/script') + h[j:]
        if dd:
            h = re.sub(r'البيانات بتاريخ \d{2}/\d{2}/\d{4}', 'البيانات بتاريخ ' + f'{dd[8:10]}/{dd[5:7]}/{dd[:4]}', h)
        open(p, 'w', encoding='utf-8').write(h)
        print('patched', p)
    files = {}
    for ds in changed:
        sh, rows = data[ds]
        if sh['fmt'] == 'csv':
            fn = f'{ds}.csv'; open(os.path.join(outdir, fn), 'w', encoding='utf-8').write(to_csv(rows, sh['columns']))
        elif sh['fmt'] == 'sor':
            fn = 'sor_critical.json'
            json.dump({'rows': data['sor_critical'][1], 'themes': data['sor_themes'][1]}, open(os.path.join(outdir, fn), 'w', encoding='utf-8'), ensure_ascii=False)
        else:
            fn = f'{ds}.json'; json.dump({'rows': rows}, open(os.path.join(outdir, fn), 'w', encoding='utf-8'), ensure_ascii=False)
        files[ds] = fn
    json.dump({'proj': sc['proj'], 'data_date': dd, 'changed': files}, open(os.path.join(outdir, 'changed.json'), 'w'), ensure_ascii=False, indent=1)
    DICT = i18n_dict(h0) or {}
    miss = set()
    for ds in changed:
        for r in data[ds][1]:
            for v in r.values():
                if isinstance(v, str) and AR.search(v) and len(v) < 400:
                    k = NUM.sub('#', v.strip())
                    if k not in DICT: miss.add(k)
    json.dump(sorted(miss), open(os.path.join(outdir, 'untranslated.json'), 'w'), ensure_ascii=False, indent=0)
    print('changed datasets:', files); print('Arabic text without English translation:', len(miss))

if __name__ == '__main__':
    main()
