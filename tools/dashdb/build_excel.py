"""Build the dashboard Excel database from the data embedded in a standalone dashboard page.
usage: python3 build_excel.py <proj> <page.html> <out.xlsx> <data_date YYYY-MM-DD>"""
import sys, json, re, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
sys.path.insert(0, __file__.rsplit('/', 1)[0])
from common import DS, load_D, labels
from labels_ar import L

PROJ = {'2242': ('AP418', 'NWC 511250274', 'تنفيذ خط الصرف الصحي الرئيسي الثاني (منى – العزيزية) من منطقة منى إلى محطة معالجة عرنة'),
        '2094': ('P1545', 'NWC 511250201', 'إعادة إعمار واستبدال خطوط النقل الدائرية الرئيسية لمياه الشرب بمكة المكرمة (المرحلة الثالثة)')}
DESC = {
 'project':'البيانات العامة: الأرقام والتواريخ وقيمة العقد والأطراف.',
 'weekly':'مؤشرات الأداء الأسبوعية: النسب المخططة والفعلية والكميات والمالية والتصاريح والسلامة والجودة. الأسبوع المُبلّغ عنه = yes.',
 'daily_reports':'قراءات التقارير اليومية: نسبة الإنجاز التراكمية والكميات المنفذة.',
 'components':'المكونات الرئيسية للمشروع وكمياتها وقيمتها.',
 'invoices':'المستخلصات الجارية وحالتها.',
 'issues':'القضايا والمعوقات المذكورة في التقارير.',
 'reg_sor_hse':'سجل ملاحظات السلامة (SOR). أيام الإغلاق والأيام المفتوحة والشهر تُحسب تلقائياً.',
 'sor_critical':'الملاحظات الحرجة (عالية الخطورة): المفتوحة والمغلقة المتكررة.',
 'sor_themes':'أنواع المخاطر الحرجة وعدد مرات ورودها في ملاحظات السلامة.',
 'reg_ncr':'سجل تقارير عدم المطابقة (NCR).',
 'reg_incident':'سجل الحوادث.',
 'reg_ptw':'سجل تصاريح العمل (PTW).',
 'permits':'تصاريح الحفر من الأمانة (بلدي) وحالتها.',
 'reg_sor_q':'سجل ملاحظات الجودة.',
 'reg_wir':'سجل طلبات استلام الأعمال (WIR).',
 'reg_mir':'سجل فحص المواد الموردة (MIR).',
 'reg_mar':'سجل اعتماد المواد (MAT).',
 'reg_sub':'سجل اعتماد مقاولي الباطن.',
 'reg_doc':'سجل تقديمات المستندات (خطط، طرق تنفيذ، ITP، تقارير).',
 'reg_shd':'سجل المخططات التنفيذية.',
 'supplies':'التوريدات للموقع.',
 'white_helmet':'تقارير الموقع بتصوير 360° من منصة White Helmet.',
 'works_segments':'موقف الأعمال على المسار بين كل منهولين (منتهٍ / جارٍ / لم يبدأ).',
 'team':'سجل أعضاء فريق المشروع: شركة المياه والاستشاري والمقاول، بالمنصب والبريد والجوال. الجهة: NWC أو Consultant أو Contractor.',
}
DERIVED = {'*': ['status_name'],
           'reg_sor_hse': ['days_to_close', 'days_open', 'month', 'complete'],
           'reg_sor_q': ['days_to_close', 'days_open', 'month', 'complete'],
           'sor_critical': ['n', 'days_open', 'theme'],
           'works_segments': ['kind_name']}
FORMULAS = {'weekly': {'variance_pct': 'IF(AND(ISNUMBER({actual_pct}),ISNUMBER({planned_pct})),ROUND({actual_pct}-{planned_pct},4),"")',
                       'spi': 'IF(AND(ISNUMBER({actual_pct}),ISNUMBER({planned_pct})),IF({planned_pct}<>0,{actual_pct}/{planned_pct},IF({actual_pct}=0,0,"")),"")'}}
DROPDOWN = {'party', 'party_name', 'status', 'reported', 'kind', 'theme_name', 'is_latest', 'row_status', 'complete', 'current', 'mh_known'}
DATE_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')

def is_num_str(s):
    try: float(s); return True
    except: return False

def col_type(vals, fmt):
    v = [x for x in vals if x is not None and x != '']
    if not v: return 'str'
    if all(isinstance(x, str) for x in v):
        if all(DATE_RE.match(x) for x in v): return 'date'
        if fmt == 'csv' and all(is_num_str(x) for x in v): return 'numstr'
        return 'str'
    if all(isinstance(x, bool) for x in v): return 'bool'
    if all(isinstance(x, int) and not isinstance(x, bool) for x in v): return 'int'
    if all(isinstance(x, (int, float)) and not isinstance(x, bool) for x in v): return 'num'
    return 'mixed'

def to_cell(v, t):
    if v is None or v == '': return None
    if t == 'date': return datetime.date.fromisoformat(v)
    if t == 'numstr':
        f = float(v); return int(f) if f.is_integer() and 'e' not in v.lower() and abs(f) < 1e15 else f
    return v

HEAD = PatternFill('solid', fgColor='1F3864'); HFONT = Font(name='Arial', bold=True, color='FFFFFF', size=10)
CALC = PatternFill('solid', fgColor='E7E6E6'); BODY = Font(name='Arial', size=10)
THIN = Side(style='thin', color='BFBFBF'); BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def main(proj, page, out, data_date):
    h, i, j, D = load_D(page)
    LB = labels(h)
    wb = Workbook(); g = wb.active; g.title = 'دليل الاستخدام'; g.sheet_view.rightToLeft = True
    schema = {'proj': proj, 'version': 1, 'sheets': []}
    sheets = []
    for ds, (sname, fmt) in DS.items():
        if ds not in D: continue
        cols = D[ds].get('columns') or []
        rows = D[ds].get('data') or []
        if ds == 'reg_incident' and not cols:
            cols = ['id', 'ref', 'date', 'location', 'type', 'desc', 'status', 'status_name']
        types = {c: col_type([r.get(c) for r in rows], fmt) for c in cols}
        empty = {c: '' for c in cols if sum(r.get(c) == '' for r in rows) > sum(r.get(c) is None for r in rows)}
        maps = {}; consistent = set()
        for a, b in (('status', 'status_name'), ('kind', 'kind_name'), ('theme_name', 'theme')):
            if a in cols and b in cols:
                m = {}; ok = True
                for r in rows:
                    if r.get(a) is None or r.get(b) is None: continue
                    k = str(r[a])
                    if k in m and m[k] != r[b]: ok = False
                    m[k] = r[b]
                maps[b] = {'from': a, 'map': m}
                if ok: consistent.add(b)
        derived = [c for c in cols if (c in DERIVED['*'] + DERIVED.get(ds, [])) and (c not in maps or c in consistent)]
        edit = [c for c in cols if c not in derived]
        schema['sheets'].append({'ds': ds, 'sheet': sname, 'fmt': fmt, 'columns': cols, 'edit': edit,
                                 'derived': derived, 'types': types, 'maps': maps, 'empty': empty, 'kv': ds == 'project'})
        ws = wb.create_sheet(sname); ws.sheet_view.rightToLeft = True
        lab = lambda c: LB.get(ds, {}).get(c) or L.get(c) or c
        if ds == 'project':
            ws['A1'], ws['B1'], ws['C1'] = 'key', 'البيان', 'القيمة'
            for c in 'ABC': ws[c + '1'].font = HFONT; ws[c + '1'].fill = HEAD
            r0 = rows[0] if rows else {}
            for k, c in enumerate(edit, start=2):
                ws.cell(k, 1, c); ws.cell(k, 2, lab(c)).font = Font(name='Arial', bold=True, size=10)
                v = to_cell(r0.get(c), types[c]); cell = ws.cell(k, 3, v); cell.font = BODY
                if types[c] == 'date': cell.number_format = 'dd/mm/yyyy'
                elif isinstance(v, str): cell.data_type = 's'
                for cc in range(1, 4): ws.cell(k, cc).border = BOX
            ws.column_dimensions['A'].hidden = True
            ws.column_dimensions['B'].width = 34; ws.column_dimensions['C'].width = 60
            ws.freeze_panes = 'A2'
            sheets.append((sname, ds, 1)); continue
        letters = {c: get_column_letter(k) for k, c in enumerate(edit, start=1)}
        for k, c in enumerate(edit, start=1):
            ws.cell(1, k, c).font = Font(name='Arial', size=8, color='808080')
            hc = ws.cell(2, k, lab(c)); hc.font = HFONT; hc.fill = HEAD
            hc.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); hc.border = BOX
        ws.row_dimensions[1].hidden = True; ws.row_dimensions[2].height = 32
        fml_all = FORMULAS.get(ds, {})
        last_rep = max([k for k, r in enumerate(rows) if r.get('reported') == 'yes'] or [-1])
        for rn, r in enumerate(rows, start=3):
            fml = fml_all if rn - 3 > last_rep else {}
            for k, c in enumerate(edit, start=1):
                cell = ws.cell(rn, k)
                if c in fml:
                    cell.value = '=' + fml[c].format(**{x: f'{letters[x]}{rn}' for x in letters})
                    cell.fill = CALC
                else:
                    v = to_cell(r.get(c), types[c])
                    if isinstance(v, (list, dict)): v = json.dumps(v, ensure_ascii=False)
                    cell.value = v
                    if isinstance(v, str): cell.data_type = 's'
                cell.font = BODY; cell.border = BOX
                if types[c] == 'date': cell.number_format = 'dd/mm/yyyy'
                elif types[c] in ('num', 'numstr') or c in fml_all: cell.number_format = 'General'
        for k, c in enumerate(edit, start=1):
            vals = [str(r.get(c) or '') for r in rows[:300]]
            w = max([len(lab(c))] + [min(len(x), 70) for x in vals]) + 2
            ws.column_dimensions[get_column_letter(k)].width = max(9, min(w, 60))
            if c == 'coords': ws.column_dimensions[get_column_letter(k)].hidden = True
            if max([len(x) for x in vals] + [0]) > 60:
                for rn in range(3, len(rows) + 3): ws.cell(rn, k).alignment = Alignment(wrap_text=True, vertical='top')
            opts = sorted({str(r.get(c)) for r in rows if r.get(c) not in (None, '')})
            if c in DROPDOWN and 0 < len(opts) <= 12 and all(',' not in o and len(o) < 60 for o in opts):
                dv = DataValidation(type='list', formula1='"' + ','.join(opts) + '"', allow_blank=True, showErrorMessage=False)
                ws.add_data_validation(dv); dv.add(f'{letters[c]}3:{letters[c]}{max(len(rows) + 300, 500)}')
        ws.freeze_panes = 'B3'
        ws.auto_filter.ref = f'A2:{get_column_letter(len(edit))}{len(rows) + 2}'
        sheets.append((sname, ds, len(rows)))
    # guide sheet
    code, nwc, name = PROJ[proj]
    g['B2'] = f'قاعدة بيانات داشبورد المشروع {proj} – {code} – {nwc}'; g['B2'].font = Font(name='Arial', bold=True, size=14, color='1F3864')
    g['B3'] = name; g['B3'].font = Font(name='Arial', size=11)
    g['B5'] = 'تاريخ البيانات'; g['B5'].font = Font(name='Arial', bold=True, size=11)
    g['C5'] = datetime.date.fromisoformat(data_date); g['C5'].number_format = 'dd/mm/yyyy'
    g['C5'].font = Font(name='Arial', bold=True, size=11, color='C00000'); g['C5'].fill = PatternFill('solid', fgColor='FFF2CC')
    g['D5'] = 'يظهر في رأس الداشبورد. غيّره عند كل تحديث.'; g['D5'].font = Font(name='Arial', size=9, color='808080')
    notes = [
        'طريقة الاستخدام',
        '1. عدّل البيانات في الأوراق مباشرة: أضف صفوفاً جديدة أسفل الجدول، أو غيّر القيم الموجودة.',
        '2. لا تغيّر عناوين الأعمدة، ولا تحذف الصف الأول المخفي في كل ورقة (فيه أسماء الحقول التي يقرأها الداشبورد).',
        '3. التواريخ بصيغة يوم/شهر/سنة. النسب المئوية تُكتب رقماً (12.5 وليس 12.5%).',
        '4. الخلايا الرمادية محسوبة تلقائياً (الانحراف وSPI في البيانات الأسبوعية). اسم الحالة بالعربية وأيام الإغلاق تُحسب عند التحديث.',
        '5. في البيانات الأسبوعية: الأسبوع المُبلّغ عنه يُكتب yes في عمود «مُبلّغ عنه»، وآخر أسبوع yes هو ما يظهر افتراضياً في الداشبورد.',
        '6. بعد الحفظ أرسل الملف وقل «حدّث الداشبورد»، فتُحدَّث الأماكن الثلاثة: الداشبورد الأصلي وصفحة Claude العامة وGitHub Pages.',
        '7. البرنامج الزمني (WBS والمسار الحرج والخطة اليومية) والخرائط والمسارات لا تُعدَّل من هذا الملف؛ تُحدَّث من ملفات Primavera وKMZ.',
    ]
    for k, t in enumerate(notes, start=7):
        c = g.cell(k, 2, t); c.font = Font(name='Arial', bold=(k == 7), size=11 if k == 7 else 10)
        g.merge_cells(start_row=k, start_column=2, end_row=k, end_column=4); c.alignment = Alignment(wrap_text=True, vertical='top')
        if k > 7: g.row_dimensions[k].height = 28
    for k in (2, 3): g.merge_cells(start_row=k, start_column=2, end_row=k, end_column=4)
    r = 7 + len(notes) + 1
    for c, t in zip('BCD', ['الورقة', 'المحتوى', 'عدد الصفوف']):
        g[f'{c}{r}'] = t; g[f'{c}{r}'].font = HFONT; g[f'{c}{r}'].fill = HEAD; g[f'{c}{r}'].border = BOX
    for sname, ds, n in sheets:
        r += 1
        g[f'B{r}'] = sname; g[f'B{r}'].hyperlink = f"#'{sname}'!A1"; g[f'B{r}'].font = Font(name='Arial', size=10, color='0563C1', underline='single')
        g[f'C{r}'] = DESC.get(ds, ''); g[f'C{r}'].font = BODY; g[f'C{r}'].alignment = Alignment(wrap_text=True)
        g[f'D{r}'] = n; g[f'D{r}'].font = BODY
        for c in 'BCD': g[f'{c}{r}'].border = BOX
    g.column_dimensions['A'].width = 2; g.column_dimensions['B'].width = 30; g.column_dimensions['C'].width = 95; g.column_dimensions['D'].width = 14
    # hidden schema sheet
    sc = wb.create_sheet('_schema'); s = json.dumps(schema, ensure_ascii=False)
    for k in range(0, len(s), 30000): sc.cell(k // 30000 + 1, 1, s[k:k + 30000]).data_type = 's'
    sc.sheet_state = 'hidden'
    wb.save(out); print('saved', out, [(a, n) for a, _, n in sheets])

if __name__ == '__main__':
    main(*sys.argv[1:5])
