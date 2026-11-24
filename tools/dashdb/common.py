import json,re,html as H
DS={  # dataset id -> (sheet name, original format)
 'project':('بيانات المشروع','csv'),
 'weekly':('البيانات الأسبوعية','csv'),
 'daily_reports':('التقارير اليومية','json'),
 'components':('مكونات المشروع','csv'),
 'invoices':('المستخلصات','csv'),
 'issues':('القضايا والمعوقات','json'),
 'reg_sor_hse':('ملاحظات السلامة SOR','json'),
 'sor_critical':('الملاحظات الحرجة','sor'),
 'sor_themes':('أنواع المخاطر الحرجة','sor'),
 'reg_ncr':('عدم المطابقة NCR','json'),
 'reg_incident':('الحوادث','json'),
 'reg_ptw':('تصاريح العمل PTW','json'),
 'permits':('تصاريح الحفر - بلدي','json'),
 'reg_sor_q':('ملاحظات الجودة SOR','json'),
 'reg_wir':('طلبات الاستلام WIR','json'),
 'reg_mir':('فحص المواد MIR','json'),
 'reg_mar':('اعتماد المواد MAT','json'),
 'reg_sub':('مقاولو الباطن','json'),
 'reg_doc':('تقديمات المستندات','json'),
 'reg_shd':('المخططات التنفيذية','json'),
 'supplies':('التوريدات','json'),
 'white_helmet':('White Helmet','json'),
 'works_segments':('موقف الأعمال بين المناهل','json'),
}
def load_D(path):
    h=open(path,encoding='utf-8').read()
    i=h.find('const D=')+8; j=h.find(',C={},R=',i)
    return h,i,j,json.loads(h[i:j])
def labels(h):
    out={}
    for t in re.finditer(r'<table[^>]*data-source="([^"]+)"[^>]*>(.*?)</table>',h,re.S):
        ds=t.group(1).split()[0]
        for th in re.finditer(r'<th[^>]*data-field="([^"]+)"[^>]*>(.*?)</th>',t.group(2),re.S):
            src=re.search(r'data-source="([^"]+)"',th.group(0)); d=src.group(1) if src else ds
            out.setdefault(d,{}).setdefault(th.group(1),H.unescape(re.sub('<[^>]+>','',th.group(2))).strip())
    return out
