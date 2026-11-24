# Arabic column labels used when the dashboard markup has no table header for a field
L = {
 # project
 'code':'الرمز','project_no':'رقم المشروع (شركة المياه)','po_no':'رقم أمر الشراء','name_en':'اسم المشروع بالإنجليزية',
 'start':'تاريخ البدء','finish':'تاريخ الانتهاء','revised_finish':'تاريخ الانتهاء المعدل','extension_days':'أيام التمديد',
 'contract_value_excl_vat':'قيمة العقد بدون الضريبة','vat_rate_pct':'نسبة الضريبة %','design_capacity_m3_day':'السعة التصميمية م³/يوم',
 'beneficiaries':'عدد المستفيدين','target_year':'سنة الهدف','owner':'المالك','consultant':'الاستشاري','contractor':'المقاول',
 # weekly
 'week':'الأسبوع','week_start':'بداية الأسبوع','week_end':'نهاية الأسبوع','reported':'مُبلّغ عنه (yes/no)',
 'planned_pct':'المخطط %','actual_pct':'الفعلي %','variance_pct':'الانحراف % (محسوب)','spi':'SPI (محسوب)',
 'docs_submitted':'مستندات مقدمة','docs_approved':'مستندات معتمدة','review_days':'أيام المراجعة',
 'trench_open_m':'الحفر م.ط','tunneling_m':'الدفع النفقي م.ط','pipe_installed_m':'تركيب الأنابيب م.ط','tested_pipes_m':'أنابيب مختبرة م.ط',
 'rehab_pipes_m':'أنابيب مؤهلة م.ط','bedding_m':'الفرشة/الردم م.ط','asphalt_m':'الأسفلت م.ط','connections':'التوصيلات',
 'valve_chambers':'غرف المحابس','valves':'المحابس','manholes':'المناهل','staff_approved':'الكادر المعتمد','staff_mobilized':'الكادر المتواجد',
 'contractor_staff':'كادر المقاول','contract_value':'قيمة العقد','revised_contract_value':'قيمة العقد المعدلة','earned_value':'القيمة المكتسبة',
 'invoices':'المستخلصات (تراكمي)','consultant_penalties':'غرامات الاستشاري','authority_penalties':'غرامات الجهات',
 'ptw_opened':'تصاريح عمل صادرة','ptw_active':'تصاريح عمل سارية','ptw_closed':'تصاريح عمل مغلقة','ptw_not_started':'تصاريح لم تبدأ',
 'fatalities':'الوفيات','near_misses':'الحوادث الوشيكة','hse_sor_opened':'ملاحظات سلامة صادرة','hse_sor_closed':'ملاحظات سلامة مغلقة',
 'hse_sor_open':'ملاحظات سلامة مفتوحة','hse_ncr_opened':'NCR سلامة صادرة','hse_ncr_closed':'NCR سلامة مغلقة','hse_ncr_open':'NCR سلامة مفتوحة',
 'q_sor_opened':'ملاحظات جودة صادرة','q_sor_closed':'ملاحظات جودة مغلقة','q_sor_open':'ملاحظات جودة مفتوحة','q_ncr_opened':'NCR جودة صادرة',
 'q_ncr_closed':'NCR جودة مغلقة','q_ncr_open':'NCR جودة مفتوحة','status':'الحالة','thrust_pits':'غرف الدفع والاستقبال',
 # daily reports
 'date':'التاريخ','report_no':'رقم التقرير','day_pct':'إنجاز اليوم %','pipes_count':'عدد المواسير المركبة','staff':'الكادر',
 'labour':'العمالة','equipment':'المعدات','source':'المصدر',
 # components
 'id':'م','icon':'الأيقونة','component':'المكون','detail':'التفاصيل','qty':'الكمية','unit':'الوحدة','boq_ref':'بند جدول الكميات',
 'value':'القيمة','pct':'النسبة',
 # invoices
 'no':'الرقم','period_from':'من','period_to':'إلى','vat':'الضريبة','total':'الإجمالي','cumulative':'التراكمي','submitted':'تاريخ الرفع',
 'approved':'تاريخ الاعتماد','approval_basis':'أساس الاعتماد',
 # registers
 'ref':'الرقم المرجعي','inspector':'المفتش','issued':'تاريخ الإصدار','location':'الموقع','desc':'الوصف','close_request':'تاريخ طلب الإغلاق',
 'closed':'تاريخ الإغلاق','days_to_close':'أيام الإغلاق','days_open':'أيام مفتوحة','month':'الشهر','complete':'مكتمل',
 'received':'تاريخ الاستلام','resubmitted':'تاريخ إعادة التقديم','days_resubmit':'أيام الرد','corrective_date':'تاريخ الإجراء التصحيحي',
 'days_corrective':'أيام الإجراء التصحيحي','category':'التصنيف','desc_ar':'الوصف بالعربية','type':'النوع','rev':'المراجعة',
 'element':'العنصر','returned':'تاريخ الرد','days':'الأيام','remarks':'ملاحظات','discipline':'التخصص','supplier':'المورد',
 'mat_approval':'اعتماد المادة','storage':'التخزين','inspected':'تاريخ الفحص','is_latest':'آخر مراجعة','material_ref':'رقم المادة',
 'letter':'الخطاب','replied':'تاريخ الرد','date_check':'تدقيق التاريخ','item_ref':'رقم البند','title':'العنوان','row_status':'حالة السطر',
 'subject':'الموضوع','company':'الشركة','line_of_work':'مجال العمل','letter_reply':'خطاب الرد',
 # permits
 'permit_no':'رقم التصريح','work':'نوع العمل','method':'طريقة التنفيذ','version':'الإصدار','issue':'رقم الإصدار','nature':'طبيعة التصريح',
 'equip':'المعدات','cleared':'شهادة الإخلاء','warranty_to':'الضمان حتى','paths':'عدد المسارات','municipality':'البلدية',
 'supervisor':'المشرف','balady_id':'رقم بلدي','file':'الملف','district':'الحي','street':'الشارع',
 # supplies
 'accepted':'مقبول','key':'المفتاح','material':'المادة','variant':'النوع الفرعي','len_per_unit':'الطول لكل وحدة','need_m':'المطلوب م','qty_text':'نص الكمية',
 # white helmet
 'name':'الاسم','camera':'الكاميرا','link':'الرابط',
 # works segments
 'sector':'القطاع','mh_from':'من منهل','mh_to':'إلى منهل','ch_from':'من محطة','ch_to':'إلى محطة','length_m':'الطول م','kind':'طريقة التنفيذ',
 'stage':'المرحلة','wir_count':'عدد طلبات الاستلام','last_wir':'آخر طلب استلام','mh_known':'موقع المنهل معروف','coords':'الإحداثيات',
 # critical
 'n':'رقم الملاحظة','theme':'رمز نوع الخطر','theme_name':'نوع الخطر','summary':'البند الحرج','repeats':'ورد سابقاً (مرة)','last_prior':'آخر ملاحظة سابقة',
 'count':'إجمالي الملاحظات التي ورد فيها','open':'مفتوحة منها','critical':'ملاحظات حرجة','critical_open':'حرجة مفتوحة',
 'first':'أول ظهور','last':'آخر ظهور','refs':'آخر الأرقام',
 # issues
 'text':'النص','first_week':'أول أسبوع','last_week':'آخر أسبوع','current':'قائمة حالياً',
}
