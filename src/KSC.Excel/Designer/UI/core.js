/* KSC.Excel designer - core: state, i18n, helpers, template model operations. No external dependencies. */
(function () {
'use strict';
const KXL = window.KXL = { S: {}, ui: {}, fn: {} };

/* ------------------------------------------------------------------ i18n */
const I18N = {
  fa: {
    appName: 'طراح قالب اکسل', newT: 'جدید', open: 'باز کردن', save: 'ذخیره', preview: 'پیش‌نمایش', download: 'دریافت xlsx',
    undo: 'بازگردانی (Ctrl+Z)', redo: 'جلو (Ctrl+Y)', templateName: 'نام قالب', untitled: 'بدون‌نام', saved: 'ذخیره شد', saveFailed: 'ذخیره ناموفق بود',
    unsaved: 'تغییرات ذخیره‌نشده', sheet: 'برگه', addSheet: 'برگه جدید', rename: 'تغییر نام', duplicate: 'کپی', del: 'حذف', cancel: 'انصراف', ok: 'تأیید', close: 'بستن',
    fields: 'فیلدها', data: 'داده', sample: 'داده‌ی نمونه', sampleHelp: 'JSON نمونه‌ای از داده‌ای که به قالب می‌دهید. فقط برای پیش‌نمایش در طراح است.',
    model: 'مدل داده', noModel: '(بدون مدل — از داده‌ی نمونه)', searchFields: 'جستجوی فیلد…', specials: 'متغیرهای ویژه', insertHint: 'دوبار کلیک = درج در سلول · بکشید و رها کنید روی سلول',
    cell: 'سلول', sheetP: 'برگه', bands: 'نوارها', rules: 'قوانین', page: 'چاپ', images: 'تصاویر و نمودار',
    value: 'مقدار / فرمول', link: 'پیوند', typeAuto: 'خودکار (عدد/متن)', typeText: 'همیشه متن', valueType: 'نوع مقدار',
    font: 'فونت', size: 'اندازه', bold: 'ضخیم (Ctrl+B)', italic: 'مورب (Ctrl+I)', underline: 'زیرخط (Ctrl+U)', strike: 'خط‌خورده', fontColor: 'رنگ متن', fillColor: 'رنگ پس‌زمینه',
    borders: 'کادر', alignH: 'تراز افقی', alignV: 'تراز عمودی', left: 'چپ', center: 'وسط', right: 'راست', top: 'بالا', middle: 'وسط', bottom: 'پایین', wrap: 'شکستن متن', merge: 'ادغام', unmerge: 'جداسازی',
    numFmt: 'قالب عدد', indent: 'تورفتگی', outdent: 'کاهش تورفتگی', clearFmt: 'پاک کردن قالب', rtlText: 'جهت متن راست‌به‌چپ',
    insRowAbove: 'درج سطر بالا', insRowBelow: 'درج سطر پایین', delRow: 'حذف سطر', insColBefore: 'درج ستون قبل', insColAfter: 'درج ستون بعد', delCol: 'حذف ستون', clearCells: 'پاک کردن محتوا',
    freeze: 'ثابت‌سازی', unfreeze: 'لغو ثابت‌سازی', freezeAt: 'ثابت‌سازی تا سلول فعال', addBand: 'افزودن نوار تکرار', addImage: 'درج تصویر', condFmt: 'قالب شرطی', validation: 'اعتبارسنجی',
    bandName: 'نام نوار', bandSource: 'مسیر مجموعه', bandRows: 'سطرها', bandAlt: 'رنگ ردیف‌های زوج', bandKeep: 'اگر خالی بود یک سطر خالی بماند', bandHelp: 'نوار: سطرهای انتخابی برای هر آیتم مجموعه تکرار می‌شوند.',
    noBands: 'نواری تعریف نشده. چند سطر را انتخاب و «افزودن نوار تکرار» را بزنید، یا یک فیلد از مجموعه را روی یک سطر بکشید.', from: 'از', to: 'تا',
    sheetName: 'نام برگه', rtl: 'راست‌به‌چپ', gridlines: 'خطوط شبکه', tabColor: 'رنگ تب', zoom: 'بزرگ‌نمایی', defCol: 'عرض پیش‌فرض ستون', defRow: 'ارتفاع پیش‌فرض سطر', repeatFor: 'تکرار برگه برای هر آیتم', protect: 'محافظت از برگه',
    autoFilter: 'فیلتر خودکار روی انتخاب', clearFilter: 'حذف فیلتر', orientation: 'جهت صفحه', portrait: 'عمودی', landscape: 'افقی', paper: 'کاغذ', fitWidth: 'جا دادن عرض در یک صفحه', margins: 'حاشیه‌ها (اینچ)',
    header: 'سرصفحه', footer: 'پاصفحه', printTitles: 'تکرار سطرها در هر صفحه (مثل 1:3)', printArea: 'محدوده‌ی چاپ', center2: 'وسط‌چین افقی',
    rangeLbl: 'محدوده', type: 'نوع', operator: 'عملگر', value1: 'مقدار ۱', value2: 'مقدار ۲', textLbl: 'متن', add: 'افزودن', edit: 'ویرایش', format: 'قالب', fill: 'پس‌زمینه',
    noRules: 'قانونی نیست.', cfCellIs: 'مقدار سلول', cfExpr: 'فرمول', cfText: 'شامل متن', cfDup: 'مقادیر تکراری', cfTop: 'بیشترین/کمترین', cfScale: 'طیف رنگ', cfBar: 'نوار داده',
    opGt: 'بزرگ‌تر از', opLt: 'کوچک‌تر از', opGe: 'بزرگ‌تر یا مساوی', opLe: 'کوچک‌تر یا مساوی', opEq: 'مساوی', opNe: 'نامساوی', opBetween: 'بین', opNotBetween: 'خارج از',
    vList: 'فهرست', vWhole: 'عدد صحیح', vDecimal: 'عدد اعشاری', vDate: 'تاریخ', vLen: 'طول متن', vCustom: 'فرمول سفارشی', vListHelp: 'با کاما جدا کنید یا =$H$1:$H$5', errMsg: 'پیام خطا',
    imgBinding: 'مسیر داده‌ی تصویر (اختیاری)', imgPos: 'موقعیت', imgSize: 'اندازه (px)', imgStatic: 'تصویر ثابت', pickImage: 'انتخاب تصویر…', imgDynamicHelp: 'اگر مسیر داده بدهید (byte[] یا base64) تصویر از داده خوانده می‌شود.',
    previewTitle: 'پیش‌نمایش با داده‌ی نمونه', warnings: 'هشدارها', noWarn: 'مشکلی پیدا نشد.', loading: 'در حال بارگذاری…',
    openTitle: 'باز کردن قالب', noTemplates: 'هنوز قالبی ذخیره نشده.', confirmDel: 'حذف شود؟', confirmDiscard: 'تغییرات ذخیره‌نشده از بین می‌رود. ادامه می‌دهید؟',
    newTitle: 'شروع قالب جدید', blank: 'قالب خالی', blankRtl: 'خالی (راست‌به‌چپ)', starter_table: 'گزارش جدولی (عنوان، نوار تکرار، جمع)', starter_invoice: 'فاکتور',
    exportJson: 'خروجی JSON', importJson: 'ورود JSON', readOnly: 'حالت فقط‌خواندنی', invalidName: 'نام نامعتبر: فقط حروف، عدد، فاصله، _ - .',
    ctxInsert: 'درج', menuCut: 'برش', menuCopy: 'کپی', menuPaste: 'چسباندن', alignment: 'تراز', numberGroup: 'عدد', structure: 'ساختار', dataGroup: 'داده', rulesGroup: 'قواعد',
    sumBtn: 'درج جمع (Σ) این ستون', jalaliBtn: 'درج به‌صورت تاریخ شمسی', bordersAll: 'همه‌ی کادرها', bordersOuter: 'کادر بیرونی', bordersNone: 'بدون کادر', bordersTop: 'بالا', bordersBottom: 'پایین', bordersLeft: 'چپ', bordersRight: 'راست',
    bordersThickOuter: 'کادر بیرونی ضخیم', bordersDouble: 'خط دوبل پایین', bStyle: 'سبک خط', bColor: 'رنگ خط', customFmt: 'سفارشی…', customFmtPrompt: 'قالب عدد اکسل (مثل #,##0.00)',
    modelFromSample: 'ساخت از داده‌ی نمونه', applyJson: 'اعمال', invalidJson: 'JSON نامعتبر است', rowsSel: 'سطر', colsSel: 'ستون', bandLbl: 'نوار', hintFormula: 'فرمول: =SUM([Items.Price]) جمع ستون Price در نوار Items؛ =[Qty]*[Price] برای همان سطر',
    help: 'راهنما', tipsTitle: 'نکته‌ها', tips: ['متن ثابت را مستقیم بنویسید؛ داده را با {{Field}} بگذارید.', 'سطری که فیلد یک مجموعه در آن است برای هر آیتم تکرار می‌شود (نوار).', 'جمع ستون: =SUM([Items.Price]).', 'تاریخ شمسی: {{Date|fa}}.'],
    sortBy: 'مرتب‌سازی بر اساس', sortDesc: 'نزولی', groupBy: 'گروه‌بندی بر اساس', groupHelp: 'با گروه‌بندی، هر تکرار یک گروه است (فیلدهای Key, Count, Items). یک نوار تو در تو با مسیر Items داخلش بگذارید و جمع هر گروه را با =SUM([نوار داخلی.فیلد]) بنویسید.',
    charts: 'نمودارها', addChart: 'افزودن نمودار', chartType: 'نوع نمودار', chartTitle: 'عنوان', categories: 'محدوده‌ی دسته‌ها (مثل B9:B9)', series: 'سری‌ها', seriesName: 'نام سری', seriesValues: 'محدوده‌ی مقادیر (مثل D9:D9)', addSeries: 'افزودن سری', stacked: 'انباشته', legend: 'راهنما', dataLabels: 'برچسب مقدار', chartHelp: 'محدوده‌ای که یک نوار را بپوشاند، با داده بزرگ می‌شود.', tColumn: 'ستونی', tBar: 'میله‌ای', tLine: 'خطی', tArea: 'ناحیه‌ای', tPie: 'دایره‌ای', imagesCharts: 'تصاویر و نمودار', noCharts: 'نموداری نیست.',
    orientation2: 'جهت تکرار', orRows: 'سطرها (عمودی)', orCols: 'ستون‌ها (افقی - جدول متقاطع)', colsRange: 'ستون‌ها', colBandHelp: 'نوار ستونی: ستون‌های انتخابی به تعداد آیتم‌های مجموعه تکثیر می‌شوند. داخل آن: {{$item.Key}} آیتم جاری، {{Cells[$col].Value}} برای داده‌ی ردیف.',
    sets: 'مجموعه و پارامتر', dataSets: 'مجموعه‌های داده (جدول متقاطع، SQL، JSON)', addPivot: 'افزودن جدول متقاطع', dsName: 'نام', dsSource: 'مسیر مجموعه‌ی منبع', dsRow: 'فیلد سطر', dsCol: 'فیلد ستون', dsValue: 'فیلد مقدار', dsAgg: 'تجمیع', params: 'پارامترها', addParam: 'افزودن پارامتر', pName: 'نام', pLabel: 'عنوان', pType: 'نوع', pDefault: 'مقدار پیش‌فرض', paramHelp: 'در قالب: {{$p.Name}} . برنامه‌نویس هنگام تولید مقدار می‌دهد، وگرنه پیش‌فرض.', pivotHelp: 'خروجی: Name.Columns (Key, Total) ، Name.Rows (Key, Total, Cells[Value]) ، Name.GrandTotal',
    addSql: 'افزودن کوئری SQL', dsConnection: 'اتصال', dsSql: 'متن SELECT (پارامتر: @Name)', testQuery: 'آزمایش کوئری', sqlDisabled: 'اجرای SQL در سرور غیرفعال است (AllowSql). در پیش‌نمایش از داده‌ی نمونه استفاده می‌شود.', sqlHint: 'فقط SELECT / WITH. مقدار @Name از «پارامتر» هم‌نام گرفته می‌شود. از کاربر فقط‌خواندنی دیتابیس استفاده کنید.', rowsFound: 'ردیف', exprHint: 'محاسبه: {{Qty * Price}} · {{IF(Total > 1000, "VIP", "-")}} · {{ROUND(Price * 1.09, 0)}}',
    pdf: 'دریافت PDF', pdfFail: 'تولید PDF ممکن نشد',
    history: 'تاریخچه', historyTip: 'نسخه‌های قبلی این قالب', historyHelp: 'نسخه‌ی قبل از هر ذخیره نگه داشته می‌شود. «باز کردن» نسخه را در طراح می‌گذارد؛ با «ذخیره» جای نسخه‌ی فعلی می‌نشیند (نسخه‌ی فعلی هم در تاریخچه می‌ماند).', historyEmpty: 'هنوز نسخه‌ی قبلی‌ای نیست.', historyOpen: 'باز کردن', historyLoaded: 'نسخه بارگذاری شد؛ برای بازگردانی ذخیره کنید', historyNeedSave: 'ابتدا قالب را ذخیره کنید',
    comboType: 'نوع این سری (ترکیبی)', secondaryAxis: 'محور دوم (سمت راست)',
    gwNone: '(بدون سطح دوم)', gwGroup2: 'سطح دوم گروه‌بندی', gwExpr: 'یا عبارت (مثل =JYM(Date))', gwExprHelp: 'عبارت‌ها: =JYM(Date) سال/ماه شمسی · =JYEAR(Date) · =JMONTHNAME(Date) · =IF(Qty>10,"big","small")', gwGrand: 'سطر جمع کل', gwGrandLbl: 'جمع کل',
    barcode: 'بارکد', bcNone: '(بدون بارکد)', bcQr: 'QR Code', bcCode128: 'Code 128', bcHelp: 'متن سلول (مثلاً {{InvoiceNo}}) به بارکد تبدیل می‌شود و کل سلول را پر می‌کند. سلول را بزرگ‌تر یا ادغام کنید.',
    pageBreak: 'هر آیتم در صفحه‌ی جدید', watermark: 'واترمارک (فقط PDF)', addJson: 'افزودن داده‌ی JSON / REST', addProvider: 'افزودن ارائه‌دهنده‌ی کد', dsJsonInline: 'JSON ثابت (اختیاری)', dsEndpoint: 'نقطه‌ی REST (ثبت‌شده در کد)', dsProvider: 'ارائه‌دهنده (ثبت‌شده در کد)', dsUrl: 'مسیر / کوئری ({Name} = پارامتر)', dsJsonPath: 'مسیر آرایه در JSON (مثل data.items)', testData: 'آزمایش', jsonHelp: 'یا JSON ثابت بنویسید، یا یک نقطه‌ی REST که برنامه‌نویس ثبت کرده انتخاب کنید (فقط مسیر را شما می‌دهید، نه آدرس سرور).', providerHelp: 'ارائه‌دهنده یک تابع در کد برنامه است که با پارامترهای گزارش، داده را برمی‌گرداند (EF، Dapper، سرویس...).', noEndpoints: 'هیچ نقطه‌ی REST ثبت نشده (AddHttpEndpoint).', noProviders: 'هیچ ارائه‌دهنده‌ای ثبت نشده (AddProvider).',
    sections: 'بخش‌های گزارش', secReportHeader: 'سرگزارش', secPageHeader: 'سرصفحه (تکرار در هر صفحه)', secReportFooter: 'پاگزارش', addSection: 'افزودن بخش روی سطرهای انتخابی', secHelp: 'برچسب بخش روی سطرها. «سرصفحه» در چاپ و PDF روی هر صفحه تکرار می‌شود.',
    groupWizard: 'ویزارد گروه‌بندی', gwList: 'لیست', gwGroup: 'گروه‌بندی بر اساس', gwFields: 'ستون‌های جدول', gwTotals: 'جمع هر گروه', gwTotalLbl: 'جمع', gwNeedList: 'ابتدا یک مدل یا داده‌ی نمونه با لیست لازم است.', gwHelp: 'چهار سطر می‌سازد: عنوان ستون‌ها، سرگروه، سطر داده، و جمع گروه — با نوار گروه‌بندی و نوار داده.',
    importXlsx: 'ورود از فایل Excel', importDone: 'فایل Excel وارد شد. سلول‌هایی که باید از داده بیایند را با {{Field}} جایگزین کنید.', importFail: 'ورود فایل ممکن نشد',
    selSummary: 'انتخاب', restore: 'بازیابی', discard: 'دور انداختن', restoreAsk: 'پیش‌نویس ذخیره‌نشده‌ای از جلسه‌ی قبل پیدا شد. بازیابی شود؟'
  },
  en: {
    appName: 'Excel Template Designer', newT: 'New', open: 'Open', save: 'Save', preview: 'Preview', download: 'Download xlsx',
    undo: 'Undo (Ctrl+Z)', redo: 'Redo (Ctrl+Y)', templateName: 'Template name', untitled: 'Untitled', saved: 'Saved', saveFailed: 'Save failed',
    unsaved: 'Unsaved changes', sheet: 'Sheet', addSheet: 'Add sheet', rename: 'Rename', duplicate: 'Duplicate', del: 'Delete', cancel: 'Cancel', ok: 'OK', close: 'Close',
    fields: 'Fields', data: 'Data', sample: 'Sample data', sampleHelp: 'Sample JSON of the data you will pass to the template. Used only for the designer preview.',
    model: 'Data model', noModel: '(no model - use sample data)', searchFields: 'Search fields…', specials: 'Special variables', insertHint: 'Double-click = insert into cell · drag onto a cell',
    cell: 'Cell', sheetP: 'Sheet', bands: 'Bands', rules: 'Rules', page: 'Print', images: 'Images & charts',
    value: 'Value / formula', link: 'Link', typeAuto: 'Auto (number/text)', typeText: 'Always text', valueType: 'Value type',
    font: 'Font', size: 'Size', bold: 'Bold (Ctrl+B)', italic: 'Italic (Ctrl+I)', underline: 'Underline (Ctrl+U)', strike: 'Strikethrough', fontColor: 'Font color', fillColor: 'Fill color',
    borders: 'Borders', alignH: 'Horizontal align', alignV: 'Vertical align', left: 'Left', center: 'Center', right: 'Right', top: 'Top', middle: 'Middle', bottom: 'Bottom', wrap: 'Wrap text', merge: 'Merge', unmerge: 'Unmerge',
    numFmt: 'Number format', indent: 'Indent', outdent: 'Outdent', clearFmt: 'Clear formatting', rtlText: 'Right-to-left text',
    insRowAbove: 'Insert row above', insRowBelow: 'Insert row below', delRow: 'Delete row', insColBefore: 'Insert column before', insColAfter: 'Insert column after', delCol: 'Delete column', clearCells: 'Clear contents',
    freeze: 'Freeze', unfreeze: 'Unfreeze', freezeAt: 'Freeze up to active cell', addBand: 'Add repeat band', addImage: 'Insert image', condFmt: 'Conditional format', validation: 'Validation',
    bandName: 'Band name', bandSource: 'Collection path', bandRows: 'Rows', bandAlt: 'Even-row fill', bandKeep: 'Keep one blank block when empty', bandHelp: 'A band repeats the selected rows for every item of a collection.',
    noBands: 'No bands. Select rows and press "Add repeat band", or drag a collection field onto a row.', from: 'From', to: 'To',
    sheetName: 'Sheet name', rtl: 'Right-to-left', gridlines: 'Gridlines', tabColor: 'Tab color', zoom: 'Zoom', defCol: 'Default column width', defRow: 'Default row height', repeatFor: 'Repeat sheet per item of', protect: 'Protect sheet',
    autoFilter: 'AutoFilter on selection', clearFilter: 'Remove filter', orientation: 'Orientation', portrait: 'Portrait', landscape: 'Landscape', paper: 'Paper', fitWidth: 'Fit width to one page', margins: 'Margins (inches)',
    header: 'Header', footer: 'Footer', printTitles: 'Repeat rows on each page (e.g. 1:3)', printArea: 'Print area', center2: 'Center horizontally',
    rangeLbl: 'Range', type: 'Type', operator: 'Operator', value1: 'Value 1', value2: 'Value 2', textLbl: 'Text', add: 'Add', edit: 'Edit', format: 'Format', fill: 'Fill',
    noRules: 'No rules.', cfCellIs: 'Cell value', cfExpr: 'Formula', cfText: 'Contains text', cfDup: 'Duplicates', cfTop: 'Top / bottom', cfScale: 'Color scale', cfBar: 'Data bar',
    opGt: 'greater than', opLt: 'less than', opGe: 'greater or equal', opLe: 'less or equal', opEq: 'equal to', opNe: 'not equal', opBetween: 'between', opNotBetween: 'not between',
    vList: 'List', vWhole: 'Whole number', vDecimal: 'Decimal', vDate: 'Date', vLen: 'Text length', vCustom: 'Custom formula', vListHelp: 'Comma separated or =$H$1:$H$5', errMsg: 'Error message',
    imgBinding: 'Image data path (optional)', imgPos: 'Position', imgSize: 'Size (px)', imgStatic: 'Static image', pickImage: 'Choose image…', imgDynamicHelp: 'With a data path (byte[] or base64) the image comes from the data.',
    previewTitle: 'Preview with sample data', warnings: 'Warnings', noWarn: 'No problems found.', loading: 'Loading…',
    openTitle: 'Open template', noTemplates: 'No templates saved yet.', confirmDel: 'Delete?', confirmDiscard: 'Unsaved changes will be lost. Continue?',
    newTitle: 'Start a new template', blank: 'Blank', blankRtl: 'Blank (right-to-left)', starter_table: 'Table report (title, repeat band, totals)', starter_invoice: 'Invoice',
    exportJson: 'Export JSON', importJson: 'Import JSON', readOnly: 'Read-only mode', invalidName: 'Invalid name: letters, digits, space, _ - . only',
    ctxInsert: 'Insert', menuCut: 'Cut', menuCopy: 'Copy', menuPaste: 'Paste', alignment: 'Align', numberGroup: 'Number', structure: 'Structure', dataGroup: 'Data', rulesGroup: 'Rules',
    sumBtn: 'Insert sum (Σ) of this column', jalaliBtn: 'Insert as Persian (Jalali) date', bordersAll: 'All borders', bordersOuter: 'Outside borders', bordersNone: 'No border', bordersTop: 'Top', bordersBottom: 'Bottom', bordersLeft: 'Left', bordersRight: 'Right',
    bordersThickOuter: 'Thick outside border', bordersDouble: 'Double bottom border', bStyle: 'Line style', bColor: 'Line color', customFmt: 'Custom…', customFmtPrompt: 'Excel number format (e.g. #,##0.00)',
    modelFromSample: 'Build from sample data', applyJson: 'Apply', invalidJson: 'Invalid JSON', rowsSel: 'rows', colsSel: 'cols', bandLbl: 'Band', hintFormula: 'Formula: =SUM([Items.Price]) sums column Price of band Items; =[Qty]*[Price] for the same row',
    help: 'Help', tipsTitle: 'Tips', tips: ['Type fixed text directly; bind data with {{Field}}.', 'A row holding a collection field repeats for every item (a band).', 'Column total: =SUM([Items.Price]).', 'Persian date: {{Date|fa}}.'],
    sortBy: 'Sort by', sortDesc: 'Descending', groupBy: 'Group by', groupHelp: 'With grouping each repeat is one group (fields Key, Count, Items). Put a nested band bound to Items inside and total each group with =SUM([InnerBand.Field]).',
    charts: 'Charts', addChart: 'Add chart', chartType: 'Chart type', chartTitle: 'Title', categories: 'Categories range (e.g. B9:B9)', series: 'Series', seriesName: 'Series name', seriesValues: 'Values range (e.g. D9:D9)', addSeries: 'Add series', stacked: 'Stacked', legend: 'Legend', dataLabels: 'Data labels', chartHelp: 'A range covering a band grows with the data.', tColumn: 'Column', tBar: 'Bar', tLine: 'Line', tArea: 'Area', tPie: 'Pie', imagesCharts: 'Images & charts', noCharts: 'No charts.',
    orientation2: 'Repeat direction', orRows: 'Rows (vertical)', orCols: 'Columns (horizontal - cross-tab)', colsRange: 'Columns', colBandHelp: 'Column band: the selected columns are repeated for every item. Inside: {{$item.Key}} is the current item, {{Cells[$col].Value}} reads the row data.',
    sets: 'Data sets & parameters', dataSets: 'Data sets (cross-tab, SQL, JSON)', addPivot: 'Add cross-tab', dsName: 'Name', dsSource: 'Source collection path', dsRow: 'Row field', dsCol: 'Column field', dsValue: 'Value field', dsAgg: 'Aggregate', params: 'Parameters', addParam: 'Add parameter', pName: 'Name', pLabel: 'Label', pType: 'Type', pDefault: 'Default value', paramHelp: 'In the template: {{$p.Name}}. The developer passes values at render time, otherwise the default is used.', pivotHelp: 'Result: Name.Columns (Key, Total), Name.Rows (Key, Total, Cells[Value]), Name.GrandTotal',
    addSql: 'Add SQL query', dsConnection: 'Connection', dsSql: 'SELECT text (parameters: @Name)', testQuery: 'Test query', sqlDisabled: 'SQL execution is disabled on the server (AllowSql). The preview uses sample data.', sqlHint: 'SELECT / WITH only. @Name is bound from the report parameter of the same name. Use a read-only database login.', rowsFound: 'rows', exprHint: 'Calculated: {{Qty * Price}} · {{IF(Total > 1000, "VIP", "-")}} · {{ROUND(Price * 1.09, 0)}}',
    pdf: 'Download PDF', pdfFail: 'PDF export failed',
    history: 'History', historyTip: 'Earlier versions of this template', historyHelp: 'The version before every save is kept. "Open" puts it in the designer; "Save" makes it current (the replaced one stays in the history).', historyEmpty: 'No earlier versions yet.', historyOpen: 'Open', historyLoaded: 'Version loaded - save to restore it', historyNeedSave: 'Save the template first',
    comboType: 'Series type (combo)', secondaryAxis: 'Secondary axis (right)',
    gwNone: '(no second level)', gwGroup2: 'Second grouping level', gwExpr: 'or expression (e.g. =JYM(Date))', gwExprHelp: 'Expressions: =JYM(Date) Jalali year/month · =JYEAR(Date) · =JMONTHNAME(Date) · =IF(Qty>10,"big","small")', gwGrand: 'Grand total row', gwGrandLbl: 'Grand total',
    barcode: 'Barcode', bcNone: '(no barcode)', bcQr: 'QR Code', bcCode128: 'Code 128', bcHelp: 'The cell text (e.g. {{InvoiceNo}}) is encoded as a barcode that fills the cell. Enlarge or merge cells for a bigger code.',
    pageBreak: 'Start each item on a new page', watermark: 'Watermark (PDF only)', addJson: 'Add JSON / REST data', addProvider: 'Add code provider', dsJsonInline: 'Fixed JSON (optional)', dsEndpoint: 'REST endpoint (registered in code)', dsProvider: 'Provider (registered in code)', dsUrl: 'Path / query ({Name} = parameter)', dsJsonPath: 'Array path in the JSON (e.g. data.items)', testData: 'Test', jsonHelp: 'Either type fixed JSON, or pick a REST endpoint your developer registered (you only choose the path, never the server).', providerHelp: 'A provider is a function in the application code that returns the data for the report parameters (EF, Dapper, a service...).', noEndpoints: 'No REST endpoint is registered (AddHttpEndpoint).', noProviders: 'No provider is registered (AddProvider).',
    sections: 'Report sections', secReportHeader: 'Report header', secPageHeader: 'Page header (repeats on every page)', secReportFooter: 'Report footer', addSection: 'Add section on selected rows', secHelp: 'A label over rows. "Page header" repeats on every printed / PDF page.',
    groupWizard: 'Group wizard', gwList: 'List', gwGroup: 'Group by', gwFields: 'Table columns', gwTotals: 'Group totals', gwTotalLbl: 'Total', gwNeedList: 'A model or sample data with a list is needed first.', gwHelp: 'Builds four rows: column titles, group header, detail row and group total - with the group band and detail band.',
    importXlsx: 'Import from Excel file', importDone: 'Excel file imported. Replace the cells that should come from data with {{Field}} placeholders.', importFail: 'Import failed',
    selSummary: 'Selection', restore: 'Restore', discard: 'Discard', restoreAsk: 'An unsaved draft from your previous session was found. Restore it?'
  }
};
const LANG = (document.documentElement.lang === 'en') ? 'en' : 'fa';
KXL.LANG = LANG;
KXL.RTL = LANG === 'fa';
const T = KXL.T = (k) => (I18N[LANG] && I18N[LANG][k] !== undefined) ? I18N[LANG][k] : (I18N.en[k] !== undefined ? I18N.en[k] : k);

/* ------------------------------------------------------------------ http */
const BASE = (document.querySelector('base') ? document.querySelector('base').getAttribute('href') : './').replace(/\/$/, '');
KXL.BASE = BASE;
KXL.api = async function (method, path, body, raw) {
  const opt = { method, headers: { 'X-KXL-Request': '1' }, credentials: 'same-origin' };
  if (body !== undefined) { opt.headers['Content-Type'] = 'application/json'; opt.body = typeof body === 'string' ? body : JSON.stringify(body); }
  const res = await fetch(BASE + '/api/' + path, opt);
  if (raw) { if (!res.ok) throw new Error(await errText(res)); return res; }
  if (!res.ok) throw new Error(await errText(res));
  return res.json();
};
async function errText(res) { try { const j = await res.json(); return j.error || res.statusText; } catch (e) { return res.status + ' ' + res.statusText; } }

/* ------------------------------------------------------------------ helpers */
const $ = KXL.$ = (sel, root) => (root || document).querySelector(sel);
const $$ = KXL.$$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));
KXL.esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
KXL.el = function (tag, attrs, ...kids) {
  const e = document.createElement(tag);
  if (attrs) for (const k in attrs) {
    if (k === 'class') e.className = attrs[k]; else if (k === 'html') e.innerHTML = attrs[k];
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), attrs[k]); else if (attrs[k] !== false && attrs[k] != null) e.setAttribute(k, attrs[k]);
  }
  for (const k of kids.flat()) if (k != null) e.append(k.nodeType ? k : document.createTextNode(k));
  return e;
};
KXL.colName = function (c) { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
KXL.colIndex = function (s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
KXL.a1 = (r, c) => KXL.colName(c) + r;
KXL.parseRange = function (s) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec((s || '').trim());
  if (!m) return null;
  let c1 = KXL.colIndex(m[1]), r1 = +m[2], c2 = m[3] ? KXL.colIndex(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
};
KXL.rangeStr = (g) => (g.r1 === g.r2 && g.c1 === g.c2) ? KXL.a1(g.r1, g.c1) : KXL.a1(g.r1, g.c1) + ':' + KXL.a1(g.r2, g.c2);
KXL.toast = function (msg, kind) {
  const host = $('#toasts'); const t = KXL.el('div', { class: 'toast ' + (kind || '') }, msg); host.append(t);
  setTimeout(() => t.classList.add('out'), kind === 'err' ? 5000 : 2200); setTimeout(() => t.remove(), kind === 'err' ? 5600 : 2800);
};

/* ------------------------------------------------------------------ state */
const S = KXL.S;
S.cfg = null; S.tpl = null; S.sheetIdx = 0; S.dirty = false; S.savedName = null; S.undo = []; S.redo = []; S.blobs = {}; S.styleSeq = 0; S.styleKey = {};
S.sel = { r1: 1, c1: 1, r2: 1, c2: 1, ar: 1, ac: 1 };
S.clip = null; S.panelTab = 'cell'; S.border = { style: 'thin', color: '#000000' };
S.readOnly = false;

KXL.cur = () => S.tpl.sheets[S.sheetIdx];

/* ---- working model: each sheet gets a Map of cells for O(1) access (serialized back to a list) */
function hydrateSheet(sh) {
  sh.columns = sh.columns || []; sh.rows = sh.rows || []; sh.cells = sh.cells || []; sh.merges = sh.merges || []; sh.bands = sh.bands || [];
  sh.charts = sh.charts || []; sh.sections = sh.sections || []; sh.conditionalFormats = sh.conditionalFormats || []; sh.validations = sh.validations || []; sh.images = sh.images || [];
  sh.page = sh.page || {}; if (sh.showGridLines === undefined) sh.showGridLines = true; if (!sh.zoom) sh.zoom = 100;
  if (!sh.defaultColWidth) sh.defaultColWidth = 12; if (!sh.defaultRowHeight) sh.defaultRowHeight = 18;
  sh._m = new Map(); for (const c of sh.cells) sh._m.set(c.r + ',' + c.c, c);
  sh.cells = undefined;
  sh._colw = new Map(); for (const c of sh.columns) if (c.width) sh._colw.set(c.index, c.width); else if (c.hidden) sh._colw.set(c.index, null);
  sh._colh = new Set(sh.columns.filter(c => c.hidden).map(c => c.index));
  sh._rowh = new Map(); for (const r of sh.rows) if (r.height) sh._rowh.set(r.index, r.height);
  sh._rowhid = new Set(sh.rows.filter(r => r.hidden).map(r => r.index));
  sh.columns = undefined; sh.rows = undefined;
  for (const im of sh.images) { if (im.data && !im.dataRef) { im.dataRef = 'b' + (++S.blobSeq || (S.blobSeq = 1)); S.blobs[im.dataRef] = im.data; } delete im.data; }
}
KXL.fn.hydrateOne = hydrateSheet;
function dehydrateSheet(sh) {
  const o = {};
  for (const k in sh) if (k[0] !== '_') o[k] = sh[k];
  o.cells = Array.from(sh._m.values()).filter(c => c.v || c.s || c.l).sort((a, b) => a.r - b.r || a.c - b.c).map(c => { const x = { r: c.r, c: c.c }; if (c.v) x.v = c.v; if (c.s) x.s = c.s; if (c.t) x.t = c.t; if (c.l) x.l = c.l; if (c.bc && c.v) x.bc = c.bc; return x; });
  o.columns = Array.from(new Set([...sh._colw.keys(), ...sh._colh])).sort((a, b) => a - b).map(i => { const x = { index: i }; if (sh._colw.get(i)) x.width = sh._colw.get(i); if (sh._colh.has(i)) x.hidden = true; return x; });
  o.rows = Array.from(new Set([...sh._rowh.keys(), ...sh._rowhid])).sort((a, b) => a - b).map(i => { const x = { index: i }; if (sh._rowh.get(i)) x.height = sh._rowh.get(i); if (sh._rowhid.has(i)) x.hidden = true; return x; });
  o.images = sh.images.map(im => { const x = Object.assign({}, im); x.data = S.blobs[im.dataRef] || undefined; delete x.dataRef; return x; });
  return o;
}
KXL.serialize = function () {
  const t = Object.assign({}, S.tpl);
  t.sheets = S.tpl.sheets.map(dehydrateSheet);
  // drop unreferenced styles
  const used = new Set(); for (const sh of t.sheets) for (const c of sh.cells) if (c.s) used.add(c.s);
  t.styles = {}; for (const id of used) if (S.tpl.styles[id]) t.styles[id] = S.tpl.styles[id];
  return t;
};
KXL.load = function (tpl, name) {
  tpl.sheets = tpl.sheets && tpl.sheets.length ? tpl.sheets : [{ name: 'Sheet1', rtl: KXL.RTL }];
  tpl.styles = tpl.styles || {};
  S.blobSeq = 0; S.blobs = {};
  for (const sh of tpl.sheets) hydrateSheet(sh);
  S.tpl = tpl; S.sheetIdx = 0; S.savedName = name || null; S.dirty = false; S.undo = []; S.redo = [];
  S.styleKey = {}; S.styleSeq = 0;
  for (const id in tpl.styles) { S.styleKey[styleKeyOf(tpl.styles[id])] = id; const n = parseInt(id.replace(/\D/g, ''), 10); if (n > S.styleSeq) S.styleSeq = n; }
  S.sel = { r1: 1, c1: 1, r2: 1, c2: 1, ar: 1, ac: 1 };
  for (const sh of tpl.sheets) { sh._gr = 0; sh._gc = 0; }
  KXL.fn.afterLoad && KXL.fn.afterLoad();
};

/* ---- styles */
function normStyle(s) {
  const o = {};
  for (const k of Object.keys(s || {}).sort()) {
    let v = s[k];
    if (v === false || v === null || v === undefined || v === '' || (k === 'indent' && !v) || (k === 'rotation' && !v)) continue;
    if (v && typeof v === 'object') { v = normStyle(v); if (!Object.keys(v).length) continue; }
    o[k] = v;
  }
  return o;
}
function styleKeyOf(s) { return JSON.stringify(normStyle(s)); }
KXL.styleOf = (cell) => (cell && cell.s && S.tpl.styles[cell.s]) || {};
KXL.styleId = function (obj) {
  const n = normStyle(obj);
  if (!Object.keys(n).length) return undefined;
  const key = JSON.stringify(n);
  if (S.styleKey[key]) return S.styleKey[key];
  const id = 's' + (++S.styleSeq);
  S.tpl.styles[id] = n; S.styleKey[key] = id;
  return id;
};

/* ---- history */
KXL.snapshot = () => JSON.stringify(KXL.serialize());
KXL.commit = function () {   // call AFTER a mutation; stores the state BEFORE via S.pending
  S.dirty = true; S.redo = [];
  if (S.pendingSnap) { S.undo.push(S.pendingSnap); if (S.undo.length > 80) S.undo.shift(); S.pendingSnap = null; }
  KXL.fn.changed && KXL.fn.changed();
};
KXL.begin = function () { if (!S.pendingSnap) S.pendingSnap = KXL.snapshot(); };   // call BEFORE a mutation
KXL.mutate = function (fn) { KXL.begin(); fn(); KXL.commit(); };
function restore(json) {
  const t = JSON.parse(json); const name = S.savedName; const idx = S.sheetIdx; const sel = S.sel; const undo = S.undo, redo = S.redo;
  const blobs = S.blobs; const sc = document.getElementById('gridscroll'); const st = sc ? sc.scrollTop : 0, sl = sc ? sc.scrollLeft : 0;
  KXL.load(t, name); S.blobs = blobs; S.undo = undo; S.redo = redo; S.dirty = true; S.sheetIdx = Math.min(idx, S.tpl.sheets.length - 1); S.sel = sel;
  for (const sh of t.sheets) for (const im of sh.images) { /* dataRef already assigned in hydrate from data */ }
  KXL.fn.changed && KXL.fn.changed(true);
  if (sc) { sc.scrollTop = st; sc.scrollLeft = sl; }
}
KXL.undoOp = function () { if (!S.undo.length) return; S.redo.push(KXL.snapshot()); restore(S.undo.pop()); };
KXL.redoOp = function () { if (!S.redo.length) return; S.undo.push(KXL.snapshot()); restore(S.redo.pop()); };

/* ---- cells */
KXL.cell = (r, c, sh) => (sh || KXL.cur())._m.get(r + ',' + c);
KXL.ensureCell = function (r, c, sh) {
  sh = sh || KXL.cur(); const k = r + ',' + c; let x = sh._m.get(k);
  if (!x) { x = { r, c }; sh._m.set(k, x); } return x;
};
KXL.setValue = function (r, c, v, t) {
  const sh = KXL.cur(); const k = r + ',' + c; let x = sh._m.get(k);
  if (v === '' || v == null) { if (x) { delete x.v; delete x.t; if (!x.s && !x.l) sh._m.delete(k); } return; }
  if (!x) x = KXL.ensureCell(r, c);
  x.v = v; if (t !== undefined) { if (t) x.t = t; else delete x.t; }
};
KXL.mergeAt = function (r, c, sh) {   // merge covering (r,c): {r1,c1,r2,c2}
  sh = sh || KXL.cur();
  for (const m of sh.merges) { const g = KXL.parseRange(m); if (g && r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) return g; }
  return null;
};
KXL.maxUsed = function (sh) {
  sh = sh || KXL.cur(); let mr = 0, mc = 0;
  for (const x of sh._m.values()) { if (x.v || x.s) { mr = Math.max(mr, x.r); mc = Math.max(mc, x.c); } }
  for (const m of sh.merges) { const g = KXL.parseRange(m); if (g) { mr = Math.max(mr, g.r2); mc = Math.max(mc, g.c2); } }
  for (const b of sh.bands) { if (b.orientation === 'columns') mc = Math.max(mc, b.lastCol || 0); else mr = Math.max(mr, b.lastRow || 0); }
  for (const k of sh._rowh.keys()) mr = Math.max(mr, k); for (const k of sh._colw.keys()) mc = Math.max(mc, k);
  return { r: mr, c: mc };
};
KXL.colWidth = (c, sh) => { sh = sh || KXL.cur(); const w = sh._colw.get(c); return w ? w : sh.defaultColWidth; };
KXL.rowHeight = (r, sh) => { sh = sh || KXL.cur(); const h = sh._rowh.get(r); return h ? h : sh.defaultRowHeight; };

/* ---- structural edits: insert/delete rows/cols, shifting every dependent structure */
function shiftRangeStr(str, axis, at, delta, count) {  // returns new range string or null (removed)
  const g = KXL.parseRange(str); if (!g) return str;
  const a = axis === 'r' ? ['r1', 'r2'] : ['c1', 'c2'];
  if (delta > 0) { if (g[a[0]] >= at) g[a[0]] += delta; if (g[a[1]] >= at) g[a[1]] += delta; }
  else {
    const end = at + count - 1;
    const f = (v, isStart) => v < at ? v : (v > end ? v - count : (isStart ? at : at - 1));
    const n1 = f(g[a[0]], true), n2 = f(g[a[1]], false);
    if (n2 < n1) return null; g[a[0]] = n1; g[a[1]] = n2;
  }
  return KXL.rangeStr(g);
}
function shiftFormulaRefs(f, axis, at, delta, count) { return f; }
KXL.structural = function (axis, at, delta, count) {   // delta>0 insert `delta` before `at`; delta<0 delete `count` starting at `at`
  const sh = KXL.cur(); const isR = axis === 'r';
  const newM = new Map();
  for (const x of sh._m.values()) {
    const k = isR ? x.r : x.c;
    if (delta > 0) { if (k >= at) { if (isR) x.r += delta; else x.c += delta; } newM.set(x.r + ',' + x.c, x); }
    else { const end = at + count - 1; if (k >= at && k <= end) continue; if (k > end) { if (isR) x.r -= count; else x.c -= count; } newM.set(x.r + ',' + x.c, x); }
  }
  sh._m = newM;
  sh.merges = sh.merges.map(m => shiftRangeStr(m, axis, at, delta, count)).filter(Boolean);
  for (const cf of sh.conditionalFormats) cf.range = shiftRangeStr(cf.range, axis, at, delta, count);
  sh.conditionalFormats = sh.conditionalFormats.filter(c => c.range);
  for (const v of sh.validations) v.range = shiftRangeStr(v.range, axis, at, delta, count);
  sh.validations = sh.validations.filter(v => v.range);
  for (const ch of sh.charts) {
    if (ch.categories) ch.categories = shiftRangeStr(ch.categories, axis, at, delta, count) || undefined;
    for (const se of ch.series || []) { if (se.values) se.values = shiftRangeStr(se.values, axis, at, delta, count) || undefined; if (se.nameCell) se.nameCell = shiftRangeStr(se.nameCell, axis, at, delta, count) || undefined; }
    ch.series = (ch.series || []).filter(se => se.values);
  }
  if (sh.autoFilter) sh.autoFilter = shiftRangeStr(sh.autoFilter, axis, at, delta, count) || undefined;
  if (isR) {
    for (const b of sh.bands) {
      if (delta > 0) { if (b.firstRow >= at) b.firstRow += delta; if (b.lastRow >= at) b.lastRow += delta; }
      else { const end = at + count - 1; const f = (v, s) => v < at ? v : (v > end ? v - count : (s ? at : at - 1)); b.firstRow = f(b.firstRow, true); b.lastRow = f(b.lastRow, false); }
    }
    sh.bands = sh.bands.filter(b => b.orientation === 'columns' || b.lastRow >= b.firstRow);
    for (const x of sh.sections) {
      if (delta > 0) { if (x.firstRow >= at) x.firstRow += delta; if (x.lastRow >= at) x.lastRow += delta; }
      else { const end = at + count - 1; const f = (v, st) => v < at ? v : (v > end ? v - count : (st ? at : at - 1)); x.firstRow = f(x.firstRow, true); x.lastRow = f(x.lastRow, false); }
    }
    sh.sections = sh.sections.filter(x => x.lastRow >= x.firstRow);
    for (const im of sh.images) { if (delta > 0 ? im.row >= at : im.row > at + count - 1) im.row += delta > 0 ? delta : -count; }
    for (const ch of sh.charts) { if (delta > 0 ? ch.row >= at : ch.row > at + count - 1) ch.row += delta > 0 ? delta : -count; }
    const nh = new Map(); for (const [k, v] of sh._rowh) { let n = k; if (delta > 0) { if (k >= at) n = k + delta; } else { const end = at + count - 1; if (k >= at && k <= end) continue; if (k > end) n = k - count; } nh.set(n, v); } sh._rowh = nh;
    if (delta > 0 ? at <= sh.freezeRows : at <= sh.freezeRows) sh.freezeRows = Math.max(0, (sh.freezeRows || 0) + (delta > 0 ? delta : -Math.min(count, sh.freezeRows - at + 1)));
  } else {
    for (const b of sh.bands) if (b.orientation === 'columns') {
      if (delta > 0) { if (b.firstCol >= at) b.firstCol += delta; if (b.lastCol >= at) b.lastCol += delta; }
      else { const end = at + count - 1; const f = (v, st) => v < at ? v : (v > end ? v - count : (st ? at : at - 1)); b.firstCol = f(b.firstCol, true); b.lastCol = f(b.lastCol, false); }
    }
    sh.bands = sh.bands.filter(b => b.orientation !== 'columns' || b.lastCol >= b.firstCol);
    for (const im of sh.images) { if (delta > 0 ? im.col >= at : im.col > at + count - 1) im.col += delta > 0 ? delta : -count; }
    for (const ch of sh.charts) { if (delta > 0 ? ch.col >= at : ch.col > at + count - 1) ch.col += delta > 0 ? delta : -count; }
    const nw = new Map(); for (const [k, v] of sh._colw) { let n = k; if (delta > 0) { if (k >= at) n = k + delta; } else { const end = at + count - 1; if (k >= at && k <= end) continue; if (k > end) n = k - count; } nw.set(n, v); } sh._colw = nw;
    const nh = new Set(); for (const k of sh._colh) { if (delta > 0) nh.add(k >= at ? k + delta : k); else { const end = at + count - 1; if (k >= at && k <= end) continue; nh.add(k > end ? k - count : k); } } sh._colh = nh;
  }
  if (sh.freezeCols && !isR && at <= sh.freezeCols) sh.freezeCols = Math.max(0, sh.freezeCols + (delta > 0 ? delta : -Math.min(count, sh.freezeCols - at + 1)));
};

/* ---- model field tree */
KXL.fieldTree = function () {
  const cfg = S.cfg || { models: [] };
  const m = cfg.models.find(x => x.name === S.tpl.modelName);
  const base = m ? (m.fields || []) : inferFields(S.tpl.sampleData);
  const sets = (S.tpl.dataSets || []).filter(d => d.name).map(d => (d.type === 'sql' || d.type === 'json' || d.type === 'provider')
    ? { name: d.name, type: 'list', children: (d.columns || []).map(c => ({ name: c.name, type: c.type || 'text' })) }
    : ({
    name: d.name, type: 'object', children: [
      { name: 'Columns', type: 'list', children: [{ name: 'Key', type: 'text' }, { name: 'Total', type: 'number' }] },
      { name: 'Rows', type: 'list', children: [{ name: 'Key', type: 'text' }, { name: 'Total', type: 'number' }, { name: 'Cells', type: 'list', children: [{ name: 'Value', type: 'number' }, { name: 'Count', type: 'number' }] }] },
      { name: 'GrandTotal', type: 'number' }] }));
  return base.concat(sets);
};
function inferFields(t) {
  const obj = Array.isArray(t) ? t[0] : t;
  if (!obj || typeof obj !== 'object') return [];
  return Object.keys(obj).map(k => {
    const v = obj[k];
    if (Array.isArray(v)) return { name: k, type: 'list', children: inferFields(v) };
    if (v && typeof v === 'object') return { name: k, type: 'object', children: inferFields(v) };
    if (typeof v === 'number') return { name: k, type: 'number' };
    if (typeof v === 'boolean') return { name: k, type: 'bool' };
    if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return { name: k, type: 'date' };
    return { name: k, type: 'text' };
  });
}
KXL.inferFields = inferFields;
})();
