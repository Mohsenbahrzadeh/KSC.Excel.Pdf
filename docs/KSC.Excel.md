<div dir="rtl">

# KSC.Excel — طراح قالب اکسل و تولید فایل اکسل بدون کتابخانه

قالب اکسل را **یک‌بار در یک طراح تصویری** (شبیه Stimulsoft/Excel) می‌کشید و می‌سازید؛ برنامه‌نویس فقط داده را می‌دهد و **فایل `.xlsx` زیبا و کامل** می‌گیرد — بدون درگیر شدن با EPPlus / ClosedXML / OpenXML SDK، بدون نصب Excel روی سرور، و بدون هیچ وابستگی به CDN (مناسب شبکه‌ی داخلی).

```csharp
return this.ExcelFile("Invoice", model, "invoice.xlsx");   // همین!
```

> طراحی و تصمیمات معماری: [KSC.Excel-Design.md](KSC.Excel-Design.md)

## فهرست

1. [ویژگی‌ها](#۱-ویژگی‌ها)
2. [بسته‌ها و سازگاری](#۲-بسته‌ها-و-سازگاری)
3. [شروع سریع — .NET (8/10)](#۳-شروع-سریع--net-810)
4. [شروع سریع — .NET Framework](#۴-شروع-سریع--net-framework-mvc-5)
5. [کار با طراح](#۵-کار-با-طراح)
6. [زبان قالب (داده، فیلتر، نوار تکرار، فرمول)](#۶-زبان-قالب)
7. [امکانات قالب‌بندی فایل](#۷-امکانات-قالب‌بندی-فایل)
8. [API برنامه‌نویس](#۸-api-برنامه‌نویس)
9. [نمایش در صفحه یا دانلود؟](#۹-نمایش-در-صفحه-یا-دانلود)
10. [تنظیمات](#۱۰-تنظیمات)
11. [امنیت و دسترسی](#۱۱-امنیت-و-دسترسی)
12. [استقرار](#۱۲-استقرار)
13. [کارایی و محدودیت‌ها](#۱۳-کارایی-و-محدودیت‌ها)
14. [عیب‌یابی و سؤالات متداول](#۱۴-عیب‌یابی-و-سؤالات-متداول)

---

## ۱) ویژگی‌ها

| حوزه | امکان |
|---|---|
| **طراح** | شبکه‌ی شبیه Excel، کادر/رنگ/فونت/تراز/ادغام، قالب عدد، تغییر اندازه‌ی سطر و ستون، کپی/چسباندن، Undo/Redo، چند برگه، راست‌به‌چپ، بدون نصب (داخل وب‌اپ خودتان) |
| **داده** | کشیدن فیلد روی سلول، نوار تکرار (Band) برای لیست‌ها، نوار تو در تو (master/detail)، تجمیع (`Sum/Count/...`)، فیلتر (تاریخ شمسی، گرد کردن، …) |
| **فایل** | `.xlsx` استاندارد: فرمول زنده، ادغام، فریز، فیلتر خودکار، قالب شرطی (۷ نوع)، اعتبارسنجی داده، تصویر (ثابت و پویا)، لینک، تنظیمات چاپ، راست‌به‌چپ |
| **نمایش** | `ViewComponent` (.NET) و `Html.KscExcelViewer` (.NET Framework) برای دیدن گزارش در صفحه + دکمه‌ی دانلود |
| **پلتفرم** | یک هسته‌ی `netstandard2.0` برای **.NET Framework 4.6.1+ و .NET** + آداپتر ASP.NET Core + آداپتر MVC 5 |
| **کیفیت** | فایل خروجی در **Excel واقعی** باز و بررسی شده؛ تست‌ها روی .NET 10 و .NET Framework 4.8 اجرا می‌شوند |

## ۲) بسته‌ها و سازگاری

| بسته | Target | کاربرد |
|---|---|---|
| `KSC.Excel` | `netstandard2.0` | هسته: مدل قالب، موتور رندر، نویسنده‌ی XLSX، پیش‌نمایش HTML، API طراح، UI طراح (embedded) |
| `KSC.Excel.AspNetCore` | `net8.0`, `net10.0` | `AddKscExcel`, `MapKscExcelDesigner`, `ExcelFile`, `KscExcelViewer` |
| `KSC.Excel.Mvc` | `net472`, `net48` | `KscExcel.Initialize`, `routes.MapKscExcelDesigner`, `ExcelResult`, `Html.KscExcelViewer` |
| `KSC.Excel.Pdf` | `netstandard2.0` | خروجی PDF از همان قالب و داده (SkiaSharp + HarfBuzz؛ شکل‌دهی درست فارسی/عربی و RTL)؛ روی .NET Framework و .NET |
| `KSC.Excel.Designer.Windows` | `net10.0-windows`, `net48` | برنامه‌ی ویندوزی طراح |

تنها وابستگی خارجی: `Newtonsoft.Json` (که در پروژه‌های .NET Framework معمولاً هست).

مرورگرهای طراح: Chrome/Edge/Firefox جدید (ES2018+).
Excel: فایل‌ها در Excel 2016+ (نسخه‌ی دسکتاپ) باز و بررسی شده‌اند. خروجی OOXML استاندارد است و فرمول‌ها مقدار ذخیره‌شده دارند، اما LibreOffice/Excel Online/Google Sheets در این پروژه آزمایش نشده‌اند.

## ۳) شروع سریع — .NET (8/10)

**۱) پکیج و ثبت سرویس**

```csharp
// Program.cs
using KSC.Excel.AspNetCore;

builder.Services.AddControllersWithViews();
builder.Services.AddKscExcel(o =>
{
    o.AddModel<Invoice>("Invoice");      // طراح فیلدهای Invoice را نشان می‌دهد و داده‌ی نمونه می‌سازد
});

var app = builder.Build();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

app.MapKscExcelDesigner("/excel-designer").RequireAuthorization();   // طراح — حتماً پشت لاگین
app.MapKscExcelDownloads();                                           // فقط اگر از Viewer استفاده می‌کنید
app.MapDefaultControllerRoute();
```

**۲) طراحی قالب**: به `/excel-designer` بروید، قالب را بکشید و با نام (مثلاً `Invoice`) ذخیره کنید.
قالب‌ها فایل‌های JSON در `{ContentRoot}/ExcelTemplates/Invoice.kxl.json` هستند (در Git نگه دارید).

**۳) تولید فایل**

```csharp
public class ReportsController : Controller
{
    public IActionResult Invoice(int id)
    {
        Invoice model = _repo.Get(id);
        return this.ExcelFile("Invoice", model, $"invoice-{id}.xlsx");   // دانلود
    }
}
```

**۴) یا نمایش در صفحه** (`Views/.../Invoice.cshtml`):

```cshtml
@model Invoice
@await Component.InvokeAsync("KscExcelViewer", new { template = "Invoice", data = Model, downloadName = "invoice" })
```

نمونه‌ی کامل و اجرایی: `samples/KSC.Excel.Sample.Web` (`dotnet run --project samples/KSC.Excel.Sample.Web` و رفتن به `/excel-designer`؛ در حالت Development بدون لاگین باز می‌شود).

**Minimal API**

```csharp
app.MapGet("/report", (IExcelReportService x) => x.File("Invoice", Invoice.Demo(), "invoice.xlsx"));
```

**بدون وب (سرویس/Job)**

```csharp
var svc = provider.GetRequiredService<IExcelReportService>();
byte[] bytes = svc.Render("Invoice", model);
File.WriteAllBytes("invoice.xlsx", bytes);
// یا مستقیم روی Stream:
svc.Render(svc.GetTemplate("Invoice"), model, outputStream);
```

## ۴) شروع سریع — .NET Framework (MVC 5)

نصب: ارجاع به `KSC.Excel` و `KSC.Excel.Mvc` (Target پروژه‌ی شما ≥ 4.7.2؛ هسته روی 4.6.1 هم کار می‌کند، فقط آداپتر 4.7.2+ است).

**Global.asax.cs**

```csharp
using KSC.Excel.Mvc;

protected void Application_Start()
{
    AreaRegistration.RegisterAllAreas();

    KscExcel.Initialize(o =>
    {
        o.TemplateDirectory = "ExcelTemplates";     // نسبت به ریشه‌ی سایت (~/ExcelTemplates)
        o.AddModel<InvoiceVm>("Invoice");
    });
    KscExcel.Authorize = ctx => ctx.User.IsInRole("Admin");   // اختیاری: قبل از آن باید کاربر لاگین باشد

    RouteConfig.RegisterRoutes(RouteTable.Routes);
}
```

**RouteConfig.cs** (قبل از مسیر پیش‌فرض MVC)

```csharp
routes.MapKscExcelDesigner();          // /excel-designer
routes.MapKscExcelDownloads();         // فقط برای Html.KscExcelViewer
routes.MapRoute("Default", "{controller}/{action}/{id}", ...);
```

**Controller**

```csharp
using KSC.Excel.Mvc;

public ActionResult Invoice(int id)
{
    var vm = _service.GetInvoice(id);
    return this.Excel("Invoice", vm, "invoice.xlsx");     // دانلود
}
```

**View** (نمایش در صفحه + دکمه‌ی دانلود)

```cshtml
@using KSC.Excel.Mvc
@Html.KscExcelViewer("Invoice", Model, "invoice")
```

**وب‌سرور**: در IIS Integrated Pipeline کار می‌کند. اگر مسیر `/excel-designer/assets/*.js` ۴۰۴ داد، در `web.config` مطمئن شوید `UrlRoutingModule` فعال است یا `<modules runAllManagedModulesForAllRequests="true" />` را بگذارید.

> **راستی‌آزمایی:** آداپتر MVC 5 (طراح، Viewer، نمایشگر PDF، دانلود Excel/PDF) روی IIS Express واقعی با .NET Framework 4.8 اجرا و بررسی شده است (`samples/KSC.Excel.Sample.Mvc5`).

## ۴-۱) برنامه‌ی ویندوزی طراح (Desktop)

اگر ترجیح می‌دهید طراح مثل یک نرم‌افزار نصبی و در پنجره‌ی مستقل باز شود (بدون مرورگر، بدون وب‌سایت، بدون پورت):

```text
dotnet publish src/KSC.Excel.Designer.Windows -c Release -r win-x64 --self-contained false -o artifacts/KSC.Excel.Designer
artifacts\KSC.Excel.Designer\KSC.Excel.Designer.exe          (یا:  KSC.Excel.Designer.exe "C:\path\Invoice.kxl.json")
```

- پیش‌نیاز: **.NET 10 Desktop Runtime** و **WebView2 Runtime** (روی ویندوز ۱۱ هست). ویندوز ۱۰ یا بالاتر.
- همان طراح وب داخل `WebView2` اجرا می‌شود و API آن **داخل همان پروسه** جواب می‌دهد (`https://kxl.local`، بدون Kestrel و بدون پورت)؛ پس ظاهر، امکانات و **فرمت قالب دقیقاً یکی است** و قالب‌ها بین دسکتاپ و برنامه‌ی وب شما جابه‌جا می‌شوند.
- قالب‌ها پیش‌فرض در `Documents\KSC Excel Templates` ذخیره می‌شوند؛ از منوی **فایل ← انتخاب پوشه‌ی قالب‌ها** می‌توانید پوشه‌ی پروژه‌ی خودتان (مثلاً `ExcelTemplates` داخل Git) را بدهید. تنظیم در `%APPDATA%\KSC.Excel.Designer\settings.json` می‌ماند.
- منوها: قالب جدید، **باز کردن فایل `.kxl.json`**، ذخیره، ذخیره به‌عنوان فایل، **خروجی Excel و PDF با داده‌ی نمونه** (همان‌جا فایل را می‌سازد و باز می‌کند)، انتخاب/نمایش پوشه، بارگذاری مجدد، DevTools.
- محدودیت فعلی: مدل‌های داده‌ی کلاس‌های برنامه‌ی شما (`AddModel<T>`) و اتصال SQL در دسکتاپ ثبت نمی‌شوند؛ داده‌ی نمونه (JSON) را در تب «داده‌ی نمونه» بدهید. (برای مدل و SQL واقعی از طراح داخل برنامه‌ی وب خودتان استفاده کنید.)
- بررسی خودکار: `KSC.Excel.Designer.exe --selftest out.txt` طراح را بالا می‌آورد، API درون‌پروسه را صدا می‌زند و نتیجه را می‌نویسد.

## ۵) کار با طراح

### نمای کلی

```
┌─ نوار بالا ─ نام قالب · جدید · باز کردن · Undo/Redo · ذخیره · پیش‌نمایش · دریافت xlsx · JSON ──┐
├─ ریبون: فونت · B I U S · رنگ · کادر · تراز · ادغام · قالب عدد · ساختار · نوار · تصویر · قالب شرطی ┤
├─ نوار فرمول:  [آدرس]  fx  [مقدار / فرمول]                                                      ┤
│ فیلدها / داده‌ی نمونه │        شبکه‌ی سلول‌ها (شبیه Excel)         │ سلول · برگه · نوارها · قوانین · چاپ · تصاویر │
├────────────────────────────── برگه‌ها ───────────────────────────────────────────────────────┘
```

### روند پیشنهادی (۵ دقیقه)

1. **جدید** ← یک قالب شروع (جدولی / فاکتور / خالی راست‌به‌چپ).
2. تب **داده‌ی نمونه**: JSON نمونه‌ی داده‌ای را که به قالب می‌دهید بچسبانید (یا در کد `AddModel<T>()` بزنید تا خودکار ساخته شود).
3. از تب **فیلدها**، فیلد را **روی سلول بکشید** (یا دوبار کلیک). اگر فیلد عضو یک لیست باشد، **نوار تکرار** همان سطر خودکار ساخته می‌شود.
4. ظاهر را بسازید: رنگ، کادر، قالب عدد، ادغام، عرض ستون (کشیدن لبه‌ی سرستون).
5. **پیش‌نمایش** با داده‌ی نمونه (و هشدارهای قالب) ← **ذخیره**.

### میان‌برها

| کلید | عمل | کلید | عمل |
|---|---|---|---|
| `Ctrl+S` | ذخیره | `F2` / `Enter` | ویرایش سلول |
| `Ctrl+Z` / `Ctrl+Y` | Undo / Redo | `Alt+Enter` | سطر جدید داخل سلول |
| `Ctrl+B / I / U` | ضخیم / مورب / زیرخط | `Tab` / `Shift+Tab` | سلول بعدی / قبلی |
| `Ctrl+C / X / V` | کپی / برش / چسباندن (از/به Excel هم) | `Delete` | پاک کردن محتوا |
| `Ctrl+A` | انتخاب همه | `Ctrl+فلش` | پرش به لبه‌ی داده |
| `Shift+کلیک/فلش` | گسترش انتخاب | کلیک راست | منوی سلول/سطر/ستون |

در برگه‌ی راست‌به‌چپ، فلش راست = ستون قبلی (مثل Excel).

### ورود از فایل Excel موجود (Import)

دکمه‌ی **xlsx** در نوار بالا (یا منوی برنامه‌ی ویندوزی: «ورود از فایل Excel») یک فایل `.xlsx` را به قالب تبدیل می‌کند: برگه‌ها، مقدارها، **فرمول‌ها**، فونت/رنگ (شامل رنگ‌های تم Excel)/کادر/تراز/قالب عدد، ادغام، عرض ستون و ارتفاع سطر، فریز، فیلتر، راست‌به‌چپ، قالب شرطی (cellIs، فرمول، شامل متن، تکراری، Top/Bottom، طیف رنگ، نوار داده)، اعتبارسنجی، لینک، **تصویر**، تنظیمات چاپ و تکرار سطر عنوان. بعد در طراح، سلول‌هایی که باید از داده بیایند را با `{{Field}}` عوض کنید و نوار تکرار بگذارید. **وارد نمی‌شود** (هشدار می‌دهد): نمودار، PivotTable، جدول Excel (به‌صورت سلول ساده می‌آید)، ماکرو؛ فرمول مشترک (shared) با مقدار فعلی می‌آید. ایمنی: XML با `DtdProcessing=Prohibit` (بدون XXE)، سقف ۵۰ مگابایت و ۵۰۰هزار سلول. در کد: `XlsxImporter.Import(bytes, "Name").Template`.

### بخش‌های گزارش و ویزارد گروه‌بندی (تب «نوارها»)

- **بخش‌ها** (برچسب روی سطرها مثل Stimulsoft): **سرگزارش**، **سرصفحه** و **پاگزارش**. سطرهای «سرصفحه» در چاپ و PDF **روی هر صفحه تکرار** می‌شوند (معادل «تکرار سطرها»، مگر خودتان `PrintTitleRows` را داده باشید)؛ سرگزارش/پاگزارش فعلاً برچسب بصری‌اند. برچسب رنگی کنار شماره‌ی سطر نشان داده می‌شود.
- **ویزارد گروه‌بندی** ✨: یک لیست و فیلد گروه‌بندی و ستون‌ها را انتخاب کنید ← چهار سطر می‌سازد (عنوان ستون‌ها، سرگروه `{{Key}} ({{Count}})`، سطر داده، **جمع هر گروه** با `=SUM([Detail.X])`) و نوار گروه‌بندی (مرتب‌شده) و نوار داده (با رنگ ردیف زوج) را خودکار تعریف می‌کند. قالب عددی `#,##0` است؛ هر چه خواستید عوض کنید.

### دکمه‌ی PDF در طراح

با `o.UsePdf()` (بسته‌ی `KSC.Excel.Pdf`) دکمه‌ی **PDF** کنار «دریافت xlsx» ظاهر می‌شود و خروجی را با داده‌ی نمونه می‌سازد (در برنامه‌ی ویندوزی خودکار فعال است).

### نوارها (Bands)

نوار = «این سطرها را برای هر آیتم لیست تکرار کن». در تب **نوارها** مسیر مجموعه (`Lines`)، نام، محدوده‌ی سطر، رنگ ردیف‌های زوج (zebra) و «اگر خالی بود» را تنظیم می‌کنید. نوارها با رنگ در کنار شماره‌ی سطر نشان داده می‌شوند.

### سایر پنل‌ها

- **برگه**: نام، راست‌به‌چپ، خطوط شبکه، رنگ تب، بزرگ‌نمایی، عرض/ارتفاع پیش‌فرض، `تکرار برگه برای هر آیتم`، فریز، فیلتر خودکار، محافظت.
- **قوانین**: قالب شرطی (مقدار سلول، فرمول، شامل متن، تکراری، Top/Bottom، طیف رنگ، نوار داده) و اعتبارسنجی (فهرست، عدد، تاریخ، طول، فرمول).
- **چاپ**: جهت، کاغذ، جا دادن عرض در یک صفحه، حاشیه‌ها، سرصفحه/پاصفحه (`&P` صفحه، `&N` تعداد، `&D` تاریخ)، تکرار سطرها در هر صفحه، محدوده‌ی چاپ.
- **تصاویر و نمودار**: نمودار با نوع، عنوان، محدوده‌ی دسته‌ها و سری‌ها (محدوده‌ای که نوار را بپوشاند با داده بزرگ می‌شود؛ اگر محدوده نوار تو در تو را هم بپوشاند، سطرهای سرگروه/جمع هم داخل نمودار می‌آیند — برای نمودار خلاصه از ستون/محدوده‌ی جدا استفاده کنید)؛ در طراح نمودار فقط به‌صورت کادر نشان داده می‌شود و در Excel رسم می‌شود. تصویر: لوگوی ثابت یا **مسیر داده** (`Company.Logo` از نوع `byte[]` یا base64).

### ذخیره‌ی خودکار

پیش‌نویس هر تغییر در مرورگر نگه داشته می‌شود؛ اگر صفحه بسته شود، دفعه‌ی بعد «بازیابی؟» می‌پرسد. ذخیره‌ی واقعی فقط با **ذخیره** روی سرور انجام می‌شود. از **خروجی JSON / ورود JSON** برای جابه‌جایی بین محیط‌ها (dev → prod) استفاده کنید.

## ۶) زبان قالب

### ۶-۱ متن، عدد، مقدار

- متن ثابت: همان‌طور که تایپ می‌کنید. عدد ثابت (`123`) عدد است؛ `0912…` (صفر ابتدایی) و رشته‌های بیش از ۱۵ رقم **متن** می‌مانند. برای متن اجباری: پنل سلول ← «همیشه متن».
- داده: `{{مسیر}}`، مثل `{{Customer.Name}}`، `{{Lines[0].Product}}`. بزرگی/کوچکی حروف مهم نیست.
- اگر کل سلول **فقط یک** `{{…}}` باشد، **نوع** حفظ می‌شود (عدد، تاریخ، منطقی)؛ ترکیب با متن، رشته می‌شود: `مشتری: {{Customer.Name}}`.
- مقدار `null`/نبودن مسیر = سلول خالی (هشداری نمی‌دهد؛ برای مقدار جایگزین از `|default:` استفاده کنید).

### ۶-۲ متغیرهای ویژه

| متغیر | معنی |
|---|---|
| `{{$index}}` | شماره‌ی ردیف (از ۱) داخل نوار |
| `{{$count}}` | تعداد آیتم‌های نوار |
| `{{$today}}` / `{{$now}}` | تاریخ امروز / اکنون (سرور) |
| `{{$root}}` / `{{$parent}}` / `{{$item}}` | داده‌ی ریشه / آیتم والد / آیتم جاری |

داخل نوار، اگر نامی روی آیتم جاری نبود، به آیتم والد و در نهایت ریشه نگاه می‌شود؛ پس `{{Title}}` داخل نوار هم کار می‌کند.

### ۶-۳ فیلترها

`{{مسیر | فیلتر1 | فیلتر2:آرگومان}}`

| فیلتر | کار | مثال → نتیجه |
|---|---|---|
| `fa` (یا `persian`, `jalali`) | تاریخ شمسی (متن) | `{{Date\|fa}}` → `1404/01/01` ؛ `{{Date\|fa:yyyy/MM/dd HH:mm}}` |
| `fadigits` (`fanum`) | ارقام فارسی | `{{N\|fanum}}` → `۱۲۳۴` |
| `format:الگو` | قالب عدد/تاریخ .NET | `{{Date\|format:yyyy-MM-dd}}` |
| `round:n` | گرد کردن | `{{P\|round:1}}` |
| `upper` `lower` `trim` | حروف / فاصله | |
| `default:متن` | جایگزین مقدار خالی | `{{X\|default:-}}` |
| `yesno:بله/خیر` | منطقی به متن | `{{Active\|yesno:بله/خیر}}` |
| `join:، ` | اتصال عناصر لیست | `{{Tags\|join:, }}` |
| `truncate:n` | کوتاه کردن با … | |
| `text` | تبدیل به متن | |

> تاریخ در Excel بدون فیلتر، **تاریخ میلادیِ واقعی** است (قابل مرتب‌سازی/فیلتر) با قالب `yyyy/mm/dd`؛ `|fa` متن شمسی می‌دهد (برای نمایش). اگر هم‌زمان هر دو لازم است دو ستون بگذارید.

### ۶-۴ تجمیع داخل متن

`{{Count(Lines)}}`، `{{Sum(Lines.Qty)}}`، `{{Avg(...)}}`، `{{Min(...)}}`، `{{Max(...)}}` — روی کل لیست، مقدار ثابت (نه فرمول Excel). برای جمع **زنده** از فرمول (بخش بعد) استفاده کنید.

### ۶-۵ نوارهای تکرار

| سناریو | قالب |
|---|---|
| لیست ساده | سطر ۹ شامل `{{Product}}`/`{{Qty}}`… و نوار `Lines` روی سطر ۹ ← به تعداد آیتم‌ها تکثیر می‌شود |
| چند سطر در هر آیتم | نوار روی چند سطر (مثلاً ۹ تا ۱۰) ← هر آیتم دو سطر |
| **تو در تو** | نوار `Groups` روی سطرهای ۱–۳ و نوار `Lines` روی سطر ۲ ← `Lines` نسبت به آیتم `Group` خوانده می‌شود (مسیر `Lines`، نه `Groups.Lines`) |
| لیست خالی | به‌صورت پیش‌فرض سطرها حذف می‌شوند؛ «اگر خالی بود یک سطر بماند» یک بلوک خالی نگه می‌دارد |
| ردیف‌های زوج | رنگ زوج (zebra): فقط سلول‌های بدون پس‌زمینه‌ی صریح |

قواعد: نوارها نباید جزئی هم‌پوشان باشند (یا کامل داخل هم یا جدا)؛ ادغام سلولی که از لبه‌ی نوار رد شود نادیده گرفته و هشدار داده می‌شود.

### مرتب‌سازی و گروه‌بندی با جمع هر گروه

روی نوار دو گزینه هست: **مرتب‌سازی بر اساس** (`SortBy`, `SortDesc`) و **گروه‌بندی بر اساس** (`GroupBy`). با گروه‌بندی، هر تکرار نوار یک **گروه** است با فیلدهای `Key` (مقدار گروه)، `Count` (تعداد) و `Items` (اعضا). الگو:

| سطر | محتوا | نوار |
|---|---|---|
| ۱ | `{{Key}} ({{Count}})` — سرگروه | نوار `Groups` (مسیر `Rows`، گروه‌بندی `Category`) روی سطرهای ۱ تا ۳ |
| ۲ | `{{Item}}` · `{{Qty}}` — اعضا | نوار `Detail` (مسیر `Items`) روی سطر ۲ |
| ۳ | `جمع` · `=SUM([Detail.Qty])` — جمع هر گروه | (جزو نوار `Groups`) |

هر گروه جمع خودش را دارد؛ جمع کل را با `{{Sum(Rows.Qty)}}` (مقدار ثابت) بنویسید. در کد: `new BandDef { Source = "Rows", GroupBy = "Category", SortBy = "Item", ... }`.

### جدول متقاطع (Cross-tab / Pivot) — ستون‌های پویا

برای گزارش «منطقه × محصول» که تعداد ستون‌ها از داده می‌آید:

1. در طراح، تب چپ **مجموعه و پارامتر** ← «افزودن جدول متقاطع»: نام `Matrix`، مسیر منبع (`Sales`)، فیلد سطر (`Region`)، فیلد ستون (`Product`)، فیلد مقدار (`Amt`) و تجمیع (`sum/count/avg/min/max`). خروجی: `Matrix.Columns` (`Key`, `Total`)، `Matrix.Rows` (`Key`, `Total`, `Cells[].Value`)، `Matrix.GrandTotal`.
2. نوار **ردیف** `Rows` (مسیر `Matrix.Rows`) روی سطر داده.
3. نوار **ستونی** `Cols` (مسیر `Matrix.Columns`، جهت «ستون‌ها»، ستون ۲ تا ۲): ستون‌های انتخابی به تعداد محصول‌ها تکثیر می‌شوند.

| سلول | محتوا |
|---|---|
| سرستون (داخل نوار ستونی) | `{{$item.Key}}` ← نام محصول |
| ردیف داده (داخل نوار ستونی) | `{{Cells[$col].Value}}` ← مقدار همان منطقه/محصول (`$col` = شماره‌ی کپی) |
| جمع هر ردیف | `=SUM([Cols.Value])` ← بازه‌ی افقی همان سطر |
| جمع هر ستون (سطر جمع، داخل نوار ستونی) | `=SUM([Rows.Value])` ← هر کپی، ستون خودش |
| جمع کل | `=SUM([Rows:C])` (ستون `C` قالب = ستون «جمع») |

ستون‌های بعد از نوار ستونی خودکار به راست/چپ می‌روند؛ عرض ستون، ادغام (داخل یک کپی)، قالب شرطی، اعتبارسنجی، نمودار و فیلتر روی ستون‌های جدید نگاشت می‌شوند. ارجاع‌های ثابت `A1` در فرمول جابه‌جا **نمی‌شوند**؛ از توکن‌ها استفاده کنید. در کد: `Pivot.Build(data, "Region", "Product", "Amt", "sum")` (برای داده‌ای که از پیش آماده کرده‌اید، و آن را زیر هر نامی به قالب بدهید).

### پارامترها

تب **مجموعه و پارامتر** ← «افزودن پارامتر»: نام، نوع (`text/number/date/bool`)، پیش‌فرض. در قالب: `{{$p.Title}}`. در کد مقدار بدهید (نام بدون حساسیت به حروف):

```csharp
svc.Render("Sales", data, new Dictionary<string, object> { ["Title"] = "فروش بهار", ["From"] = DateTime.Today });
return this.ExcelFile("Sales", data, new Dictionary<string, object> { ["Title"] = "…" }, "sales.xlsx");   // Core
return this.Excel("Sales", data, new Dictionary<string, object> { ["Title"] = "…" }, "sales.xlsx");      // Framework
```

### فیلد محاسباتی (Expression) داخل `{{ }}`

هر چیزی که فقط «مسیر» نباشد یک عبارت است و روی داده‌ی همان سطر حساب می‌شود:

```text
{{Qty * UnitPrice}}                    {{ROUND(Price * 1.09, 0)}}
{{IF(Total > 1000, 'VIP', 'عادی')}}    {{FirstName & ' ' & LastName}}
{{COALESCE(Nick, Name)}}               {{Sum(Items.Qty) + 1}}
{{FA(Date, 'yyyy/MM/dd')}}             {{Tax(Price)}}        ← تابع سفارشی
```

عملگرها `+ - * / % ^ & = <> < > <= >=` و توابع `IF IFERROR COALESCE ROUND ABS MIN MAX SUM AVG LEN UPPER LOWER TRIM CONCAT AND OR NOT TODAY NOW YEAR MONTH DAY FA`. مقدار گم‌شده در جمع/ضرب صفر حساب می‌شود، تقسیم بر صفر یا عدد نامعتبر = سلول خالی. فیلترها بعدش کار می‌کنند: `{{Qty * Price | round:0}}`. نتیجه‌ی عددی، عدد واقعی در Excel است (نه فرمول). برای مقدار **زنده** که کاربر در Excel تغییر دهد از فرمول (`=C9*D9`) استفاده کنید.

### فیلتر و تابع سفارشی (کد)

```csharp
builder.Services.AddKscExcel(o =>
{
    o.AddFunction("Tax", a => Convert.ToDecimal(a[0]) * 0.09m);                 // {{Tax(Price)}}  و  =TAX(B2) (مقدار ذخیره‌شده)
    o.AddFunction("SumSq", a => ((object[])a[0]).Sum(v => Math.Pow(Convert.ToDouble(v), 2)));   // بازه به‌صورت object[]
    o.AddFilter("stars", (value, arg) => new string('*', Convert.ToInt32(value)));              // {{Rating|stars}}
});
```

- تابع سفارشی در **عبارت‌ها** و برای **مقدار ذخیره‌شده‌ی فرمول‌های Excel** (`=TAX(B3)`) کار می‌کند؛ متن فرمول دست‌نخورده در فایل می‌ماند و Excel خودش آن را فقط اگر تابع در Excel تعریف شده باشد (مثلاً افزونه) حساب می‌کند؛ در غیر این‌صورت `#NAME?` — برای فرمول‌های Excel فقط از توابع داخلی Excel استفاده کنید و سفارشی‌ها را در `{{ }}`.
- تابع شما ورودی‌ها را به‌صورت `object[]` می‌گیرد؛ خروجی عدد/متن/تاریخ/منطقی.

### منبع داده‌ی SQL

قالب می‌تواند خودش داده را از SQL بگیرد (برای گزارش‌های مستقل از کد):

1. **در برنامه** (یک‌بار، به‌صورت صریح):
   ```csharp
   o.AllowSql = true;                                                  // پیش‌فرض false
   o.AddConnection("Sales", SqlClientFactory.Instance, connectionString);   // Microsoft.Data.SqlClient؛ هر DbProviderFactory
   o.SqlCommandTimeoutSeconds = 30;  o.SqlMaxRows = 100000;
   ```
2. **در طراح**: تب چپ «مجموعه و پارامتر» ← «افزودن کوئری SQL»: نام (`Orders`)، اتصال، متن `SELECT`. `@Min` از **پارامتر** هم‌نام (`{{$p.Min}}`) بسته می‌شود (Parameter binding؛ هیچ‌وقت الحاق رشته). «آزمایش کوئری» ستون‌ها را پیدا و در درخت فیلد می‌گذارد.
3. نوار تکرار با مسیر `Orders` بسازید؛ یا جدول متقاطع روی آن (`Source = Orders`).

```csharp
svc.Render("Sales", null, new Dictionary<string, object> { ["Min"] = 100, ["From"] = DateTime.Today.AddDays(-30) });   // داده‌ی ورودی لازم نیست
```

**ایمنی:** فقط یک دستور `SELECT`/`WITH`؛ `INSERT/UPDATE/DELETE/DROP/EXEC/INTO/;` و … رد می‌شوند (حتی داخل توضیح/رشته مشکلی ندارند). این فقط لایه‌ی دوم است — **حتماً از یک لاگین فقط‌خواندنی** در رشته‌ی اتصال استفاده کنید، چون هر کس بتواند قالب ذخیره کند، می‌تواند روی اتصال‌های ثبت‌شده SELECT اجرا کند. اگر `AllowSql=false` و قالب SQL داشته باشد، تولید با خطای روشن متوقف می‌شود؛ **پیش‌نمایش طراح** در این حالت کوئری را اجرا نمی‌کند و از `داده‌ی نمونه` (عضو هم‌نام) استفاده می‌کند. نتیجه‌ی بیش از `SqlMaxRows` ردیف خطا می‌دهد.

### ۶-۶ فرمول‌ها

هر سلولی که با `=` شروع شود فرمول Excel است و **زنده** می‌ماند. علاوه بر ارجاع‌های معمول (`A1`, `$B$2:$B$9`)، سه «توکن» مخصوص قالب هست که هنگام تولید به آدرس واقعی تبدیل می‌شوند:

| توکن | معنی | مثال |
|---|---|---|
| `[Field]` | سلول همان فیلد در **همان سطر** نوار | `=[Qty]*[UnitPrice]` ← `=C9*D9` |
| `[Band.Field]` | **کل ستون** سلولی که `{{Field}}` در نوار `Band` دارد | `=SUM([Lines.Qty])` ← `=SUM($C$9:$C$11)` |
| `[Band:E]` | کل ستون `E` در نوار `Band` (برای ستون‌های فرمولی که فیلد ندارند) | `=SUM([Lines:E])` |

همچنین `{{…}}` داخل فرمول با مقدارش جایگزین می‌شود: `=C9*{{TaxRate}}`.
اگر نوار خالی باشد، توکن‌ها `0` می‌شوند (فرمول خطا نمی‌دهد).

**مقدار ذخیره‌شده:** موتور جواب فرمول‌های رایج را محاسبه و در فایل می‌گذارد تا نمایشگرهایی که محاسبه نمی‌کنند (موبایل، پیش‌نمایش، pandas/openpyxl) هم عدد ببینند؛ Excel هنگام باز شدن همه را دوباره حساب می‌کند.
توابع پشتیبانی‌شده برای مقدار ذخیره‌شده: `SUM AVERAGE MIN MAX COUNT COUNTA ABS ROUND IF IFERROR AND OR NOT CONCATENATE CONCAT LEN UPPER LOWER TRIM SUMIF COUNTIF` و عملگرهای `+ - * / ^ & % = <> < > <= >=`.
هر تابع دیگر (مثل `VLOOKUP`) **در فایل نوشته می‌شود و در Excel درست کار می‌کند**؛ فقط مقدار ذخیره‌شده ندارد.

### ۶-۷ تکرار برگه

پنل برگه ← «تکرار برگه برای هر آیتم» = `Branches` ← برای هر شعبه یک برگه. نام برگه می‌تواند `{{Name}}` داشته باشد؛ نام‌ها یکتا و مطابق قوانین Excel (۳۱ نویسه، بدون `[]:*?/\`) ساخته می‌شوند.

### ۶-۸ لینک

پنل سلول ← «پیوند»: `https://site/{{Id}}` یا `#Sheet2!A1` (داخل فایل). در پیش‌نمایش فقط http/https/mailto فعال است.

## ۷) امکانات قالب‌بندی فایل

| دسته | پشتیبانی |
|---|---|
| فونت | نام، اندازه، ضخیم/مورب/زیرخط/خط‌خورده، رنگ (فونت باید روی سیستم کاربر باشد؛ پیش‌فرض Tahoma) |
| سلول | رنگ پس‌زمینه، کادر چهار طرف (thin/medium/thick/dashed/dotted/double/hair + رنگ)، ادغام |
| تراز | افقی/عمودی، شکستن متن، تورفتگی، چرخش، shrink، جهت خواندن راست‌به‌چپ |
| عدد | `General`، `0.00`، `#,##0`، `0%`، تاریخ/ساعت، و هر قالب سفارشی Excel |
| ساختار | عرض ستون، ارتفاع سطر، فریز سطر/ستون، فیلتر خودکار، محافظت برگه (بدون رمز) |
| قالب شرطی | `cellIs`, `expression`, `containsText`, `duplicates`, `top10`, `colorScale`, `dataBar` |
| اعتبارسنجی | `list`, `whole`, `decimal`, `date`, `textLength`, `custom` + پیام خطا/راهنما |
| تصویر | PNG/JPEG/GIF؛ ثابت یا از داده |
| **نمودار** | ستونی، میله‌ای، خطی، ناحیه‌ای، دایره‌ای؛ چند سری، انباشته، برچسب مقدار؛ **محدوده‌ها با داده بزرگ می‌شوند** (مثلاً `D9:D9` روی نوار → همه‌ی سطرهای تولیدشده) |
| چاپ | جهت، کاغذ، مقیاس/جا دادن عرض، حاشیه، سر/پاصفحه، تکرار سطر، محدوده‌ی چاپ |
| برگه | راست‌به‌چپ، رنگ تب، زوم، خطوط شبکه |

در حال حاضر **پشتیبانی نمی‌شود** (نقشه‌ی راه): زیرگزارش به‌صورت قالب جدا (نوار تو در تو + چند مجموعه‌ی داده جایگزین است)، نمودار ترکیبی/پراکندگی، جدول‌های ساختاری Excel (ListObject)، PivotTable، ماکرو، کامنت سلول، رمزگذاری. منابع داده‌ی REST/JSON/کد (`AddHttpEndpoint` / `AddProvider`)، بارکد/QR، واترمارک و شکست صفحه اکنون پشتیبانی می‌شوند. جزئیات: [KSC.Excel-Design.md §۱۶](KSC.Excel-Design.md).

## ۸) API برنامه‌نویس

### `IExcelReportService`

```csharp
byte[]  Render(string templateName, object data);
byte[]  Render(ExcelTemplate template, object data);
void    Render(ExcelTemplate template, object data, Stream output);
string  RenderHtml(string templateName, object data, bool showHeaders = false);
RenderedWorkbook Bind(ExcelTemplate template, object data);   // هشدارها (Warnings) و دیباگ
ExcelTemplate GetTemplate(string name);
ITemplateStore Store { get; }
```

`data` می‌تواند هر یک از این‌ها باشد: **کلاس/DTO**، `Dictionary<string, object>`، `ExpandoObject`، **`JObject/JToken`**، `DataTable`/`DataRow`/`DataView`، یا ترکیبی (مثلاً کلاس که یک `DataTable` دارد). نام‌ها case-insensitive هستند.

### کد به‌جای طراح (Code-first)

```csharp
var b = new TemplateBuilder("Report", "گزارش");
var head = b.Style("h", new StyleDef { Bold = true, Fill = "#2E75B6", Color = "#FFFFFF", H = "center" });
b.Sheet("گزارش", s =>
{
    s.Cell("A1", "نام", head).Cell("B1", "مبلغ", head);
    s.Cell("A2", "{{Name}}").Cell("B2", "{{Amount}}");
    s.Band("Rows", "Rows", 2, 2, "#F5F9FD");
    s.Cell("A3", "جمع").Cell("B3", "=SUM([Rows.Amount])");
    s.Freeze(1);
}, rtl: true);
byte[] xlsx = svc.Render(b.Build(), new { Rows = data });
```

مناسب برای قالب‌های پیش‌فرضی که همراه برنامه می‌آیند: `store.Save(template)` در شروع برنامه.

### `ITemplateStore` (ذخیره‌سازی دلخواه)

پیش‌فرض: پوشه‌ی `ExcelTemplates` (`*.kxl.json`). برای دیتابیس:

```csharp
public sealed class DbTemplateStore : ITemplateStore
{
    public IList<TemplateInfo> List() => /* SELECT Name, Title, Modified */;
    public ExcelTemplate Get(string name) { var json = /* SELECT Json */; return json == null ? null : ExcelTemplate.FromJson(json); }
    public void Save(ExcelTemplate t) { /* UPSERT t.Name, t.ToJson() */ }
    public bool Delete(string name) => /* DELETE */;
}
builder.Services.AddKscExcel(o => o.Store = new DbTemplateStore(...));
```

نام قالب را همیشه با `TemplateNames.IsValid(name)` اعتبارسنجی کنید.

### `ViewComponent` / Html helper

| پارامتر | پیش‌فرض | توضیح |
|---|---|---|
| `template` | — | نام قالب |
| `data` | — | داده |
| `downloadName` | نام قالب | نام فایل دانلود |
| `showDownload` | `true` | دکمه‌ی دانلود |
| `showHeaders` | `false` | سرستون A,B,C و شماره‌ی سطر مثل Excel |
| `downloadLabel` | «دریافت فایل اکسل» | متن دکمه |

لینک دانلود **یک‌بار مصرف‌نما** (۱۵ دقیقه اعتبار، توکن ۱۴۴ بیتی، متصل به کاربر لاگین) است و فایل از حافظه‌ی سرور داده می‌شود؛ حداکثر ۶۴ فایل هم‌زمان نگه داشته می‌شود. در .NET برای استفاده از `<vc:ksc-excel-viewer>` باید `@addTagHelper *, KSC.Excel.AspNetCore` را به `_ViewImports` اضافه کنید.

### JSON قالب

قالب یک سند JSON ساده است (`schemaVersion`, `sheets[].cells[]`, `styles{}`, …). می‌توان آن را در Git نگه داشت، diff کرد و با `ExcelTemplate.FromJson / ToJson / LoadFile / SaveFile` خواند و نوشت. شِمای کامل: [KSC.Excel-Design.md §۵](KSC.Excel-Design.md).

### گزارش‌های بسیار بزرگ — نویسنده‌ی Streaming

برای صدها هزار تا میلیون‌ها سطر که در حافظه جا نمی‌شوند:

```csharp
// داده را «تنبل» بدهید (IEnumerable با yield، DataReader، EF AsEnumerable) تا حافظه تا آخر ثابت بماند
IEnumerable<Row> rows = repo.StreamRows();
bool streamed = svc.RenderStreaming("BigReport", new { Title = "…", Items = rows }, outputStream);
```

- **کنترلر .NET:** `return this.ExcelFileStreaming("BigReport", new { Items = rows }, "big.xlsx");` (داخل یک فایل موقت می‌نویسد و ارسال می‌کند). **.NET Framework:** `return this.ExcelStreaming("BigReport", data, "big.xlsx");`
- اندازه‌گیری: **۱٬۰۰۰٬۰۰۰ سطر × ۶ ستون ≈ ۱۲٫۶ ثانیه، ≈ ۴۷ مگابایت حافظه‌ی اوج** (فایل ۳۰ مگابایت)؛ `MaxCells` در این حالت اعمال نمی‌شود (فقط `MaxRows`).
- SQL: داده‌ست SQL در این حالت **خط‌به‌خط خوانده می‌شود** (اتصال تا پایان گزارش باز می‌ماند).
- اگر قالب «قابل Streaming» نباشد، `RenderStreaming` خودش به رندر عادی برمی‌گردد و `false` می‌دهد؛ دلیل‌ها: `StreamAnalyzer.Check(template)`. شروط: جمع/فرمول هر نوار فقط **زیر** نوار باشد (نه بالا یا داخل آن)، `[Field]` به سلول سطر بعدی ارجاع ندهد، `{{$count}}` نباشد، سطرهای فریزشده بالای همه‌ی نوارها باشند. مرتب‌سازی/گروه‌بندی همان نوار را در حافظه نگه می‌دارد (فقط همان نوار).
- تفاوت‌ها با خروجی عادی: متن‌ها `inlineStr` و فرمول‌ها **بدون مقدار ذخیره‌شده** (Excel هنگام باز شدن حساب می‌کند؛ نمایشگرهایی که حساب نمی‌کنند صفر/خالی می‌بینند). اگر هنگام نوشتن خطا بیاید، فایل ناقص در `output` می‌ماند — در فایل موقت بنویسید (آداپترها همین کار را می‌کنند).

## ۹) خروجی PDF

پکیج `KSC.Excel.Pdf` همان قالب و داده را PDF می‌کند؛ نیازی به Excel نیست:

```csharp
using KSC.Excel.Pdf;
byte[] pdf = svc.RenderPdf("Invoice", model);                       // یا با پارامتر: svc.RenderPdf("Sales", data, parameters)
return File(pdf, "application/pdf", "invoice.pdf");

ExcelPdf.Render(svc.Bind(template, data, null), stream, new PdfOptions { FontFile = "/fonts/Vazirmatn.ttf" });
```

| | پشتیبانی PDF |
|---|---|
| متن فارسی/عربی | شکل‌دهی درست حروف (HarfBuzz)، ترتیب متن مختلط فارسی + عدد + لاتین (Bidi)، برگه‌ی راست‌به‌چپ (ستون A سمت راست) |
| سلول | رنگ پس‌زمینه، کادر، فونت/اندازه/ضخیم/مورب/زیرخط/خط‌خورده، رنگ متن، تراز افقی/عمودی، شکستن متن، تورفتگی، ادغام، قالب عدد، فرمول (مقدار محاسبه‌شده) |
| صفحه | اندازه‌ی کاغذ (A4/A3/A5/Letter/Legal) و جهت، حاشیه، جا دادن عرض در یک صفحه / مقیاس، تقسیم به صفحه (پایین سپس کنار)، **تکرار سطرهای عنوان**، سرصفحه/پاصفحه (`&P &N &D &T &A &F &L &C &R`) |
| تصویر و نمودار | تصویر ثابت/پویا؛ نمودار ستونی، میله‌ای، خطی، ناحیه‌ای، دایره‌ای (با عنوان فارسی، محور، راهنما، برچسب مقدار) |
| لینک | `http/https/mailto` قابل کلیک |

**نکته‌ی نسخه‌ی جدید:** قالب شرطی، چرخش متن، سرریز متن، بارکد/QR (برداری)، واترمارک (`PdfOptions.Watermark`) و شکست صفحه برای هر گروه (`BandDef.PageBreak`) هم پشتیبانی می‌شود؛ نمایشگر PDF داخل صفحه: `KscPdfViewer` / `Html.KscPdfViewer`؛ جزئیات در [راهنمای کامل](guide/KSC.Excel-Guide.fa.md) بخش ۱۰.

**ناپشتیبانی در PDF:** خطوط شبکه (به‌صورت پیش‌فرض چاپ نمی‌شود؛ `PrintGridLines`)، پیکان فیلتر خودکار. چون موتور رسم متفاوت از Excel است، ظاهر PDF بسیار نزدیک ولی نه پیکسل‌به‌پیکسل مثل «ذخیره به‌عنوان PDF» در Excel است.

**فونت:** از فونت سلول، سپس `FallbackFontFamily` (پیش‌فرض Tahoma)، سپس Segoe UI/Arial/Noto/DejaVu استفاده می‌شود؛ اگر هیچ‌کدام حروف فارسی نداشت از فونت‌منیجر سیستم کمک می‌گیرد. روی **لینوکس/Docker** معمولاً فونت فارسی نصب نیست: فونتی مثل Vazirmatn را نصب کنید یا `PdfOptions.FontFile` بدهید، و بسته‌ی `SkiaSharp.NativeAssets.Linux.NoDependencies` را به پروژه اضافه کنید. سقف صفحه: `PdfOptions.MaxPages` (۲۰۰۰). برای پیش‌نمایش تصویری: `ExcelPdf.RenderPng(wb, page)`.

## ۹-۱) نمایش در صفحه یا دانلود؟

| | نمایش در صفحه (Viewer) | دانلود (`ExcelFile` / `Excel`) |
|---|---|---|
| تجربه | گزارش بلافاصله دیده می‌شود، چند برگه با تب | فایل مستقیم |
| حجم داده | تا ۵۰۰ سطر نمایش (قابل تغییر `PreviewMaxRows`)؛ دانلود همیشه کامل است | هر اندازه (تا حد Excel) |
| دقت | تقریب HTML از Excel | خود Excel |
| پیشنهاد | گزارش‌های کوتاه (فاکتور، رسید، خلاصه) | گزارش‌های بزرگ/لیست‌ها |

تفاوت‌های شناخته‌شده‌ی پیش‌نمایش با Excel: قالب شرطی `expression/top10/duplicates` و تصاویر نمایش داده نمی‌شوند؛ چرخش متن و بعضی قالب‌های عددی خاص تقریبی‌اند؛ فرمول‌های خارج از مجموعه‌ی توابع بالا خالی می‌مانند.

## ۱۰) تنظیمات

`AddKscExcel(o => …)` / `KscExcel.Initialize(o => …)` — شیء `KscExcelOptions`:

| گزینه | پیش‌فرض | توضیح |
|---|---|---|
| `TemplateDirectory` | `ExcelTemplates` | پوشه‌ی قالب‌ها (نسبی به ContentRoot / ریشه‌ی سایت) |
| `Store` | فایل | ذخیره‌سازی سفارشی |
| `DesignerTitle` | `KSC Excel Designer` | عنوان طراح |
| `Language` | `fa` | `fa` (راست‌به‌چپ) یا `en` |
| `ReadOnlyDesigner` | `false` | طراح فقط باز/پیش‌نمایش/دانلود می‌کند، ذخیره/حذف ممنوع (مناسب Production) |
| `AllowAnonymousDesigner` | `false` | بدون آن، طراح خارج از Development نیاز به کاربر لاگین دارد |
| `DownloadPath` | `/kxl-download` | مسیر لینک‌های Viewer |
| `MaxRows` | ۱٬۰۴۸٬۵۷۶ | سقف سطر هر برگه |
| `MaxCells` | ۳٬۰۰۰٬۰۰۰ | سقف سلول (حافظه). بالاتر ببرید فقط اگر لازم است |
| `AllowSql` | `false` | اجازه‌ی اجرای SQL داده‌ست‌ها (§ منبع داده‌ی SQL) |
| `AddConnection(name, …)` | — | ثبت اتصال نام‌دار برای SQL |
| `SqlCommandTimeoutSeconds` / `SqlMaxRows` | ۳۰ / ۱۰۰٬۰۰۰ | محدودیت اجرا |
| `AddFilter` / `AddFunction` | — | فیلتر و تابع سفارشی |
| `PreviewMaxRows` | ۵۰۰ | سطرهای پیش‌نمایش |
| `AddModel<T>(name)` | — | ثبت مدل برای فهرست فیلد و داده‌ی نمونه‌ی خودکار |
| `AddModelFromJson(name, json)` | — | مدل از JSON نمونه |

## ۱۱) امنیت و دسترسی

- **طراح می‌تواند قالب بنویسد.** حتماً `RequireAuthorization("پالیسی")` (در .NET) یا `KscExcel.Authorize` (در Framework) بگذارید. خارج از Development، کاربر ناشناس رد می‌شود مگر `AllowAnonymousDesigner`.
- در پروژه‌های KSC برای هدایت کاربر ناشناس به صفحه‌ی ورود، مسیر را به `IAuthenticatedPathProvider` بدهید (مثل `KSC.ErrorLog`):
  ```csharp
  services.AddSingleton<IAuthenticatedPathProvider, ExcelPaths>();
  class ExcelPaths : IAuthenticatedPathProvider { public IEnumerable<string> GetAuthenticatedPaths() => new[] { "/excel-designer" }; }
  ```
- در Production معمولاً `ReadOnlyDesigner = true` و قالب‌ها از Git/Deploy بیایند.
- درخواست‌های نوشتن (PUT/POST/DELETE) هدر `X-KXL-Request: 1` می‌خواهند (محافظت CSRF بدون Cookie اضافه).
- نام قالب: فقط حرف/عدد/فاصله/`_ - .`؛ عبور از مسیر (`..`) رد می‌شود.
- داده‌ی کاربر هرگز **فرمول** نمی‌شود: مقدارهای رشته‌ای (حتی با `=`) به‌صورت متن ثابت نوشته می‌شوند → تزریق فرمول/CSV ندارد.
- خروجی HTML پیش‌نمایش کاملاً encode شده است؛ لینک‌ها فقط `http/https/mailto`.
- سقف اندازه‌ی درخواست طراح ۱۶ مگابایت؛ سقف سطر/سلول در برابر داده‌ی بسیار بزرگ.

## ۱۲) استقرار

1. پوشه‌ی `ExcelTemplates` را در پروژه بگذارید و به خروجی کپی کنید:
   ```xml
   <Content Include="ExcelTemplates\**" CopyToOutputDirectory="PreserveNewest" />
   ```
2. چند سرور (Web Farm): قالب‌ها را در پوشه‌ی مشترک یا دیتابیس (`ITemplateStore`) نگه دارید؛ لینک‌های Viewer در **حافظه‌ی همان سرور** هستند (Sticky session یا فقط دانلود مستقیم).
3. Docker: پوشه‌ی قالب را Volume کنید یا داخل Image بسازید و `ReadOnlyDesigner` بگذارید.
4. فونت: اگر قالب فونت خاصی (مثل `B Nazanin`) دارد، روی سیستم کاربر باید نصب باشد؛ فایل خودش فونت را جاسازی نمی‌کند.

## ۱۳) کارایی و محدودیت‌ها

اندازه‌گیری روی لپ‌تاپ معمولی (قالب «جدولی» با ۶ ستون، .NET 10، Release، تولید کامل xlsx):

| سطر × ستون | زمان | حجم فایل | حافظه‌ی اضافه |
|---|---|---|---|
| ۱٬۰۰۰ × ۶ | ≈ ۴۰ ms | ۳۷ KB | ≈ ۱ MB |
| ۱۰٬۰۰۰ × ۶ | ≈ ۰٫۶ s | ۳۳۷ KB | ≈ ۱۵ MB |
| ۱۰۰٬۰۰۰ × ۶ | ≈ ۲٫۸ s | ۳٫۳ MB | ≈ ۸۰ MB |

- کل گزارش در حافظه ساخته می‌شود؛ برای بیش از چند میلیون سلول، داده را به چند برگه/فایل بشکنید یا فیلتر کنید (`MaxCells` محافظ است).
- Excel حداکثر ۱٬۰۴۸٬۵۷۶ سطر و ۱۶٬۳۸۴ ستون و ۳۱ نویسه برای نام برگه دارد.
- رندر هم‌زمان چند درخواست امن است (بدون حالت مشترک).
- برای گزارش‌های سنگین از `Render(template, data, stream)` روی `Response` یا یک Job پس‌زمینه استفاده کنید.

## ۱۴) عیب‌یابی و سؤالات متداول

**فایل باز می‌شود ولی مقدار سلول خالی است.** نام فیلد را در پیش‌نمایش با داده‌ی نمونه بررسی کنید؛ `{{Customer.Name}}` با مسیر کامل. اگر مقدار `null` باشد سلول خالی است (`|default:`).

**«Formula references unknown band» / `#REF!`.** نام نوار در `[Band.Field]` باید دقیقاً نام نوار باشد و سلولی در آن نوار `{{Field}}` داشته باشد؛ برای ستون فرمولی از `[Band:E]` استفاده کنید. هشدارها در پنجره‌ی پیش‌نمایش (`هشدارها`) و در `RenderedWorkbook.Warnings` هستند.

**جمع ستون صفر یا ۰ می‌شود.** نوار خالی است (لیست داده خالی). عمداً خطا نمی‌دهد.

**ستون‌ها در Excel راست‌به‌چپ نیستند.** پنل برگه ← «راست‌به‌چپ».

**ارقام فارسی لازم است.** `{{N|fanum}}` (متن)، یا قالب سفارشی عدد (Excel خودش با تنظیمات ویندوز ارقام را نشان می‌دهد).

**`The report has more than … rows/cells`.** داده را فیلتر کنید یا `MaxCells` را (آگاهانه) بالا ببرید.

**ذخیره در طراح خطای 403 می‌دهد.** `ReadOnlyDesigner` فعال است یا مجوز ندارید.

**۴۰۱ در طراح.** خارج از Development کاربر باید لاگین باشد (یا `AllowAnonymousDesigner`).

**طراح بعد از Publish قالب را پیدا نمی‌کند.** پوشه‌ی `ExcelTemplates` را `CopyToOutputDirectory/Publish` کنید و `TemplateDirectory` را بررسی کنید.

**می‌خواهم UI طراح را در توسعه بدون Build عوض کنم.** متغیر محیطی `KSC_EXCEL_UI_DIR` را روی پوشه‌ی `src/KSC.Excel/Designer/UI` بگذارید (فقط برای توسعه‌ی خود KSC.Excel).

**آیا به Excel روی سرور نیاز است؟** خیر؛ هیچ COM/Interop استفاده نمی‌شود.

**آیا `.xls` قدیمی هم می‌دهد؟** خیر؛ فقط `.xlsx` (OOXML).

</div>
