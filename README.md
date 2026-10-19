# KSC.Excel.Pdf

**Persian/Arabic-correct PDF reports from visual templates — for .NET 8/10 *and* .NET Framework 4.6.1+.**
Design the report once in the [KSC.Excel](https://github.com/Mohsenbahrzadeh/KSC.Excel) visual designer (Stimulsoft-style, Excel-like grid), supply data from code, and get a PDF with correct letter shaping (HarfBuzz), bidi ordering and RTL layout — plus an **in-page PDF viewer**.

**گزارش PDF با شکل‌دهی درست فارسی/عربی و راست‌به‌چپ، برای .NET 8/10 و .NET Framework.** این مخزن بسته‌ی `KSC.Excel.Pdf` را همراه همه‌ی وابستگی‌هایش (موتور KSC.Excel، آداپتر ASP.NET Core و MVC 5) و نمونه‌ها دارد.

```csharp
byte[] pdf = svc.RenderPdf("Invoice", model);                       // from any code
return this.PdfFile("Invoice", model, "invoice.pdf");               // ASP.NET Core download
return this.Pdf("Invoice", model, "invoice.pdf");                   // MVC 5 (.NET Framework)
```

```cshtml
@await Component.InvokeAsync("KscPdfViewer", new { template = "Invoice", data = Model })   @* PDF viewer in the page *@
@Html.KscPdfViewer("Invoice", Model, "invoice")                                            @* same on MVC 5 *@
```

## What is in this repository

| Project | Target | Role |
|---|---|---|
| `src/KSC.Excel.Pdf` | netstandard2.0 | **the PDF renderer** (SkiaSharp + HarfBuzz): `RenderPdf`, `ExcelPdf.Render/RenderPng`, `PdfOptions`, `UsePdf()` |
| `src/KSC.Excel` | netstandard2.0 | dependency: template model, report engine, designer API + UI, `.xlsx` writer |
| `src/KSC.Excel.AspNetCore` | net8.0, net10.0 | dependency: ASP.NET Core integration, `PdfFile`, `KscPdfViewer` |
| `src/KSC.Excel.Mvc` | net472, net48 | dependency: MVC 5 integration, `Pdf`, `Html.KscPdfViewer` |
| `samples/KSC.Excel.Sample.Web`, `samples/KSC.Excel.Sample.Mvc5` | net10.0 / net48 | runnable samples (designer + PDF viewer + downloads) |
| `tests/KSC.Excel.Pdf.Tests`, `tests/KSC.Excel.Tests` | net10.0 + net48 | PDF pixel tests and engine tests |

The Windows desktop designer lives in the [KSC.Excel](https://github.com/Mohsenbahrzadeh/KSC.Excel) repository.

## PDF features

- **Persian / Arabic**: correct joining forms, mixed Persian + digits + Latin ordering, RTL sheets (column A on the right).
- **Cells**: fills, borders, fonts, alignment, wrap, indent, merges, number formats, formulas, **conditional formatting** (cell value, expression, text, duplicates, top/bottom, colour scale, data bar), **rotated and vertical text**, **text spill** into empty neighbours, shrink-to-fit.
- **Pages**: A3/A4/A5/Letter/Legal, orientation, margins, fit-to-width/scale, **repeating header rows**, header/footer codes (`&P &N &D &T &A &F &L &C &R`), **page break per group**, **watermark** (`PdfOptions.Watermark` or the template).
- **Graphics**: images, charts (column, bar, line, area, pie), **QR codes and Code 128 barcodes as crisp vector shapes**, clickable links.
- **Viewer**: browser PDF reader inside your page with Download PDF / Download Excel / Print buttons; no JavaScript library, no CDN.
- **Data**: objects, JSON, `DataTable`, SQL, REST endpoints, your own providers.

## Quick start

```bash
git clone https://github.com/Mohsenbahrzadeh/KSC.Pdf.git
cd KSC.Pdf
dotnet run --project samples/KSC.Excel.Sample.Web --urls http://localhost:5199 -- --environment Development
# http://localhost:5199/Home/PdfViewer   (PDF viewer)    http://localhost:5199/excel-designer   (designer)
```

```csharp
builder.Services.AddKscExcel(o => { o.UsePdf(); o.AddModel<Invoice>("Invoice"); });
app.MapKscExcelDesigner("/excel-designer").RequireAuthorization();
app.MapKscExcelDownloads();            // links used by the viewer
```

.NET Framework / MVC 5: `samples/KSC.Excel.Sample.Mvc5` (verified on IIS Express); call `o.UsePdf()` inside `KscExcel.Initialize`. SkiaSharp needs binding redirects on .NET Framework (already in that sample's `web.config`).

## Fonts

Cell font → `PdfOptions.FallbackFontFamily` (Tahoma) → Segoe UI/Arial/Noto/DejaVu → OS font manager. On **Linux/Docker** install a Persian-capable font (or set `PdfOptions.FontFile`) and add `SkiaSharp.NativeAssets.Linux.NoDependencies`. Page cap: `PdfOptions.MaxPages` (2000).

## Documentation, build, status

Complete guides (English / Persian): [docs/guide](docs/guide) — PDF is section 10. Build: `dotnet build`, `dotnet test tests/KSC.Excel.Pdf.Tests` (net10.0 and net48, Windows).

Honest limits: the output is very close to, but not pixel-identical with, Excel's own *Save as PDF*; auto-filter arrows and cell comments are not drawn; Linux/macOS hosts and LibreOffice were not tested.

## Security, contributing, license

See [SECURITY.md](SECURITY.md) and [CONTRIBUTING.md](CONTRIBUTING.md). Licensed under the [MIT License](LICENSE).
