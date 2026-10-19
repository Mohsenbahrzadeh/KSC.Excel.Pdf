using System.Text;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Pdf;
using KSC.Excel.Services;
using Newtonsoft.Json.Linq;
using SkiaSharp;
using Xunit;

namespace KSC.Excel.Pdf.Tests;

public class PdfTests
{
    static IExcelReportService Svc() => new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore());

    static ExcelTemplate PersianInvoice()
    {
        var b = new TemplateBuilder("Inv", "فاکتور").Font("Tahoma", 10);
        var title = b.Style("t", new StyleDef { Size = 18, Bold = true, Color = "#FFFFFF", Fill = "#1F4E79", H = "center", V = "center" });
        var head = b.Style("h", new StyleDef { Bold = true, Color = "#FFFFFF", Fill = "#2E75B6", H = "center", Left = new BorderDef { Style = "thin", Color = "#1F4E79" }, Right = new BorderDef { Style = "thin", Color = "#1F4E79" }, Top = new BorderDef { Style = "thin", Color = "#1F4E79" }, Bottom = new BorderDef { Style = "thin", Color = "#1F4E79" } });
        var cell = b.Style("c", new StyleDef { Bottom = new BorderDef { Style = "thin", Color = "#BFC9D4" }, Format = "#,##0" });
        var tot = b.Style("tot", new StyleDef { Bold = true, Fill = "#DDEBF7", Format = "#,##0" });
        b.Sheet("فاکتور", s =>
        {
            s.Col(1, 8).Col(2, 34).Col(3, 10).Col(4, 18).Col(5, 20).Row(1, 36);
            s.Cell("A1", "فاکتور فروش شماره {{Number}}", title).StyleRange("A1:E1", title).Merge("A1:E1");
            s.Cell("A3", "مشتری:").Cell("B3", "{{Customer}} (تهران)");
            s.Cell("A5", "ردیف", head).Cell("B5", "شرح کالا", head).Cell("C5", "تعداد", head).Cell("D5", "قیمت واحد", head).Cell("E5", "مبلغ", head);
            s.Cell("A6", "{{$index}}", cell).Cell("B6", "{{Product}}", cell).Cell("C6", "{{Qty}}", cell).Cell("D6", "{{UnitPrice}}", cell).Cell("E6", "=[Qty]*[UnitPrice]", cell);
            s.Band("L", "Lines", 6, 6, "#F5F9FD");
            s.Cell("A7", "جمع کل", tot).StyleRange("A7:D7", tot).Merge("A7:D7").Cell("E7", "=SUM([L:E])", tot);
            s.Raw(sh => { sh.Page.FitToWidth = true; sh.Page.Footer = "&C&P / &N"; sh.Page.PrintTitleRows = "5:5"; });
        }, rtl: true);
        return b.Build();
    }

    static JObject Data(int lines) => JObject.FromObject(new
    {
        Number = "۱۴۰۳-۰۲۱", Customer = "شرکت نمونه",
        Lines = Enumerable.Range(1, lines).Select(i => new { Product = "ورق فولادی " + i + " میلی‌متر", Qty = i, UnitPrice = 1000 * i })
    });

    [Fact]
    public void Renders_a_valid_pdf_with_persian_text()
    {
        var svc = Svc();
        var pdf = svc.RenderPdf(PersianInvoice(), Data(5));
        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(pdf, 0, 8));
        Assert.True(pdf.Length > 2000);
    }

    [Fact]
    public void Long_reports_paginate_and_repeat_headers()
    {
        var svc = Svc();
        var wb = svc.Bind(PersianInvoice(), Data(120), null);
        Assert.True(ExcelPdf.CountPages(wb) >= 3);
        var pdf = ExcelPdf.Render(wb);
        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(pdf, 0, 8));
    }

    [Fact]
    public void Png_preview_is_not_blank_and_has_the_page_size()
    {
        var svc = Svc();
        var wb = svc.Bind(PersianInvoice(), Data(8), null);
        var png = ExcelPdf.RenderPng(wb, 1, 1f);
        using var bmp = SKBitmap.Decode(png);
        Assert.Equal(595, bmp.Width); Assert.Equal(842, bmp.Height);
        int dark = 0; for (int y = 0; y < bmp.Height; y += 3) for (int x = 0; x < bmp.Width; x += 3) if (bmp.GetPixel(x, y).Red < 120) dark++;
        Assert.True(dark > 200, "page looks blank");
        var dir = Environment.GetEnvironmentVariable("KXL_PDF_PNG_DIR");
        if (!string.IsNullOrEmpty(dir)) File.WriteAllBytes(Path.Combine(dir, "invoice-page1.png"), ExcelPdf.RenderPng(wb, 1, 2f));
    }

    [Fact]
    public void Charts_images_and_wide_sheets_render()
    {
        var b = new TemplateBuilder("W", "wide");
        b.Sheet("S", s =>
        {
            s.Cell("A1", "{{Name}}").Cell("B1", "{{Qty}}");
            for (int c = 3; c <= 14; c++) s.Cell(CellRef.A1(1, c), "col" + c);
            s.Band("R", "Rows", 1, 1);
        });
        var t = b.Build(); var sh = t.Sheets[0];
        for (int c = 1; c <= 14; c++) sh.Columns.Add(new ColumnDef { Index = c, Width = 22 });
        sh.Charts.Add(new ChartDef { Name = "c", Type = "column", Title = "نمودار تعداد", Row = 8, Col = 1, Categories = "A1:A1", Series = { new ChartSeriesDef { Name = "تعداد", Values = "B1:B1" } } });
        sh.Charts.Add(new ChartDef { Name = "p", Type = "pie", Row = 8, Col = 5, Categories = "A1:A1", Series = { new ChartSeriesDef { Values = "B1:B1" } } });
        var data = JObject.FromObject(new { Rows = Enumerable.Range(1, 6).Select(i => new { Name = "n" + i, Qty = i * 3 }) });
        var wb = Svc().Bind(t, data, null);
        Assert.True(ExcelPdf.CountPages(wb) >= 2);                  // 14 wide columns do not fit one page
        var dir = Environment.GetEnvironmentVariable("KXL_PDF_PNG_DIR");
        if (!string.IsNullOrEmpty(dir)) File.WriteAllBytes(Path.Combine(dir, "wide-page1.png"), ExcelPdf.RenderPng(wb, 1, 2f));
        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(ExcelPdf.Render(wb), 0, 8));
    }

    [Fact]
    public void Page_limit_is_enforced()
    {
        var wb = Svc().Bind(PersianInvoice(), Data(300), null);
        Assert.Throws<ExcelRenderException>(() => ExcelPdf.Render(wb, new PdfOptions { MaxPages = 2 }));
    }
}
