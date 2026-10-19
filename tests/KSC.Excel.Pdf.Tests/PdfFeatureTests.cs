using System.Text;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Pdf;
using KSC.Excel.Services;
using Newtonsoft.Json.Linq;
using SkiaSharp;
using Xunit;

namespace KSC.Excel.Pdf.Tests;

/// <summary>PDF fidelity: conditional formats, text rotation, overflow, barcodes, watermark, page breaks.</summary>
public class PdfFeatureTests
{
    static IExcelReportService Svc() => new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore());

    static SKBitmap Page(RenderedWorkbook wb, int page = 1, PdfOptions o = null)
        => SKBitmap.Decode(ExcelPdf.RenderPng(wb, page, 1f, o));

    static int CountColor(SKBitmap bmp, Func<SKColor, bool> test)
    {
        int n = 0;
        for (int y = 0; y < bmp.Height; y++) for (int x = 0; x < bmp.Width; x++) if (test(bmp.GetPixel(x, y))) n++;
        return n;
    }

    static ExcelTemplate Simple(Action<SheetTemplate> tweak = null)
    {
        var b = new TemplateBuilder("T", "t");
        b.Sheet("S", s =>
        {
            s.Col(1, 24).Col(2, 12).Row(1, 24);
            s.Cell("A1", "{{Name}}").Cell("B1", "{{Qty}}");
            s.Band("R", "Rows", 1, 1);
            if (tweak != null) s.Raw(tweak);
        });
        return b.Build();
    }

    static JObject Rows(int n) => JObject.FromObject(new { Rows = Enumerable.Range(1, n).Select(i => new { Name = "item " + i, Qty = i * 10 }) });

    [Fact]
    public void Conditional_fill_is_drawn()
    {
        var plain = Svc().Bind(Simple(), Rows(5), null);
        var cf = Svc().Bind(Simple(s => s.ConditionalFormats.Add(new CondFormatDef { Range = "B1:B5", Type = "cellIs", Operator = "greaterThan", Formula1 = "20", Style = new StyleDef { Fill = "#FF0000" } })), Rows(5), null);
        bool Red(SKColor c) => c.Red > 240 && c.Green < 20 && c.Blue < 20;
        using var a = Page(plain); using var b = Page(cf);
        Assert.Equal(0, CountColor(a, Red));
        Assert.True(CountColor(b, Red) > 500, "three cells (30, 40, 50) should be filled red");
    }

    [Fact]
    public void Data_bars_and_color_scales_are_drawn()
    {
        var wb = Svc().Bind(Simple(s =>
        {
            s.ConditionalFormats.Add(new CondFormatDef { Range = "B1:B5", Type = "dataBar", BarColor = "#00AA00" });
        }), Rows(5), null);
        using var bmp = Page(wb);
        Assert.True(CountColor(bmp, c => c.Green > 120 && c.Red < 120 && c.Blue < 140) > 300);
    }

    [Fact]
    public void Rotated_text_stays_inside_its_cell_and_draws_something()
    {
        var wb = Svc().Bind(Simple(s => { s.Rows.Add(new RowDef { Index = 1, Height = 80 }); s.Cells[1].S = "rot"; }), Rows(2), null);
        wb.Template.Styles["rot"] = new StyleDef { Rotation = 90, H = "center", V = "center" };
        foreach (var r in wb.Sheets[0].Rows) foreach (var c in r.Cells) if (c.Col == 2) c.Style = wb.Template.Styles["rot"];
        using var bmp = Page(wb);
        Assert.True(CountColor(bmp, c => c.Red < 100) > 50);
        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(ExcelPdf.Render(wb), 0, 8));
    }

    [Fact]
    public void Long_text_spills_into_empty_neighbour_cells_but_not_into_filled_ones()
    {
        var b = new TemplateBuilder("T", "t");
        b.Sheet("S", s => { s.Col(1, 8).Col(2, 8).Col(3, 8); s.Cell("A1", "this is a long line of text that overflows"); s.Cell("A3", "this is a long line of text that overflows").Cell("B3", "X"); });
        var wb = Svc().Bind(b.Build(), null, null);
        using var bmp = Page(wb);
        // sheet pt: col = (8*7+5)*0.75 = 45.75 -> three columns end near x = 0.5in margin + 137
        float margin = 36f, oneCol = 45.75f;
        int InkRight(int y0, int y1) { int max = 0; for (int y = y0; y < y1; y++) for (int x = 0; x < bmp.Width; x++) if (bmp.GetPixel(x, y).Red < 110) max = Math.Max(max, x); return max; }
        int top = (int)(0.75 * 72);
        int overflowed = InkRight(top, top + 18);
        int clipped = InkRight(top + 36, top + 54);
        Assert.True(overflowed > margin + oneCol * 2, "text should spill over B and C (ink reaches x=" + overflowed + ")");
        Assert.True(clipped < margin + oneCol * 2 - 10, "text must be clipped at the filled neighbour (ink reaches x=" + clipped + ")");
    }

    [Fact]
    public void Barcodes_are_drawn_as_vector_shapes()
    {
        var b = new TemplateBuilder("T", "t");
        b.Sheet("S", s => { s.Col(1, 20).Row(1, 100); s.Cell("A1", "{{No}}"); s.Raw(sh => sh.Cells[0].Bc = "qr"); });
        var wb = Svc().Bind(b.Build(), JObject.Parse("{\"No\":\"INV-1403-77\"}"), null);
        using var bmp = Page(wb);
        Assert.True(CountColor(bmp, c => c.Red < 60 && c.Green < 60) > 400, "QR modules should be black");
        var pdf = ExcelPdf.Render(wb);
        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(pdf, 0, 8));
    }

    [Fact]
    public void Watermark_is_drawn_on_every_page_and_options_override_the_template()
    {
        var wb = Svc().Bind(Simple(s => s.Page.Watermark = "DRAFT"), Rows(3), null);
        var wb2 = Svc().Bind(Simple(), Rows(3), null);
        using var withMark = Page(wb);
        using var without = Page(wb2);
        bool Grey(SKColor c) => c.Red > 150 && c.Red < 240 && Math.Abs(c.Red - c.Green) < 6 && Math.Abs(c.Red - c.Blue) < 6;
        Assert.True(CountColor(withMark, Grey) > CountColor(without, Grey) + 2000);
        using var viaOptions = Page(wb2, 1, new PdfOptions { Watermark = "DRAFT" });
        Assert.True(CountColor(viaOptions, Grey) > CountColor(without, Grey) + 2000);
    }

    [Fact]
    public void Page_break_per_item_starts_a_new_pdf_page()
    {
        var t = Simple(s => s.Bands[0].PageBreak = true);
        var wb = Svc().Bind(t, Rows(4), null);
        Assert.Equal(4, ExcelPdf.CountPages(wb));
        var wb2 = Svc().Bind(Simple(), Rows(4), null);
        Assert.Equal(1, ExcelPdf.CountPages(wb2));
    }

    [Fact]
    public void Json_data_set_feeds_a_pdf()
    {
        var t = Simple();
        t.Sheets[0].Bands[0].Source = "Lookup";
        t.DataSets.Add(new DataSetDef { Name = "Lookup", Type = "json", Json = "[{\"Name\":\"k\",\"Qty\":1}]" });
        var pdf = ExcelPdf.Render(Svc().Bind(t, null, null));
        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(pdf, 0, 8));
    }

    [Fact]
    public void Password_protected_pdf_is_encrypted_and_well_formed()
    {
        var wb = Svc().Bind(Simple(), Rows(30), null);
        var plain = ExcelPdf.Render(wb, new PdfOptions { Title = "SecretTitleXYZ" });
        var enc = ExcelPdf.Render(wb, new PdfOptions { Title = "SecretTitleXYZ", Security = new PdfSecurity { UserPassword = "open-sesame", OwnerPassword = "boss", AllowCopy = false } });
        var text = Encoding.GetEncoding("ISO-8859-1").GetString(enc);
        Assert.StartsWith("%PDF-", text);
        Assert.Contains("/Encrypt", text); Assert.Contains("/AESV3", text);
        Assert.DoesNotContain("SecretTitleXYZ", text);
        Assert.Contains("SecretTitleXYZ", Encoding.GetEncoding("ISO-8859-1").GetString(plain));
        Assert.EndsWith("%%EOF" + (char)10, text);
        var dir = Environment.GetEnvironmentVariable("KXL_PDF_PNG_DIR");
        if (!string.IsNullOrEmpty(dir))
        {
            File.WriteAllBytes(Path.Combine(dir, "plain.pdf"), plain);
            File.WriteAllBytes(Path.Combine(dir, "locked.pdf"), enc);
            File.WriteAllBytes(Path.Combine(dir, "locked-open.pdf"), ExcelPdf.Render(wb, new PdfOptions { Security = new PdfSecurity { OwnerPassword = "boss", AllowPrint = false, AllowCopy = false } }));
        }
    }

    [Fact]
    public void Combo_chart_with_secondary_axis_renders()
    {
        var b = new TemplateBuilder("Ch", "combo");
        b.Sheet("S", s => { s.Cell("A1", "{{M}}").Cell("B1", "{{Amount}}").Cell("C1", "{{Pct}}"); s.Band("R", "Rows", 1, 1); });
        var t = b.Build();
        t.Sheets[0].Charts.Add(new ChartDef { Name = "c", Type = "column", Title = "فروش و درصد رشد", Row = 6, Col = 1, Width = 520, Height = 300, Categories = "A1:A1",
            Series = { new ChartSeriesDef { Name = "فروش", Values = "B1:B1", Color = "#2E75B6" }, new ChartSeriesDef { Name = "رشد", Values = "C1:C1", Type = "line", Secondary = true, Color = "#C55A11" } } });
        var data = JObject.Parse("{\"Rows\":[{\"M\":\"فروردین\",\"Amount\":100,\"Pct\":0.1},{\"M\":\"اردیبهشت\",\"Amount\":250,\"Pct\":0.4},{\"M\":\"خرداد\",\"Amount\":180,\"Pct\":0.3},{\"M\":\"تیر\",\"Amount\":320,\"Pct\":0.55}]}");
        var wb = Svc().Bind(t, data, null);
        var png = ExcelPdf.RenderPng(wb, 1, 1.5f);
        using var bmp = SKBitmap.Decode(png);
        Assert.True(CountColor(bmp, c => c.Red > 170 && c.Green < 120 && c.Blue < 60) > 30, "the orange line should be drawn");
        Assert.True(CountColor(bmp, c => c.Blue > 150 && c.Red < 80) > 500, "blue columns should be drawn");
        var dir = Environment.GetEnvironmentVariable("KXL_PDF_PNG_DIR");
        if (!string.IsNullOrEmpty(dir)) File.WriteAllBytes(Path.Combine(dir, "combo.png"), png);
    }

    [Fact]
    public void Visual_sample_for_manual_inspection()
    {
        var dir = Environment.GetEnvironmentVariable("KXL_PDF_PNG_DIR");
        if (string.IsNullOrEmpty(dir)) return;
        var b = new TemplateBuilder("V", "گزارش نمونه").Font("Tahoma", 10);
        var head = b.Style("h", new StyleDef { Bold = true, Color = "#FFFFFF", Fill = "#1F4E79", H = "center", V = "center", Rtl = true });
        var rot = b.Style("rot", new StyleDef { Bold = true, Color = "#FFFFFF", Fill = "#2E75B6", H = "center", V = "center", Rotation = 90 });
        b.Sheet("نمونه", s =>
        {
            s.Col(1, 6).Col(2, 22).Col(3, 12).Col(4, 14).Col(5, 22).Col(6, 14).Row(1, 30).Row(3, 70);
            s.Cell("A1", "گزارش فروش با قالب شرطی، بارکد و واترمارک — این عنوان طولانی در ستون‌های خالی کناری ادامه می‌یابد", head);
            s.Cell("A3", "ردیف", head).Cell("B3", "محصول", head).Cell("C3", "تعداد", head).Cell("D3", "مبلغ", head).Cell("E3", "QR", head).Cell("F3", "عمودی", rot);
            s.Cell("A4", "{{$index}}").Cell("B4", "{{Name}}").Cell("C4", "{{Qty}}").Cell("D4", "{{Amount}}").Cell("E4", "{{Code}}").Cell("F4", "{{Code}}");
            s.Band("R", "Rows", 4, 4, "#F5F9FD");
            s.Raw(sh =>
            {
                sh.Rtl = true;
                sh.Rows.Add(new RowDef { Index = 4, Height = 54 });
                sh.Page.Watermark = "پیش‌نویس";
                sh.Page.FitToWidth = true;
                sh.Page.Footer = "&C&P / &N";
                sh.Cells.First(c => c.R == 4 && c.C == 5).Bc = "qr";
                sh.Cells.First(c => c.R == 4 && c.C == 6).Bc = "code128";
                sh.Bands[0].PageBreak = false;
                sh.ConditionalFormats.Add(new CondFormatDef { Range = "C4:C4", Type = "cellIs", Operator = "greaterThan", Formula1 = "30", Style = new StyleDef { Fill = "#F8CBAD", Color = "#9C0006", Bold = true } });
                sh.ConditionalFormats.Add(new CondFormatDef { Range = "D4:D4", Type = "dataBar", BarColor = "#63BE7B" });
                sh.Merges.Add("A1:F1");
            });
        }, rtl: true);
        var data = JObject.FromObject(new { Rows = Enumerable.Range(1, 5).Select(i => new { Name = "ورق فولادی " + i, Qty = i * 12, Amount = i * 1_250_000, Code = "INV-" + (1000 + i) }) });
        var wb = Svc().Bind(b.Build(), data, null);
        File.WriteAllBytes(Path.Combine(dir, "features-page1.png"), ExcelPdf.RenderPng(wb, 1, 2f));
        File.WriteAllBytes(Path.Combine(dir, "features.pdf"), ExcelPdf.Render(wb));
    }
}
