using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Services;
using KSC.Excel.Writer;
using Newtonsoft.Json.Linq;
using Xunit;

namespace KSC.Excel.Tests;

/// <summary>CSV, Word (docx), stand-alone HTML and the unified export API.</summary>
public class OutputTests
{
    static ExcelTemplate Report()
    {
        var b = new TemplateBuilder("R", "گزارش");
        var head = b.Style("h", new StyleDef { Bold = true, Fill = "#2E75B6", Color = "#FFFFFF", H = "center", Top = new BorderDef { Style = "thin", Color = "#000000" } });
        b.Sheet("نمونه", s =>
        {
            s.Col(1, 20).Col(2, 12).Row(1, 30);
            s.Cell("A1", "عنوان گزارش", head).StyleRange("A1:B1", head).Merge("A1:B1");
            s.Cell("A2", "{{Name}}").Cell("B2", "{{Qty}}");
            s.Band("R", "Rows", 2, 2);
            s.Cell("A3", "جمع").Cell("B3", "=SUM([R.Qty])");
            s.Raw(sh => sh.Rtl = true);
        }, rtl: true);
        return b.Build();
    }

    const string Data = "{\"Rows\":[{\"Name\":\"alpha, \\\"quoted\\\"\",\"Qty\":5},{\"Name\":\"=cmd|' /C calc'!A0\",\"Qty\":7},{\"Name\":\"فارسی\",\"Qty\":1.5}]}";

    static RenderedWorkbook Bind(JObject data = null)
        => new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore()).Bind(Report(), data ?? JObject.Parse(Data), null);

    // ------------------------------------------------------------------ CSV
    [Fact]
    public void Csv_quotes_values_and_neutralises_formula_injection()
    {
        var bytes = CsvWriter.ToBytes(Bind());
        Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, bytes.Take(3).ToArray());           // BOM: Excel shows Persian correctly
        var lines = Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3).Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal("عنوان گزارش,", lines[0]);                                              // merged range: only the anchor has text
        Assert.Equal("\"alpha, \"\"quoted\"\"\",5", lines[1]);
        Assert.StartsWith("'=cmd", lines[2]);                                                // text starting with = is neutralised
        Assert.Equal("فارسی,1.5", lines[3]);
        Assert.Equal("جمع,13.5", lines[4]);                                                  // formula cached value
    }

    [Fact]
    public void Csv_rejects_a_missing_sheet_and_supports_other_delimiters()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CsvWriter.ToBytes(Bind(), 5));
        Assert.Contains(";", CsvWriter.ToString(Bind(), 0, ';'));
    }

    // ------------------------------------------------------------------ DOCX
    static Dictionary<string, XDocument> Open(byte[] docx)
    {
        var res = new Dictionary<string, XDocument>();
        using var zip = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        foreach (var e in zip.Entries.Where(e => e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels")))
            using (var s = e.Open()) res[e.FullName] = XDocument.Load(s);          // throws on malformed XML
        return res;
    }

    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void Docx_is_wellformed_with_table_merge_rtl_and_styles()
    {
        var parts = Open(DocxWriter.ToBytes(Bind()));
        foreach (var p in new[] { "[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/styles.xml", "word/_rels/document.xml.rels" }) Assert.True(parts.ContainsKey(p), p);
        var doc = parts["word/document.xml"];
        Assert.Single(doc.Descendants(W + "tbl"));
        Assert.Contains("عنوان گزارش", doc.Descendants(W + "t").Select(t => t.Value));
        Assert.Equal("2", (string)doc.Descendants(W + "gridSpan").First().Attribute(W + "val"));      // A1:B1 merged
        Assert.NotEmpty(doc.Descendants(W + "bidiVisual"));                                        // RTL table
        Assert.Equal("2E75B6", (string)doc.Descendants(W + "shd").First().Attribute(W + "fill"));
        Assert.NotEmpty(doc.Descendants(W + "b"));                                                  // bold header
        Assert.Contains("13.5", doc.Descendants(W + "t").Select(t => t.Value));                    // formula value
    }

    [Fact]
    public void Docx_embeds_barcode_images_inside_their_cell()
    {
        var b = new TemplateBuilder("B");
        b.Sheet("S", s => { s.Col(1, 20).Row(1, 80); s.Cell("A1", "{{No}}"); s.Raw(sh => sh.Cells[0].Bc = "qr"); });
        var wb = new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore()).Bind(b.Build(), JObject.Parse("{\"No\":\"INV-1\"}"), null);
        var bytes = DocxWriter.ToBytes(wb);
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        Assert.Contains(zip.Entries, e => e.FullName == "word/media/image1.png");
        var parts = Open(bytes);
        Assert.Contains("rIdImg1", parts["word/_rels/document.xml.rels"].ToString());
        Assert.NotEmpty(parts["word/document.xml"].Descendants(XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing") + "inline"));
    }

    [Fact]
    public void Docx_multiple_sheets_get_headings_and_page_settings()
    {
        var b = new TemplateBuilder("M");
        b.Sheet("One", s => s.Cell("A1", "x").Landscape());
        b.Sheet("Two", s => s.Cell("A1", "y"));
        var wb = new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore()).Bind(b.Build(), null, null);
        var doc = Open(DocxWriter.ToBytes(wb))["word/document.xml"];
        Assert.Equal(2, doc.Descendants(W + "tbl").Count());
        Assert.Equal(new[] { "One", "Two" }, doc.Descendants(W + "pStyle").Select(p => (string)p.Parent.Parent.Descendants(W + "t").First()).ToArray());
        Assert.Equal("landscape", (string)doc.Descendants(W + "pgSz").First().Attribute(W + "orient"));
    }

    // ------------------------------------------------------------------ HTML + unified export
    [Fact]
    public void Html_document_is_a_complete_page_with_direction()
    {
        var html = ReportExport.HtmlDocument(Bind());
        Assert.StartsWith("<!DOCTYPE html>", html); Assert.Contains("dir=\"rtl\"", html); Assert.Contains("<meta charset=\"utf-8\">", html);
        Assert.Contains("عنوان گزارش", html);
    }

    [Theory]
    [InlineData("xlsx", ReportFormat.Xlsx)] [InlineData("excel", ReportFormat.Xlsx)] [InlineData("PDF", ReportFormat.Pdf)] [InlineData("csv", ReportFormat.Csv)]
    [InlineData("word", ReportFormat.Docx)] [InlineData("html", ReportFormat.Html)] [InlineData("png", ReportFormat.Png)]
    public void Formats_parse(string text, ReportFormat expected)
    {
        Assert.True(ReportExport.TryParseFormat(text, out var f)); Assert.Equal(expected, f);
    }

    [Fact]
    public void Export_dispatches_every_format_and_explains_missing_pdf()
    {
        var svc = new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore());
        var o = new KscExcelOptions();
        var data = JObject.Parse(Data);
        foreach (var (fmt, ext, ct) in new[] { (ReportFormat.Xlsx, "xlsx", ExcelContent.Xlsx), (ReportFormat.Csv, "csv", "text/csv; charset=utf-8"), (ReportFormat.Docx, "docx", ReportExport.DocxContentType), (ReportFormat.Html, "html", "text/html; charset=utf-8") })
        {
            var f = svc.Export(o, Report(), data, null, fmt);
            Assert.Equal(ext, f.Extension); Assert.Equal(ct, f.ContentType); Assert.True(f.Bytes.Length > 50);
        }
        Assert.Contains("UsePdf", Assert.Throws<InvalidOperationException>(() => svc.Export(o, Report(), data, null, ReportFormat.Pdf)).Message);
        Assert.Contains("UsePdf", Assert.Throws<InvalidOperationException>(() => svc.Export(o, Report(), data, null, ReportFormat.Png)).Message);
        o.PdfExporter = wb => new byte[] { 1, 2, 3 }; o.PngExporter = (wb, p) => new byte[] { (byte)p };
        Assert.Equal(new byte[] { 1, 2, 3 }, svc.Export(o, Report(), data, null, ReportFormat.Pdf).Bytes);
        Assert.Equal(new byte[] { 4 }, svc.Export(o, Report(), data, null, ReportFormat.Png, 4).Bytes);
    }

    [Fact]
    public void File_names_are_safe_and_keep_persian()
    {
        Assert.Equal("گزارش_فروش.pdf", ReportExport.FileName("گزارش فروش", "pdf"));
        Assert.Equal("a_b_c.csv", ReportExport.FileName("a/b\\c", "csv"));
        Assert.Equal("report.xlsx", ReportExport.FileName("..", "xlsx"));
        Assert.Equal("x.docx", ReportExport.FileName("x.docx", "docx"));
        var dump = Environment.GetEnvironmentVariable("KXL_DUMP_DIR");
        if (!string.IsNullOrEmpty(dump)) { File.WriteAllBytes(Path.Combine(dump, "report.docx"), DocxWriter.ToBytes(Bind())); File.WriteAllBytes(Path.Combine(dump, "report.csv"), CsvWriter.ToBytes(Bind())); }
    }
}
