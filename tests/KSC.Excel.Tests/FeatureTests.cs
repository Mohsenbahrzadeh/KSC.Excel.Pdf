using System.Net;
using System.Text;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Services;
using Newtonsoft.Json.Linq;
using Xunit;
using ZXing;
using ZXing.Common;

namespace KSC.Excel.Tests;

/// <summary>Conditional formats, page breaks, JSON / provider / REST data sets, barcodes and the PDF viewer markup.</summary>
public class FeatureTests
{
    // ------------------------------------------------------------------ helpers
    static RenderedWorkbook Bind(ExcelTemplate t, object data, KscExcelOptions o = null)
        => new ExcelReportService(o ?? new KscExcelOptions(), new MemoryTemplateStore()).Bind(t, data, null);

    static ExcelTemplate Numbers(Action<SheetTemplate> tweak)
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s =>
        {
            s.Cell("A1", "{{Name}}").Cell("B1", "{{Qty}}");
            s.Band("Items", "Items", 1, 1);
            s.Raw(tweak);
        });
        return b.Build();
    }

    const string Items = "{\"Items\":[{\"Name\":\"a\",\"Qty\":5},{\"Name\":\"b\",\"Qty\":50},{\"Name\":\"c\",\"Qty\":500},{\"Name\":\"b\",\"Qty\":7}]}";

    static CfResult Cf(RenderedWorkbook wb, int row, int col)
    {
        var sh = wb.Sheets[0];
        var cell = sh.Get(row, col);
        return new CondFormatEvaluator(sh).Evaluate(row, col, cell?.Value, Convert.ToString(cell?.Value));
    }

    static StyleDef Red => new StyleDef { Fill = "#FF0000" };

    // ------------------------------------------------------------------ conditional formats
    [Fact]
    public void CondFormat_cellIs_expression_top10_duplicates_and_text()
    {
        var wb = Bind(Numbers(s =>
        {
            s.ConditionalFormats.Add(new CondFormatDef { Range = "B1:B4", Type = "cellIs", Operator = "greaterThan", Formula1 = "100", Style = Red });
            s.ConditionalFormats.Add(new CondFormatDef { Range = "A1:A4", Type = "duplicates", Style = new StyleDef { Color = "#00FF00" } });
            s.ConditionalFormats.Add(new CondFormatDef { Range = "B1:B4", Type = "top10", Rank = 2, Style = new StyleDef { Bold = true } });
            s.ConditionalFormats.Add(new CondFormatDef { Range = "A1:A4", Type = "expression", Formula1 = "$B1=50", Style = new StyleDef { Fill = "#0000FF" } });
            s.ConditionalFormats.Add(new CondFormatDef { Range = "A1:A4", Type = "cellIs", Operator = "equal", Formula1 = "\"c\"", Style = new StyleDef { Italic = true } });
        }), Make.Json(Items));

        Assert.Equal("#FF0000", Cf(wb, 3, 2).Style.Fill);          // 500 > 100
        Assert.Null(Cf(wb, 1, 2)?.Style?.Fill);                      // 5
        Assert.Equal("#00FF00", Cf(wb, 2, 1).Style.Color);           // "b" appears twice
        Assert.Equal("#00FF00", Cf(wb, 4, 1).Style.Color);
        Assert.Null(Cf(wb, 1, 1)?.Style?.Color);                     // "a" is unique
        Assert.True(Cf(wb, 3, 2).Style.Bold);                        // top 2: 500 and 50
        Assert.True(Cf(wb, 2, 2).Style.Bold);
        Assert.False(Cf(wb, 4, 2)?.Style?.Bold ?? false);
        Assert.Equal("#0000FF", Cf(wb, 2, 1).Style.Fill);            // expression relative to the first cell: row 2 has Qty 50
        Assert.NotEqual("#0000FF", Cf(wb, 1, 1)?.Style?.Fill);
        Assert.True(Cf(wb, 3, 1).Style.Italic);                      // text equality
    }

    [Fact]
    public void CondFormat_color_scale_and_data_bar()
    {
        var wb = Bind(Numbers(s =>
        {
            s.ConditionalFormats.Add(new CondFormatDef { Range = "B1:B4", Type = "colorScale", Colors = new List<string> { "#000000", "#FFFFFF" } });
            s.ConditionalFormats.Add(new CondFormatDef { Range = "B1:B4", Type = "dataBar", BarColor = "#123456" });
        }), Make.Json(Items));
        Assert.Equal("#000000", Cf(wb, 1, 2).ScaleFill);             // minimum
        Assert.Equal("#FFFFFF", Cf(wb, 3, 2).ScaleFill);             // maximum
        Assert.Equal(1.0, Cf(wb, 3, 2).BarFraction.Value, 3);
        Assert.Equal("#123456", Cf(wb, 3, 2).BarColor);
    }

    [Fact]
    public void CellRef_Shift_moves_relative_columns_and_rows_only()
        => Assert.Equal("$B3+C$1+$A$1", CellRef.Shift("$B2+B$1+$A$1", 1, 1));

    // ------------------------------------------------------------------ page breaks
    [Fact]
    public void Band_page_break_records_breaks_and_writes_rowBreaks()
    {
        var t = Numbers(s => s.Bands[0].PageBreak = true);
        var wb = Bind(t, Make.Json(Items));
        Assert.Equal(new[] { 1, 2, 3 }, wb.Sheets[0].RowBreaks.OrderBy(x => x).ToArray());
        var x = Make.Render(t, Make.Json(Items));
        var brks = x.Sheet().Descendants(Xlsx.Main + "brk").Select(b => (string)b.Attribute("id")).ToList();
        Assert.Equal(new[] { "1", "2", "3" }, brks);
    }

    // ------------------------------------------------------------------ JSON / provider / REST data sets
    static ExcelTemplate WithDataSet(DataSetDef d)
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => { s.Cell("A1", "{{Name}}").Cell("B1", "{{Qty}}"); s.Band("Rows", d.Name, 1, 1); });
        var t = b.Build();
        t.DataSets.Add(d);
        return t;
    }

    [Fact]
    public void Json_data_set_inline_with_path()
    {
        var t = WithDataSet(new DataSetDef { Name = "Rows", Type = "json", Json = "{\"data\":{\"items\":[{\"Name\":\"x\",\"Qty\":1},{\"Name\":\"y\",\"Qty\":2}]}}", JsonPath = "data.items" });
        var x = Make.Render(t, null);
        Assert.Equal("x", x.Value("A1")); Assert.Equal("y", x.Value("A2")); Assert.Equal("2", x.Value("B2"));
    }

    [Fact]
    public void Provider_data_set_receives_report_parameters()
    {
        var o = new KscExcelOptions();
        IDictionary<string, object> seen = null;
        o.AddProvider("Orders", p => { seen = p; return new[] { new { Name = "p", Qty = 9 } }; });
        var t = WithDataSet(new DataSetDef { Name = "Rows", Type = "provider", Connection = "Orders" });
        t.Parameters.Add(new ParameterDef { Name = "Year", Type = "number", Default = "1403" });
        var wb = Bind(t, null, o);
        Assert.Equal("p", wb.Sheets[0].Get(1, 1).Value);
        Assert.Equal(1403d, seen["Year"]);
    }

    [Fact]
    public void Unregistered_provider_or_endpoint_fails_clearly_and_templates_cannot_choose_hosts()
    {
        var ex = Assert.Throws<ExcelRenderException>(() => Bind(WithDataSet(new DataSetDef { Name = "Rows", Type = "provider", Connection = "Nope" }), null));
        Assert.Contains("not registered", ex.Message);

        var o = new KscExcelOptions().AddHttpEndpoint("Api", "http://127.0.0.1:9/");
        var evil = Assert.Throws<ExcelRenderException>(() => Bind(WithDataSet(new DataSetDef { Name = "Rows", Type = "json", Connection = "Api", Url = "http://evil.example/steal" }), null, o));
        Assert.Contains("registered endpoint host", evil.Message);
    }

    [Fact]
    public async Task Rest_endpoint_is_called_with_encoded_parameters_headers_and_path()
    {
        int port;
        var probe = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0); probe.Start(); port = ((IPEndPoint)probe.LocalEndpoint).Port; probe.Stop();
        var listener = new HttpListener();
        listener.Prefixes.Add("http://127.0.0.1:" + port + "/");
        listener.Start();
        string seenUrl = null, seenKey = null;
        var server = Task.Run(() =>
        {
            var ctx = listener.GetContext();
            seenUrl = ctx.Request.Url.PathAndQuery; seenKey = ctx.Request.Headers["X-Api-Key"];
            var body = Encoding.UTF8.GetBytes("{\"result\":[{\"Name\":\"rest\",\"Qty\":3}]}");
            ctx.Response.ContentType = "application/json"; ctx.Response.OutputStream.Write(body, 0, body.Length); ctx.Response.Close();
        });
        try
        {
            var o = new KscExcelOptions().AddHttpEndpoint("Api", "http://127.0.0.1:" + port + "/v1/", new Dictionary<string, string> { ["X-Api-Key"] = "k1" });
            var t = WithDataSet(new DataSetDef { Name = "Rows", Type = "json", Connection = "Api", Url = "orders?q={Q}", JsonPath = "result" });
            t.Parameters.Add(new ParameterDef { Name = "Q", Type = "text", Default = "a b&c" });
            var wb = Bind(t, null, o);
            Assert.Same(server, await Task.WhenAny(server, Task.Delay(5000)));
            Assert.Equal("rest", wb.Sheets[0].Get(1, 1).Value);
            Assert.Equal("/v1/orders?q=a%20b%26c", seenUrl);
            Assert.Equal("k1", seenKey);
        }
        finally { listener.Close(); }
    }

    // ------------------------------------------------------------------ barcodes
    static string Decode(BarcodeMatrix m, BarcodeFormat fmt)
    {
        int k = 6, q = m.Quiet, cols = m.Cols + 2 * q, rows = m.Is2D ? m.Rows + 2 * q : 1;
        int w = cols * k, h = m.Is2D ? rows * k : 60;
        var px = new byte[w * h];
        for (int i = 0; i < px.Length; i++) px[i] = 255;
        for (int r = 0; r < (m.Is2D ? m.Rows : 1); r++)
            for (int c = 0; c < m.Cols; c++)
                if (m.Modules[r, c])
                    for (int y = m.Is2D ? (r + q) * k : 0; y < (m.Is2D ? (r + q + 1) * k : h); y++)
                        for (int x = (c + q) * k; x < (c + q + 1) * k; x++) px[y * w + x] = 0;
        var reader = new BarcodeReaderGeneric { Options = new DecodingOptions { TryHarder = true, PossibleFormats = new[] { fmt } } };
        return reader.Decode(new RGBLuminanceSource(px, w, h, RGBLuminanceSource.BitmapFormat.Gray8))?.Text;
    }

    [Theory]
    [InlineData("qr", "https://example.com/invoice/1403-021")]
    [InlineData("qr:L", "Hello")]
    [InlineData("qr:H", "سلام دنیا ۱۲۳")]
    [InlineData("qr:Q", "abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz")]
    public void Qr_codes_decode_with_an_independent_reader(string type, string text)
        => Assert.Equal(text, Decode(Barcodes.Encode(type, text), BarcodeFormat.QR_CODE));

    [Fact]
    public void Qr_large_versions_with_alignment_patterns_decode()
    {
        var text = new string('x', 150) + "END";     // version 8, ECC M
        Assert.Equal(text, Decode(Barcodes.Encode("qr:M", text), BarcodeFormat.QR_CODE));
        var big = string.Concat(Enumerable.Range(0, 260).Select(i => (char)('A' + i % 26)));   // version 10, ECC L
        Assert.Equal(big, Decode(Barcodes.Encode("qr:L", big), BarcodeFormat.QR_CODE));
    }

    [Theory]
    [InlineData("INV-2024-0042")]
    [InlineData("1234567890")]
    [InlineData("Hello World!")]
    public void Code128_decodes(string text) => Assert.Equal(text, Decode(Barcodes.Encode("code128", text), BarcodeFormat.CODE_128));

    [Theory]
    [InlineData("ean13", "5901234123457", BarcodeFormat.EAN_13)]
    [InlineData("ean13", "590123412345", BarcodeFormat.EAN_13)]
    [InlineData("ean13", "4006381333931", BarcodeFormat.EAN_13)]
    [InlineData("ean8", "96385074", BarcodeFormat.EAN_8)]
    [InlineData("ean8", "9638507", BarcodeFormat.EAN_8)]
    [InlineData("code39", "INV-2024 0042", BarcodeFormat.CODE_39)]
    [InlineData("code39", "ABC123$/+%", BarcodeFormat.CODE_39)]
    public void Ean_and_code39_decode(string type, string text, BarcodeFormat fmt)
    {
        var expected = type == "code39" ? text : (type == "ean13" ? (text.Length == 12 ? text + "7" : text) : (text.Length == 7 ? text + "4" : text));
        Assert.Equal(expected, Decode(Barcodes.Encode(type, text), fmt));
    }

    [Fact]
    public void Ean_rejects_bad_input()
    {
        Assert.Throws<ArgumentException>(() => Barcodes.Encode("ean13", "5901234123450"));   // wrong check digit
        Assert.Throws<ArgumentException>(() => Barcodes.Encode("ean13", "12345"));
        Assert.Throws<ArgumentException>(() => Barcodes.Encode("ean8", "abcdefgh"));
        Assert.Throws<ArgumentException>(() => Barcodes.Encode("code39", "lower*star"));
    }

    [Fact]
    public void Code128_patterns_are_well_formed_and_bad_input_is_rejected()
    {
        var p = Barcodes.Code128Patterns;
        for (int i = 0; i < p.Length; i++) Assert.Equal(i == 106 ? 13 : 11, p[i].Sum(c => c - '0'));
        Assert.Throws<ArgumentException>(() => Barcodes.Encode("code128", "سلام"));
        Assert.Throws<ArgumentException>(() => Barcodes.Encode("qr", new string('x', 400)));
        Assert.Throws<ArgumentException>(() => Barcodes.Encode("nope", "x"));
    }

    [Fact]
    public void Png_output_is_a_valid_png()
    {
        var png = Barcodes.ToPng(Barcodes.Encode("qr", "x"), 120, 120, out int w, out int h);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, png.Take(8).ToArray());
        Assert.True(w <= 120 && h <= 120 && w == h);
    }

    [Fact]
    public void Barcode_cell_becomes_an_image_in_the_workbook_and_the_xlsx()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s =>
        {
            s.Col(1, 20).Row(1, 90);
            s.Cell("A1", "{{No}}");
            s.Raw(sh => sh.Cells[0].Bc = "qr");
        });
        var t = b.Build();
        var wb = Bind(t, Make.Json("{\"No\":\"INV-7\"}"));
        var img = Assert.Single(wb.Sheets[0].Images);
        Assert.Equal("qr", img.BarType); Assert.Equal("INV-7", img.BarText); Assert.Equal(1, img.Row);
        Assert.Null(wb.Sheets[0].Get(1, 1).Value);        // only the image is shown
        var x = Make.Render(t, Make.Json("{\"No\":\"INV-7\"}"));
        Assert.Contains(x.Raw.Keys, k => k.StartsWith("xl/media/"));
        Assert.NotEmpty(StreamAnalyzer.Check(t));          // barcodes are not streamable
    }

    // ------------------------------------------------------------------ viewer markup + downloads
    [Fact]
    public void Pdf_viewer_markup_embeds_the_pdf_and_escapes_urls()
    {
        var html = ViewerMarkup.PdfViewer("/kxl-download/tok\"en", "/kxl-download/dl", "/kxl-download/x", false);
        Assert.Contains("<iframe", html); Assert.Contains("#view=FitH", html);
        Assert.DoesNotContain("tok\"en", html);
        Assert.Contains("&quot;", html);
        Assert.Contains("Download PDF", ViewerMarkup.PdfViewer("/a", "/b", null, true));
    }

    [Fact]
    public void Download_store_keeps_content_type_and_inline_flag()
    {
        var store = new DownloadStore();
        var tok = store.Add(new byte[] { 1 }, "r.pdf", null, ExcelContent.Pdf, true);
        var item = store.Take(tok, null);
        Assert.Equal(ExcelContent.Pdf, item.ContentType); Assert.True(item.Inline);
        Assert.Equal(ExcelContent.Xlsx, store.Take(store.Add(new byte[] { 1 }, "r.xlsx", null), null).ContentType);
    }

    [Fact]
    public void Pdf_export_requires_the_pdf_package()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => ViewerMarkup.ToPdf(new KscExcelOptions(), new RenderedWorkbook()));
        Assert.Contains("UsePdf", ex.Message);
    }
}
