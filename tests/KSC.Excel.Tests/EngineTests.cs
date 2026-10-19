using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Preview;
using Xunit;

namespace KSC.Excel.Tests;

public class EngineTests
{
    const string Sample = "{\"Title\":\"Report\",\"Items\":[{\"Name\":\"A\",\"Qty\":2,\"Price\":10.5},{\"Name\":\"B\",\"Qty\":3,\"Price\":20},{\"Name\":\"C\",\"Qty\":5,\"Price\":1}]}";

    static ExcelTemplate Basic()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s =>
        {
            s.Cell("A1", "{{Title}}").Merge("A1:C1");
            s.Cell("A2", "Name").Cell("B2", "Qty").Cell("C2", "Price");
            s.Cell("A3", "{{Name}}").Cell("B3", "{{Qty}}").Cell("C3", "{{Price}}");
            s.Band("Items", "Items", 3, 3);
            s.Cell("A4", "Total").Cell("B4", "=SUM([Items.Qty])").Cell("C4", "=SUM([Items.Price])");
        });
        return b.Build();
    }

    [Fact]
    public void Package_is_wellformed_and_complete()
    {
        var x = Make.Render(Basic(), Make.Json(Sample));
        foreach (var part in new[] { "[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml" })
            Assert.True(x.Xml.ContainsKey(part), part);
    }

    [Fact]
    public void Band_repeats_rows_and_totals_cover_generated_range()
    {
        var x = Make.Render(Basic(), Make.Json(Sample));
        Assert.Equal("Report", x.Value("A1"));
        Assert.Equal("A", x.Value("A3")); Assert.Equal("B", x.Value("A4")); Assert.Equal("C", x.Value("A5"));
        Assert.Equal("2", x.Value("B3")); Assert.Equal("10.5", x.Value("C3"));
        Assert.Equal("Total", x.Value("A6"));
        Assert.Equal("SUM($B$3:$B$5)", x.Formula("B6"));
        Assert.Equal("10", x.Value("B6"));          // cached value
        Assert.Equal("31.5", x.Value("C6"));
    }

    [Fact]
    public void Empty_band_removes_rows_and_keeps_formulas_valid()
    {
        var x = Make.Render(Basic(), Make.Json("{\"Title\":\"T\",\"Items\":[]}"));
        Assert.Equal("Total", x.Value("A3"));
        Assert.Equal("SUM(0)", x.Formula("B3"));
    }

    [Fact]
    public void Numbers_are_numbers_and_text_stays_text()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "123").Cell("B1", "0912").Cell("C1", "{{N}}").Cell("D1", "Hi {{N}}").Cell("E1", "{{Flag}}"));
        var x = Make.Render(b.Build(), Make.Json("{\"N\":42,\"Flag\":true}"));
        Assert.Null(x.Cell("A1").Attribute("t"));
        Assert.Equal("123", x.Value("A1"));
        Assert.Equal("s", (string)x.Cell("B1").Attribute("t"));
        Assert.Equal("42", x.Value("C1"));
        Assert.Equal("Hi 42", x.Value("D1"));
        Assert.Equal("b", (string)x.Cell("E1").Attribute("t"));
    }

    [Fact]
    public void Filters_work()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s
            .Cell("A1", "{{D|fa}}").Cell("A2", "{{Name|upper}}").Cell("A3", "{{Missing|default:n/a}}")
            .Cell("A4", "{{Price|round:1}}").Cell("A5", "{{Flag|yesno:بله/خیر}}").Cell("A6", "{{N|fanum}}"));
        var x = Make.Render(b.Build(), Make.Json("{\"D\":\"2025-03-21\",\"Name\":\"ali\",\"Price\":3.14159,\"Flag\":true,\"N\":1234}"));
        Assert.Equal("1404/01/01", x.Value("A1"));
        Assert.Equal("ALI", x.Value("A2"));
        Assert.Equal("n/a", x.Value("A3"));
        Assert.Equal("3.1", x.Value("A4"));
        Assert.Equal("بله", x.Value("A5"));
        Assert.Equal("۱۲۳۴", x.Value("A6"));
    }

    [Fact]
    public void Dates_become_real_dates_with_a_format()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{D}}"));
        var x = Make.Render(b.Build(), Make.Json("{\"D\":\"2025-03-21T00:00:00\"}"));
        // serial for 2025-03-21 = 45737
        Assert.Equal("45737", x.Value("A1"));
        Assert.Contains("yyyy/mm/dd", x.Xml["xl/styles.xml"].ToString());
    }

    [Fact]
    public void Aggregate_placeholders_and_special_variables()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{Count(Items)}} items, sum {{Sum(Items.Qty)}}").Cell("A2", "{{$index}}:{{Name}}").Cell("A3", "{{Title}}/{{Name}}"));
        var t = b.Build();
        t.Sheets[0].Bands.Add(new BandDef { Name = "X", Source = "Items", FirstRow = 2, LastRow = 3 });
        var x = Make.Render(t, Make.Json(Sample));
        Assert.Equal("3 items, sum 10", x.Value("A1"));
        Assert.Equal("1:A", x.Value("A2"));
        Assert.Equal("Report/A", x.Value("A3"));     // parent scope fallback
        Assert.Equal("3:C", x.Value("A6"));
    }

    [Fact]
    public void Nested_bands_and_scoped_formulas()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s
            .Cell("A1", "{{Name}}").Cell("A2", "{{Item}}").Cell("B2", "{{Qty}}").Cell("A3", "sub").Cell("B3", "=SUM([Lines.Qty])"));
        var t = b.Build();
        var sh = t.Sheets[0];
        sh.Bands.Add(new BandDef { Name = "Groups", Source = "Groups", FirstRow = 1, LastRow = 3 });
        sh.Bands.Add(new BandDef { Name = "Lines", Source = "Lines", FirstRow = 2, LastRow = 2 });
        var x = Make.Render(t, Make.Json("{\"Groups\":[{\"Name\":\"G1\",\"Lines\":[{\"Item\":\"a\",\"Qty\":1},{\"Item\":\"b\",\"Qty\":2}]},{\"Name\":\"G2\",\"Lines\":[{\"Item\":\"c\",\"Qty\":10}]}]}"));
        Assert.Equal("G1", x.Value("A1")); Assert.Equal("a", x.Value("A2")); Assert.Equal("b", x.Value("A3"));
        Assert.Equal("SUM($B$2:$B$3)", x.Formula("B4")); Assert.Equal("3", x.Value("B4"));
        Assert.Equal("G2", x.Value("A5"));
        Assert.Equal("SUM($B$6:$B$6)", x.Formula("B7")); Assert.Equal("10", x.Value("B7"));
    }

    [Fact]
    public void Same_row_tokens_and_column_letter_tokens()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{Qty}}").Cell("B1", "{{Price}}").Cell("C1", "=[Qty]*[Price]").Cell("A2", "=SUM([Items:C])"));
        var t = b.Build();
        t.Sheets[0].Bands.Add(new BandDef { Name = "Items", Source = "Items", FirstRow = 1, LastRow = 1 });
        var x = Make.Render(t, Make.Json(Sample));
        Assert.Equal("A1*B1", x.Formula("C1"));
        Assert.Equal("21", x.Value("C1"));
        Assert.Equal("SUM($C$1:$C$3)", x.Formula("A4"));
        Assert.Equal("21", x.Value("C1")); Assert.Equal("60", x.Value("C2")); Assert.Equal("5", x.Value("C3"));
        Assert.Equal("86", x.Value("A4"));
    }

    [Fact]
    public void Merges_replicate_per_band_occurrence()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{Name}}").Merge("A1:B1"));
        var t = b.Build();
        t.Sheets[0].Bands.Add(new BandDef { Name = "Items", Source = "Items", FirstRow = 1, LastRow = 1 });
        var x = Make.Render(t, Make.Json(Sample));
        var merges = x.Sheet().Descendants(Xlsx.Main + "mergeCell").Select(m => (string)m.Attribute("ref")).ToList();
        Assert.Equal(new[] { "A1:B1", "A2:B2", "A3:B3" }, merges);
    }

    [Fact]
    public void Conditional_format_range_follows_band()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{Qty}}").Cell("A2", "end"));
        var t = b.Build();
        var sh = t.Sheets[0];
        sh.Bands.Add(new BandDef { Name = "Items", Source = "Items", FirstRow = 1, LastRow = 1 });
        sh.ConditionalFormats.Add(new CondFormatDef { Range = "A1:A1", Type = "cellIs", Operator = "greaterThan", Formula1 = "2", Style = new StyleDef { Fill = "#FF0000" } });
        sh.ConditionalFormats.Add(new CondFormatDef { Range = "A1:A1", Type = "expression", Formula1 = "A1>2", Style = new StyleDef { Bold = true } });
        var x = Make.Render(t, Make.Json(Sample));
        var cfs = x.Sheet().Descendants(Xlsx.Main + "conditionalFormatting").Select(c => (string)c.Attribute("sqref")).ToList();
        Assert.Equal(new[] { "A1:A3", "A1:A3" }, cfs);
    }

    [Fact]
    public void Repeat_sheet_per_item_with_unique_names()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("{{Name}}", s => s.Cell("A1", "{{Name}}"));
        var t = b.Build();
        t.Sheets[0].RepeatFor = "Items";
        var wb = new TemplateRenderer().Render(t, Make.Json("{\"Items\":[{\"Name\":\"X\"},{\"Name\":\"X\"},{\"Name\":\"bad/name:*\"}]}"));
        Assert.Equal(new[] { "X", "X (2)", "bad_name__" }, wb.Sheets.Select(s => s.Name).ToArray());
    }

    [Fact]
    public void Poco_dictionary_and_datatable_sources()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{Name}}").Cell("B1", "{{Qty}}"));
        var t = b.Build();
        t.Sheets[0].Bands.Add(new BandDef { Name = "Rows", Source = "Rows", FirstRow = 1, LastRow = 1 });

        var table = new System.Data.DataTable();
        table.Columns.Add("Name"); table.Columns.Add("Qty", typeof(int));
        table.Rows.Add("dt", 7);
        var x1 = Make.Render(t, new { Rows = table });
        Assert.Equal("dt", x1.Value("A1")); Assert.Equal("7", x1.Value("B1"));

        var x2 = Make.Render(t, new Dictionary<string, object> { ["Rows"] = new[] { new { Name = "poco", Qty = 3m } } });
        Assert.Equal("poco", x2.Value("A1")); Assert.Equal("3", x2.Value("B1"));
    }

    [Fact]
    public void Row_limit_is_enforced()
    {
        var t = Basic();
        var svc = new KSC.Excel.Services.ExcelReportService(new KSC.Excel.Services.KscExcelOptions { MaxRows = 3 }, new KSC.Excel.Services.MemoryTemplateStore());
        Assert.Throws<ExcelRenderException>(() => svc.Render(t, Make.Json(Sample)));
    }

    [Fact]
    public void Invalid_xml_characters_and_long_text_do_not_corrupt_the_file()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{X}}").Cell("A2", "{{Y}}"));
        var x = Make.Render(b.Build(), Make.Json("{\"X\":\"a\\u0001b\\u000bc\",\"Y\":\"  padded \\n line\"}"));
        Assert.Equal("abc", x.Value("A1"));
        Assert.Equal("  padded \n line", x.Value("A2"));
    }

    [Fact]
    public void Styles_fonts_fills_borders_are_written()
    {
        var b = new TemplateBuilder("T").Font("Tahoma", 10);
        var st = b.Style("h", new StyleDef { Bold = true, Fill = "#2E75B6", Color = "#FFFFFF", H = "center", Wrap = true, Format = "#,##0.00", Bottom = new BorderDef { Style = "thin", Color = "#000000" } });
        b.Sheet("S", s => s.Cell("A1", "1234.5", st), rtl: true);
        var x = Make.Render(b.Build(), Make.Json("{}"));
        var styles = x.Xml["xl/styles.xml"].ToString();
        Assert.Contains("FF2E75B6", styles);
        Assert.Contains("wrapText=\"1\"", styles);
        Assert.Contains("numFmtId=\"4\"", styles);   // built-in #,##0.00
        Assert.Contains("Tahoma", styles);
        Assert.Equal("1", (string)x.Sheet().Descendants(Xlsx.Main + "sheetView").First().Attribute("rightToLeft"));
        Assert.NotNull(x.Cell("A1").Attribute("s"));
    }

    [Fact]
    public void Images_hyperlinks_validations_freeze_are_packaged()
    {
        var png = Convert.ToBase64String(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 0 });
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "go").Freeze(1, 1));
        var t = b.Build();
        var sh = t.Sheets[0];
        sh.Cells[0].L = "https://example.com/{{Id}}";
        sh.Images.Add(new ImageDef { Row = 2, Col = 2, Data = png });
        sh.Validations.Add(new ValidationDef { Range = "B1:B5", Type = "list", Formula1 = "Yes,No" });
        var x = Make.Render(t, Make.Json("{\"Id\":7}"));
        Assert.Contains(x.Raw.Keys, k => k.StartsWith("xl/media/image1"));
        Assert.True(x.Xml.ContainsKey("xl/drawings/drawing1.xml"));
        Assert.Contains("https://example.com/7", x.Xml["xl/worksheets/_rels/sheet1.xml.rels"].ToString());
        Assert.Equal("frozen", (string)x.Sheet().Descendants(Xlsx.Main + "pane").First().Attribute("state"));
        Assert.Equal("\"Yes,No\"", x.Sheet().Descendants(Xlsx.Main + "formula1").First().Value);
    }

    [Fact]
    public void Template_json_roundtrip_is_lossless()
    {
        var t = KSC.Excel.Model.StarterTemplates.Invoice();
        var json = t.ToJson();
        var t2 = ExcelTemplate.FromJson(json);
        Assert.Equal(json, t2.ToJson());
    }

    [Fact]
    public void Newer_schema_is_rejected()
    {
        Assert.Throws<InvalidDataException>(() => ExcelTemplate.FromJson("{\"schemaVersion\":99,\"sheets\":[]}"));
    }

    [Fact]
    public void Starter_templates_render_and_compute()
    {
        var x = Make.Render(StarterTemplates.SimpleTable(), StarterTemplates.SimpleTable().SampleData);
        Assert.Equal("Sales report", x.Value("A1"));
        Assert.Equal("1", x.Value("A5"));
        Assert.Equal("16", x.Value("D8"));
        Assert.Equal("2290000", x.Value("F8"));

        var inv = StarterTemplates.Invoice();
        var y = Make.Render(inv, inv.SampleData);
        Assert.Equal("54000000", y.Value("E8"));
        Assert.Equal("INV-1403-021", y.Value("C3"));
    }

    [Fact]
    public void Formula_evaluator_basics()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "10").Cell("A2", "20").Cell("B1", "=SUM(A1:A2)*2").Cell("B2", "=IF(A1>5,\"big\",\"small\")")
            .Cell("B3", "=A1/0").Cell("B4", "=ROUND(A1/3,2)&\"x\"").Cell("B5", "=SUMIF(A1:A2,\">10\")").Cell("B6", "=COUNT(A1:A2)+AVERAGE(A1:A2)").Cell("B7", "=UNKNOWNFN(1)"));
        var x = Make.Render(b.Build(), null);
        Assert.Equal("60", x.Value("B1"));
        Assert.Equal("big", x.Value("B2"));
        Assert.Equal("#DIV/0!", x.Value("B3"));
        Assert.Equal("3.33x", x.Value("B4"));
        Assert.Equal("20", x.Value("B5"));
        Assert.Equal("17", x.Value("B6"));
        Assert.Null(x.Value("B7"));          // unsupported -> no cached value, Excel recalculates
    }

    [Fact]
    public void Html_preview_shows_values_formulas_and_merges()
    {
        var t = Basic();
        var wb = new TemplateRenderer().Render(t, Make.Json(Sample));
        var html = HtmlPreviewRenderer.Render(wb);
        Assert.Contains("Report", html);
        Assert.Contains("colspan=\"3\"", html);
        Assert.Contains(">31.5<", html);       // computed total
    }

    [Fact]
    public void Number_formatter()
    {
        Assert.Equal("1,234,567", NumberFormatter.Format(1234567.0, "#,##0"));
        Assert.Equal("1,234.50", NumberFormatter.Format(1234.5, "#,##0.00"));
        Assert.Equal("25%", NumberFormatter.Format(0.25, "0%"));
        Assert.Equal("2025/03/21", NumberFormatter.Format(new DateTime(2025, 3, 21), "yyyy/mm/dd"));
        Assert.Equal("14:05", NumberFormatter.Format(new DateTime(2025, 3, 21, 14, 5, 0), "hh:mm"));
        Assert.Equal("-5", NumberFormatter.Format(-5.0, "0"));
        Assert.Equal("$1,000", NumberFormatter.Format(1000.0, "\"$\"#,##0"));
        Assert.Equal("150.", NumberFormatter.Format(150.0, "#,##0.##"));      // Excel keeps the dot
        Assert.Equal("150.5", NumberFormatter.Format(150.5, "#,##0.##"));
        Assert.Equal("1.5", NumberFormatter.Format(1.5, "0.0#"));
    }

    [Fact]
    public void ShiftRows_rebases_relative_references_only()
    {
        Assert.Equal("A5>2", CellRef.ShiftRows("A1>2", 4));
        Assert.Equal("$A$1>2", CellRef.ShiftRows("$A$1>2", 4));
        Assert.Equal("A$1+B3", CellRef.ShiftRows("A$1+B1", 2));
        Assert.Equal("\"A1\"&B3", CellRef.ShiftRows("\"A1\"&B1", 2));
    }

    [Fact]
    public void Column_widths_include_excel_cell_padding()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Col(1, 36).Cell("A1", "x"));
        var x = Make.Render(b.Build(), null);
        var col = x.Sheet().Descendants(Xlsx.Main + "col").First();
        Assert.Equal("36.710938", (string)col.Attribute("width")); // 36 chars + 5px padding, as Excel stores it
    }

    [Fact]
    public void Cell_limit_is_enforced()
    {
        var svc = new KSC.Excel.Services.ExcelReportService(new KSC.Excel.Services.KscExcelOptions { MaxCells = 5 }, new KSC.Excel.Services.MemoryTemplateStore());
        Assert.Throws<ExcelRenderException>(() => svc.Render(Basic(), Make.Json(Sample)));
    }

    [Fact]
    public void Concurrent_renders_do_not_corrupt_each_other()
    {
        var t = Basic();
        var results = new string[16];
        Parallel.For(0, results.Length, i =>
        {
            var data = Make.Json("{\"Title\":\"R" + i + "\",\"Items\":[{\"Name\":\"n" + i + "\",\"Qty\":" + i + ",\"Price\":1}]}");
            var x = Make.Render(t, data);
            results[i] = x.Value("A1") + "/" + x.Value("A3") + "/" + x.Value("B3");
        });
        for (int i = 0; i < results.Length; i++) Assert.Equal("R" + i + "/n" + i + "/" + i, results[i]);
    }

    const string Sales = "{\"Rows\":[{\"Cat\":\"B\",\"Item\":\"b2\",\"Qty\":5},{\"Cat\":\"A\",\"Item\":\"a1\",\"Qty\":1},{\"Cat\":\"B\",\"Item\":\"b1\",\"Qty\":2},{\"Cat\":\"A\",\"Item\":\"a2\",\"Qty\":10}]}";

    [Fact]
    public void Group_by_with_sort_subtotals_and_group_fields()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s
            .Cell("A1", "{{Key}} ({{Count}})")
            .Cell("A2", "{{Item}}").Cell("B2", "{{Qty}}")
            .Cell("A3", "sub").Cell("B3", "=SUM([Detail.Qty])"));
        var t = b.Build();
        var sh = t.Sheets[0];
        sh.Bands.Add(new BandDef { Name = "Groups", Source = "Rows", GroupBy = "Cat", SortBy = "Item", FirstRow = 1, LastRow = 3 });
        sh.Bands.Add(new BandDef { Name = "Detail", Source = "Items", FirstRow = 2, LastRow = 2 });
        var x = Make.Render(t, Make.Json(Sales));
        // sorted by Item => a1,a2,b1,b2 => groups A then B
        Assert.Equal("A (2)", x.Value("A1")); Assert.Equal("a1", x.Value("A2")); Assert.Equal("a2", x.Value("A3"));
        Assert.Equal("SUM($B$2:$B$3)", x.Formula("B4")); Assert.Equal("11", x.Value("B4"));
        Assert.Equal("B (2)", x.Value("A5")); Assert.Equal("b1", x.Value("A6")); Assert.Equal("b2", x.Value("A7"));
        Assert.Equal("7", x.Value("B8"));
    }

    [Fact]
    public void Sort_desc_without_grouping()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{Item}}"));
        var t = b.Build();
        t.Sheets[0].Bands.Add(new BandDef { Name = "R", Source = "Rows", SortBy = "Qty", SortDesc = true, FirstRow = 1, LastRow = 1 });
        var x = Make.Render(t, Make.Json(Sales));
        Assert.Equal(new[] { "a2", "b2", "b1", "a1" }, new[] { x.Value("A1"), x.Value("A2"), x.Value("A3"), x.Value("A4") });
    }

    [Fact]
    public void Chart_ranges_follow_the_band_and_parts_are_written()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "{{Item}}").Cell("B1", "{{Qty}}"));
        var t = b.Build();
        var sh = t.Sheets[0];
        sh.Bands.Add(new BandDef { Name = "R", Source = "Rows", FirstRow = 1, LastRow = 1 });
        sh.Charts.Add(new ChartDef { Name = "c1", Type = "column", Title = "Qty", Row = 2, Col = 4, Categories = "A1:A1", Series = { new ChartSeriesDef { Name = "Qty", Values = "B1:B1", Color = "#2E75B6" } } });
        sh.Charts.Add(new ChartDef { Name = "c2", Type = "pie", Categories = "A1:A1", Series = { new ChartSeriesDef { Name = "Q", Values = "B1:B1" } } });
        var x = Make.Render(t, Make.Json(Sales));
        var c1 = x.Xml["xl/charts/chart1.xml"].ToString();
        Assert.Contains("'S'!$B$1:$B$4", c1);
        Assert.Contains("'S'!$A$1:$A$4", c1);
        Assert.Contains("barChart", c1);
        Assert.Contains("pieChart", x.Xml["xl/charts/chart2.xml"].ToString());
        Assert.Contains("chart1.xml", x.Xml["xl/drawings/_rels/drawing1.xml.rels"].ToString());
        Assert.Contains("graphicFrame", x.Xml["xl/drawings/drawing1.xml"].ToString());
    }

    const string SalesFlat = "{\"Sales\":[{\"Region\":\"North\",\"Product\":\"A\",\"Amt\":10},{\"Region\":\"North\",\"Product\":\"B\",\"Amt\":5},{\"Region\":\"South\",\"Product\":\"A\",\"Amt\":7},{\"Region\":\"North\",\"Product\":\"A\",\"Amt\":1},{\"Region\":\"South\",\"Product\":\"C\",\"Amt\":4}]}";

    static ExcelTemplate CrossTab()
    {
        var b = new TemplateBuilder("X");
        b.Sheet("S", s => s
            .Cell("A1", "Region").Cell("B1", "{{$item.Key}}").Cell("C1", "Total")
            .Cell("A2", "{{Key}}").Cell("B2", "{{Cells[$col].Value}}").Cell("C2", "=SUM([Cols.Value])")
            .Cell("A3", "All").Cell("B3", "=SUM([Rows.Value])").Cell("C3", "=SUM([Rows:C])")
            .Col(2, 9));
        var t = b.Build();
        t.DataSets.Add(new DataSetDef { Name = "Matrix", Source = "Sales", RowPath = "Region", ColPath = "Product", ValuePath = "Amt", Agg = "sum" });
        var sh = t.Sheets[0];
        sh.Bands.Add(new BandDef { Name = "Rows", Source = "Matrix.Rows", FirstRow = 2, LastRow = 2 });
        sh.Bands.Add(new BandDef { Name = "Cols", Source = "Matrix.Columns", Orientation = "columns", FirstCol = 2, LastCol = 2 });
        sh.ConditionalFormats.Add(new CondFormatDef { Range = "B2:B2", Type = "cellIs", Operator = "greaterThan", Formula1 = "5", Style = new StyleDef { Fill = "#FF0000" } });
        return t;
    }

    [Fact]
    public void Cross_tab_with_column_band_pivot_and_totals()
    {
        var x = Make.Render(CrossTab(), Make.Json(SalesFlat));
        // columns: A | B C D (products A,B,C) | E (total)
        Assert.Equal("A", x.Value("B1")); Assert.Equal("B", x.Value("C1")); Assert.Equal("C", x.Value("D1")); Assert.Equal("Total", x.Value("E1"));
        Assert.Equal("North", x.Value("A2")); Assert.Equal("11", x.Value("B2")); Assert.Equal("5", x.Value("C2")); Assert.Null(x.Value("D2"));
        Assert.Equal("South", x.Value("A3")); Assert.Equal("7", x.Value("B3")); Assert.Null(x.Value("C3")); Assert.Equal("4", x.Value("D3"));
        Assert.Equal("SUM($B2:$D2)", x.Formula("E2")); Assert.Equal("16", x.Value("E2")); Assert.Equal("11", x.Value("E3"));
        Assert.Equal("SUM(B2:B3)".Replace("B2:B3", "$B$2:$B$3"), x.Formula("B4")); Assert.Equal("18", x.Value("B4")); Assert.Equal("27", x.Value("E4"));
        var cf = x.Sheet().Descendants(Xlsx.Main + "conditionalFormatting").Select(c => (string)c.Attribute("sqref")).Single();
        Assert.Equal("B2:D3", cf);
        var widths = x.Sheet().Descendants(Xlsx.Main + "col").Select(c => (string)c.Attribute("min")).ToList();
        Assert.Equal(new[] { "2", "3", "4" }, widths);
    }

    [Fact]
    public void Parameters_defaults_and_overrides()
    {
        var b = new TemplateBuilder("P");
        b.Sheet("S", s => s.Cell("A1", "{{$p.Title}}").Cell("B1", "{{$p.Limit}}").Cell("C1", "{{$p.Missing|default:none}}"));
        var t = b.Build();
        t.Parameters.Add(new ParameterDef { Name = "Title", Default = "Default title" });
        t.Parameters.Add(new ParameterDef { Name = "Limit", Type = "number", Default = "5" });
        var svc = Make.Service();
        var x1 = Xlsx.Open(svc.XlsxBytes(t, null, null));
        Assert.Equal("Default title", x1.Value("A1")); Assert.Equal("5", x1.Value("B1")); Assert.Equal("none", x1.Value("C1"));
        var x2 = Xlsx.Open(svc.XlsxBytes(t, null, new Dictionary<string, object> { ["Title"] = "Mine", ["limit"] = 9 }));
        Assert.Equal("Mine", x2.Value("A1")); Assert.Equal("9", x2.Value("B1"));
    }

    [Fact]
    public void Pivot_helper_works_on_pocos_and_aggregations()
    {
        var items = new[] { new { R = "x", C = "a", V = 1 }, new { R = "x", C = "a", V = 3 }, new { R = "y", C = "b", V = 5 } };
        var avg = Pivot.Build(items, "R", "C", "V", "avg");
        Assert.Equal(2.0, avg.Rows[0].Cells[0].Value);
        Assert.Null(avg.Rows[0].Cells[1].Value);
        var cnt = Pivot.Build(items, "R", "C", "V", "count");
        Assert.Equal(3.0, cnt.GrandTotal);
    }
}

static class ServiceTestExt
{
    public static byte[] XlsxBytes(this KSC.Excel.Services.IExcelReportService svc, ExcelTemplate t, object data, IDictionary<string, object> p)
        => KSC.Excel.Writer.XlsxWriter.ToBytes(svc.Bind(t, data, p));
}

public class StreamingTests
{
    const string Sample = "{\"Items\":[{\"Qty\":2},{\"Qty\":3},{\"Qty\":5}]}";
    const string Sales = "{\"Rows\":[{\"Qty\":5},{\"Qty\":1}]}";

    static IEnumerable<object> Lazy(int n, Action<int> onItem = null)
    {
        for (int i = 1; i <= n; i++) { onItem?.Invoke(i); yield return new { Name = "Item " + i, Category = "C" + (i % 3), Qty = i % 7, Price = 10m * i, Total = (decimal)(i * 3) }; }
    }

    [Fact]
    public void Streaming_matches_buffered_output_cell_for_cell()
    {
        var t = StarterTemplates.SimpleTable();
        var data = new { Title = "T", Period = "P", Items = Lazy(25).ToList() };
        var svc = Make.Service();
        var buffered = Xlsx.Open(svc.Render(t, data));
        var ms = new MemoryStream();
        Assert.True(svc.RenderStreaming(t, data, ms));
        var streamed = Xlsx.Open(ms.ToArray());                      // also proves the package is well-formed
        foreach (var addr in new[] { "A1", "B4", "B5", "C5", "D5", "F5", "B29", "A30" })
            Assert.Equal(buffered.Value(addr), streamed.Value(addr));
        Assert.Equal(buffered.Formula("D30"), streamed.Formula("D30"));
        Assert.Equal(buffered.Xml["xl/worksheets/sheet1.xml"].Descendants(Xlsx.Main + "mergeCell").Count(), streamed.Xml["xl/worksheets/sheet1.xml"].Descendants(Xlsx.Main + "mergeCell").Count());
        Assert.Equal("frozen", (string)streamed.Sheet().Descendants(Xlsx.Main + "pane").First().Attribute("state"));
        Assert.Null(streamed.Cell("D30").Element(Xlsx.Main + "v"));    // no cached value while streaming
    }

    [Fact]
    public void Streaming_does_not_materialize_the_collection()
    {
        var t = StarterTemplates.SimpleTable();
        int produced = 0, maxAhead = 0, written = 0;
        var items = Lazy(1000, i => { produced = i; maxAhead = Math.Max(maxAhead, produced - written); });
        var data = new { Title = "T", Period = "P", Items = items };
        var svc = Make.Service();
        using (var counting = new CountingStream(n => written = (int)(n / 60)))
            svc.RenderStreaming(t, data, counting);
        Assert.Equal(1000, produced);
        Assert.True(maxAhead < 1000, "all items were produced before anything was written");
    }

    [Fact]
    public void Unstreamable_templates_fall_back_and_the_reasons_are_reported()
    {
        var b = new TemplateBuilder("T");
        b.Sheet("S", s => s.Cell("A1", "=SUM([Items.Qty])").Cell("A2", "{{Qty}}"));
        var t = b.Build();
        t.Sheets[0].Bands.Add(new BandDef { Name = "Items", Source = "Items", FirstRow = 2, LastRow = 2 });
        Assert.NotEmpty(StreamAnalyzer.Check(t));                    // total above the band
        var svc = Make.Service();
        var ms = new MemoryStream();
        Assert.False(svc.RenderStreaming(t, Make.Json(Sample), ms));
        Assert.Equal("SUM($A$2:$A$4)", Xlsx.Open(ms.ToArray()).Formula("A1"));   // buffered fallback still correct
        Assert.Empty(StreamAnalyzer.Check(StarterTemplates.SimpleTable()));
    }

    [Fact]
    public void Zebra_fill_now_applies_per_cell_and_still_skips_explicit_fills()
    {
        var b = new TemplateBuilder("Z");
        var filled = b.Style("f", new StyleDef { Fill = "#00FF00" });
        b.Sheet("S", s => s.Cell("A1", "{{Qty}}").Cell("B1", "{{Qty}}", filled));
        var t = b.Build();
        t.Sheets[0].Bands.Add(new BandDef { Name = "R", Source = "Rows", FirstRow = 1, LastRow = 1, AlternateFill = "#EEEEEE" });
        var wb = new TemplateRenderer().Render(t, Make.Json(Sales));
        Assert.Null(wb.Sheets[0].Get(1, 1).Style);                    // odd occurrence
        Assert.Equal("#EEEEEE", wb.Sheets[0].Get(2, 1).Style.Fill);   // even occurrence
        Assert.Equal("#00FF00", wb.Sheets[0].Get(2, 2).Style.Fill);   // explicit fill kept
    }
}

sealed class CountingStream : Stream
{
    readonly Action<long> _on; long _n;
    public CountingStream(Action<long> on) { _on = on; }
    public override bool CanRead => false; public override bool CanSeek => false; public override bool CanWrite => true;
    public override long Length => _n; public override long Position { get => _n; set { } }
    public override void Flush() { }
    public override int Read(byte[] b, int o, int c) => 0; public override long Seek(long o, SeekOrigin s) => 0; public override void SetLength(long v) { }
    public override void Write(byte[] b, int o, int c) { _n += c; _on(_n); }

}

public class ImportTests
{
    static ExcelTemplate Sample()
    {
        var b = new TemplateBuilder("Src", "Src").Font("Tahoma", 10);
        var h = b.Style("h", new StyleDef { Bold = true, Fill = "#2E75B6", Color = "#FFFFFF", H = "center", Format = "#,##0.00", Bottom = new BorderDef { Style = "medium", Color = "#1F4E79" }, Wrap = true, Indent = 1 });
        b.Sheet("گزارش", s => s.Col(1, 30).Col(2, 14).Row(1, 30)
            .Cell("A1", "عنوان", h).Cell("B1", "1234.5", h).Cell("A2", "0912").Cell("B2", "7").Cell("B3", "=B2*2").Cell("A3", "x").Merge("A5:B5").Cell("A5", "ادغام")
            .Freeze(1).AutoFilter("A1:B1"), rtl: true);
        var t = b.Build(); var sh = t.Sheets[0];
        sh.Validations.Add(new ValidationDef { Range = "A2:A4", Type = "list", Formula1 = "بله,خیر" });
        sh.ConditionalFormats.Add(new CondFormatDef { Range = "B2:B3", Type = "cellIs", Operator = "greaterThan", Formula1 = "5", Style = new StyleDef { Fill = "#FFC7CE", Color = "#9C0006", Bold = true } });
        sh.Cells.First(c => c.R == 3 && c.C == 1).L = "https://example.com";
        sh.Page.Orientation = "landscape"; sh.Page.PrintTitleRows = "1:1";
        sh.Images.Add(new ImageDef { Row = 7, Col = 2, Width = 40, Height = 30, Data = Convert.ToBase64String(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 0 }) });
        return t;
    }

    [Fact]
    public void Roundtrip_xlsx_to_template_keeps_values_styles_and_structure()
    {
        var bytes = Make.Service().Render(Sample(), null);
        var imp = KSC.Excel.Import.XlsxImporter.Import(bytes, "Imported");
        var t = imp.Template; var sh = t.Sheets[0];
        Assert.Equal("گزارش", sh.Name); Assert.True(sh.Rtl);
        Assert.Equal("Tahoma", t.DefaultFont);
        string V(int r, int c) => sh.Cells.FirstOrDefault(x => x.R == r && x.C == c)?.V;
        Assert.Equal("عنوان", V(1, 1)); Assert.Equal("1234.5", V(1, 2)); Assert.Equal("0912", V(2, 1)); Assert.Equal("=B2*2", V(3, 2));
        Assert.Equal("text", sh.Cells.First(x => x.R == 2 && x.C == 1).T);                     // leading zero stays text
        var st = t.Styles[sh.Cells.First(x => x.R == 1 && x.C == 1).S];
        Assert.True(st.Bold); Assert.Equal("#2E75B6", st.Fill); Assert.Equal("#FFFFFF", st.Color); Assert.Equal("center", st.H); Assert.True(st.Wrap);
        Assert.Equal("#,##0.00", st.Format); Assert.Equal("medium", st.Bottom.Style); Assert.Equal(1, st.Indent);
        Assert.Contains("A5:B5", sh.Merges); Assert.Equal(1, sh.FreezeRows); Assert.Equal("A1:B1", sh.AutoFilter);
        Assert.Equal(30, sh.Columns.First(c => c.Index == 1).Width.Value, 1);                  // width round-trips (padding removed)
        Assert.Equal(30, sh.Rows.First(r => r.Index == 1).Height.Value, 1);
        Assert.Equal("list", sh.Validations[0].Type); Assert.Equal("بله,خیر", sh.Validations[0].Formula1);
        Assert.Equal("cellIs", sh.ConditionalFormats[0].Type); Assert.Equal("#FFC7CE", sh.ConditionalFormats[0].Style.Fill);
        Assert.Equal("https://example.com", sh.Cells.First(x => x.R == 3 && x.C == 1).L);
        Assert.Equal("landscape", sh.Page.Orientation); Assert.Equal("1:1", sh.Page.PrintTitleRows);
        Assert.Single(sh.Images); Assert.Equal(7, sh.Images[0].Row); Assert.Equal(2, sh.Images[0].Col);

        // and the imported template renders again
        var again = Xlsx.Open(Make.Service().Render(t, null));
        Assert.Equal("عنوان", again.Value("A1")); Assert.Equal("14", again.Value("B3"));
    }

    [Fact]
    public void Garbage_and_oversized_input_is_rejected_cleanly()
    {
        Assert.Throws<InvalidDataException>(() => KSC.Excel.Import.XlsxImporter.Import(new byte[0], "x"));
        Assert.ThrowsAny<Exception>(() => KSC.Excel.Import.XlsxImporter.Import(new byte[] { 1, 2, 3, 4 }, "x"));
    }

    [Fact]
    public void Section_page_header_becomes_print_titles()
    {
        var b = new TemplateBuilder("S");
        b.Sheet("S", s => s.Cell("A1", "title").Cell("A2", "{{N}}"));
        var t = b.Build(); var sh = t.Sheets[0];
        sh.Bands.Add(new BandDef { Name = "R", Source = "Rows", FirstRow = 2, LastRow = 2 });
        sh.Sections.Add(new SectionDef { Kind = "pageHeader", FirstRow = 1, LastRow = 1 });
        var wb = new TemplateRenderer().Render(t, Make.Json("{\"Rows\":[{\"N\":1},{\"N\":2},{\"N\":3}]}"));
        Assert.Equal("1:1", wb.Sheets[0].PrintTitleRows);
    }
}
