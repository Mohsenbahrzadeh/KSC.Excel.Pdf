using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Services;
using Newtonsoft.Json.Linq;
using Xunit;

namespace KSC.Excel.Tests;

/// <summary>Multi-level grouping with subtotals, Jalali calendar, calculated columns.</summary>
public class AdvancedReportTests
{
    static RenderedWorkbook Bind(ExcelTemplate t, object data, IDictionary<string, object> p = null)
        => new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore()).Bind(t, data, p);

    static object V(RenderedWorkbook wb, int r, int c) => wb.Sheets[0].Get(r, c)?.Value;
    static object Cached(RenderedWorkbook wb, int r, int c) { var x = wb.Sheets[0].Get(r, c); return x?.Cached; }

    // ------------------------------------------------------------------ two-level group with SUBTOTAL (what the wizard builds)
    static ExcelTemplate TwoLevel()
    {
        var b = new TemplateBuilder("G");
        b.Sheet("S", s =>
        {
            s.Cell("A1", "Cat").Cell("B1", "Amount");
            s.Cell("A2", "{{Key}}");                                         // level 1 header
            s.Cell("A3", "{{Key}}");                                         // level 2 header
            s.Cell("A4", "{{Cat}}").Cell("B4", "{{Amount}}");                // detail
            s.Cell("A5", "month total").Cell("B5", "=SUBTOTAL(9,[Detail.Amount])");
            s.Cell("A6", "year total").Cell("B6", "=SUBTOTAL(9,[Detail.Amount])");
            s.Cell("A7", "grand").Cell("B7", "=SUBTOTAL(9,[Detail.Amount])");
            s.Raw(sh =>
            {
                sh.Bands.Add(new BandDef { Name = "Groups", Source = "Rows", GroupBy = "=JYEAR(Date)", SortBy = "=JYEAR(Date)", FirstRow = 2, LastRow = 6 });
                sh.Bands.Add(new BandDef { Name = "SubGroups", Source = "Items", GroupBy = "=JMONTH(Date)", SortBy = "=JMONTH(Date)", FirstRow = 3, LastRow = 5 });
                sh.Bands.Add(new BandDef { Name = "Detail", Source = "Items", FirstRow = 4, LastRow = 4 });
            });
        });
        return b.Build();
    }

    const string Sales = "{\"Rows\":[" +
        "{\"Date\":\"2025-03-25\",\"Cat\":\"x\",\"Amount\":10}," +    // 1404/01/05
        "{\"Date\":\"2025-04-26\",\"Cat\":\"y\",\"Amount\":7}," +     // 1404/02/06
        "{\"Date\":\"2025-03-30\",\"Cat\":\"z\",\"Amount\":20}," +    // 1404/01/10
        "{\"Date\":\"2026-03-25\",\"Cat\":\"w\",\"Amount\":100}," +   // 1405/01/05
        "{\"Date\":\"2025-04-25\",\"Cat\":\"v\",\"Amount\":5}" +      // 1404/02/05
        "]}";

    [Fact]
    public void Two_level_grouping_by_jalali_expressions_has_correct_nested_subtotals()
    {
        var wb = Bind(TwoLevel(), JObject.Parse(Sales));
        // 1404 / month 1: 10, 20 -> 30 ; month 2: 7, 5 -> 12 ; year 42 ; 1405 / month 1: 100 ; grand 142
        Assert.Equal("1404", Convert.ToString(V(wb, 2, 1)));
        Assert.Equal(30d, Cached(wb, 6, 2));
        Assert.Equal(12d, Cached(wb, 10, 2));
        Assert.Equal(42d, Cached(wb, 11, 2));
        Assert.Equal("1405", Convert.ToString(V(wb, 12, 1)));
        Assert.Equal(100d, Cached(wb, 15, 2));
        Assert.Equal(100d, Cached(wb, 16, 2));
        Assert.Equal(142d, Cached(wb, 17, 2));
        Assert.Equal("grand", V(wb, 17, 1));
        var dump = Environment.GetEnvironmentVariable("KXL_DUMP_DIR");
        if (!string.IsNullOrEmpty(dump)) File.WriteAllBytes(Path.Combine(dump, "twolevel.xlsx"), new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore()).Render(TwoLevel(), JObject.Parse(Sales)));
    }

    [Fact]
    public void Subtotal_formulas_are_written_so_excel_ignores_nested_subtotals()
    {
        var x = Make.Render(TwoLevel(), JObject.Parse(Sales));
        var f = x.Formula("B17");
        Assert.StartsWith("SUBTOTAL(9,$B$", f);
        // the grand total range spans every detail row (rows 4..14), month/year subtotals inside it are ignored by SUBTOTAL
        Assert.Equal("SUBTOTAL(9,$B$4:$B$14)", f);
    }

    // ------------------------------------------------------------------ calculated columns
    [Fact]
    public void Running_sum_percent_previous_diff_rank_and_moving_average()
    {
        var b = new TemplateBuilder("C");
        b.Sheet("S", s =>
        {
            s.Cell("A1", "{{Qty}}").Cell("B1", "{{RUNSUM(Qty)}}").Cell("C1", "{{PCT(Qty)}}").Cell("D1", "{{PREV(Qty)}}")
             .Cell("E1", "{{DIFF(Qty)}}").Cell("F1", "{{RANK(Qty)}}").Cell("G1", "{{MOVAVG(Qty, 2)}}");
            s.Band("R", "Rows", 1, 1);
        });
        var wb = Bind(b.Build(), JObject.Parse("{\"Rows\":[{\"Qty\":10},{\"Qty\":20},{\"Qty\":30},{\"Qty\":40}]}"));
        Assert.Equal(new object[] { 10d, 30d, 60d, 100d }, Enumerable.Range(1, 4).Select(r => V(wb, r, 2)).ToArray());
        Assert.Equal(new object[] { 0.1, 0.2, 0.3, 0.4 }, Enumerable.Range(1, 4).Select(r => (object)Math.Round((double)V(wb, r, 3), 6)).ToArray());
        Assert.Null(V(wb, 1, 4)); Assert.Equal(30d, V(wb, 4, 4));
        Assert.Null(V(wb, 1, 5)); Assert.Equal(10d, V(wb, 3, 5));
        Assert.Equal(new object[] { 4d, 3d, 2d, 1d }, Enumerable.Range(1, 4).Select(r => V(wb, r, 6)).ToArray());
        Assert.Equal(new object[] { 10d, 15d, 25d, 35d }, Enumerable.Range(1, 4).Select(r => V(wb, r, 7)).ToArray());
    }

    [Fact]
    public void Percent_of_total_works_inside_groups_relative_to_the_group()
    {
        var b = new TemplateBuilder("C");
        b.Sheet("S", s =>
        {
            s.Cell("A1", "{{Key}}");
            s.Cell("A2", "{{Qty}}").Cell("B2", "{{PCT(Qty)}}");
            s.Raw(sh =>
            {
                sh.Bands.Add(new BandDef { Name = "G", Source = "Rows", GroupBy = "Cat", FirstRow = 1, LastRow = 2 });
                sh.Bands.Add(new BandDef { Name = "D", Source = "Items", FirstRow = 2, LastRow = 2 });
            });
        });
        var wb = Bind(b.Build(), JObject.Parse("{\"Rows\":[{\"Cat\":\"a\",\"Qty\":1},{\"Cat\":\"a\",\"Qty\":3},{\"Cat\":\"b\",\"Qty\":5}]}"));
        Assert.Equal(0.25, (double)V(wb, 2, 2), 6); Assert.Equal(0.75, (double)V(wb, 3, 2), 6); Assert.Equal(1.0, (double)V(wb, 5, 2), 6);
    }

    // ------------------------------------------------------------------ Jalali calendar
    [Fact]
    public void Jalali_functions_and_filters()
    {
        var b = new TemplateBuilder("J");
        b.Sheet("S", s => s
            .Cell("A1", "{{JYEAR(D)}}").Cell("B1", "{{JMONTH(D)}}").Cell("C1", "{{JDAY(D)}}").Cell("D1", "{{JMONTHNAME(D)}}")
            .Cell("E1", "{{JYM(D)}}").Cell("F1", "{{D|jmonthname}}").Cell("G1", "{{JQUARTER(D)}}").Cell("H1", "{{JDATE(1404,1,1)|fa}}")
            .Cell("I1", "{{JPARSE('1404/02/10')|fa}}"));
        var wb = Bind(b.Build(), JObject.Parse("{\"D\":\"2025-04-26\"}"));   // 1404/02/06
        Assert.Equal(1404d, V(wb, 1, 1)); Assert.Equal(2d, V(wb, 1, 2)); Assert.Equal(6d, V(wb, 1, 3));
        Assert.Equal("اردیبهشت", V(wb, 1, 4)); Assert.Equal("1404/02", V(wb, 1, 5)); Assert.Equal("اردیبهشت", V(wb, 1, 6));
        Assert.Equal(1d, V(wb, 1, 7)); Assert.Equal("1404/01/01", V(wb, 1, 8)); Assert.Equal("1404/02/10", V(wb, 1, 9));
    }

    [Fact]
    public void Jalali_parameter_accepts_persian_text_and_digits()
    {
        var b = new TemplateBuilder("P");
        b.Sheet("S", s => s.Cell("A1", "{{$p.From|fa}}").Cell("B1", "{{$p.To|fa}}"));
        var t = b.Build();
        t.Parameters.Add(new ParameterDef { Name = "From", Type = "jdate", Default = "1404/01/05" });
        t.Parameters.Add(new ParameterDef { Name = "To", Type = "jdate", Default = "1404/12/29" });
        var wb = Bind(t, null, new Dictionary<string, object> { ["To"] = "۱۴۰۴/۰۲/۱۰" });
        Assert.Equal("1404/01/05", V(wb, 1, 1));
        Assert.Equal("1404/02/10", V(wb, 1, 2));
    }

    [Fact]
    public void Jalali_parse_rejects_invalid_dates()
    {
        DateTime d;
        Assert.False(Jalali.TryParse("1404/13/01", out d));
        Assert.False(Jalali.TryParse("2025/03/25", out d));      // Gregorian-looking
        Assert.False(Jalali.TryParse("abc", out d));
        Assert.True(Jalali.TryParse("1403/12/30", out d));       // 1403 is a leap year
        Assert.False(Jalali.TryParse("1404/12/30", out d));
    }

    // ------------------------------------------------------------------ combo chart
    static ExcelTemplate Combo()
    {
        var b = new TemplateBuilder("Ch");
        b.Sheet("S", s => { s.Cell("A1", "{{M}}").Cell("B1", "{{Amount}}").Cell("C1", "{{Pct}}"); s.Band("R", "Rows", 1, 1); });
        var t = b.Build();
        t.Sheets[0].Charts.Add(new ChartDef
        {
            Name = "c", Type = "column", Title = "combo", Row = 8, Col = 1, Categories = "A1:A1",
            Series = { new ChartSeriesDef { Name = "Amount", Values = "B1:B1" }, new ChartSeriesDef { Name = "Pct", Values = "C1:C1", Type = "line", Secondary = true } }
        });
        return t;
    }

    const string ComboData = "{\"Rows\":[{\"M\":\"a\",\"Amount\":100,\"Pct\":0.1},{\"M\":\"b\",\"Amount\":250,\"Pct\":0.4},{\"M\":\"c\",\"Amount\":180,\"Pct\":0.3}]}";

    [Fact]
    public void Combo_chart_writes_bar_and_line_with_a_secondary_axis()
    {
        var x = Make.Render(Combo(), JObject.Parse(ComboData));
        var chart = x.Xml.Keys.First(k => k.StartsWith("xl/charts/chart")); var xml = x.Xml[chart].ToString();
        Assert.Contains("c:barChart", xml); Assert.Contains("c:lineChart", xml);
        Assert.Contains("444444", xml);                              // secondary value axis
        Assert.Contains("<c:crosses val=\"max\"", xml.Replace(" />", "/>").Replace("\"/>", "\"/>"));
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(xml, "<c:ser>").Count);
        var dump = Environment.GetEnvironmentVariable("KXL_DUMP_DIR");
        if (!string.IsNullOrEmpty(dump)) File.WriteAllBytes(Path.Combine(dump, "combo.xlsx"), new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore()).Render(Combo(), JObject.Parse(ComboData)));
    }

    [Fact]
    public void Series_type_roundtrips_in_template_json()
    {
        var t = ExcelTemplate.FromJson(Combo().ToJson());
        var s = t.Sheets[0].Charts[0].Series;
        Assert.Equal("line", s[1].Type); Assert.True(s[1].Secondary); Assert.Null(s[0].Type);
    }
}
