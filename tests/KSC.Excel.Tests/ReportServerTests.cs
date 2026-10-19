using System.Text;
using KSC.Excel.Designer;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Services;
using Newtonsoft.Json.Linq;
using Xunit;

namespace KSC.Excel.Tests;

public class ReportServerTests
{
    static (ReportServer server, KscExcelOptions opt, IDictionary<string, object>[] seen) Setup(bool withPdf = false)
    {
        var store = new MemoryTemplateStore();
        var b = new TemplateBuilder("Sales", "گزارش فروش");
        b.Sheet("S", s =>
        {
            s.Cell("A1", "{{$p.Title}}").Cell("B1", "{{$p.Min}}").Cell("C1", "{{$p.From|fa}}");
            s.Cell("A2", "{{Name}}").Cell("B2", "{{Qty}}");
            s.Band("R", "Rows", 2, 2);
            s.Cell("C2", "{{Name}}", null);
            s.Raw(sh => sh.Cells.First(c => c.R == 2 && c.C == 3).L = "report:Detail?Id={{Qty}}&Name={{Name|urlenc}}");
        });
        var t = b.Build();
        t.Parameters.Add(new ParameterDef { Name = "Title", Type = "text", Default = "Default title", Label = "عنوان" });
        t.Parameters.Add(new ParameterDef { Name = "Min", Type = "number", Default = "0" });
        t.Parameters.Add(new ParameterDef { Name = "From", Type = "jdate", Default = "1404/01/05" });
        t.Parameters.Add(new ParameterDef { Name = "Kind", Type = "text", Default = "a", Options = "a:Alpha|b:Beta" });
        t.Parameters.Add(new ParameterDef { Name = "Active", Type = "bool", Default = "true" });
        store.Save(t);
        var opt = new KscExcelOptions();
        var seen = new IDictionary<string, object>[1];
        opt.AddReport("Sales", p => { seen[0] = p; return new { Rows = new[] { new { Name = "a b", Qty = 5 }, new { Name = "c&d", Qty = 7 } } }; }, title: "Sales report");
        if (withPdf) { opt.PdfExporter = wb => new byte[] { 37, 80, 68, 70 }; opt.PngExporter = (wb, p) => new byte[] { 1 }; }
        return (new ReportServer(opt, new ExcelReportService(opt, store)), opt, seen);
    }

    static string Body(DesignerResponse r) => Encoding.UTF8.GetString(r.Body);

    static ReportRequest Req(string name, params string[] kv)
    {
        var r = new ReportRequest { Name = name, BasePath = "/reports" };
        for (int i = 0; i + 1 < kv.Length; i += 2) r.Query[kv[i]] = kv[i + 1];
        return r;
    }

    [Fact]
    public void Catalog_lists_registered_reports_only()
    {
        var (srv, _, _) = Setup();
        var html = Body(srv.Handle(Req(null)));
        Assert.Contains("Sales report", html); Assert.Contains("href=\"/reports/Sales\"", html);
        Assert.Equal(404, srv.Handle(Req("Nope")).Status);
        Assert.Equal(404, srv.Handle(Req("../x")).Status);
    }

    [Fact]
    public void Query_parameters_are_typed_and_reach_the_data_provider_and_the_template()
    {
        var (srv, _, seen) = Setup();
        var r = srv.Handle(Req("Sales", "p_Title", "My <b>title</b>", "p_Min", "12.5", "p_From", "۱۴۰۴/۰۲/۱۰", "p_Active", "0"));
        Assert.Equal(200, r.Status);
        Assert.Equal(12.5, seen[0]["Min"]); Assert.Equal(false, seen[0]["Active"]); Assert.Equal("My <b>title</b>", seen[0]["Title"]);
        var html = Body(r);
        Assert.Contains("My &lt;b&gt;title&lt;/b&gt;", html);              // encoded, never injected
        Assert.DoesNotContain("My <b>title</b>", html);
        Assert.Contains("1404/02/10", html);                                  // jdate converted
    }

    [Fact]
    public void Parameter_form_shows_current_values_labels_dropdowns_and_checkboxes()
    {
        var (srv, _, _) = Setup();
        var html = Body(srv.Handle(Req("Sales", "p_Kind", "b", "p_Min", "9")));
        Assert.Contains("<form method=\"get\" action=\"/reports/Sales\"", html);
        Assert.Contains("name=\"p_Title\" value=\"Default title\"", html);
        Assert.Contains("name=\"p_Min\" value=\"9\"", html); Assert.Contains("type=\"number\"", html);
        Assert.Contains("<option value=\"b\" selected>Beta</option>", html);
        Assert.Contains("type=\"checkbox\" name=\"p_Active\" value=\"1\" checked", html);
        Assert.Contains("placeholder=\"1404/01/01\"", html);
        Assert.Contains("عنوان", html);
    }

    [Fact]
    public void Drill_down_links_point_to_the_report_server_with_encoded_values()
    {
        var (srv, _, _) = Setup();
        var html = Body(srv.Handle(Req("Sales")));
        Assert.Contains("href=\"/reports/Detail?p_Id=5&amp;p_Name=a%20b\"", html);
        Assert.Contains("p_Name=c%26d", html);
        Assert.Equal("about:blank", ReportLinks.Resolve("report:../../x?a=1", "/reports"));
        Assert.Equal("https://x", ReportLinks.Resolve("https://x", "/reports"));
        Assert.Equal("report:Detail", ReportLinks.Resolve("report:Detail", null));    // no server: untouched
    }

    [Fact]
    public void Formats_are_served_with_the_right_headers_and_unavailable_ones_fail_cleanly()
    {
        var (srv, _, _) = Setup(false);
        foreach (var (fmt, ct, disp) in new[] { ("xlsx", ExcelContent.Xlsx, "attachment"), ("docx", ReportExport.DocxContentType, "attachment"), ("csv", "text/csv; charset=utf-8", "attachment"), ("html", "text/html; charset=utf-8", "inline") })
        {
            var r = srv.Handle(Req("Sales", "format", fmt));
            Assert.Equal(200, r.Status); Assert.Equal(ct, r.ContentType);
            Assert.StartsWith(disp, r.Headers["Content-Disposition"]);
            Assert.True(r.Body.Length > 20);
        }
        Assert.Equal(500, srv.Handle(Req("Sales", "format", "pdf")).Status);     // no PDF package: generic error, no details leaked
        Assert.Equal(400, srv.Handle(Req("Sales", "format", "exe")).Status);

        var (srv2, _, _) = Setup(true);
        var pdf = srv2.Handle(Req("Sales", "format", "pdf"));
        Assert.StartsWith("inline", pdf.Headers["Content-Disposition"]);
        Assert.StartsWith("attachment", srv2.Handle(Req("Sales", "format", "pdf", "download", "1")).Headers["Content-Disposition"]);
        Assert.Equal("image/png", srv2.Handle(Req("Sales", "format", "png")).ContentType);
        var page = Body(srv2.Handle(Req("Sales", "view", "pdf", "p_Min", "3")));
        Assert.Contains("<iframe src=\"/reports/Sales?p_Min=3&amp;format=pdf#view=FitH\"", page);
        Assert.Contains("name=\"view\" value=\"pdf\"", page);                       // form keeps the PDF view
    }

    [Fact]
    public void Errors_are_reported_to_the_hook_but_not_to_the_browser()
    {
        var store = new MemoryTemplateStore();
        var b = new TemplateBuilder("Boom"); b.Sheet("S", s => s.Cell("A1", "x")); store.Save(b.Build());
        var opt = new KscExcelOptions(); Exception logged = null;
        opt.OnReportError = (n, e) => logged = e;
        opt.AddReport("Boom", p => throw new InvalidOperationException("secret connection string"));
        var srv = new ReportServer(opt, new ExcelReportService(opt, store));
        var r = srv.Handle(Req("Boom"));
        Assert.Equal(500, r.Status);
        Assert.DoesNotContain("secret", Body(r));
        Assert.Contains("secret", logged.Message);
        Assert.Equal(404, new ReportServer(opt, new ExcelReportService(opt, new MemoryTemplateStore())).Handle(Req("Boom")).Status);   // template missing
    }

    [Fact]
    public void Parameter_parsing_ignores_garbage()
    {
        var t = new TemplateBuilder("T").Build();
        t.Parameters.Add(new ParameterDef { Name = "N", Type = "number" }); t.Parameters.Add(new ParameterDef { Name = "D", Type = "date" });
        var q = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["p_N"] = "abc", ["p_D"] = "not a date", ["p_Other"] = "x" };
        Assert.Empty(ReportParameters.FromQuery(t, q));
        q["p_N"] = "1e3"; q["p_D"] = "2025-03-01";
        var r = ReportParameters.FromQuery(t, q);
        Assert.Equal(1000d, r["N"]); Assert.Equal(new DateTime(2025, 3, 1), r["D"]);
        Assert.Equal("?p_N=1e3&p_D=2025-03-01&x=1", ReportParameters.QueryString(q.Where(k => k.Key != "p_Other").ToDictionary(k => k.Key, k => k.Value), "x=1"));
    }

    [Fact]
    public void Urlenc_filter_and_same_site_links_in_the_html_preview()
    {
        var b = new TemplateBuilder("U");
        b.Sheet("S", s => s.Cell("A1", "go").Cell("B1", "{{V|urlenc}}"));
        var t = b.Build(); t.Sheets[0].Cells[0].L = "/reports/X";
        var wb = new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore()).Bind(t, JObject.Parse("{\"V\":\"a b/c\"}"), null);
        Assert.Equal("a%20b%2Fc", wb.Sheets[0].Get(1, 2).Value);
        var html = KSC.Excel.Preview.HtmlPreviewRenderer.Render(wb);
        Assert.Contains("href=\"/reports/X\"", html);
        t.Sheets[0].Cells[0].L = "//evil.example/x";
        Assert.Contains("about:blank", KSC.Excel.Preview.HtmlPreviewRenderer.Render(new ExcelReportService(new KscExcelOptions(), new MemoryTemplateStore()).Bind(t, null, null)));
    }
}
