using System.Text;
using KSC.Excel.Designer;
using KSC.Excel.Model;
using KSC.Excel.Services;
using Newtonsoft.Json.Linq;
using Xunit;

namespace KSC.Excel.Tests;

public class DesignerHostTests
{
    class Order { public string Number { get; set; } public DateTime Date { get; set; } public List<Line> Lines { get; set; } public Customer Customer { get; set; } }
    class Line { public string Product { get; set; } public int Qty { get; set; } public decimal Price { get; set; } }
    class Customer { public string Name { get; set; } }

    static (DesignerHost host, MemoryTemplateStore store) Create(bool readOnly = false)
    {
        var opt = new KscExcelOptions { ReadOnlyDesigner = readOnly }.AddModel<Order>("Order");
        var store = new MemoryTemplateStore();
        return (new DesignerHost(opt, new ExcelReportService(opt, store)), store);
    }

    static DesignerRequest Req(string method, string path, string body = null, bool csrf = true)
    {
        var r = new DesignerRequest { Method = method, Path = path, BasePath = "/excel-designer", Body = body == null ? null : Encoding.UTF8.GetBytes(body) };
        if (csrf) r.Headers["X-KXL-Request"] = "1";
        return r;
    }

    static JToken Parse(DesignerResponse r) => JToken.Parse(Encoding.UTF8.GetString(r.Body));

    [Fact]
    public void Serves_index_with_base_path_substituted()
    {
        var (h, _) = Create();
        var r = h.Handle(Req("GET", "/"));
        Assert.Equal(200, r.Status);
        Assert.StartsWith("text/html", r.ContentType);
        var html = Encoding.UTF8.GetString(r.Body);
        Assert.Contains("/excel-designer", html);
        Assert.DoesNotContain("{{BASE}}", html);
    }

    [Theory]
    [InlineData("/designer.js")]
    [InlineData("/designer.css")]
    [InlineData("/core.js")]
    [InlineData("/grid.js")]
    [InlineData("/panels.js")]
    public void Serves_embedded_assets(string path)
    {
        var (h, _) = Create();
        Assert.Equal(200, h.Handle(Req("GET", path)).Status);
    }

    [Fact]
    public void Path_traversal_is_blocked()
    {
        var (h, _) = Create();
        Assert.Equal(404, h.Handle(Req("GET", "/..%2Fsecret")).Status);
        Assert.Equal(404, h.Handle(Req("GET", "/assets/../x")).Status);
    }

    [Fact]
    public void Config_lists_models_with_field_tree()
    {
        var (h, _) = Create();
        var cfg = Parse(h.Handle(Req("GET", "/api/config")));
        var fields = cfg["models"][0]["fields"];
        Assert.Equal("Order", (string)cfg["models"][0]["name"]);
        Assert.Contains(fields, f => (string)f["name"] == "Lines" && (string)f["type"] == "list" && f["children"].Any(c => (string)c["name"] == "Qty" && (string)c["type"] == "number"));
        Assert.Contains(fields, f => (string)f["name"] == "Customer" && (string)f["type"] == "object");
    }

    [Fact]
    public void Save_get_list_delete_roundtrip()
    {
        var (h, store) = Create();
        var t = StarterTemplates.SimpleTable();
        Assert.Equal(200, h.Handle(Req("PUT", "/api/templates/My%20Report", t.ToJson())).Status);
        Assert.NotNull(store.Get("My Report"));
        Assert.Equal("My Report", (string)Parse(h.Handle(Req("GET", "/api/templates")))[0]["name"]);
        Assert.Equal("My Report", (string)Parse(h.Handle(Req("GET", "/api/templates/My%20Report")))["name"]);
        Assert.Equal(200, h.Handle(Req("DELETE", "/api/templates/My%20Report")).Status);
        Assert.Equal(404, h.Handle(Req("GET", "/api/templates/My%20Report")).Status);
    }

    [Fact]
    public void Mutations_need_the_csrf_header_and_a_valid_name_and_not_readonly()
    {
        var (h, _) = Create();
        var json = StarterTemplates.SimpleTable().ToJson();
        Assert.Equal(403, h.Handle(Req("PUT", "/api/templates/x", json, csrf: false)).Status);
        Assert.Equal(400, h.Handle(Req("PUT", "/api/templates/..%2F..%2Fevil", json)).Status);
        Assert.Equal(400, h.Handle(Req("PUT", "/api/templates/x", "not json")).Status);

        var (ro, _) = Create(readOnly: true);
        Assert.Equal(403, ro.Handle(Req("PUT", "/api/templates/x", json)).Status);
    }

    [Fact]
    public void Preview_returns_html_and_download_returns_a_valid_xlsx()
    {
        var (h, _) = Create();
        var t = StarterTemplates.Invoice();
        var body = new JObject { ["template"] = JObject.Parse(t.ToJson()) }.ToString();
        var p = Parse(h.Handle(Req("POST", "/api/preview", body)));
        Assert.Contains("Acme Co.", (string)p["html"]);

        var d = h.Handle(Req("POST", "/api/download", body));
        Assert.Equal(ExcelContent.Xlsx, d.ContentType);
        Assert.Contains("attachment", d.Headers["Content-Disposition"]);
        Assert.Equal("Acme Co.", Xlsx.Open(d.Body).Value("C5"));
    }

    [Fact]
    public void Preview_uses_model_generated_sample_when_template_has_none()
    {
        var (h, _) = Create();
        var b = new TemplateBuilder("T"); b.Sheet("S", s => s.Cell("A1", "{{Product}}").Band("Lines", "Lines", 1, 1));
        var t = b.Build(); t.ModelName = "Order";
        var p = Parse(h.Handle(Req("POST", "/api/preview", new JObject { ["template"] = JObject.Parse(t.ToJson()) }.ToString())));
        Assert.Contains("Product 1", (string)p["html"]);
        Assert.Contains("Product 3", (string)p["html"]);
    }

    [Fact]
    public void English_language_switches_direction()
    {
        var opt = new KscExcelOptions { Language = "en" };
        var h = new DesignerHost(opt, new ExcelReportService(opt, new MemoryTemplateStore()));
        var html = Encoding.UTF8.GetString(h.Handle(Req("GET", "/")).Body);
        Assert.Contains("lang=\"en\" dir=\"ltr\"", html);
    }
}
