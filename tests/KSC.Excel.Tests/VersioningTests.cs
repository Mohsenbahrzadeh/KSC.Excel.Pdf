using System.Text;
using KSC.Excel.Designer;
using KSC.Excel.Model;
using KSC.Excel.Services;
using Newtonsoft.Json.Linq;
using Xunit;

namespace KSC.Excel.Tests;

public class VersioningTests
{
    static ExcelTemplate T(string name, string title, string cell = "x")
    {
        var b = new TemplateBuilder(name, title);
        b.Sheet("S", s => s.Cell("A1", cell));
        return b.Build();
    }

    static string TempDir() { var d = Path.Combine(Path.GetTempPath(), "kxl-ver-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(d); return d; }

    [Fact]
    public void File_store_keeps_the_replaced_version_and_restores_it()
    {
        var dir = TempDir();
        try
        {
            var store = new FileTemplateStore(dir);
            store.Save(T("Inv", "v1", "one"));
            Assert.Empty(store.ListVersions("Inv"));                    // first save: nothing to keep
            store.Save(T("Inv", "v2", "two"));
            store.Save(T("Inv", "v2", "two"));                          // identical save: no new version
            store.Save(T("Inv", "v3", "three"));
            var versions = store.ListVersions("Inv");
            Assert.Equal(2, versions.Count);
            Assert.Equal("v2", versions[0].Title);                      // newest first
            Assert.Equal("v1", versions[1].Title);
            var old = store.GetVersion("Inv", versions[1].Id);
            Assert.Equal("one", old.Sheets[0].Cells[0].V);
            Assert.Equal("Inv", old.Name);
            Assert.Equal("v3", store.Get("Inv").Title);                 // current untouched
            Assert.DoesNotContain(store.List(), t => t.Name.StartsWith("."));   // history never shows up as a template
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Deleting_a_template_keeps_its_history_and_the_cap_is_enforced()
    {
        var dir = TempDir();
        try
        {
            var store = new FileTemplateStore(dir) { MaxVersions = 3 };
            for (int i = 0; i < 8; i++) store.Save(T("Cap", "t" + i, "c" + i));
            Assert.Equal(3, store.ListVersions("Cap").Count);
            Assert.True(store.Delete("Cap"));
            Assert.Null(store.Get("Cap"));
            Assert.Equal(3, store.ListVersions("Cap").Count);          // the deleted state joins the history, the cap still applies
            Assert.Equal("t7", store.ListVersions("Cap").Select(v => store.GetVersion("Cap", v.Id).Title).First());
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void History_can_be_disabled_and_ids_are_validated()
    {
        var dir = TempDir();
        try
        {
            var store = new FileTemplateStore(dir) { MaxVersions = 0 };
            store.Save(T("N", "a")); store.Save(T("N", "b"));
            Assert.Empty(store.ListVersions("N"));
            Assert.Throws<ArgumentException>(() => store.GetVersion("N", "../../secret"));
            Assert.Throws<ArgumentException>(() => store.ListVersions("../x"));
            Assert.Null(store.GetVersion("N", "20250101000000000"));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Memory_store_versions()
    {
        var store = new MemoryTemplateStore();
        store.Save(T("M", "a", "1")); store.Save(T("M", "b", "2")); store.Save(T("M", "c", "3"));
        var v = store.ListVersions("M");
        Assert.Equal(new[] { "b", "a" }, v.Select(x => x.Title).ToArray());
        Assert.Equal("1", store.GetVersion("M", v[1].Id).Sheets[0].Cells[0].V);
        Assert.Null(store.GetVersion("M", "999"));
    }

    [Fact]
    public void Designer_api_lists_and_returns_versions()
    {
        var store = new MemoryTemplateStore();
        var svc = new ExcelReportService(new KscExcelOptions(), store);
        var host = new DesignerHost(new KscExcelOptions(), svc);
        store.Save(T("D", "a", "1")); store.Save(T("D", "b", "2"));
        DesignerResponse Get(string path) => host.Handle(new DesignerRequest { Method = "GET", Path = path, BasePath = "" });

        var cfg = JObject.Parse(Encoding.UTF8.GetString(Get("/api/config").Body));
        Assert.True((bool)cfg["versions"]);
        var list = JArray.Parse(Encoding.UTF8.GetString(Get("/api/templates/D/versions").Body));
        Assert.Single(list);
        var one = Get("/api/templates/D/versions/" + (string)list[0]["id"]);
        Assert.Equal(200, one.Status);
        Assert.Equal("a", (string)JObject.Parse(Encoding.UTF8.GetString(one.Body))["title"]);
        Assert.Equal(404, Get("/api/templates/D/versions/42").Status);
    }
}
