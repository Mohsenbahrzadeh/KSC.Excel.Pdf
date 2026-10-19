using System.IO.Compression;
using System.Xml.Linq;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Services;
using KSC.Excel.Writer;
using Newtonsoft.Json.Linq;

namespace KSC.Excel.Tests;

/// <summary>Opens an .xlsx and gives typed access to its parts (and fails on any malformed XML).</summary>
public sealed class Xlsx
{
    public static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public Dictionary<string, byte[]> Raw = new();
    public Dictionary<string, XDocument> Xml = new();

    public static Xlsx Open(byte[] bytes)
    {
        var x = new Xlsx();
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        foreach (var e in zip.Entries)
        {
            using var s = e.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            x.Raw[e.FullName] = ms.ToArray();
            if (e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels"))
                x.Xml[e.FullName] = XDocument.Load(new MemoryStream(ms.ToArray())); // throws if malformed
        }
        return x;
    }

    public XDocument Sheet(int n = 1) => Xml[$"xl/worksheets/sheet{n}.xml"];

    public List<string> SharedStrings()
        => Xml["xl/sharedStrings.xml"].Root!.Elements(Main + "si").Select(si => string.Concat(si.Descendants(Main + "t").Select(t => t.Value))).ToList();

    public XElement Cell(string addr, int sheet = 1)
        => Sheet(sheet).Descendants(Main + "c").FirstOrDefault(c => (string)c.Attribute("r") == addr);

    /// <summary>Display value of a cell: shared string text, or the raw value.</summary>
    public string Value(string addr, int sheet = 1)
    {
        var c = Cell(addr, sheet);
        if (c == null) return null;
        if ((string)c.Attribute("t") == "inlineStr") return string.Concat(c.Descendants(Main + "t").Select(t => t.Value));
        var v = c.Element(Main + "v")?.Value;
        if (v == null) return null;
        return (string)c.Attribute("t") == "s" ? SharedStrings()[int.Parse(v)] : v;
    }

    public string Formula(string addr, int sheet = 1) => Cell(addr, sheet)?.Element(Main + "f")?.Value;
}

public static class Make
{
    public static IExcelReportService Service(ITemplateStore store = null)
        => new ExcelReportService(new KscExcelOptions(), store ?? new MemoryTemplateStore());

    public static Xlsx Render(ExcelTemplate t, object data) => Xlsx.Open(Service().Render(t, data));

    public static JObject Json(string json) => JObject.Parse(json);
}
