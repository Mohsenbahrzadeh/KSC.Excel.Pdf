using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using KSC.Excel.Model;
using KSC.Excel.Preview;
using KSC.Excel.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace KSC.Excel.Designer
{
    /// <summary>Framework-neutral request (adapters translate ASP.NET Core / System.Web requests into this).</summary>
    public sealed class DesignerRequest
    {
        public string Method { get; set; } = "GET";
        /// <summary>Path relative to the designer root, starting with "/" (e.g. "/api/templates").</summary>
        public string Path { get; set; } = "/";
        /// <summary>Public URL of the designer root without trailing slash (e.g. "/excel-designer").</summary>
        public string BasePath { get; set; } = "";
        public byte[] Body { get; set; }
        public IDictionary<string, string> Headers { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class DesignerResponse
    {
        public int Status { get; set; } = 200;
        public string ContentType { get; set; } = "application/json; charset=utf-8";
        public byte[] Body { get; set; } = new byte[0];
        public IDictionary<string, string> Headers { get; } = new Dictionary<string, string>();
    }

    /// <summary>
    /// The designer back-end: serves the embedded UI and a small JSON API. Host-neutral on purpose so the same code runs
    /// under ASP.NET Core and classic ASP.NET (.NET Framework). Authentication/authorization is the host's job.
    /// </summary>
    public sealed class DesignerHost
    {
        const int MaxBodyBytes = 16 * 1024 * 1024;

        static readonly JsonSerializerSettings Json = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore,
            DateParseHandling = DateParseHandling.None
        };

        readonly KscExcelOptions _opt;
        readonly IExcelReportService _svc;

        public DesignerHost(KscExcelOptions options, IExcelReportService service)
        {
            _opt = options ?? throw new ArgumentNullException(nameof(options));
            _svc = service ?? throw new ArgumentNullException(nameof(service));
        }

        public DesignerResponse Handle(DesignerRequest req)
        {
            try
            {
                var path = string.IsNullOrEmpty(req.Path) ? "/" : req.Path;
                var method = (req.Method ?? "GET").ToUpperInvariant();

                if (path.StartsWith("/api/", StringComparison.Ordinal))
                {
                    if (method != "GET" && method != "HEAD")
                    {
                        // CSRF guard: browsers cannot add a custom header cross-site without a CORS preflight.
                        string h;
                        if (!req.Headers.TryGetValue("X-KXL-Request", out h) || h != "1") return Error(403, "Missing X-KXL-Request header.");
                        if (req.Body != null && req.Body.Length > MaxBodyBytes) return Error(413, "Request too large.");
                    }
                    return Api(method, path, req);
                }
                if (method != "GET" && method != "HEAD") return Error(405, "Method not allowed.");
                return Static(path, req);
            }
            catch (JsonException ex) { return Error(400, "Invalid JSON: " + ex.Message); }
            catch (Engine.ExcelRenderException ex) { return Error(422, ex.Message); }
            catch (ArgumentException ex) { return Error(400, ex.Message); }
            catch (InvalidDataException ex) { return Error(400, ex.Message); }
            catch (FormatException ex) { return Error(400, "Invalid file: " + ex.Message); }
            catch (System.Xml.XmlException ex) { return Error(400, "Invalid XML in the file: " + ex.Message); }
        }

        // ---------------------------------------------------------------- API
        DesignerResponse Api(string method, string path, DesignerRequest req)
        {
            var segs = path.Substring(5).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries).Select(Uri.UnescapeDataString).ToArray();
            if (segs.Length == 0) return Error(404, "Not found.");

            switch (segs[0])
            {
                case "config" when method == "GET":
                    return Ok(new
                    {
                        title = _opt.DesignerTitle,
                        language = _opt.Language,
                        readOnly = _opt.ReadOnlyDesigner,
                        version = typeof(DesignerHost).Assembly.GetName().Version.ToString(3),
                        models = _opt.Models.Values.Select(m => new { name = m.Name, fields = m.Fields }).ToList(),
                        allowSql = _opt.AllowSql,
                        pdf = _opt.PdfExporter != null,
                        connections = _opt.ConnectionNames.ToList(),
                        endpoints = _opt.Endpoints.Keys.ToList(),
                        providers = _opt.Providers.Keys.ToList(),
                        starters = StarterTemplates.All().Select(k => new { key = k.Key, title = k.Value.Title }).ToList()
                    });

                case "templates":
                    if (segs.Length == 1 && method == "GET") return Ok(_svc.Store.List());
                    if (segs.Length == 2)
                    {
                        var name = segs[1];
                        if (method == "GET")
                        {
                            var t = _svc.Store.Get(name);
                            return t == null ? Error(404, "Template not found.") : Raw(t.ToJson());
                        }
                        if (method == "PUT")
                        {
                            if (_opt.ReadOnlyDesigner) return Error(403, "The designer is read-only.");
                            var t = ExcelTemplate.FromJson(Encoding.UTF8.GetString(req.Body ?? new byte[0]));
                            t.Name = name;
                            Validate(t);
                            _svc.Store.Save(t);
                            return Ok(new { saved = true, name });
                        }
                        if (method == "DELETE")
                        {
                            if (_opt.ReadOnlyDesigner) return Error(403, "The designer is read-only.");
                            return _svc.Store.Delete(name) ? Ok(new { deleted = true }) : Error(404, "Template not found.");
                        }
                    }
                    break;

                case "models" when method == "GET" && segs.Length == 2:
                {
                    ModelInfo m;
                    if (!_opt.Models.TryGetValue(segs[1], out m)) return Error(404, "Model not registered.");
                    return Ok(new { name = m.Name, fields = m.Fields, sample = m.Sample });
                }

                case "starters" when method == "GET" && segs.Length == 2:
                {
                    var s = StarterTemplates.All().FirstOrDefault(k => k.Key == segs[1]);
                    return s.Value == null ? Error(404, "Starter not found.") : Raw(s.Value.ToJson());
                }

                case "preview" when method == "POST":
                {
                    var body = JObject.Parse(Encoding.UTF8.GetString(req.Body ?? new byte[0]));
                    var tpl = ExcelTemplate.FromJson(body["template"].ToString(Formatting.None));
                    var wb = _svc.BindPreview(tpl, DataFrom(body, tpl), null);
                    var html = HtmlPreviewRenderer.Render(wb, new PreviewOptions { MaxRows = _opt.PreviewMaxRows, ShowHeaders = true });
                    return Ok(new { html, warnings = wb.Warnings, truncated = wb.Truncated });
                }

                case "download" when method == "POST":
                {
                    var body = JObject.Parse(Encoding.UTF8.GetString(req.Body ?? new byte[0]));
                    var tpl = ExcelTemplate.FromJson(body["template"].ToString(Formatting.None));
                    var bytes = Writer.XlsxWriter.ToBytes(_svc.BindPreview(tpl, DataFrom(body, tpl), null));
                    var resp = new DesignerResponse { ContentType = ExcelContent.Xlsx, Body = bytes };
                    var file = (string.IsNullOrWhiteSpace(tpl.Name) ? "report" : tpl.Name) + ".xlsx";
                    resp.Headers["Content-Disposition"] = "attachment; filename=\"report.xlsx\"; filename*=UTF-8''" + Uri.EscapeDataString(file);
                    return resp;
                }

                case "import" when method == "POST":
                {
                    if (_opt.ReadOnlyDesigner) return Error(403, "The designer is read-only.");
                    var body = JObject.Parse(Encoding.UTF8.GetString(req.Body ?? new byte[0]));
                    var name = System.IO.Path.GetFileNameWithoutExtension((string)body["name"] ?? "Imported");
                    if (!TemplateNames.IsValid(name)) name = "Imported";
                    var res = Import.XlsxImporter.Import(Convert.FromBase64String((string)body["data"] ?? ""), name);
                    return Raw(JsonConvert.SerializeObject(new { template = JObject.Parse(res.Template.ToJson()), warnings = res.Warnings }, Json));
                }

                case "pdf" when method == "POST":
                {
                    if (_opt.PdfExporter == null) return Error(404, "PDF export is not enabled (add the KSC.Excel.Pdf package and call options.UsePdf()).");
                    var body = JObject.Parse(Encoding.UTF8.GetString(req.Body ?? new byte[0]));
                    var tpl = ExcelTemplate.FromJson(body["template"].ToString(Formatting.None));
                    var bytes = _opt.PdfExporter(_svc.BindPreview(tpl, DataFrom(body, tpl), null));
                    var resp = new DesignerResponse { ContentType = "application/pdf", Body = bytes };
                    var file = (string.IsNullOrWhiteSpace(tpl.Name) ? "report" : tpl.Name) + ".pdf";
                    resp.Headers["Content-Disposition"] = "attachment; filename=\"report.pdf\"; filename*=UTF-8''" + Uri.EscapeDataString(file);
                    return resp;
                }

                case "validate" when method == "POST":
                {
                    var tpl = ExcelTemplate.FromJson(Encoding.UTF8.GetString(req.Body ?? new byte[0]));
                    var wb = _svc.BindPreview(tpl, ParseData(tpl.SampleData), null);
                    return Ok(new { warnings = wb.Warnings });
                }

                case "data" when method == "POST" && segs.Length == 2 && segs[1] == "test":
                {
                    // "Test" button of json / provider data sets: runs the data set (a registered endpoint or provider) and reports the columns
                    var body = JObject.Parse(Encoding.UTF8.GetString(req.Body ?? new byte[0]));
                    var ds = body["dataSet"] == null ? null : body["dataSet"].ToObject<DataSetDef>();
                    if (ds == null) return Error(400, "dataSet is required.");
                    bool inline = string.Equals(ds.Type, "json", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(ds.Json);
                    if (!inline && _opt.ReadOnlyDesigner) return Error(403, "The designer is read-only.");
                    var pars = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    var pj = body["params"] as JObject;
                    if (pj != null) foreach (var p in pj.Properties()) pars[p.Name] = ((JValue)p.Value).Value;
                    var warnings = new List<string>();
                    var ro = new Engine.RenderOptions { Providers = _opt.Providers, Endpoints = _opt.Endpoints };
                    var data = Engine.ExternalData.Load(ds, ro, pars, warnings.Add);
                    var items = Engine.DataAccess.Items(data).Take(200).ToList();
                    var cols = new List<object>();
                    if (items.Count > 0)
                    {
                        JToken first;
                        try { first = items[0] as JToken ?? JToken.FromObject(items[0]); } catch (Exception) { first = null; }
                        var jo = first as JObject;
                        if (jo != null)
                            foreach (var pr in jo.Properties())
                                cols.Add(new { name = pr.Name, type = pr.Value.Type == JTokenType.Integer || pr.Value.Type == JTokenType.Float ? "number" : pr.Value.Type == JTokenType.Boolean ? "bool" : pr.Value.Type == JTokenType.Array ? "list" : "text" });
                    }
                    return Ok(new { columns = cols, count = items.Count, warnings });
                }

                case "sql" when method == "POST" && segs.Length == 2 && segs[1] == "test":
                {
                    if (!_opt.AllowSql) return Error(403, "SQL data sets are disabled (KscExcelOptions.AllowSql).");
                    if (_opt.ReadOnlyDesigner) return Error(403, "The designer is read-only.");
                    var body = JObject.Parse(Encoding.UTF8.GetString(req.Body ?? new byte[0]));
                    var con = _opt.OpenFactory((string)body["connection"]);
                    if (con == null) return Error(400, "Unknown connection.");
                    var pars = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    var pj = body["params"] as JObject;
                    if (pj != null) foreach (var p in pj.Properties()) pars[p.Name] = ((JValue)p.Value).Value;
                    var warnings = new List<string>();
                    var rows = Engine.SqlDataSource.Query(con, (string)body["sql"], pars, _opt.SqlCommandTimeoutSeconds, 200, warnings.Add, true);
                    var cols = new List<object>();
                    if (rows.Count > 0)
                        foreach (var k in rows[0].Keys)
                            cols.Add(new { name = k, type = Engine.SqlDataSource.TypeName(rows.Select(r => r[k]).FirstOrDefault(v => v != null)) });
                    return Ok(new { columns = cols, rows = rows.Take(5).ToList(), count = rows.Count, warnings });
                }
            }
            return Error(404, "Not found.");
        }

        object DataFrom(JObject body, ExcelTemplate tpl)
        {
            JToken data = body["data"];
            if (data == null || data.Type == JTokenType.Null)
            {
                ModelInfo m;
                if (!string.IsNullOrEmpty(tpl.ModelName) && _opt.Models.TryGetValue(tpl.ModelName, out m) && tpl.SampleData == null) return m.Sample;
                return ParseData(tpl.SampleData);
            }
            return ParseData(data);
        }

        // Re-parse so ISO date strings become real dates (the template JSON keeps them as strings).
        static object ParseData(JToken t)
        {
            if (t == null) return null;
            return JToken.Parse(t.ToString(Formatting.None));
        }

        static void Validate(ExcelTemplate t)
        {
            if (t.Sheets.Count == 0) throw new ArgumentException("A template needs at least one sheet.");
            if (t.Sheets.Count > 64) throw new ArgumentException("Too many sheets.");
            long cells = t.Sheets.Sum(s => (long)s.Cells.Count);
            if (cells > 500000) throw new ArgumentException("Template is too large.");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in t.Sheets)
                if (!string.IsNullOrWhiteSpace(s.Name) && !names.Add(s.Name) && string.IsNullOrWhiteSpace(s.RepeatFor)) throw new ArgumentException("Duplicate sheet name '" + s.Name + "'.");
        }

        // ---------------------------------------------------------------- static UI
        DesignerResponse Static(string path, DesignerRequest req)
        {
            string file = path == "/" || path == "" ? "index.html" : path.TrimStart('/');
            if (file.StartsWith("assets/", StringComparison.Ordinal)) file = file.Substring(7);
            if (file.Contains("/") || file.Contains("\\") || file.Contains("..")) return Error(404, "Not found.");

            var bytes = ReadResource(file);
            if (bytes == null) return Error(404, "Not found.");

            var resp = new DesignerResponse { Body = bytes, ContentType = ContentTypeOf(file) };
            resp.Headers["Cache-Control"] = "no-cache";
            resp.Headers["X-Content-Type-Options"] = "nosniff";
            if (file == "index.html")
            {
                var html = Encoding.UTF8.GetString(bytes)
                    .Replace("{{BASE}}", System.Net.WebUtility.HtmlEncode((req.BasePath ?? "").TrimEnd('/')))
                    .Replace("{{TITLE}}", System.Net.WebUtility.HtmlEncode(_opt.DesignerTitle ?? "KSC Excel Designer"))
                    .Replace("{{LANG}}", _opt.Language == "en" ? "en" : "fa")
                    .Replace("{{DIR}}", _opt.Language == "en" ? "ltr" : "rtl");
                resp.Body = Encoding.UTF8.GetBytes(html);
                resp.Headers["Referrer-Policy"] = "same-origin";
                resp.Headers["X-Frame-Options"] = "SAMEORIGIN";
            }
            return resp;
        }

        static byte[] ReadResource(string file)
        {
            // Developer convenience (editing the UI without rebuilding): KSC_EXCEL_UI_DIR=<path to Designer/UI>
            var dev = Environment.GetEnvironmentVariable("KSC_EXCEL_UI_DIR");
            if (!string.IsNullOrEmpty(dev))
            {
                var p = Path.Combine(dev, file);
                return File.Exists(p) ? File.ReadAllBytes(p) : null;
            }
            var asm = typeof(DesignerHost).Assembly;
            using (var s = asm.GetManifestResourceStream("KSC.Excel.UI." + file))
            {
                if (s == null) return null;
                using (var ms = new MemoryStream()) { s.CopyTo(ms); return ms.ToArray(); }
            }
        }

        static string ContentTypeOf(string file)
        {
            switch (Path.GetExtension(file).ToLowerInvariant())
            {
                case ".html": return "text/html; charset=utf-8";
                case ".js": return "application/javascript; charset=utf-8";
                case ".css": return "text/css; charset=utf-8";
                case ".svg": return "image/svg+xml";
                case ".png": return "image/png";
                default: return "application/octet-stream";
            }
        }

        // ---------------------------------------------------------------- helpers
        static DesignerResponse Ok(object o) => Raw(JsonConvert.SerializeObject(o, Json));

        static DesignerResponse Raw(string json) => new DesignerResponse { Body = Encoding.UTF8.GetBytes(json) };

        static DesignerResponse Error(int status, string message)
            => new DesignerResponse { Status = status, Body = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(new { error = message }, Json)) };
    }
}
