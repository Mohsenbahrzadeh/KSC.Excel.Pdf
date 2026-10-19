using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using KSC.Excel.Designer;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Preview;

namespace KSC.Excel.Services
{
    /// <summary>A report published by the report server: a stored template plus the code that supplies its data.</summary>
    public sealed class ReportDef
    {
        /// <summary>Public name (URL segment). Letters, digits, space, _ - .</summary>
        public string Name { get; set; }
        /// <summary>Stored template to render (defaults to <see cref="Name"/>).</summary>
        public string Template { get; set; }
        public string Title { get; set; }
        /// <summary>Returns the data for the report parameters (any object / list / DataTable / JSON). May return null when the template reads its own data sets.</summary>
        public Func<IDictionary<string, object>, object> Data { get; set; }
    }

    /// <summary>Framework-neutral request of the report server.</summary>
    public sealed class ReportRequest
    {
        /// <summary>Report name, or null for the catalog.</summary>
        public string Name { get; set; }
        /// <summary>Raw query-string values (names case-insensitive).</summary>
        public IDictionary<string, string> Query { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Public URL of the report server root without trailing slash, e.g. "/reports".</summary>
        public string BasePath { get; set; } = "/reports";
    }

    /// <summary>Turns query-string values into typed report parameters.</summary>
    public static class ReportParameters
    {
        public const string Prefix = "p_";

        /// <summary>Reads <c>p_Name</c> values for every parameter of the template. Missing or empty values keep the template default.</summary>
        public static Dictionary<string, object> FromQuery(ExcelTemplate template, IDictionary<string, string> query)
        {
            var res = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in template.Parameters)
            {
                string raw;
                if (string.IsNullOrWhiteSpace(p.Name) || !query.TryGetValue(Prefix + p.Name, out raw)) continue;
                raw = (raw ?? "").Trim();
                switch ((p.Type ?? "text").ToLowerInvariant())
                {
                    case "number":
                        double d; if (raw.Length > 0 && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) res[p.Name] = d;
                        break;
                    case "date":
                        DateTime dt; if (raw.Length > 0 && DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) res[p.Name] = dt;
                        break;
                    case "bool":
                        res[p.Name] = raw == "1" || raw.Equals("true", StringComparison.OrdinalIgnoreCase) || raw.Equals("on", StringComparison.OrdinalIgnoreCase);
                        break;
                    case "jdate":
                        if (raw.Length > 0) res[p.Name] = raw;       // converted by the renderer
                        break;
                    default:
                        if (raw.Length > 0) res[p.Name] = raw;
                        break;
                }
            }
            return res;
        }

        /// <summary>Only the <c>p_*</c> pairs of a query, re-encoded for links.</summary>
        public static string QueryString(IDictionary<string, string> query, params string[] extra)
        {
            var parts = new List<string>();
            foreach (var kv in query)
                if (kv.Key.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase)) parts.Add(Uri.EscapeDataString(kv.Key) + "=" + Uri.EscapeDataString(kv.Value ?? ""));
            parts.AddRange(extra);
            return parts.Count == 0 ? "" : "?" + string.Join("&", parts);
        }
    }

    /// <summary>Rewrites <c>report:Name?Param={{Id|urlenc}}</c> cell links into report-server URLs (drill-down).</summary>
    public static class ReportLinks
    {
        public static string Resolve(string link, string reportsBase)
        {
            if (string.IsNullOrEmpty(link) || reportsBase == null || !link.StartsWith("report:", StringComparison.OrdinalIgnoreCase)) return link;
            var rest = link.Substring(7).Trim();
            string name = rest, query = "";
            int q = rest.IndexOf('?');
            if (q >= 0) { name = rest.Substring(0, q); query = rest.Substring(q + 1); }
            if (!TemplateNames.IsValid(name)) return "about:blank";
            var pairs = new List<string>();
            foreach (var part in query.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int eq = part.IndexOf('=');
                var key = eq < 0 ? part : part.Substring(0, eq); var val = eq < 0 ? "" : part.Substring(eq + 1);
                pairs.Add(Uri.EscapeDataString(ReportParameters.Prefix + key.Trim()) + "=" + val);   // value is already encoded by |urlenc
            }
            return reportsBase.TrimEnd('/') + "/" + Uri.EscapeDataString(name) + (pairs.Count > 0 ? "?" + string.Join("&", pairs) : "");
        }
    }

    /// <summary>
    /// A small report server: <c>/reports</c> lists the reports you registered with <c>options.AddReport(...)</c>;
    /// <c>/reports/{name}</c> shows one with a parameter form and download buttons; <c>?format=pdf|xlsx|docx|csv|html|png</c> returns a file.
    /// Host-neutral (ASP.NET Core and System.Web adapters call <see cref="Handle"/>). Authorization is the host's job.
    /// </summary>
    public sealed class ReportServer
    {
        readonly KscExcelOptions _opt;
        readonly IExcelReportService _svc;

        public ReportServer(KscExcelOptions options, IExcelReportService service) { _opt = options; _svc = service; }

        public DesignerResponse Handle(ReportRequest req)
        {
            try
            {
                if (string.IsNullOrEmpty(req.Name)) return Page("Reports", Catalog(req), req);
                ReportDef def;
                if (!_opt.Reports.TryGetValue(req.Name, out def)) return Text(404, "Report not found.");
                ExcelTemplate tpl;
                try { tpl = _svc.GetTemplate(def.Template ?? def.Name); }
                catch (System.IO.FileNotFoundException) { return Text(404, "The template of this report was not found."); }

                var pars = ReportParameters.FromQuery(tpl, req.Query);
                string fmtText; req.Query.TryGetValue("format", out fmtText);
                ReportFormat fmt;
                bool isView = string.IsNullOrEmpty(fmtText) || fmtText.Equals("view", StringComparison.OrdinalIgnoreCase);
                if (!isView)
                {
                    if (!ReportExport.TryParseFormat(fmtText, out fmt)) return Text(400, "Unknown format. Use pdf, xlsx, docx, csv, html or png.");
                    int page = 1; string pg; if (req.Query.TryGetValue("page", out pg)) int.TryParse(pg, out page);
                    var data = def.Data == null ? null : def.Data(pars);
                    var file = _svc.Export(_opt, tpl, data, pars, fmt, Math.Max(1, page));
                    var res = new DesignerResponse { ContentType = file.ContentType, Body = file.Bytes };
                    string dl; bool forceDownload = req.Query.TryGetValue("download", out dl) && dl == "1";
                    bool inlineOk = (fmt == ReportFormat.Pdf || fmt == ReportFormat.Png || fmt == ReportFormat.Html) && !forceDownload;
                    var fileName = ReportExport.FileName(def.Name, file.Extension);
                    res.Headers["Content-Disposition"] = (inlineOk ? "inline" : "attachment") + "; filename=\"report." + file.Extension + "\"; filename*=UTF-8''" + Uri.EscapeDataString(fileName);
                    res.Headers["Cache-Control"] = "no-store";
                    return res;
                }
                return Page(def.Title ?? tpl.Title ?? def.Name, View(def, tpl, pars, req), req);
            }
            catch (ExcelRenderException ex) { return Text(422, ex.Message); }
            catch (Exception ex)
            {
                if (_opt.OnReportError != null) { try { _opt.OnReportError(req.Name, ex); } catch (Exception) { } }
                return Text(500, "The report could not be produced.");
            }
        }

        string Catalog(ReportRequest req)
        {
            var sb = new StringBuilder("<ul style=\"line-height:2\">");
            foreach (var r in _opt.Reports.Values.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase))
                sb.Append("<li><a href=\"").Append(WebUtility.HtmlEncode(req.BasePath.TrimEnd('/') + "/" + Uri.EscapeDataString(r.Name))).Append("\">").Append(WebUtility.HtmlEncode(r.Title ?? r.Name)).Append("</a></li>");
            if (_opt.Reports.Count == 0) sb.Append("<li>—</li>");
            return sb.Append("</ul>").ToString();
        }

        string View(ReportDef def, ExcelTemplate tpl, IDictionary<string, object> pars, ReportRequest req)
        {
            bool en = _opt.Language == "en";
            var self = req.BasePath.TrimEnd('/') + "/" + Uri.EscapeDataString(def.Name);
            string view; req.Query.TryGetValue("view", out view);
            bool pdfView = string.Equals(view, "pdf", StringComparison.OrdinalIgnoreCase) && _opt.PdfExporter != null;

            var sb = new StringBuilder();
            sb.Append(ViewerMarkup.ParameterForm(tpl, req.Query, self, en, pdfView ? "pdf" : null));

            // toolbar
            Func<string, string, string, string> btn = (fmt, label, color) => "<a href=\"" + WebUtility.HtmlEncode(self + ReportParameters.QueryString(req.Query, "format=" + fmt)) + "\" style=\"display:inline-block;padding:5px 12px;margin-inline-end:6px;background:" + color + ";color:#fff;border-radius:5px;text-decoration:none;font:13px Tahoma,sans-serif\">" + WebUtility.HtmlEncode(label) + "</a>";
            sb.Append("<div style=\"margin:10px 0\">");
            if (_opt.PdfExporter != null)
            {
                var other = pdfView ? "view=html" : "view=pdf";
                sb.Append("<a href=\"").Append(WebUtility.HtmlEncode(self + ReportParameters.QueryString(req.Query, other))).Append("\" style=\"display:inline-block;padding:5px 12px;margin-inline-end:6px;background:#555;color:#fff;border-radius:5px;text-decoration:none;font:13px Tahoma,sans-serif\">").Append(pdfView ? (en ? "HTML view" : "نمای HTML") : (en ? "PDF view" : "نمای PDF")).Append("</a>");
                sb.Append(btn("pdf&download=1", "PDF", "#b8312f"));
            }
            sb.Append(btn("xlsx", "Excel", "#1f7a46")).Append(btn("docx", "Word", "#2b579a")).Append(btn("csv", "CSV", "#6b6b6b"));
            if (_opt.PngExporter != null) sb.Append(btn("png&download=1", "PNG", "#8a5a00"));
            sb.Append("</div>");

            if (pdfView)
                sb.Append("<iframe src=\"").Append(WebUtility.HtmlEncode(self + ReportParameters.QueryString(req.Query, "format=pdf") + "#view=FitH")).Append("\" title=\"PDF\" style=\"width:100%;height:80vh;border:1px solid #ccd;border-radius:4px\"></iframe>");
            else
            {
                var data = def.Data == null ? null : def.Data(pars);
                var wb = _svc.Bind(tpl, data, pars);
                sb.Append(HtmlPreviewRenderer.Render(wb, new PreviewOptions { MaxRows = _opt.PreviewMaxRows, ShowHeaders = false }));
            }
            return sb.ToString();
        }

        DesignerResponse Page(string title, string body, ReportRequest req)
        {
            bool rtl = _opt.Language != "en";
            var html = "<!DOCTYPE html><html lang=\"" + (rtl ? "fa" : "en") + "\" dir=\"" + (rtl ? "rtl" : "ltr") + "\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>"
                + WebUtility.HtmlEncode(title) + "</title><style>body{font-family:Tahoma,Segoe UI,sans-serif;background:#f3f5f8;margin:0}main{max-width:1100px;margin:20px auto;padding:0 16px}h1{font-size:20px}a{color:#1f6fb5}</style></head><body><main><p><a href=\""
                + WebUtility.HtmlEncode(req.BasePath) + "\">" + (rtl ? "همه‌ی گزارش‌ها" : "All reports") + "</a></p><h1>" + WebUtility.HtmlEncode(title) + "</h1>" + body + "</main></body></html>";
            var r = new DesignerResponse { ContentType = "text/html; charset=utf-8", Body = new UTF8Encoding(false).GetBytes(html) };
            r.Headers["Cache-Control"] = "no-store";
            return r;
        }

        static DesignerResponse Text(int status, string message)
            => new DesignerResponse { Status = status, ContentType = "text/plain; charset=utf-8", Body = new UTF8Encoding(false).GetBytes(message) };
    }
}
