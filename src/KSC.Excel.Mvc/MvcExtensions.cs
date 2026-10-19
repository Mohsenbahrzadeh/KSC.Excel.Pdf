using System;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;
using KSC.Excel.Services;

namespace KSC.Excel.Mvc
{
    /// <summary>Action result that sends a rendered template as an .xlsx download.</summary>
    public sealed class ExcelResult : ActionResult
    {
        readonly string _template;
        readonly object _data;
        readonly string _fileName;

        readonly System.Collections.Generic.IDictionary<string, object> _params;

        public ExcelResult(string templateName, object data, string fileName = null, System.Collections.Generic.IDictionary<string, object> parameters = null)
        {
            _template = templateName; _data = data; _params = parameters;
            _fileName = EnsureXlsx(fileName ?? templateName);
        }

        internal static string EnsureXlsx(string n) => n.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? n : n + ".xlsx";

        public override void ExecuteResult(ControllerContext context)
        {
            var bytes = _params == null ? KscExcel.Service.Render(_template, _data) : KscExcel.Service.Render(_template, _data, _params);
            var r = context.HttpContext.Response;
            r.Clear();
            r.ContentType = ExcelContent.Xlsx;
            r.AddHeader("Content-Disposition", "attachment; filename=\"report.xlsx\"; filename*=UTF-8''" + Uri.EscapeDataString(_fileName));
            r.Cache.SetCacheability(HttpCacheability.NoCache);
            r.BinaryWrite(bytes);
        }
    }

    /// <summary>Sends a rendered template as a PDF (download, or shown in the browser when inline). Needs options.UsePdf() from KSC.Excel.Pdf.</summary>
    public sealed class PdfResult : ActionResult
    {
        readonly string _template; readonly object _data; readonly string _fileName; readonly System.Collections.Generic.IDictionary<string, object> _params; readonly bool _inline;

        public PdfResult(string templateName, object data, string fileName = null, System.Collections.Generic.IDictionary<string, object> parameters = null, bool inline = false)
        { _template = templateName; _data = data; _params = parameters; _inline = inline; _fileName = EnsurePdf(fileName ?? templateName); }

        internal static string EnsurePdf(string n) => n.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? n : n + ".pdf";

        internal static byte[] Render(string template, object data, System.Collections.Generic.IDictionary<string, object> parameters)
        {
            var svc = KscExcel.Service;
            return ViewerMarkup.ToPdf(KscExcel.Options, svc.Bind(svc.GetTemplate(template), data, parameters));
        }

        public override void ExecuteResult(ControllerContext context)
        {
            var bytes = Render(_template, _data, _params);
            var r = context.HttpContext.Response;
            r.Clear();
            r.ContentType = ExcelContent.Pdf;
            r.AddHeader("Content-Disposition", (_inline ? "inline" : "attachment") + "; filename=\"report.pdf\"; filename*=UTF-8''" + Uri.EscapeDataString(_fileName));
            r.Cache.SetCacheability(HttpCacheability.NoCache);
            r.BinaryWrite(bytes);
        }
    }

    /// <summary>Streams a huge report straight to the response (flat memory). See IExcelReportService.RenderStreaming.</summary>
    public sealed class ExcelStreamingResult : ActionResult
    {
        readonly string _template; readonly object _data; readonly string _fileName; readonly System.Collections.Generic.IDictionary<string, object> _params;

        public ExcelStreamingResult(string templateName, object data, string fileName = null, System.Collections.Generic.IDictionary<string, object> parameters = null)
        { _template = templateName; _data = data; _params = parameters; _fileName = ExcelResult.EnsureXlsx(fileName ?? templateName); }

        public override void ExecuteResult(ControllerContext context)
        {
            // render into a temp file first: if rendering fails nothing half-written reaches the browser
            var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "kxl-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                using (var fs = new System.IO.FileStream(path, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None, 81920))
                    KscExcel.Service.RenderStreaming(_template, _data, fs, _params);
                var r = context.HttpContext.Response;
                r.Clear();
                r.ContentType = ExcelContent.Xlsx;
                r.AddHeader("Content-Disposition", "attachment; filename=\"report.xlsx\"; filename*=UTF-8''" + Uri.EscapeDataString(_fileName));
                r.Cache.SetCacheability(HttpCacheability.NoCache);
                r.TransmitFile(path);
                r.Flush();
            }
            finally { try { System.IO.File.Delete(path); } catch (System.IO.IOException) { } }
        }
    }

    public static class ControllerExtensions
    {
        /// <summary><c>return this.Report("Invoice", model, ReportFormat.Docx);</c> - xlsx, pdf, csv, docx, html or png (pdf/png need <c>o.UsePdf()</c>).</summary>
        public static ActionResult Report(this Controller controller, string templateName, object data, ReportFormat format, string fileName = null, System.Collections.Generic.IDictionary<string, object> parameters = null, int page = 1)
        {
            var f = KscExcel.Service.Export(KscExcel.Options, templateName, data, parameters, format, page);
            return new FileContentResult(f.Bytes, f.ContentType) { FileDownloadName = ReportExport.FileName(fileName ?? templateName, f.Extension) };
        }

        /// <summary>Typed report parameters from the query string (<c>?p_From=2025-01-01</c>) for the template's declared parameters.</summary>
        public static System.Collections.Generic.Dictionary<string, object> GetReportParameters(this Controller controller, string templateName)
        {
            var tpl = KscExcel.Service.GetTemplate(templateName);
            var q = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string key in controller.Request.QueryString.AllKeys) if (key != null) q[key] = controller.Request.QueryString[key];
            return ReportParameters.FromQuery(tpl, q);
        }

        /// <summary>One-liner: <c>return this.Pdf("Invoice", model, "invoice.pdf");</c> (needs <c>o.UsePdf()</c> from KSC.Excel.Pdf).</summary>
        public static PdfResult Pdf(this Controller controller, string templateName, object data, string fileName = null, System.Collections.Generic.IDictionary<string, object> parameters = null, bool inline = false)
            => new PdfResult(templateName, data, fileName, parameters, inline);

        /// <summary>For very large reports: <c>return this.ExcelStreaming("Big", reader, "big.xlsx");</c></summary>
        public static ExcelStreamingResult ExcelStreaming(this Controller controller, string templateName, object data, string fileName = null, System.Collections.Generic.IDictionary<string, object> parameters = null)
            => new ExcelStreamingResult(templateName, data, fileName, parameters);

        /// <summary>One-liner: <c>return this.Excel("Invoice", model, "invoice.xlsx");</c></summary>
        public static ExcelResult Excel(this Controller controller, string templateName, object data, string fileName = null)
            => new ExcelResult(templateName, data, fileName);

        /// <summary>Same, passing report parameters ({{$p.Name}} in the template).</summary>
        public static ExcelResult Excel(this Controller controller, string templateName, object data, System.Collections.Generic.IDictionary<string, object> parameters, string fileName = null)
            => new ExcelResult(templateName, data, fileName, parameters);
    }

    public static class HtmlHelperExtensions
    {
        /// <summary>
        /// Shows the report as a PDF on the page (browser PDF reader) with Download / Print buttons:
        /// <c>@Html.KscPdfViewer("Invoice", Model, "invoice")</c>. Needs <c>o.UsePdf()</c> and <c>routes.MapKscExcelDownloads()</c>.
        /// </summary>
        public static MvcHtmlString KscPdfViewer(this HtmlHelper html, string template, object data, string downloadName = null, System.Collections.Generic.IDictionary<string, object> parameters = null, string height = "80vh", bool showToolbar = true, bool showExcel = false)
        {
            var svc = KscExcel.Service; var o = KscExcel.Options;
            var wb = svc.Bind(svc.GetTemplate(template), data, parameters);
            var ctx = html.ViewContext.HttpContext;
            var user = ctx.User != null && ctx.User.Identity.IsAuthenticated ? ctx.User.Identity.Name : null;
            var baseName = downloadName ?? template;
            Func<string, string> url = token => VirtualPathUtility.ToAbsolute("~/" + o.DownloadPath.Trim('/') + "/" + token);

            var pdf = ViewerMarkup.ToPdf(o, wb);
            var inlineHref = url(KscExcel.Downloads.Add(pdf, PdfResult.EnsurePdf(baseName), user, ExcelContent.Pdf, true));
            var pdfHref = url(KscExcel.Downloads.Add(pdf, PdfResult.EnsurePdf(baseName), user, ExcelContent.Pdf, false));
            string xlsxHref = showExcel ? url(KscExcel.Downloads.Add(Writer.XlsxWriter.ToBytes(wb), ExcelResult.EnsureXlsx(baseName), user)) : null;
            return MvcHtmlString.Create(ViewerMarkup.PdfViewer(inlineHref, pdfHref, xlsxHref, o.Language == "en", height, showToolbar));
        }

        /// <summary>
        /// Renders the report on the page (sheet tabs) and a download button:
        /// <c>@Html.KscExcelViewer("Invoice", Model, "invoice")</c>
        /// </summary>
        public static MvcHtmlString KscExcelViewer(this HtmlHelper html, string template, object data, string downloadName = null, bool showDownload = true, bool showHeaders = false, string downloadLabel = null)
        {
            var svc = KscExcel.Service;
            var o = KscExcel.Options;
            var tpl = svc.GetTemplate(template);
            var wb = svc.Bind(tpl, data);
            var preview = Preview.HtmlPreviewRenderer.Render(wb, new Preview.PreviewOptions { MaxRows = o.PreviewMaxRows, ShowHeaders = showHeaders });

            var sb = new StringBuilder("<div class=\"ksc-excel-viewer\">");
            if (showDownload)
            {
                var ctx = html.ViewContext.HttpContext;
                var user = ctx.User != null && ctx.User.Identity.IsAuthenticated ? ctx.User.Identity.Name : null;
                var token = KscExcel.Downloads.Add(Writer.XlsxWriter.ToBytes(wb), ExcelResult.EnsureXlsx(downloadName ?? template), user);
                var href = VirtualPathUtility.ToAbsolute("~/" + o.DownloadPath.Trim('/') + "/" + token);
                var label = downloadLabel ?? (o.Language == "en" ? "Download Excel" : "دریافت فایل اکسل");
                sb.Append("<div style=\"margin-bottom:8px\"><a href=\"").Append(WebUtility.HtmlEncode(href))
                  .Append("\" style=\"display:inline-block;padding:6px 14px;background:#1f7a46;color:#fff;border-radius:5px;text-decoration:none;font:13px Tahoma,sans-serif\">")
                  .Append(WebUtility.HtmlEncode(label)).Append("</a></div>");
            }
            sb.Append(preview).Append("</div>");
            return MvcHtmlString.Create(sb.ToString());
        }
    }
}
