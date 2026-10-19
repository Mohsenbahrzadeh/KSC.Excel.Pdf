using System.Net;
using KSC.Excel.Services;
using Microsoft.AspNetCore.Html;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ViewComponents;

namespace KSC.Excel.AspNetCore;

/// <summary>
/// Shows a report on the page (all sheets as tabs) with a download button:
/// <code>@await Component.InvokeAsync("KscExcelViewer", new { template = "Invoice", data = Model, downloadName = "invoice" })</code>
/// or <c>&lt;vc:ksc-excel-viewer template="Invoice" data="Model" /&gt;</c>.
/// </summary>
[ViewComponent(Name = "KscExcelViewer")]
public sealed class KscExcelViewerViewComponent : ViewComponent
{
    private readonly IExcelReportService _svc;
    private readonly KscExcelOptions _options;
    private readonly DownloadStore _downloads;

    public KscExcelViewerViewComponent(IExcelReportService svc, KscExcelOptions options, DownloadStore downloads)
    {
        _svc = svc; _options = options; _downloads = downloads;
    }

    /// <param name="template">Stored template name.</param>
    /// <param name="data">The data to bind (any object, dictionary, JSON, DataTable...).</param>
    /// <param name="downloadName">File name for the download button (without extension is fine).</param>
    /// <param name="showDownload">Show the download button.</param>
    /// <param name="showHeaders">Show Excel-like A,B,C / 1,2,3 headers.</param>
    /// <param name="downloadLabel">Button text (default is Persian when the designer language is fa).</param>
    public IViewComponentResult Invoke(string template, object? data, string? downloadName = null, bool showDownload = true, bool showHeaders = false, string? downloadLabel = null)
    {
        var tpl = _svc.GetTemplate(template);
        var wb = _svc.Bind(tpl, data);
        var html = Preview.HtmlPreviewRenderer.Render(wb, new Preview.PreviewOptions { MaxRows = _options.PreviewMaxRows, ShowHeaders = showHeaders });

        var sb = new System.Text.StringBuilder();
        sb.Append("<div class=\"ksc-excel-viewer\">");
        if (showDownload)
        {
            var bytes = Writer.XlsxWriter.ToBytes(wb);
            var token = _downloads.Add(bytes, ControllerExtensions.EnsureXlsx(downloadName ?? template), HttpContext.User.Identity?.Name);
            var href = HttpContext.Request.PathBase + "/" + _options.DownloadPath.Trim('/') + "/" + token;
            var label = downloadLabel ?? (_options.Language == "en" ? "Download Excel" : "دریافت فایل اکسل");
            sb.Append("<div style=\"margin-bottom:8px\"><a href=\"").Append(WebUtility.HtmlEncode(href))
              .Append("\" style=\"display:inline-block;padding:6px 14px;background:#1f7a46;color:#fff;border-radius:5px;text-decoration:none;font:13px Tahoma,sans-serif\">")
              .Append(WebUtility.HtmlEncode(label)).Append("</a></div>");
        }
        sb.Append(html).Append("</div>");
        return new HtmlContentViewComponentResult(new HtmlString(sb.ToString()));
    }
}
