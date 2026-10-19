using KSC.Excel.Services;
using Microsoft.AspNetCore.Html;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ViewComponents;

namespace KSC.Excel.AspNetCore;

/// <summary>
/// Shows a report as a PDF on the page (browser PDF reader: pages, zoom, search, print) with Download / Print buttons:
/// <code>@await Component.InvokeAsync("KscPdfViewer", new { template = "Invoice", data = Model, downloadName = "invoice" })</code>
/// Requires <c>o.UsePdf()</c> (KSC.Excel.Pdf) and <c>app.MapKscExcelDownloads()</c>.
/// </summary>
[ViewComponent(Name = "KscPdfViewer")]
public sealed class KscPdfViewerViewComponent : ViewComponent
{
    private readonly IExcelReportService _svc;
    private readonly KscExcelOptions _options;
    private readonly DownloadStore _downloads;

    public KscPdfViewerViewComponent(IExcelReportService svc, KscExcelOptions options, DownloadStore downloads)
    {
        _svc = svc; _options = options; _downloads = downloads;
    }

    /// <param name="template">Stored template name.</param>
    /// <param name="data">The data to bind.</param>
    /// <param name="parameters">Report parameters ({{$p.Name}}).</param>
    /// <param name="downloadName">File name for downloads (extension optional).</param>
    /// <param name="height">CSS height of the viewer, e.g. "80vh" or "700px".</param>
    /// <param name="showToolbar">Show Download / Print buttons.</param>
    /// <param name="showExcel">Also offer the same report as an .xlsx download.</param>
    public IViewComponentResult Invoke(string template, object? data, IDictionary<string, object>? parameters = null, string? downloadName = null, string height = "80vh", bool showToolbar = true, bool showExcel = false, bool showParameters = false)
    {
        var tpl = _svc.GetTemplate(template);
        var query = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in HttpContext.Request.Query) query[kv.Key] = kv.Value.ToString();
        var pars = ReportParameters.FromQuery(tpl, query);
        if (parameters != null) foreach (var kv in parameters) pars[kv.Key] = kv.Value;
        var wb = _svc.Bind(tpl, data!, pars);
        var user = HttpContext.User.Identity?.Name;
        var baseName = downloadName ?? template;
        string Url(string token) => HttpContext.Request.PathBase + "/" + _options.DownloadPath.Trim('/') + "/" + token;

        var pdf = ViewerMarkup.ToPdf(_options, wb);
        var inlineHref = Url(_downloads.Add(pdf, ControllerExtensions.EnsurePdf(baseName), user, ExcelContent.Pdf, true));
        var pdfHref = Url(_downloads.Add(pdf, ControllerExtensions.EnsurePdf(baseName), user, ExcelContent.Pdf, false));
        string? xlsxHref = showExcel ? Url(_downloads.Add(Writer.XlsxWriter.ToBytes(wb), ControllerExtensions.EnsureXlsx(baseName), user)) : null;

        var html = ViewerMarkup.PdfViewer(inlineHref, pdfHref, xlsxHref, _options.Language == "en", height, showToolbar);
        if (showParameters) html = ViewerMarkup.ParameterForm(tpl, query, HttpContext.Request.PathBase + HttpContext.Request.Path, _options.Language == "en") + html;
        return new HtmlContentViewComponentResult(new HtmlString(html));
    }
}
