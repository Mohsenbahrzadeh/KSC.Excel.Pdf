using KSC.Excel.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
namespace KSC.Excel.AspNetCore;


public static class ControllerExtensions
{
    /// <summary>
    /// One-liner for controllers: <c>return this.ExcelFile("Invoice", model, "invoice.xlsx");</c>
    /// </summary>
    public static FileContentResult ExcelFile(this ControllerBase controller, string templateName, object? data, string? fileName = null)
    {
        var svc = controller.HttpContext.RequestServices.GetRequiredService<IExcelReportService>();
        return new FileContentResult(svc.Render(templateName, data), ExcelContent.Xlsx) { FileDownloadName = EnsureXlsx(fileName ?? templateName) };
    }

    /// <summary>Same, passing report parameters ({{$p.Name}} in the template).</summary>
    public static FileContentResult ExcelFile(this ControllerBase controller, string templateName, object? data, IDictionary<string, object> parameters, string? fileName = null)
    {
        var svc = controller.HttpContext.RequestServices.GetRequiredService<IExcelReportService>();
        return new FileContentResult(svc.Render(templateName, data, parameters), ExcelContent.Xlsx) { FileDownloadName = EnsureXlsx(fileName ?? templateName) };
    }

    /// <summary>
    /// For very large reports: renders with the streaming writer into a temp file (flat memory) and sends it.
    /// Pass lazily enumerated data (IEnumerable / IAsyncEnumerable is not supported / DataReader-backed) to keep memory low end-to-end.
    /// </summary>
    public static IActionResult ExcelFileStreaming(this ControllerBase controller, string templateName, object? data, string? fileName = null, IDictionary<string, object>? parameters = null)
    {
        var svc = controller.HttpContext.RequestServices.GetRequiredService<IExcelReportService>();
        var tmp = new FileStream(Path.Combine(Path.GetTempPath(), "kxl-" + Guid.NewGuid().ToString("N") + ".xlsx"), FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.DeleteOnClose);
        try { svc.RenderStreaming(templateName, data, tmp, parameters!); tmp.Position = 0; }
        catch { tmp.Dispose(); throw; }
        return new FileStreamResult(tmp, ExcelContent.Xlsx) { FileDownloadName = EnsureXlsx(fileName ?? templateName) };
    }

    /// <summary>Minimal-API flavour: <c>app.MapGet("/r", (IExcelReportService x) =&gt; ExcelResults.File(x, "Invoice", data));</c></summary>
    public static IResult File(this IExcelReportService svc, string templateName, object? data, string? fileName = null)
        => Results.File(svc.Render(templateName, data), ExcelContent.Xlsx, EnsureXlsx(fileName ?? templateName));

    // ---------------------------------------------------------------- PDF (needs the KSC.Excel.Pdf package + options.UsePdf())

    /// <summary>
    /// <c>return this.PdfFile("Invoice", model, "invoice.pdf");</c> - downloads a PDF (pass <paramref name="inline"/> = true to show it in the browser).
    /// Requires <c>AddKscExcel(o =&gt; o.UsePdf())</c> from the KSC.Excel.Pdf package.
    /// </summary>
    public static FileContentResult PdfFile(this ControllerBase controller, string templateName, object? data, string? fileName = null, IDictionary<string, object>? parameters = null, bool inline = false)
    {
        var sp = controller.HttpContext.RequestServices;
        var pdf = RenderPdf(sp.GetRequiredService<IExcelReportService>(), sp.GetRequiredService<KscExcelOptions>(), templateName, data, parameters);
        return new FileContentResult(pdf, ExcelContent.Pdf) { FileDownloadName = inline ? null : EnsurePdf(fileName ?? templateName) };
    }

    /// <summary>Minimal-API flavour: <c>app.MapGet("/r", (IExcelReportService x, KscExcelOptions o) =&gt; x.Pdf(o, "Invoice", data));</c></summary>
    public static IResult Pdf(this IExcelReportService svc, KscExcelOptions options, string templateName, object? data, string? fileName = null, IDictionary<string, object>? parameters = null)
        => Results.File(RenderPdf(svc, options, templateName, data, parameters), ExcelContent.Pdf, EnsurePdf(fileName ?? templateName));

    internal static byte[] RenderPdf(IExcelReportService svc, KscExcelOptions options, string templateName, object? data, IDictionary<string, object>? parameters)
        => ViewerMarkup.ToPdf(options, svc.Bind(svc.GetTemplate(templateName), data!, parameters!));

    // ---------------------------------------------------------------- any format / report parameters

    /// <summary>
    /// <c>return this.ReportFile("Invoice", model, ReportFormat.Docx);</c> - xlsx, pdf, csv, docx, html or png. Pdf/Png need <c>o.UsePdf()</c>.
    /// </summary>
    public static FileContentResult ReportFile(this ControllerBase controller, string templateName, object? data, ReportFormat format, string? fileName = null, IDictionary<string, object>? parameters = null, int page = 1)
    {
        var sp = controller.HttpContext.RequestServices;
        var f = sp.GetRequiredService<IExcelReportService>().Export(sp.GetRequiredService<KscExcelOptions>(), templateName, data, parameters!, format, page);
        return new FileContentResult(f.Bytes, f.ContentType) { FileDownloadName = ReportExport.FileName(fileName ?? templateName, f.Extension) };
    }

    /// <summary>Typed report parameters from the current query string (<c>?p_From=2025-01-01&amp;p_Min=100</c>) for the template's declared parameters.</summary>
    public static Dictionary<string, object> GetReportParameters(this ControllerBase controller, string templateName)
    {
        var tpl = controller.HttpContext.RequestServices.GetRequiredService<IExcelReportService>().GetTemplate(templateName);
        var q = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in controller.HttpContext.Request.Query) q[kv.Key] = kv.Value.ToString();
        return ReportParameters.FromQuery(tpl, q);
    }

    internal static string EnsurePdf(string name) => name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? name : name + ".pdf";

    internal static string EnsureXlsx(string name) => name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? name : name + ".xlsx";
}
