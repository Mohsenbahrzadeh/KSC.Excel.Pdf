using KSC.Excel.Services;

namespace KSC.Excel.Pdf;

public static class PdfExtensions
{
    /// <summary>Enables PDF export in the designer (adds a PDF button): <c>AddKscExcel(o =&gt; o.UsePdf())</c>.</summary>
    public static KscExcelOptions UsePdf(this KscExcelOptions options, PdfOptions? pdf = null)
    {
        options.PdfExporter = wb => ExcelPdf.Render(wb, pdf);
        return options;
    }

    /// <summary>Renders a stored template + data to PDF: <c>return File(svc.RenderPdf("Invoice", model), "application/pdf", "invoice.pdf");</c></summary>
    public static byte[] RenderPdf(this IExcelReportService svc, string templateName, object? data, IDictionary<string, object>? parameters = null, PdfOptions? options = null)
        => RenderPdf(svc, svc.GetTemplate(templateName), data, parameters, options);

    public static byte[] RenderPdf(this IExcelReportService svc, Model.ExcelTemplate template, object? data, IDictionary<string, object>? parameters = null, PdfOptions? options = null)
        => ExcelPdf.Render(svc.Bind(template, data!, parameters!), options);

    /// <summary>Writes the PDF to a stream.</summary>
    public static void RenderPdf(this IExcelReportService svc, string templateName, object? data, Stream output, IDictionary<string, object>? parameters = null, PdfOptions? options = null)
        => ExcelPdf.Render(svc.Bind(svc.GetTemplate(templateName), data!, parameters!), output, options);
}
