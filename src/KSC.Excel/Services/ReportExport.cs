using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Preview;
using KSC.Excel.Writer;

namespace KSC.Excel.Services
{
    /// <summary>Output formats of a rendered report.</summary>
    public enum ReportFormat
    {
        Xlsx,
        /// <summary>Needs the KSC.Excel.Pdf package (<c>options.UsePdf()</c>).</summary>
        Pdf,
        /// <summary>One sheet as CSV (UTF-8 with BOM).</summary>
        Csv,
        /// <summary>Word document: one table per sheet.</summary>
        Docx,
        /// <summary>A stand-alone HTML page (all sheets, tabs).</summary>
        Html,
        /// <summary>One page as PNG image (needs the KSC.Excel.Pdf package).</summary>
        Png
    }

    /// <summary>A rendered report file ready to send.</summary>
    public sealed class ReportFile
    {
        public byte[] Bytes;
        public string ContentType;
        /// <summary>File extension without the dot (xlsx, pdf, csv, docx, html, png).</summary>
        public string Extension;
    }

    public static class ReportExport
    {
        public const string DocxContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        public static bool TryParseFormat(string text, out ReportFormat format)
        {
            format = ReportFormat.Xlsx;
            switch ((text ?? "").Trim().ToLowerInvariant())
            {
                case "xlsx": case "excel": case "": format = ReportFormat.Xlsx; return true;
                case "pdf": format = ReportFormat.Pdf; return true;
                case "csv": format = ReportFormat.Csv; return true;
                case "docx": case "word": format = ReportFormat.Docx; return true;
                case "html": case "htm": format = ReportFormat.Html; return true;
                case "png": case "image": format = ReportFormat.Png; return true;
                default: return false;
            }
        }

        /// <summary>
        /// Renders a stored template in any format. <paramref name="page"/> is the 1-based page (png) or sheet (csv).
        /// </summary>
        public static ReportFile Export(this IExcelReportService svc, KscExcelOptions options, string templateName, object data, IDictionary<string, object> parameters, ReportFormat format, int page = 1)
            => Export(svc, options, svc.GetTemplate(templateName), data, parameters, format, page);

        public static ReportFile Export(this IExcelReportService svc, KscExcelOptions options, ExcelTemplate template, object data, IDictionary<string, object> parameters, ReportFormat format, int page = 1)
        {
            var wb = svc.Bind(template, data, parameters);
            return Export(wb, options, format, page);
        }

        public static ReportFile Export(RenderedWorkbook wb, KscExcelOptions options, ReportFormat format, int page = 1)
        {
            switch (format)
            {
                case ReportFormat.Xlsx: return new ReportFile { Bytes = XlsxWriter.ToBytes(wb), ContentType = ExcelContent.Xlsx, Extension = "xlsx" };
                case ReportFormat.Pdf: return new ReportFile { Bytes = ViewerMarkup.ToPdf(options, wb), ContentType = ExcelContent.Pdf, Extension = "pdf" };
                case ReportFormat.Csv: return new ReportFile { Bytes = CsvWriter.ToBytes(wb, Math.Max(0, page - 1)), ContentType = "text/csv; charset=utf-8", Extension = "csv" };
                case ReportFormat.Docx: return new ReportFile { Bytes = DocxWriter.ToBytes(wb), ContentType = DocxContentType, Extension = "docx" };
                case ReportFormat.Html: return new ReportFile { Bytes = new UTF8Encoding(false).GetBytes(HtmlDocument(wb, options)), ContentType = "text/html; charset=utf-8", Extension = "html" };
                case ReportFormat.Png:
                    if (options == null || options.PngExporter == null) throw new InvalidOperationException("PNG output is not enabled. Reference the KSC.Excel.Pdf package and call options.UsePdf() when you configure KSC.Excel.");
                    return new ReportFile { Bytes = options.PngExporter(wb, Math.Max(1, page)), ContentType = "image/png", Extension = "png" };
                default: throw new ArgumentOutOfRangeException("format");
            }
        }

        /// <summary>A complete HTML page (UTF-8, direction from the first sheet) containing the report.</summary>
        public static string HtmlDocument(RenderedWorkbook wb, KscExcelOptions options = null)
        {
            var max = options != null ? options.PreviewMaxRows : 500;
            var body = HtmlPreviewRenderer.Render(wb, new PreviewOptions { MaxRows = Math.Max(max, 5000), ShowHeaders = false });
            bool rtl = wb.Sheets.Count > 0 && wb.Sheets[0].Rtl;
            var title = WebUtility.HtmlEncode(wb.Template != null ? (wb.Template.Title ?? wb.Template.Name ?? "Report") : "Report");
            return "<!DOCTYPE html><html lang=\"" + (rtl ? "fa" : "en") + "\" dir=\"" + (rtl ? "rtl" : "ltr") + "\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>" + title + "</title><style>body{margin:16px;background:#f3f5f8}</style></head><body>" + body + "</body></html>";
        }

        /// <summary>Safe file name for downloads: letters, digits and a few separators; Persian letters are kept.</summary>
        public static string FileName(string name, string extension)
        {
            var sb = new StringBuilder();
            foreach (var ch in name ?? "report") sb.Append(char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' || ch == '.' ? ch : '_');
            var n = sb.ToString().Trim('.', '_');
            if (n.Length == 0) n = "report";
            return n.EndsWith("." + extension, StringComparison.OrdinalIgnoreCase) ? n : n + "." + extension;
        }
    }
}
