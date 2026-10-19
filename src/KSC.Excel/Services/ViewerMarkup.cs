using System;
using System.Net;
using System.Text;
using Engine = KSC.Excel.Engine;

namespace KSC.Excel.Services
{
    /// <summary>Markup shared by the ASP.NET Core ViewComponents and the MVC 5 HTML helpers (one place, identical output).</summary>
    public static class ViewerMarkup
    {
        const string BtnStyle = "display:inline-block;padding:6px 14px;color:#fff;border-radius:5px;text-decoration:none;font:13px Tahoma,sans-serif;border:0;cursor:pointer;margin-inline-end:6px";

        public static string DownloadButton(string href, string label, string color = "#1f7a46")
            => "<a href=\"" + WebUtility.HtmlEncode(href) + "\" style=\"" + BtnStyle + ";background:" + color + "\">" + WebUtility.HtmlEncode(label) + "</a>";

        /// <summary>
        /// A PDF viewer: the browser's built-in PDF reader (zoom, page navigation, search, print) inside an iframe, with a toolbar
        /// (download PDF, optional download Excel, print). Needs no JavaScript library and no CDN.
        /// </summary>
        public static string PdfViewer(string inlineHref, string pdfDownloadHref, string xlsxDownloadHref, bool english, string height = "80vh", bool showToolbar = true)
        {
            var id = "kxlpdf" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var sb = new StringBuilder();
            sb.Append("<div class=\"ksc-pdf-viewer\">");
            if (showToolbar)
            {
                sb.Append("<div style=\"margin-bottom:8px\">");
                if (pdfDownloadHref != null) sb.Append(DownloadButton(pdfDownloadHref, english ? "Download PDF" : "دریافت PDF", "#b8312f"));
                if (xlsxDownloadHref != null) sb.Append(DownloadButton(xlsxDownloadHref, english ? "Download Excel" : "دریافت فایل اکسل"));
                sb.Append("<button type=\"button\" style=\"" + BtnStyle + ";background:#2e75b6\" onclick=\"(function(){var f=document.getElementById('" + id + "');try{f.contentWindow.focus();f.contentWindow.print();}catch(e){window.open('")
                  .Append(WebUtility.HtmlEncode(inlineHref)).Append("');}})()\">").Append(english ? "Print" : "چاپ").Append("</button>");
                sb.Append("</div>");
            }
            sb.Append("<iframe id=\"").Append(id).Append("\" src=\"").Append(WebUtility.HtmlEncode(inlineHref))
              .Append("#view=FitH\" title=\"PDF\" style=\"width:100%;height:").Append(WebUtility.HtmlEncode(height)).Append(";border:1px solid #ccd;border-radius:4px\"></iframe></div>");
            return sb.ToString();
        }

        /// <summary>Throws a clear error when the PDF package was not enabled.</summary>
        public static byte[] ToPdf(KscExcelOptions options, Engine.RenderedWorkbook wb)
        {
            if (options == null || options.PdfExporter == null)
                throw new InvalidOperationException("PDF output is not enabled. Reference the KSC.Excel.Pdf package and call options.UsePdf() when you configure KSC.Excel.");
            return options.PdfExporter(wb);
        }
    }
}
