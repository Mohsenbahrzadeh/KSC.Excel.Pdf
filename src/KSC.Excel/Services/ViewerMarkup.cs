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

        /// <summary>
        /// A GET form with one input per report parameter (text, number, date, jdate, bool, drop-down). Values come from the
        /// current query (<c>p_Name</c>) or the template defaults. Returns "" when the template has no parameters.
        /// </summary>
        public static string ParameterForm(KSC.Excel.Model.ExcelTemplate tpl, System.Collections.Generic.IDictionary<string, string> query, string action, bool english, string view = null)
        {
            if (tpl == null || tpl.Parameters == null || tpl.Parameters.Count == 0) return "";
            var sb = new StringBuilder("<form method=\"get\" action=\"" + WebUtility.HtmlEncode(action) + "\" style=\"display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;margin:8px 0;padding:10px;background:#fff;border:1px solid #dde3ea;border-radius:6px\">");
            foreach (var p in tpl.Parameters)
            {
                if (string.IsNullOrWhiteSpace(p.Name)) continue;
                string cur; var name = "p_" + p.Name;
                var val = query != null && query.TryGetValue(name, out cur) ? cur : (p.Default ?? "");
                var label = string.IsNullOrWhiteSpace(p.Label) ? p.Name : p.Label;
                var type = (p.Type ?? "text").ToLowerInvariant();
                sb.Append("<label style=\"font:12px Tahoma,sans-serif;display:flex;flex-direction:column;gap:3px\">").Append(WebUtility.HtmlEncode(label));
                string st = " style=\"padding:5px 7px;border:1px solid #bcc6d2;border-radius:4px;font:13px Tahoma,sans-serif;min-width:120px\"";
                if (type == "bool")
                    sb.Append("<input type=\"hidden\" name=\"").Append(WebUtility.HtmlEncode(name)).Append("\" value=\"0\"><input type=\"checkbox\" name=\"").Append(WebUtility.HtmlEncode(name)).Append("\" value=\"1\"").Append(val == "1" || string.Equals(val, "true", StringComparison.OrdinalIgnoreCase) || val == "on" ? " checked" : "").Append(">");
                else if (!string.IsNullOrWhiteSpace(p.Options))
                {
                    sb.Append("<select name=\"").Append(WebUtility.HtmlEncode(name)).Append("\"").Append(st).Append(">");
                    foreach (var o in p.Options.Split('|'))
                    {
                        var parts = o.Split(new[] { ':' }, 2); var v = parts[0].Trim(); var text = parts.Length > 1 ? parts[1].Trim() : v;
                        sb.Append("<option value=\"").Append(WebUtility.HtmlEncode(v)).Append("\"").Append(v == val ? " selected" : "").Append(">").Append(WebUtility.HtmlEncode(text)).Append("</option>");
                    }
                    sb.Append("</select>");
                }
                else
                {
                    string it = type == "number" ? "number\" step=\"any" : type == "date" ? "date" : "text";
                    sb.Append("<input type=\"").Append(it).Append("\" name=\"").Append(WebUtility.HtmlEncode(name)).Append("\" value=\"").Append(WebUtility.HtmlEncode(val)).Append("\"")
                      .Append(type == "jdate" ? " placeholder=\"1404/01/01\" dir=\"ltr\" pattern=\"[0-9۰-۹]{4}[/-][0-9۰-۹]{1,2}[/-][0-9۰-۹]{1,2}\"" : "").Append(st).Append(">");
                }
                sb.Append("</label>");
            }
            if (!string.IsNullOrEmpty(view)) sb.Append("<input type=\"hidden\" name=\"view\" value=\"").Append(WebUtility.HtmlEncode(view)).Append("\">");
            sb.Append("<button type=\"submit\" style=\"" + BtnStyle + ";background:#2e75b6\">").Append(english ? "Show report" : "نمایش گزارش").Append("</button></form>");
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
