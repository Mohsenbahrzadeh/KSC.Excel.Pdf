using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Writer;

namespace KSC.Excel.Preview
{
    public sealed class PreviewOptions
    {
        /// <summary>Rows shown per sheet (preview only); the download is never truncated.</summary>
        public int MaxRows = 500;
        /// <summary>Show A,B,C... / 1,2,3... headers like Excel.</summary>
        public bool ShowHeaders = true;
        /// <summary>Emit the stylesheet with the markup (set false when the page already includes it).</summary>
        public bool IncludeCss = true;
    }

    /// <summary>Renders a <see cref="RenderedWorkbook"/> to HTML (sheet tabs, merged cells, styles, conditional formats).</summary>
    public static class HtmlPreviewRenderer
    {
        static int _seq;

        public static string Render(RenderedWorkbook wb, PreviewOptions opt = null)
        {
            opt = opt ?? new PreviewOptions();
            var id = "kxl" + Interlocked();
            var sb = new StringBuilder();

            if (opt.IncludeCss) sb.Append("<style>").Append(Css(id, wb.Sheets.Count)).Append("</style>");
            sb.Append("<div class=\"kxl\" id=\"").Append(id).Append("\">");
            for (int i = 0; i < wb.Sheets.Count; i++)
                sb.Append("<input type=\"radio\" class=\"kxl-r\" name=\"").Append(id).Append("-t\" id=\"").Append(id).Append('-').Append(i).Append('"').Append(i == 0 ? " checked" : "").Append(">");
            sb.Append("<div class=\"kxl-tabs\">");
            for (int i = 0; i < wb.Sheets.Count; i++)
                sb.Append("<label for=\"").Append(id).Append('-').Append(i).Append("\">").Append(Enc(wb.Sheets[i].Name)).Append("</label>");
            sb.Append("</div><div class=\"kxl-panes\">");
            for (int i = 0; i < wb.Sheets.Count; i++)
            {
                sb.Append("<div class=\"kxl-pane kxl-p").Append(i).Append("\">");
                RenderSheet(sb, wb, wb.Sheets[i], opt);
                sb.Append("</div>");
            }
            sb.Append("</div></div>");
            return sb.ToString();
        }

        static int Interlocked() => System.Threading.Interlocked.Increment(ref _seq);

        static string Css(string id, int sheets)
        {
            var sb = new StringBuilder();
            sb.Append('#').Append(id).Append("{font-family:Tahoma,Segoe UI,sans-serif;color:#222;background:#fff;border:1px solid #d0d7de;border-radius:6px;overflow:hidden}");
            sb.Append('#').Append(id).Append(" .kxl-r{display:none}");
            sb.Append('#').Append(id).Append(" .kxl-tabs{display:flex;gap:2px;background:#f1f3f5;border-top:1px solid #d0d7de;order:2;padding:4px 6px 0}");
            sb.Append('#').Append(id).Append(" .kxl-tabs label{padding:4px 14px;font-size:12px;cursor:pointer;border:1px solid transparent;border-radius:4px 4px 0 0;color:#555;background:#e6e9ec}");
            sb.Append('#').Append(id).Append(" .kxl-panes{overflow:auto;max-height:70vh;direction:ltr}");
            sb.Append('#').Append(id).Append(" .kxl-pane{display:none}");
            sb.Append('#').Append(id).Append(" table{border-collapse:collapse;table-layout:fixed;font-size:13px}");
            sb.Append('#').Append(id).Append(" td{padding:0 4px;overflow:hidden;text-overflow:clip;vertical-align:bottom;white-space:nowrap;box-sizing:border-box}");
            sb.Append('#').Append(id).Append(" td.kxl-h{background:#f3f4f6;color:#666;text-align:center;font-size:11px;border:1px solid #d9dce0;position:sticky}");
            for (int i = 0; i < sheets; i++)
            {
                sb.Append('#').Append(id).Append('-').Append(i).Append(":checked~.kxl-panes .kxl-p").Append(i).Append("{display:block}");
                sb.Append('#').Append(id).Append('-').Append(i).Append(":checked~.kxl-tabs label[for=\"").Append(id).Append('-').Append(i).Append("\"]{background:#fff;color:#111;border-color:#d0d7de;border-bottom-color:#fff;font-weight:600}");
            }
            return sb.ToString();
        }

        static string Enc(string s) => WebUtility.HtmlEncode(s ?? "");

        static void RenderSheet(StringBuilder sb, RenderedWorkbook wb, RenderedSheet sh, PreviewOptions opt)
        {
            int rows = Math.Min(sh.Rows.Count, opt.MaxRows);
            int cols = Math.Max(sh.MaxCol, 1);
            if (sh.Rows.Count > rows) wb.Truncated = true;

            // merge map
            var anchors = new Dictionary<long, int[]>();
            var covered = new HashSet<long>();
            foreach (var m in sh.Merges)
            {
                anchors[Key(m[0], m[1])] = m;
                for (int r = m[0]; r <= m[2]; r++)
                    for (int c = m[1]; c <= m[3]; c++)
                        if (r != m[0] || c != m[1]) covered.Add(Key(r, c));
            }

            var cf = new CfEngine(sh, rows);
            var colW = new double[cols + 1];
            for (int c = 1; c <= cols; c++) colW[c] = sh.DefaultColWidth;
            foreach (var cd in sh.Columns) if (cd.Index >= 1 && cd.Index <= cols && cd.Width.HasValue) colW[cd.Index] = cd.Width.Value;
            var hidden = new HashSet<int>(sh.Columns.Where(c => c.Hidden).Select(c => c.Index));

            var evaluator = new FormulaEvaluator(sh);
            var gridColor = sh.ShowGridLines ? "#e3e5e8" : "transparent";

            sb.Append("<table dir=\"").Append(sh.Rtl ? "rtl" : "ltr").Append("\" style=\"").Append(sh.Rtl ? "margin-left:auto;" : "").Append("width:")
              .Append((int)(Enumerable.Range(1, cols).Where(c => !hidden.Contains(c)).Sum(c => Px(colW[c])) + (opt.ShowHeaders ? 40 : 0))).Append("px\"><colgroup>");
            if (opt.ShowHeaders) sb.Append("<col style=\"width:40px\">");
            for (int c = 1; c <= cols; c++) if (!hidden.Contains(c)) sb.Append("<col style=\"width:").Append((int)Px(colW[c])).Append("px\">");
            sb.Append("</colgroup>");

            if (opt.ShowHeaders)
            {
                sb.Append("<tr style=\"height:20px\"><td class=\"kxl-h\" style=\"top:0\"></td>");
                for (int c = 1; c <= cols; c++) if (!hidden.Contains(c)) sb.Append("<td class=\"kxl-h\" style=\"top:0\">").Append(CellRef.ColName(c)).Append("</td>");
                sb.Append("</tr>");
            }

            for (int r = 1; r <= rows; r++)
            {
                var row = sh.Rows[r - 1];
                if (row.Hidden) continue;
                sb.Append("<tr style=\"height:").Append((row.Height ?? sh.DefaultRowHeight).ToString("0.##", CultureInfo.InvariantCulture)).Append("pt\">");
                if (opt.ShowHeaders) sb.Append("<td class=\"kxl-h\" style=\"left:0\">").Append(r).Append("</td>");
                int ci = 0;
                for (int c = 1; c <= cols; c++)
                {
                    if (hidden.Contains(c)) continue;
                    if (covered.Contains(Key(r, c))) continue;
                    while (ci < row.Cells.Count && row.Cells[ci].Col < c) ci++;
                    RenderedCell cell = ci < row.Cells.Count && row.Cells[ci].Col == c ? row.Cells[ci] : null;

                    sb.Append("<td");
                    int[] m;
                    if (anchors.TryGetValue(Key(r, c), out m))
                    {
                        int cs = 0; for (int k = m[1]; k <= m[3]; k++) if (!hidden.Contains(k)) cs++;
                        if (cs > 1) sb.Append(" colspan=\"").Append(cs).Append('"');
                        if (m[2] > m[0]) sb.Append(" rowspan=\"").Append(m[2] - m[0] + 1).Append('"');
                    }

                    object val = null;
                    if (cell != null)
                    {
                        if (cell.Formula != null)
                        {
                            val = cell.Cached;
                            if (val == null) { object ev; if (evaluator.TryEvaluate(cell, out ev)) val = ev; }
                        }
                        else val = cell.Value;
                    }

                    var css = new StringBuilder();
                    var style = cell?.Style;
                    StyleCss(css, style, sh, gridColor, val);
                    string text;
                    var xe = val as FormulaEvaluator.XlError;
                    if (xe != null) text = xe.Code;
                    else text = NumberFormatter.Format(val, style?.Format);

                    cf.Apply(r, c, val, css, ref text);

                    sb.Append(" style=\"").Append(css).Append("\">");
                    if (!string.IsNullOrEmpty(cell?.Link)) sb.Append("<a href=\"").Append(Enc(SafeUrl(cell.Link))).Append("\" target=\"_blank\" rel=\"noopener\" style=\"color:inherit\">").Append(Enc(text)).Append("</a>");
                    else sb.Append(Enc(text));
                    sb.Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
            foreach (var ch in sh.Charts)
                sb.Append("<div style=\"margin:6px 8px;padding:8px 12px;border:1px dashed #9bb;border-radius:6px;font-size:12px;color:#456\">&#128202; ").Append(Enc(string.IsNullOrEmpty(ch.Def.Title) ? ch.Def.Name : ch.Def.Title)).Append(" (").Append(Enc(ch.Def.Type)).Append(") - chart is shown in Excel</div>");
            if (sh.Rows.Count > rows)
                sb.Append("<div style=\"padding:6px 10px;font-size:12px;color:#8a5a00;background:#fff7e0\">Preview shows the first ").Append(rows).Append(" of ").Append(sh.Rows.Count).Append(" rows. The downloaded file contains all rows.</div>");
        }

        static string SafeUrl(string u)
        {
            u = (u ?? "").Trim();
            if (u.StartsWith("#", StringComparison.Ordinal)) return "javascript:void(0)";
            return u.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || u.StartsWith("https://", StringComparison.OrdinalIgnoreCase) || u.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ? u : "about:blank";
        }

        static long Key(int r, int c) => ((long)r << 20) | (uint)c;
        static double Px(double chars) => Math.Round(chars * 7 + 5);

        static string Hex(string c)
        {
            var rgb = StyleBook.Rgb(c);
            return rgb == null ? null : "#" + rgb.Substring(2);
        }

        static void Border(StringBuilder css, string side, BorderDef b, string grid)
        {
            string style = b?.Style;
            if (string.IsNullOrEmpty(style)) { css.Append("border-").Append(side).Append(":1px solid ").Append(grid).Append(';'); return; }
            string w = style == "medium" ? "2px" : style == "thick" ? "3px" : "1px";
            string st = style == "dashed" || style == "mediumdashed" ? "dashed" : style == "dotted" || style == "hair" ? "dotted" : style == "double" ? "double" : "solid";
            if (style == "double") w = "3px";
            css.Append("border-").Append(side).Append(':').Append(w).Append(' ').Append(st).Append(' ').Append(Hex(b.Color) ?? "#000").Append(';');
        }

        static void StyleCss(StringBuilder css, StyleDef s, RenderedSheet sh, string grid, object val)
        {
            if (s == null)
            {
                css.Append("border:1px solid ").Append(grid).Append(';');
                if (DataAccess.TryNumber(val, out _) || val is DateTime) css.Append("text-align:right;");
                return;
            }
            if (!string.IsNullOrEmpty(s.Font)) css.Append("font-family:'").Append(s.Font.Replace("'", "")).Append("',Tahoma,sans-serif;");
            if (s.Size.HasValue) css.Append("font-size:").Append(s.Size.Value.ToString("0.##", CultureInfo.InvariantCulture)).Append("pt;");
            if (s.Bold) css.Append("font-weight:700;");
            if (s.Italic) css.Append("font-style:italic;");
            if (s.Underline || s.Strike) css.Append("text-decoration:").Append(s.Underline ? "underline " : "").Append(s.Strike ? "line-through" : "").Append(';');
            var col = Hex(s.Color); if (col != null) css.Append("color:").Append(col).Append(';');
            var fill = Hex(s.Fill); if (fill != null) css.Append("background:").Append(fill).Append(';');
            Border(css, "left", s.Left, grid); Border(css, "right", s.Right, grid); Border(css, "top", s.Top, grid); Border(css, "bottom", s.Bottom, grid);

            string h = (s.H ?? "").ToLowerInvariant();
            if (h == "left" || h == "right" || h == "center") css.Append("text-align:").Append(h).Append(';');
            else if (h == "justify" || h == "distributed") css.Append("text-align:justify;");
            else if (DataAccess.TryNumber(val, out _) || val is DateTime) css.Append("text-align:right;");
            else css.Append("text-align:start;");
            string v = (s.V ?? "").ToLowerInvariant();
            css.Append("vertical-align:").Append(v == "top" ? "top" : v == "center" || v == "middle" ? "middle" : "bottom").Append(';');
            if (s.Wrap) css.Append("white-space:pre-wrap;word-break:break-word;");
            if (s.Indent > 0) css.Append("padding-").Append(sh.Rtl ? "right" : "left").Append(':').Append(4 + s.Indent * 9).Append("px;");
            if (s.Rtl.HasValue) css.Append("direction:").Append(s.Rtl.Value ? "rtl" : "ltr").Append(';');
        }

        // ---------- conditional formats (shared evaluator, see Engine/CondFormatEvaluator) ----------
        sealed class CfEngine
        {
            readonly CondFormatEvaluator _cf;

            public CfEngine(RenderedSheet sh, int rows) { _cf = new CondFormatEvaluator(sh); }

            public void Apply(int r, int c, object val, StringBuilder css, ref string text)
            {
                var res = _cf.Evaluate(r, c, val, text);
                if (res == null) return;
                if (res.ScaleFill != null) css.Append("background:").Append(res.ScaleFill).Append(';');
                if (res.BarFraction.HasValue)
                {
                    int pct = (int)(res.BarFraction.Value * 100);
                    css.Append("background:linear-gradient(to right,").Append(res.BarColor).Append(' ').Append(pct).Append("%,transparent ").Append(pct).Append("%);");
                }
                var s = res.Style;
                if (s == null) return;
                var f = Hex(s.Fill); if (f != null) css.Append("background:").Append(f).Append(';');
                var col = Hex(s.Color); if (col != null) css.Append("color:").Append(col).Append(';');
                if (s.Bold) css.Append("font-weight:700;");
                if (s.Italic) css.Append("font-style:italic;");
            }
        }
    }
}
