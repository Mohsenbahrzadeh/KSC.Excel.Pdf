using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Preview;

namespace KSC.Excel.Writer
{
    /// <summary>
    /// Writes a rendered report as a Word document (.docx): every sheet becomes a table (merged cells, fills, borders, fonts, alignment,
    /// right-to-left), with generated images (barcodes, logos) inside their cells. Charts are not included. No Word installation needed.
    /// </summary>
    public static class DocxWriter
    {
        const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static byte[] ToBytes(RenderedWorkbook wb)
        {
            using (var ms = new MemoryStream())
            {
                Write(wb, ms);
                return ms.ToArray();
            }
        }

        public static void Write(RenderedWorkbook wb, Stream output)
        {
            var media = new List<RenderedImage>();
            var doc = new StringBuilder();
            doc.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"").Append(W)
               .Append("\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><w:body>");

            string lastSect = null;
            for (int i = 0; i < wb.Sheets.Count; i++)
            {
                var sh = wb.Sheets[i];
                if (i > 0) doc.Append(lastSect != null ? "<w:p><w:pPr>" + lastSect + "</w:pPr></w:p>" : "");
                if (wb.Sheets.Count > 1) doc.Append("<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/>").Append(sh.Rtl ? "<w:bidi/>" : "").Append("</w:pPr><w:r>").Append(sh.Rtl ? "<w:rPr><w:rtl/></w:rPr>" : "").Append("<w:t>").Append(Xml.Esc(sh.Name)).Append("</w:t></w:r></w:p>");
                Table(doc, wb, sh, media);
                lastSect = SectPr(sh);
            }
            doc.Append("<w:p/>").Append(lastSect != null ? lastSect.Replace("<w:sectPr>", "<w:sectPr>") : "<w:sectPr/>");
            doc.Append("</w:body></w:document>");

            using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            {
                Put(zip, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/><Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/><Default Extension=\"gif\" ContentType=\"image/gif\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>");
                Put(zip, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
                var rels = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
                for (int m = 0; m < media.Count; m++)
                {
                    rels.Append("<Relationship Id=\"rIdImg").Append(m + 1).Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image").Append(m + 1).Append('.').Append(media[m].Ext ?? "png").Append("\"/>");
                    var e = zip.CreateEntry("word/media/image" + (m + 1) + "." + (media[m].Ext ?? "png"));
                    using (var s = e.Open()) s.Write(media[m].Bytes, 0, media[m].Bytes.Length);
                }
                rels.Append("</Relationships>");
                Put(zip, "word/_rels/document.xml.rels", rels.ToString());
                Put(zip, "word/styles.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"" + W + "\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"" + Xml.Esc(wb.Template != null ? wb.Template.DefaultFont ?? "Tahoma" : "Tahoma") + "\" w:hAnsi=\"" + Xml.Esc(wb.Template != null ? wb.Template.DefaultFont ?? "Tahoma" : "Tahoma") + "\" w:cs=\"" + Xml.Esc(wb.Template != null ? wb.Template.DefaultFont ?? "Tahoma" : "Tahoma") + "\"/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style><w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:spacing w:before=\"240\" w:after=\"120\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr></w:style></w:styles>");
                Put(zip, "word/document.xml", doc.ToString());
            }
        }

        static void Put(ZipArchive zip, string name, string content)
        {
            var e = zip.CreateEntry(name);
            using (var s = e.Open()) { var b = new UTF8Encoding(false).GetBytes(content); s.Write(b, 0, b.Length); }
        }

        static string SectPr(RenderedSheet sh)
        {
            var pg = sh.Page ?? new PageSetupDef();
            int w, h;
            switch (pg.PaperSize) { case 8: w = 16838; h = 23811; break; case 11: w = 8392; h = 11906; break; case 1: w = 12240; h = 15840; break; case 5: w = 12240; h = 20160; break; default: w = 11906; h = 16838; break; }
            bool land = string.Equals(pg.Orientation, "landscape", StringComparison.OrdinalIgnoreCase);
            if (land) { var t = w; w = h; h = t; }
            Func<double, int> tw = inches => (int)Math.Round(inches * 1440);
            return "<w:sectPr><w:pgSz w:w=\"" + w + "\" w:h=\"" + h + "\"" + (land ? " w:orient=\"landscape\"" : "") + "/><w:pgMar w:top=\"" + tw(pg.MarginTop) + "\" w:right=\"" + tw(pg.MarginRight) + "\" w:bottom=\"" + tw(pg.MarginBottom) + "\" w:left=\"" + tw(pg.MarginLeft) + "\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" + (sh.Rtl ? "<w:bidi/>" : "") + "</w:sectPr>";
        }

        static int Twips(double chars) { return (int)Math.Round((chars * 7 + 5) * 15); }

        static void Table(StringBuilder d, RenderedWorkbook wb, RenderedSheet sh, List<RenderedImage> media)
        {
            int cols = Math.Max(1, sh.MaxCol);
            var widths = new int[cols + 1];
            for (int c = 1; c <= cols; c++) widths[c] = Twips(sh.DefaultColWidth);
            foreach (var cd in sh.Columns) if (cd.Index >= 1 && cd.Index <= cols) widths[cd.Index] = cd.Hidden ? 0 : Twips(cd.Width ?? sh.DefaultColWidth);

            // page content width: shrink columns proportionally when the sheet is wider than the page
            var pg = sh.Page ?? new PageSetupDef();
            int pageW = (pg.PaperSize == 8 ? 16838 : pg.PaperSize == 1 || pg.PaperSize == 5 ? 12240 : pg.PaperSize == 11 ? 8392 : 11906);
            if (string.Equals(pg.Orientation, "landscape", StringComparison.OrdinalIgnoreCase)) pageW = pg.PaperSize == 8 ? 23811 : pg.PaperSize == 1 ? 15840 : pg.PaperSize == 5 ? 20160 : pg.PaperSize == 11 ? 11906 : 16838;
            int avail = pageW - (int)((pg.MarginLeft + pg.MarginRight) * 1440);
            int total = widths.Sum();
            if (total > avail && total > 0) for (int c = 1; c <= cols; c++) widths[c] = (int)((long)widths[c] * avail / total);

            // merge maps
            var anchors = new Dictionary<long, int[]>(); var covered = new HashSet<long>();
            foreach (var m in sh.Merges)
            {
                anchors[(long)m[0] * 100000 + m[1]] = m;
                for (int r = m[0]; r <= m[2]; r++) for (int c = m[1]; c <= m[3]; c++) if (r != m[0] || c != m[1]) covered.Add((long)r * 100000 + c);
            }
            var imgAt = new Dictionary<long, RenderedImage>();
            foreach (var im in sh.Images) imgAt[(long)im.Row * 100000 + im.Col] = im;

            var ev = new FormulaEvaluator(sh);
            d.Append("<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>").Append(sh.Rtl ? "<w:bidiVisual/>" : "").Append("<w:tblLayout w:type=\"fixed\"/><w:tblCellMar><w:left w:w=\"60\" w:type=\"dxa\"/><w:right w:w=\"60\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr><w:tblGrid>");
            for (int c = 1; c <= cols; c++) d.Append("<w:gridCol w:w=\"").Append(widths[c]).Append("\"/>");
            d.Append("</w:tblGrid>");

            int rows = Math.Max(1, sh.Rows.Count);
            for (int r = 1; r <= rows; r++)
            {
                var row = r <= sh.Rows.Count ? sh.Rows[r - 1] : null;
                if (row != null && row.Hidden) continue;
                d.Append("<w:tr><w:trPr>");
                if (row != null && row.Height.HasValue) d.Append("<w:trHeight w:val=\"").Append((int)(row.Height.Value * 20)).Append("\" w:hRule=\"atLeast\"/>");
                d.Append("<w:cantSplit/></w:trPr>");
                for (int c = 1; c <= cols; c++)
                {
                    long key = (long)r * 100000 + c;
                    if (widths[c] == 0) continue;
                    int[] span = null; bool vcont = false;
                    if (covered.Contains(key))
                    {
                        // continuation of a vertical merge in the first column of the range only
                        var owner = sh.Merges.FirstOrDefault(m => r >= m[0] && r <= m[2] && c >= m[1] && c <= m[3]);
                        if (owner == null || c != owner[1] || r == owner[0]) continue;
                        vcont = true; span = owner;
                    }
                    else anchors.TryGetValue(key, out span);
                    var cell = row == null ? null : sh.Get(r, c);
                    int gridSpan = span != null ? span[3] - span[1] + 1 : 1;
                    int w = 0; for (int k = c; k < c + gridSpan && k <= cols; k++) w += widths[k];
                    var st = cell == null ? null : cell.Style;

                    d.Append("<w:tc><w:tcPr><w:tcW w:w=\"").Append(w).Append("\" w:type=\"dxa\"/>");
                    if (gridSpan > 1) d.Append("<w:gridSpan w:val=\"").Append(gridSpan).Append("\"/>");
                    if (span != null && span[2] > span[0]) d.Append(vcont ? "<w:vMerge/>" : "<w:vMerge w:val=\"restart\"/>");
                    if (st != null) Borders(d, st);
                    if (st != null && !string.IsNullOrEmpty(st.Fill)) { var f = Hex(st.Fill); if (f != null) d.Append("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"").Append(f).Append("\"/>"); }
                    var v = (st == null ? "" : (st.V ?? "").ToLowerInvariant());
                    d.Append("<w:vAlign w:val=\"").Append(v == "top" ? "top" : v == "bottom" ? "bottom" : "center").Append("\"/></w:tcPr>");
                    if (vcont || cell == null && !imgAt.ContainsKey(key)) { d.Append("<w:p/></w:tc>"); continue; }

                    object val = cell == null ? null : cell.Value;
                    if (cell != null && cell.Formula != null) { val = cell.Cached; object e; if (val == null && ev.TryEvaluate(cell, out e)) val = e; }
                    var xe = val as FormulaEvaluator.XlError;
                    string text = xe != null ? xe.Code : NumberFormatter.Format(val, st == null ? null : st.Format);
                    bool rtl = st != null && st.Rtl.HasValue ? st.Rtl.Value : (sh.Rtl && !(val is double || val is DateTime)) || HasRtl(text);
                    string h = st == null ? "" : (st.H ?? "").ToLowerInvariant();
                    string jc = h == "center" ? "center" : h == "right" ? (rtl ? "left" : "right") : h == "left" ? (rtl ? "right" : "left") : (val is double || val is DateTime) ? (rtl ? "left" : "right") : "";
                    // in a bidi paragraph "left"/"right" are logical (start/end): "right" means start in RTL
                    if (rtl) jc = h == "center" ? "center" : h == "right" ? "left" : h == "left" ? "right" : (val is double || val is DateTime) ? "left" : "";
                    d.Append("<w:p><w:pPr>").Append(rtl ? "<w:bidi/>" : "");
                    if (jc.Length > 0) d.Append("<w:jc w:val=\"").Append(jc).Append("\"/>");
                    d.Append("</w:pPr>");
                    RenderedImage im;
                    if (imgAt.TryGetValue(key, out im) && im.Bytes != null) { media.Add(im); d.Append(Picture(media.Count, im)); }
                    if (text.Length > 0) Run(d, text, st, rtl, wb);
                    d.Append("</w:p></w:tc>");
                }
                d.Append("</w:tr>");
            }
            d.Append("</w:tbl>");
        }

        static string Picture(int index, RenderedImage im)
        {
            long cx = (long)im.Width * 9525, cy = (long)im.Height * 9525;
            return "<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"" + cx + "\" cy=\"" + cy + "\"/><wp:docPr id=\"" + (1000 + index) + "\" name=\"Picture " + index + "\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"" + index + "\" name=\"image" + index + "\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rIdImg" + index + "\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
        }

        static void Run(StringBuilder d, string text, StyleDef st, bool rtl, RenderedWorkbook wb)
        {
            var lines = text.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                d.Append("<w:r><w:rPr>");
                if (st != null && !string.IsNullOrEmpty(st.Font)) d.Append("<w:rFonts w:ascii=\"").Append(Xml.Esc(st.Font)).Append("\" w:hAnsi=\"").Append(Xml.Esc(st.Font)).Append("\" w:cs=\"").Append(Xml.Esc(st.Font)).Append("\"/>");
                if (st != null && st.Bold) d.Append("<w:b/><w:bCs/>");
                if (st != null && st.Italic) d.Append("<w:i/><w:iCs/>");
                if (st != null && st.Strike) d.Append("<w:strike/>");
                if (st != null && !string.IsNullOrEmpty(st.Color)) { var c = Hex(st.Color); if (c != null) d.Append("<w:color w:val=\"").Append(c).Append("\"/>"); }
                double size = st != null && st.Size.HasValue ? st.Size.Value : (wb.Template != null ? wb.Template.DefaultFontSize : 10);
                d.Append("<w:sz w:val=\"").Append((int)Math.Round(size * 2)).Append("\"/><w:szCs w:val=\"").Append((int)Math.Round(size * 2)).Append("\"/>");
                if (st != null && st.Underline) d.Append("<w:u w:val=\"single\"/>");
                if (rtl) d.Append("<w:rtl/>");
                d.Append("</w:rPr>");
                if (i > 0) d.Append("<w:br/>");
                d.Append("<w:t xml:space=\"preserve\">").Append(Xml.Esc(lines[i])).Append("</w:t></w:r>");
            }
        }

        static void Borders(StringBuilder d, StyleDef st)
        {
            if (st.Top == null && st.Bottom == null && st.Left == null && st.Right == null) return;
            d.Append("<w:tcBorders>");
            Side(d, "top", st.Top); Side(d, "left", st.Left); Side(d, "bottom", st.Bottom); Side(d, "right", st.Right);
            d.Append("</w:tcBorders>");
        }

        static void Side(StringBuilder d, string name, BorderDef b)
        {
            if (b == null || string.IsNullOrEmpty(b.Style)) return;
            string val = b.Style == "double" ? "double" : b.Style == "dashed" || b.Style == "mediumdashed" ? "dashed" : b.Style == "dotted" || b.Style == "hair" ? "dotted" : "single";
            int sz = b.Style == "thick" ? 18 : b.Style == "medium" ? 12 : 4;
            d.Append("<w:").Append(name).Append(" w:val=\"").Append(val).Append("\" w:sz=\"").Append(sz).Append("\" w:space=\"0\" w:color=\"").Append(Hex(b.Color) ?? "000000").Append("\"/>");
        }

        static string Hex(string c)
        {
            if (string.IsNullOrEmpty(c)) return null;
            var h = c.Trim().TrimStart('#');
            if (h.Length == 3) h = "" + h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
            if (h.Length != 6) return null;
            foreach (var ch in h) if (!Uri.IsHexDigit(ch)) return null;
            return h.ToUpperInvariant();
        }

        static bool HasRtl(string s)
        {
            foreach (var ch in s) if ((ch >= 0x0590 && ch <= 0x08FF) || (ch >= 0xFB1D && ch <= 0xFDFF) || (ch >= 0xFE70 && ch <= 0xFEFF)) return true;
            return false;
        }
    }
}
