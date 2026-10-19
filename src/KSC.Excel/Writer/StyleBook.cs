using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using KSC.Excel.Model;

namespace KSC.Excel.Writer
{
    /// <summary>Collects unique fonts / fills / borders / number formats / cell formats and writes xl/styles.xml.</summary>
    internal sealed class StyleBook
    {
        readonly string _defFont;
        readonly double _defSize;

        readonly List<string> _fonts = new List<string>();
        readonly Dictionary<string, int> _fontIdx = new Dictionary<string, int>();
        readonly List<string> _fills = new List<string>();
        readonly Dictionary<string, int> _fillIdx = new Dictionary<string, int>();
        readonly List<string> _borders = new List<string>();
        readonly Dictionary<string, int> _borderIdx = new Dictionary<string, int>();
        readonly Dictionary<string, int> _numFmts = new Dictionary<string, int>();
        readonly List<string> _xfs = new List<string>();
        readonly Dictionary<string, int> _xfIdx = new Dictionary<string, int>();
        readonly List<string> _dxfs = new List<string>();

        static readonly Dictionary<string, int> BuiltIn = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            { "General", 0 }, { "0", 1 }, { "0.00", 2 }, { "#,##0", 3 }, { "#,##0.00", 4 },
            { "0%", 9 }, { "0.00%", 10 }, { "@", 49 }
        };

        public StyleBook(string defaultFont, double defaultSize)
        {
            _defFont = string.IsNullOrWhiteSpace(defaultFont) ? "Calibri" : defaultFont;
            _defSize = defaultSize > 0 ? defaultSize : 11;
            Font(null);                                    // font 0
            _fills.Add("<fill><patternFill patternType=\"none\"/></fill>");
            _fills.Add("<fill><patternFill patternType=\"gray125\"/></fill>");
            _fillIdx["none"] = 0;
            Border(null, null, null, null);                // border 0
            Xf(null);                                      // xf 0
        }

        public static string Rgb(string color)
        {
            if (string.IsNullOrWhiteSpace(color)) return null;
            var c = color.Trim().TrimStart('#');
            if (c.Length == 3) c = "" + c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
            if (c.Length != 6) return null;
            foreach (var ch in c) if (!Uri.IsHexDigit(ch)) return null;
            return "FF" + c.ToUpperInvariant();
        }

        static string Attr(string s) => System.Security.SecurityElement.Escape(s) ?? "";

        int Font(StyleDef s)
        {
            var sb = new StringBuilder("<font>");
            if (s != null && s.Bold) sb.Append("<b/>");
            if (s != null && s.Italic) sb.Append("<i/>");
            if (s != null && s.Strike) sb.Append("<strike/>");
            if (s != null && s.Underline) sb.Append("<u/>");
            sb.Append("<sz val=\"").Append((s != null && s.Size.HasValue ? s.Size.Value : _defSize).ToString("0.##", CultureInfo.InvariantCulture)).Append("\"/>");
            var col = s == null ? null : Rgb(s.Color);
            if (col != null) sb.Append("<color rgb=\"").Append(col).Append("\"/>"); else sb.Append("<color theme=\"1\"/>");
            sb.Append("<name val=\"").Append(Attr(s != null && !string.IsNullOrWhiteSpace(s.Font) ? s.Font : _defFont)).Append("\"/>");
            sb.Append("<family val=\"2\"/></font>");
            return Intern(_fonts, _fontIdx, sb.ToString());
        }

        int Fill(string color)
        {
            var c = Rgb(color);
            if (c == null) return 0;
            return Intern(_fills, _fillIdx, "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"" + c + "\"/><bgColor indexed=\"64\"/></patternFill></fill>");
        }

        static void Side(StringBuilder sb, string tag, BorderDef b)
        {
            var style = Style(b);
            if (style == null) { sb.Append('<').Append(tag).Append("/>"); return; }
            var col = Rgb(b.Color);
            sb.Append('<').Append(tag).Append(" style=\"").Append(style).Append("\">")
              .Append(col != null ? "<color rgb=\"" + col + "\"/>" : "<color auto=\"1\"/>")
              .Append("</").Append(tag).Append('>');
        }

        static string Style(BorderDef b)
        {
            if (b == null || string.IsNullOrWhiteSpace(b.Style)) return null;
            switch (b.Style.ToLowerInvariant())
            {
                case "thin": case "medium": case "thick": case "dashed": case "dotted": case "double": case "hair":
                case "mediumdashed": case "dashdot": case "mediumdashdot": case "dashdotdot": case "slantdashdot":
                    return b.Style.ToLowerInvariant();
                default: return "thin";
            }
        }

        int Border(BorderDef l, BorderDef r, BorderDef t, BorderDef b)
        {
            var sb = new StringBuilder("<border>");
            Side(sb, "left", l); Side(sb, "right", r); Side(sb, "top", t); Side(sb, "bottom", b);
            sb.Append("<diagonal/></border>");
            return Intern(_borders, _borderIdx, sb.ToString());
        }

        int NumFmt(string fmt)
        {
            if (string.IsNullOrWhiteSpace(fmt)) return 0;
            int id;
            if (BuiltIn.TryGetValue(fmt, out id)) return id;
            if (!_numFmts.TryGetValue(fmt, out id)) { id = 164 + _numFmts.Count; _numFmts[fmt] = id; }
            return id;
        }

        static int Intern(List<string> list, Dictionary<string, int> idx, string xml)
        {
            int i;
            if (idx.TryGetValue(xml, out i)) return i;
            i = list.Count; list.Add(xml); idx[xml] = i;
            return i;
        }

        public int Xf(StyleDef s)
        {
            int font = Font(s), fill = s == null ? 0 : Fill(s.Fill);
            int border = s == null ? 0 : Border(s.Left, s.Right, s.Top, s.Bottom);
            int fmt = s == null ? 0 : NumFmt(s.Format);

            var sb = new StringBuilder("<xf numFmtId=\"").Append(fmt).Append("\" fontId=\"").Append(font).Append("\" fillId=\"").Append(fill)
                .Append("\" borderId=\"").Append(border).Append("\" xfId=\"0\"");
            if (fmt != 0) sb.Append(" applyNumberFormat=\"1\"");
            if (font != 0) sb.Append(" applyFont=\"1\"");
            if (fill != 0) sb.Append(" applyFill=\"1\"");
            if (border != 0) sb.Append(" applyBorder=\"1\"");

            string align = s == null ? null : Alignment(s);
            if (align != null) sb.Append(" applyAlignment=\"1\">").Append(align).Append("</xf>");
            else sb.Append("/>");
            return Intern(_xfs, _xfIdx, sb.ToString());
        }

        static string Alignment(StyleDef s)
        {
            var sb = new StringBuilder("<alignment");
            bool any = false;
            string h = HAlign(s.H), v = VAlign(s.V);
            if (h != null) { sb.Append(" horizontal=\"").Append(h).Append('"'); any = true; }
            if (v != null) { sb.Append(" vertical=\"").Append(v).Append('"'); any = true; }
            if (s.Wrap) { sb.Append(" wrapText=\"1\""); any = true; }
            if (s.Indent > 0) { sb.Append(" indent=\"").Append(Math.Min(s.Indent, 15)).Append('"'); any = true; }
            if (s.Rotation != 0 && ((s.Rotation >= -90 && s.Rotation <= 90) || s.Rotation == 255))
            { sb.Append(" textRotation=\"").Append(s.Rotation < 0 ? 90 - s.Rotation : s.Rotation).Append('"'); any = true; }
            if (s.Shrink) { sb.Append(" shrinkToFit=\"1\""); any = true; }
            if (s.Rtl.HasValue) { sb.Append(" readingOrder=\"").Append(s.Rtl.Value ? 2 : 1).Append('"'); any = true; }
            if (!any) return null;
            return sb.Append("/>").ToString();
        }

        static string HAlign(string h)
        {
            switch ((h ?? "").ToLowerInvariant())
            {
                case "left": case "center": case "right": case "justify": case "fill": case "distributed": return h.ToLowerInvariant();
                default: return null;
            }
        }

        static string VAlign(string v)
        {
            switch ((v ?? "").ToLowerInvariant())
            {
                case "top": return "top";
                case "center": case "middle": return "center";
                case "bottom": return "bottom";
                case "justify": return "justify";
                default: return null;
            }
        }

        /// <summary>Differential format for conditional formatting rules.</summary>
        public int Dxf(StyleDef s)
        {
            var sb = new StringBuilder("<dxf>");
            if (s != null)
            {
                if (s.Bold || s.Italic || s.Strike || s.Underline || Rgb(s.Color) != null)
                {
                    sb.Append("<font>");
                    if (s.Bold) sb.Append("<b/>");
                    if (s.Italic) sb.Append("<i/>");
                    if (s.Strike) sb.Append("<strike/>");
                    if (s.Underline) sb.Append("<u/>");
                    var c = Rgb(s.Color); if (c != null) sb.Append("<color rgb=\"").Append(c).Append("\"/>");
                    sb.Append("</font>");
                }
                if (!string.IsNullOrWhiteSpace(s.Format))
                {
                    int id = NumFmt(s.Format);
                    sb.Append("<numFmt numFmtId=\"").Append(id).Append("\" formatCode=\"").Append(Attr(s.Format)).Append("\"/>");
                }
                var f = Rgb(s.Fill);
                if (f != null) sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"").Append(f).Append("\"/><bgColor rgb=\"").Append(f).Append("\"/></patternFill></fill>");
                if (Style(s.Left) != null || Style(s.Right) != null || Style(s.Top) != null || Style(s.Bottom) != null)
                {
                    sb.Append("<border>");
                    Side(sb, "left", s.Left); Side(sb, "right", s.Right); Side(sb, "top", s.Top); Side(sb, "bottom", s.Bottom);
                    sb.Append("</border>");
                }
            }
            sb.Append("</dxf>");
            _dxfs.Add(sb.ToString());
            return _dxfs.Count - 1;
        }

        public void Write(XmlWriter w)
        {
            w.WriteStartElement("styleSheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            if (_numFmts.Count > 0)
            {
                w.WriteStartElement("numFmts"); w.WriteAttributeString("count", _numFmts.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var kv in _numFmts)
                {
                    w.WriteStartElement("numFmt");
                    w.WriteAttributeString("numFmtId", kv.Value.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("formatCode", kv.Key);
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }
            Block(w, "fonts", _fonts);
            Block(w, "fills", _fills);
            Block(w, "borders", _borders);

            w.WriteStartElement("cellStyleXfs"); w.WriteAttributeString("count", "1");
            w.WriteRaw("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>");
            w.WriteEndElement();

            Block(w, "cellXfs", _xfs);

            w.WriteStartElement("cellStyles"); w.WriteAttributeString("count", "1");
            w.WriteRaw("<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>");
            w.WriteEndElement();

            Block(w, "dxfs", _dxfs);
            w.WriteRaw("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>");
            w.WriteEndElement();
        }

        static void Block(XmlWriter w, string name, List<string> items)
        {
            w.WriteStartElement(name);
            w.WriteAttributeString("count", items.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var x in items) w.WriteRaw(x);
            w.WriteEndElement();
        }
    }
}
