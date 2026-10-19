using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using KSC.Excel.Engine;
using KSC.Excel.Model;

namespace KSC.Excel.Writer
{
    /// <summary>
    /// Writes a <see cref="RenderedWorkbook"/> as a standards-compliant .xlsx (SpreadsheetML) package.
    /// No third-party library: only System.IO.Compression + System.Xml.
    /// </summary>
    public static class XlsxWriter
    {
        const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        const long EmuPerPixel = 9525;

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
            var tpl = wb.Template ?? new ExcelTemplate();
            var styles = new StyleBook(tpl.DefaultFont, tpl.DefaultFontSize);
            var sst = new SharedStrings();

            using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true, new UTF8Encoding(false)))
            {
                // images across all sheets: unique media files
                var media = new List<RenderedImage>();
                foreach (var s in wb.Sheets) foreach (var im in s.Images) media.Add(im);

                WriteContentTypes(zip, wb, media);
                WriteDocParts(zip, tpl);

                // sheets
                var sheetInfos = new List<SheetInfo>();
                int chartCounter = 0;
                for (int i = 0; i < wb.Sheets.Count; i++)
                {
                    var info = new SheetInfo { Index = i + 1, Sheet = wb.Sheets[i], ChartBase = chartCounter };
                    chartCounter += wb.Sheets[i].Charts.Count;
                    WriteSheet(zip, info, styles, sst, media);
                    sheetInfos.Add(info);
                }

                // media
                for (int i = 0; i < media.Count; i++)
                    using (var s = zip.CreateEntry("xl/media/image" + (i + 1) + "." + Ext(media[i]), CompressionLevel.NoCompression).Open())
                        s.Write(media[i].Bytes, 0, media[i].Bytes.Length);

                WriteWorkbook(zip, sheetInfos);
                Part(zip, "xl/styles.xml", styles.Write);
                Part(zip, "xl/sharedStrings.xml", sst.Write);
            }
        }

        /// <summary>_rels/.rels and docProps (shared by the buffered and the streaming writer).</summary>
        internal static void WriteDocParts(ZipArchive zip, ExcelTemplate tpl)
        {
        Part(zip, "_rels/.rels", w =>
        {
            w.WriteStartElement("Relationships", NsPkgRel);
            Rel(w, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
            Rel(w, "rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
            Rel(w, "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
            w.WriteEndElement();
        });
        Part(zip, "docProps/core.xml", w =>
        {
            w.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            w.WriteAttributeString("xmlns", "dc", null, "http://purl.org/dc/elements/1.1/");
            w.WriteAttributeString("xmlns", "dcterms", null, "http://purl.org/dc/terms/");
            w.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
            if (!string.IsNullOrEmpty(tpl.Title)) w.WriteElementString("dc", "title", null, Xml.Clean(tpl.Title));
            w.WriteElementString("dc", "creator", null, Xml.Clean(string.IsNullOrEmpty(tpl.Author) ? "KSC.Excel" : tpl.Author));
            w.WriteStartElement("dcterms", "created", null);
            w.WriteAttributeString("xsi", "type", null, "dcterms:W3CDTF");
            w.WriteString(DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteEndElement();
        });
        Part(zip, "docProps/app.xml", w =>
        {
            w.WriteStartElement("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
            w.WriteElementString("Application", "KSC.Excel");
            w.WriteEndElement();
        });
        }

        // ------------------------------------------------------------------
        internal sealed class SheetInfo
        {
            public int Index;
            public RenderedSheet Sheet;
            public bool HasDrawing;
            public int ChartBase;
        }

        internal static string ImageExt(RenderedImage im) => Ext(im);

        static string Ext(RenderedImage im) => im.Ext == "jpeg" ? "jpeg" : im.Ext == "gif" ? "gif" : "png";

        internal static void WriteContentTypes(ZipArchive zip, RenderedWorkbook wb, List<RenderedImage> media)
        {
            Part(zip, "[Content_Types].xml", w =>
            {
                w.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");
                Def(w, "rels", "application/vnd.openxmlformats-package.relationships+xml");
                Def(w, "xml", "application/xml");
                foreach (var e in media.Select(Ext).Distinct())
                    Def(w, e, e == "jpeg" ? "image/jpeg" : e == "gif" ? "image/gif" : "image/png");
                Over(w, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                for (int i = 1; i <= wb.Sheets.Count; i++)
                {
                    Over(w, "/xl/worksheets/sheet" + i + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                    if (wb.Sheets[i - 1].Images.Count > 0 || wb.Sheets[i - 1].Charts.Count > 0)
                        Over(w, "/xl/drawings/drawing" + i + ".xml", "application/vnd.openxmlformats-officedocument.drawing+xml");
                }
                int totalCharts = wb.Sheets.Sum(x => x.Charts.Count);
                for (int k = 1; k <= totalCharts; k++)
                    Over(w, "/xl/charts/chart" + k + ".xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
                Over(w, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                Over(w, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                Over(w, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
                Over(w, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
                w.WriteEndElement();
            });
        }

        static void Def(XmlWriter w, string ext, string ct) { w.WriteStartElement("Default"); w.WriteAttributeString("Extension", ext); w.WriteAttributeString("ContentType", ct); w.WriteEndElement(); }
        static void Over(XmlWriter w, string part, string ct) { w.WriteStartElement("Override"); w.WriteAttributeString("PartName", part); w.WriteAttributeString("ContentType", ct); w.WriteEndElement(); }
        static void Rel(XmlWriter w, string id, string type, string target, bool external = false)
        {
            w.WriteStartElement("Relationship");
            w.WriteAttributeString("Id", id); w.WriteAttributeString("Type", type); w.WriteAttributeString("Target", target);
            if (external) w.WriteAttributeString("TargetMode", "External");
            w.WriteEndElement();
        }

        internal static void Part(ZipArchive zip, string name, Action<XmlWriter> body)
        {
            var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var s = entry.Open())
            using (var w = XmlWriter.Create(s, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false, CheckCharacters = false }))
            {
                w.WriteStartDocument(true);
                body(w);
                w.WriteEndDocument();
                w.Flush();
            }
        }

        // ------------------------------------------------------------------
        static string QuoteSheet(string name) => "'" + name.Replace("'", "''") + "'";

        static string AbsRange(string a1range)
        {
            int r1, c1, r2, c2;
            if (!CellRef.TryParseRange(a1range, out r1, out c1, out r2, out c2)) return null;
            return "$" + CellRef.ColName(c1) + "$" + r1 + ":$" + CellRef.ColName(c2) + "$" + r2;
        }

        internal static void WriteWorkbook(ZipArchive zip, List<SheetInfo> sheets)
        {
            Part(zip, "xl/workbook.xml", w =>
            {
                w.WriteStartElement("workbook", NsMain);
                w.WriteAttributeString("xmlns", "r", null, NsRel);
                w.WriteStartElement("workbookPr"); w.WriteAttributeString("defaultThemeVersion", "124226"); w.WriteEndElement();
                w.WriteStartElement("bookViews"); w.WriteStartElement("workbookView");
                w.WriteAttributeString("xWindow", "0"); w.WriteAttributeString("yWindow", "0");
                w.WriteAttributeString("windowWidth", "28800"); w.WriteAttributeString("windowHeight", "12000");
                w.WriteEndElement(); w.WriteEndElement();

                w.WriteStartElement("sheets");
                foreach (var s in sheets)
                {
                    w.WriteStartElement("sheet");
                    w.WriteAttributeString("name", Xml.Clean(s.Sheet.Name));
                    w.WriteAttributeString("sheetId", s.Index.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("id", NsRel, "rId" + s.Index);
                    w.WriteEndElement();
                }
                w.WriteEndElement();

                var names = new List<Tuple<string, int, string, bool>>();
                foreach (var s in sheets)
                {
                    var sh = s.Sheet; int local = s.Index - 1;
                    if (!string.IsNullOrEmpty(sh.PrintArea)) { var a = AbsRange(sh.PrintArea); if (a != null) names.Add(Tuple.Create("_xlnm.Print_Area", local, QuoteSheet(sh.Name) + "!" + a, false)); }
                    if (!string.IsNullOrEmpty(sh.PrintTitleRows))
                    {
                        var p = sh.PrintTitleRows.Split(':');
                        names.Add(Tuple.Create("_xlnm.Print_Titles", local, QuoteSheet(sh.Name) + "!$" + p[0] + ":$" + p[1], false));
                    }
                    if (!string.IsNullOrEmpty(sh.AutoFilter)) { var a = AbsRange(sh.AutoFilter); if (a != null) names.Add(Tuple.Create("_xlnm._FilterDatabase", local, QuoteSheet(sh.Name) + "!" + a, true)); }
                }
                if (names.Count > 0)
                {
                    w.WriteStartElement("definedNames");
                    foreach (var n in names)
                    {
                        w.WriteStartElement("definedName");
                        w.WriteAttributeString("name", n.Item1);
                        w.WriteAttributeString("localSheetId", n.Item2.ToString(CultureInfo.InvariantCulture));
                        if (n.Item4) w.WriteAttributeString("hidden", "1");
                        w.WriteString(Xml.Clean(n.Item3));
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                }
                w.WriteStartElement("calcPr"); w.WriteAttributeString("calcId", "191029"); w.WriteAttributeString("fullCalcOnLoad", "1"); w.WriteEndElement();
                w.WriteEndElement();
            });

            Part(zip, "xl/_rels/workbook.xml.rels", w =>
            {
                w.WriteStartElement("Relationships", NsPkgRel);
                foreach (var s in sheets) Rel(w, "rId" + s.Index, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet" + s.Index + ".xml");
                Rel(w, "rId" + (sheets.Count + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
                Rel(w, "rId" + (sheets.Count + 2), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml");
                w.WriteEndElement();
            });
        }

        // ------------------------------------------------------------------
        static void WriteSheet(ZipArchive zip, SheetInfo info, StyleBook styles, SharedStrings sst, List<RenderedImage> allMedia)
        {
            var sh = info.Sheet;
            var links = new List<KeyValuePair<string, string>>(); // ref -> url
            var xfCache = new Dictionary<StyleDef, int>();

            Part(zip, "xl/worksheets/sheet" + info.Index + ".xml", w =>
            {
                WriteSheetHead(w, info, false);
                w.WriteStartElement("sheetData");
                foreach (var row in sh.Rows) WriteRow(w, row, styles, sst, xfCache, links);
                w.WriteEndElement();
                WriteSheetTail(w, info, styles, links);
            });
            WriteSheetExtras(zip, info, links, allMedia);
        }

        internal static void WriteRow(XmlWriter w, RenderedRow row, StyleBook styles, SharedStrings sst, Dictionary<StyleDef, int> xfCache, List<KeyValuePair<string, string>> links)
        {
            bool hasData = row.Cells.Count > 0 || row.Height.HasValue || row.Hidden;
            if (!hasData) return;
            w.WriteStartElement("row");
            w.WriteAttributeString("r", row.Index.ToString(CultureInfo.InvariantCulture));
            if (row.Height.HasValue) { w.WriteAttributeString("ht", Num(row.Height.Value)); w.WriteAttributeString("customHeight", "1"); }
            if (row.Hidden) w.WriteAttributeString("hidden", "1");
            foreach (var c in row.Cells) WriteCell(w, row.Index, c, styles, sst, xfCache, links);
            w.WriteEndElement();
        }

        /// <summary>worksheet start .. cols (everything before sheetData).</summary>
        internal static void WriteSheetHead(XmlWriter w, SheetInfo info, bool streaming)
        {
            var sh = info.Sheet;
        w.WriteStartElement("worksheet", NsMain);
        w.WriteAttributeString("xmlns", "r", null, NsRel);

        var pg = sh.Page ?? new PageSetupDef();
        var tab = StyleBook.Rgb(sh.TabColor);
        if (tab != null || pg.FitToWidth)
        {
            w.WriteStartElement("sheetPr");
            if (tab != null) { w.WriteStartElement("tabColor"); w.WriteAttributeString("rgb", tab); w.WriteEndElement(); }
            if (pg.FitToWidth) { w.WriteStartElement("pageSetUpPr"); w.WriteAttributeString("fitToPage", "1"); w.WriteEndElement(); }
            w.WriteEndElement();
        }

        int lastRow = Math.Max(sh.Rows.Count, 1), lastCol = Math.Max(sh.MaxCol, 1);
        if (!streaming) { w.WriteStartElement("dimension"); w.WriteAttributeString("ref", "A1:" + CellRef.A1(lastRow, lastCol)); w.WriteEndElement(); }

        // views
        w.WriteStartElement("sheetViews");
        w.WriteStartElement("sheetView");
        if (info.Index == 1) w.WriteAttributeString("tabSelected", "1");
        if (sh.Rtl) w.WriteAttributeString("rightToLeft", "1");
        if (!sh.ShowGridLines) w.WriteAttributeString("showGridLines", "0");
        if (sh.Zoom != 100) { w.WriteAttributeString("zoomScale", sh.Zoom.ToString(CultureInfo.InvariantCulture)); w.WriteAttributeString("zoomScaleNormal", sh.Zoom.ToString(CultureInfo.InvariantCulture)); }
        w.WriteAttributeString("workbookViewId", "0");
        int fr = Math.Max(0, streaming ? sh.FreezeRows : Math.Min(sh.FreezeRows, lastRow)), fc = Math.Max(0, sh.FreezeCols);
        if (fr > 0 || fc > 0)
        {
            var pane = fr > 0 && fc > 0 ? "bottomRight" : fr > 0 ? "bottomLeft" : "topRight";
            w.WriteStartElement("pane");
            if (fc > 0) w.WriteAttributeString("xSplit", fc.ToString(CultureInfo.InvariantCulture));
            if (fr > 0) w.WriteAttributeString("ySplit", fr.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("topLeftCell", CellRef.A1(fr + 1, fc + 1));
            w.WriteAttributeString("activePane", pane);
            w.WriteAttributeString("state", "frozen");
            w.WriteEndElement();
            w.WriteStartElement("selection"); w.WriteAttributeString("pane", pane); w.WriteEndElement();
        }
        w.WriteEndElement(); w.WriteEndElement();

        w.WriteStartElement("sheetFormatPr");
        w.WriteAttributeString("defaultColWidth", ColWidth(sh.DefaultColWidth));
        w.WriteAttributeString("defaultRowHeight", Num(sh.DefaultRowHeight));
        w.WriteAttributeString("customHeight", "1");
        w.WriteEndElement();

        var cols = sh.Columns.Where(c => c.Width.HasValue || c.Hidden).GroupBy(c => c.Index).Select(g => g.Last()).OrderBy(c => c.Index).ToList();
        if (cols.Count > 0)
        {
            w.WriteStartElement("cols");
            foreach (var c in cols)
            {
                w.WriteStartElement("col");
                w.WriteAttributeString("min", c.Index.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("max", c.Index.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("width", ColWidth(c.Width ?? sh.DefaultColWidth));
                w.WriteAttributeString("customWidth", "1");
                if (c.Hidden) w.WriteAttributeString("hidden", "1");
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        }

        /// <summary>Everything after sheetData, including the closing worksheet tag.</summary>
        internal static void WriteSheetTail(XmlWriter w, SheetInfo info, StyleBook styles, List<KeyValuePair<string, string>> links)
        {
            var sh = info.Sheet;
            var pg = sh.Page ?? new PageSetupDef();
        if (sh.Protect) { w.WriteStartElement("sheetProtection"); w.WriteAttributeString("sheet", "1"); w.WriteAttributeString("objects", "1"); w.WriteAttributeString("scenarios", "1"); w.WriteEndElement(); }
        if (!string.IsNullOrEmpty(sh.AutoFilter)) { w.WriteStartElement("autoFilter"); w.WriteAttributeString("ref", sh.AutoFilter); w.WriteEndElement(); }

        if (sh.Merges.Count > 0)
        {
            w.WriteStartElement("mergeCells"); w.WriteAttributeString("count", sh.Merges.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var m in sh.Merges)
            { w.WriteStartElement("mergeCell"); w.WriteAttributeString("ref", CellRef.Range(m[0], m[1], m[2], m[3])); w.WriteEndElement(); }
            w.WriteEndElement();
        }

        int priority = 1;
        foreach (var cf in sh.ConditionalFormats) WriteCondFormat(w, cf, styles, ref priority);
        WriteValidations(w, sh.Validations);

        if (links.Count > 0)
        {
            w.WriteStartElement("hyperlinks");
            int rid = 0;
            foreach (var l in links)
            {
                w.WriteStartElement("hyperlink");
                w.WriteAttributeString("ref", l.Key);
                if (l.Value.StartsWith("#", StringComparison.Ordinal)) w.WriteAttributeString("location", Xml.Clean(l.Value.Substring(1)));
                else w.WriteAttributeString("id", NsRel, "rId" + (++rid));
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        if (pg.CenterHorizontally) { w.WriteStartElement("printOptions"); w.WriteAttributeString("horizontalCentered", "1"); w.WriteEndElement(); }
        w.WriteStartElement("pageMargins");
        w.WriteAttributeString("left", Num(pg.MarginLeft)); w.WriteAttributeString("right", Num(pg.MarginRight));
        w.WriteAttributeString("top", Num(pg.MarginTop)); w.WriteAttributeString("bottom", Num(pg.MarginBottom));
        w.WriteAttributeString("header", "0.3"); w.WriteAttributeString("footer", "0.3");
        w.WriteEndElement();
        w.WriteStartElement("pageSetup");
        w.WriteAttributeString("paperSize", pg.PaperSize.ToString(CultureInfo.InvariantCulture));
        if (pg.FitToWidth) { w.WriteAttributeString("fitToWidth", "1"); w.WriteAttributeString("fitToHeight", "0"); }
        else if (pg.Scale > 0 && pg.Scale != 100) w.WriteAttributeString("scale", pg.Scale.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("orientation", string.Equals(pg.Orientation, "landscape", StringComparison.OrdinalIgnoreCase) ? "landscape" : "portrait");
        w.WriteEndElement();
        if (!string.IsNullOrEmpty(pg.Header) || !string.IsNullOrEmpty(pg.Footer))
        {
            w.WriteStartElement("headerFooter");
            if (!string.IsNullOrEmpty(pg.Header)) w.WriteElementString("oddHeader", Xml.Clean(pg.Header));
            if (!string.IsNullOrEmpty(pg.Footer)) w.WriteElementString("oddFooter", Xml.Clean(pg.Footer));
            w.WriteEndElement();
        }

        if (sh.RowBreaks != null && sh.RowBreaks.Count > 0)
        {
            var breaks = sh.RowBreaks.Distinct().OrderBy(x => x).Take(1000).ToList();
            w.WriteStartElement("rowBreaks");
            w.WriteAttributeString("count", breaks.Count.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("manualBreakCount", breaks.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var b in breaks) { w.WriteStartElement("brk"); w.WriteAttributeString("id", b.ToString(CultureInfo.InvariantCulture)); w.WriteAttributeString("max", "16383"); w.WriteAttributeString("man", "1"); w.WriteEndElement(); }
            w.WriteEndElement();
        }

        if (sh.Images.Count > 0 || sh.Charts.Count > 0) { info.HasDrawing = true; w.WriteStartElement("drawing"); w.WriteAttributeString("id", NsRel, "rIdDrawing"); w.WriteEndElement(); }
        w.WriteEndElement();
        }

        /// <summary>Rels part (external hyperlinks + drawing) and the drawing/chart parts of a finished sheet.</summary>
        internal static void WriteSheetExtras(ZipArchive zip, SheetInfo info, List<KeyValuePair<string, string>> links, List<RenderedImage> allMedia)
        {
            var sh = info.Sheet;
            var external = links.Where(l => !l.Value.StartsWith("#", StringComparison.Ordinal)).ToList();
            if (external.Count > 0 || info.HasDrawing)
            {
                Part(zip, "xl/worksheets/_rels/sheet" + info.Index + ".xml.rels", w =>
                {
                    w.WriteStartElement("Relationships", NsPkgRel);
                    int rid = 0;
                    foreach (var l in external) Rel(w, "rId" + (++rid), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", Xml.Clean(l.Value), true);
                    if (info.HasDrawing) Rel(w, "rIdDrawing", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "../drawings/drawing" + info.Index + ".xml");
                    w.WriteEndElement();
                });
            }
            if (info.HasDrawing)
            {
                WriteDrawing(zip, info, allMedia);
                for (int k = 0; k < sh.Charts.Count; k++) WriteChartPart(zip, info.ChartBase + k + 1, sh.Charts[k]);
            }
        }

        /// <summary>Template widths are Excel UI "character" widths; the file stores them plus the 5px cell padding.</summary>
        static string ColWidth(double chars) => (Math.Truncate((chars * 7 + 5) / 7 * 256) / 256).ToString("0.######", CultureInfo.InvariantCulture);

        static string Num(double d) => d.ToString("0.####", CultureInfo.InvariantCulture);

        // Cells are written as raw XML strings: ~4x faster than the XmlWriter element API on big reports.
        [ThreadStatic] static StringBuilder _cellBuilder;   // thread-static: renders may run concurrently

        static void WriteCell(XmlWriter w, int row, RenderedCell c, StyleBook styles, SharedStrings sst, Dictionary<StyleDef, int> xfCache, List<KeyValuePair<string, string>> links)
        {
            int xf = 0;
            if (c.Style != null && !xfCache.TryGetValue(c.Style, out xf)) { xf = styles.Xf(c.Style); xfCache[c.Style] = xf; }
            var colName = CellRef.ColName(c.Col);
            if (!string.IsNullOrEmpty(c.Link)) links.Add(new KeyValuePair<string, string>(colName + row.ToString(CultureInfo.InvariantCulture), c.Link));

            var sb = _cellBuilder ?? (_cellBuilder = new StringBuilder(256)); sb.Clear();
            sb.Append("<c r=\"").Append(colName).Append(row.ToString(CultureInfo.InvariantCulture)).Append('"');
            if (xf != 0) sb.Append(" s=\"").Append(xf.ToString(CultureInfo.InvariantCulture)).Append('"');

            if (c.Formula != null)
            {
                object cv = c.Cached;
                if (cv is DateTime d0) cv = d0.ToOADate();
                if (cv is string) sb.Append(" t=\"str\"");
                else if (cv is bool) sb.Append(" t=\"b\"");
                else if (cv is FormulaEvaluator.XlError) sb.Append(" t=\"e\"");
                sb.Append("><f>").Append(Xml.Esc(c.Formula)).Append("</f>");
                if (cv is double dd) sb.Append("<v>").Append(dd.ToString("R", CultureInfo.InvariantCulture)).Append("</v>");
                else if (cv is string cs) sb.Append("<v>").Append(Xml.Esc(cs)).Append("</v>");
                else if (cv is bool cb) sb.Append(cb ? "<v>1</v>" : "<v>0</v>");
                else if (cv is FormulaEvaluator.XlError ce) sb.Append("<v>").Append(ce.Code).Append("</v>");
                sb.Append("</c>");
            }
            else if (c.Value is string s)
            {
                if (sst != null) sb.Append(" t=\"s\"><v>").Append(sst.Add(s).ToString(CultureInfo.InvariantCulture)).Append("</v></c>");
                else
                {
                    var t = Xml.Esc(s.Length > 32767 ? s.Substring(0, 32767) : s);
                    sb.Append(" t=\"inlineStr\"><is><t");
                    if (t.Length > 0 && (char.IsWhiteSpace(t[0]) || char.IsWhiteSpace(t[t.Length - 1]) || t.IndexOf('\n') >= 0)) sb.Append(" xml:space=\"preserve\"");
                    sb.Append('>').Append(t).Append("</t></is></c>");
                }
            }
            else if (c.Value is double dv)
            {
                if (double.IsNaN(dv) || double.IsInfinity(dv)) sb.Append(" t=\"e\"><v>#NUM!</v></c>");
                else sb.Append("><v>").Append(dv.ToString("R", CultureInfo.InvariantCulture)).Append("</v></c>");
            }
            else if (c.Value is bool b) sb.Append(" t=\"b\"><v>").Append(b ? '1' : '0').Append("</v></c>");
            else if (c.Value is DateTime dt) sb.Append("><v>").Append(dt.ToOADate().ToString("R", CultureInfo.InvariantCulture)).Append("</v></c>");
            else sb.Append("/>");
            w.WriteRaw(sb.ToString());
        }

        // ------------------------------------------------------------------
        static string Formula(string f) => Xml.Clean((f ?? "").TrimStart('='));

        static void WriteCondFormat(XmlWriter w, CondFormatDef cf, StyleBook styles, ref int priority)
        {
            if (string.IsNullOrWhiteSpace(cf.Range) || string.IsNullOrWhiteSpace(cf.Type)) return;
            int r1, c1, r2, c2;
            if (!CellRef.TryParseRange(cf.Range, out r1, out c1, out r2, out c2)) return;
            var type = cf.Type.ToLowerInvariant();
            var topLeft = CellRef.A1(r1, c1);

            string xml = null;
            switch (type)
            {
                case "cellis":
                {
                    var op = CellIsOp(cf.Operator);
                    int dxf = styles.Dxf(cf.Style);
                    xml = "<cfRule type=\"cellIs\" dxfId=\"" + dxf + "\" priority=\"" + priority + "\" operator=\"" + op + "\"><formula>" + Formula(cf.Formula1) + "</formula>"
                        + ((op == "between" || op == "notBetween") ? "<formula>" + Formula(cf.Formula2) + "</formula>" : "") + "</cfRule>";
                    break;
                }
                case "expression":
                    xml = "<cfRule type=\"expression\" dxfId=\"" + styles.Dxf(cf.Style) + "\" priority=\"" + priority + "\"><formula>" + Formula(cf.Formula1) + "</formula></cfRule>";
                    break;
                case "containstext":
                {
                    var t = Xml.Clean(cf.Text ?? "");
                    xml = "<cfRule type=\"containsText\" dxfId=\"" + styles.Dxf(cf.Style) + "\" priority=\"" + priority + "\" operator=\"containsText\" text=\"" + System.Security.SecurityElement.Escape(t) + "\">"
                        + "<formula>NOT(ISERROR(SEARCH(\"" + System.Security.SecurityElement.Escape(t.Replace("\"", "\"\"")) + "\"," + topLeft + ")))</formula></cfRule>";
                    break;
                }
                case "duplicates":
                    xml = "<cfRule type=\"duplicateValues\" dxfId=\"" + styles.Dxf(cf.Style) + "\" priority=\"" + priority + "\"/>";
                    break;
                case "top10":
                    xml = "<cfRule type=\"top10\" dxfId=\"" + styles.Dxf(cf.Style) + "\" priority=\"" + priority + "\" rank=\"" + Math.Max(1, cf.Rank) + "\"" + (cf.Bottom ? " bottom=\"1\"" : "") + "/>";
                    break;
                case "colorscale":
                {
                    var cols = (cf.Colors ?? new List<string>()).Select(StyleBook.Rgb).Where(x => x != null).ToList();
                    if (cols.Count < 2) cols = new List<string> { "FFF8696B", "FFFFEB84", "FF63BE7B" };
                    if (cols.Count > 3) cols = cols.Take(3).ToList();
                    var sb = new System.Text.StringBuilder("<cfRule type=\"colorScale\" priority=\"" + priority + "\"><colorScale>");
                    sb.Append("<cfvo type=\"min\"/>");
                    if (cols.Count == 3) sb.Append("<cfvo type=\"percentile\" val=\"50\"/>");
                    sb.Append("<cfvo type=\"max\"/>");
                    foreach (var c in cols) sb.Append("<color rgb=\"").Append(c).Append("\"/>");
                    sb.Append("</colorScale></cfRule>");
                    xml = sb.ToString();
                    break;
                }
                case "databar":
                    xml = "<cfRule type=\"dataBar\" priority=\"" + priority + "\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"" + (StyleBook.Rgb(cf.BarColor) ?? "FF638EC6") + "\"/></dataBar></cfRule>";
                    break;
            }
            if (xml == null) return;
            priority++;
            w.WriteStartElement("conditionalFormatting");
            w.WriteAttributeString("sqref", CellRef.Range(r1, c1, r2, c2));
            w.WriteRaw(xml);
            w.WriteEndElement();
        }

        static string CellIsOp(string op)
        {
            switch ((op ?? "").ToLowerInvariant())
            {
                case "notequal": return "notEqual";
                case "greaterthan": return "greaterThan";
                case "lessthan": return "lessThan";
                case "greaterthanorequal": return "greaterThanOrEqual";
                case "lessthanorequal": return "lessThanOrEqual";
                case "between": return "between";
                case "notbetween": return "notBetween";
                default: return "equal";
            }
        }

        static void WriteValidations(XmlWriter w, List<ValidationDef> list)
        {
            var ok = list.Where(v => !string.IsNullOrWhiteSpace(v.Range) && !string.IsNullOrWhiteSpace(v.Type) && !string.IsNullOrWhiteSpace(v.Formula1)).ToList();
            if (ok.Count == 0) return;
            w.WriteStartElement("dataValidations"); w.WriteAttributeString("count", ok.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var v in ok)
            {
                int r1, c1, r2, c2;
                if (!CellRef.TryParseRange(v.Range, out r1, out c1, out r2, out c2)) continue;
                var type = v.Type.ToLowerInvariant();
                var xt = type == "list" ? "list" : type == "whole" ? "whole" : type == "decimal" ? "decimal" : type == "date" ? "date" : type == "textlength" ? "textLength" : type == "custom" ? "custom" : null;
                if (xt == null) continue;
                w.WriteStartElement("dataValidation");
                w.WriteAttributeString("type", xt);
                var op = CellIsOp(v.Operator);
                if (xt != "list" && xt != "custom" && op != "between") w.WriteAttributeString("operator", op);
                if (v.AllowBlank) w.WriteAttributeString("allowBlank", "1");
                w.WriteAttributeString("showInputMessage", "1"); w.WriteAttributeString("showErrorMessage", "1");
                if (!string.IsNullOrEmpty(v.ErrorTitle)) w.WriteAttributeString("errorTitle", Xml.Clean(Truncate(v.ErrorTitle, 32)));
                if (!string.IsNullOrEmpty(v.Error)) w.WriteAttributeString("error", Xml.Clean(Truncate(v.Error, 255)));
                if (!string.IsNullOrEmpty(v.PromptTitle)) w.WriteAttributeString("promptTitle", Xml.Clean(Truncate(v.PromptTitle, 32)));
                if (!string.IsNullOrEmpty(v.Prompt)) w.WriteAttributeString("prompt", Xml.Clean(Truncate(v.Prompt, 255)));
                w.WriteAttributeString("sqref", CellRef.Range(r1, c1, r2, c2));
                string f1 = v.Formula1.Trim();
                if (xt == "list") f1 = f1.StartsWith("=", StringComparison.Ordinal) ? f1.Substring(1) : "\"" + f1.Replace("\"", "\"\"") + "\"";
                else f1 = f1.TrimStart('=');
                w.WriteElementString("formula1", Xml.Clean(f1));
                if ((op == "between" || op == "notBetween") && xt != "list" && xt != "custom" && !string.IsNullOrWhiteSpace(v.Formula2))
                    w.WriteElementString("formula2", Xml.Clean(v.Formula2.TrimStart('=')));
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        static string Truncate(string s, int n) => s.Length > n ? s.Substring(0, n) : s;

        // ------------------------------------------------------------------
        static void WriteDrawing(ZipArchive zip, SheetInfo info, List<RenderedImage> allMedia)
        {
            var imgs = info.Sheet.Images;
            Part(zip, "xl/drawings/drawing" + info.Index + ".xml", w =>
            {
                w.WriteStartElement("xdr", "wsDr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                w.WriteAttributeString("xmlns", "a", null, "http://schemas.openxmlformats.org/drawingml/2006/main");
                int n = 0;
                foreach (var im in imgs)
                {
                    n++;
                    w.WriteStartElement("xdr", "oneCellAnchor", null);
                    w.WriteStartElement("xdr", "from", null);
                    w.WriteElementString("xdr", "col", null, (im.Col - 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteElementString("xdr", "colOff", null, "0");
                    w.WriteElementString("xdr", "row", null, (im.Row - 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteElementString("xdr", "rowOff", null, "0");
                    w.WriteEndElement();
                    w.WriteStartElement("xdr", "ext", null);
                    w.WriteAttributeString("cx", (im.Width * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("cy", (im.Height * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();

                    w.WriteStartElement("xdr", "pic", null);
                    w.WriteStartElement("xdr", "nvPicPr", null);
                    w.WriteStartElement("xdr", "cNvPr", null);
                    w.WriteAttributeString("id", (n + 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("name", Xml.Clean(string.IsNullOrEmpty(im.Name) ? "Picture " + n : im.Name));
                    w.WriteEndElement();
                    w.WriteStartElement("xdr", "cNvPicPr", null);
                    w.WriteStartElement("a", "picLocks", null); w.WriteAttributeString("noChangeAspect", "1"); w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteStartElement("xdr", "blipFill", null);
                    w.WriteStartElement("a", "blip", null);
                    w.WriteAttributeString("xmlns", "r", null, NsRel);
                    w.WriteAttributeString("embed", NsRel, "rId" + n);
                    w.WriteEndElement();
                    w.WriteStartElement("a", "stretch", null); w.WriteElementString("a", "fillRect", null, null); w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteStartElement("xdr", "spPr", null);
                    w.WriteStartElement("a", "xfrm", null);
                    w.WriteStartElement("a", "off", null); w.WriteAttributeString("x", "0"); w.WriteAttributeString("y", "0"); w.WriteEndElement();
                    w.WriteStartElement("a", "ext", null);
                    w.WriteAttributeString("cx", (im.Width * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("cy", (im.Height * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteStartElement("a", "prstGeom", null); w.WriteAttributeString("prst", "rect"); w.WriteElementString("a", "avLst", null, null); w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement(); // pic
                    w.WriteElementString("xdr", "clientData", null, null);
                    w.WriteEndElement(); // anchor
                }
                int k2 = 0;
                foreach (var ch in info.Sheet.Charts)
                {
                    k2++;
                    w.WriteStartElement("xdr", "oneCellAnchor", null);
                    w.WriteStartElement("xdr", "from", null);
                    w.WriteElementString("xdr", "col", null, (ch.Col - 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteElementString("xdr", "colOff", null, "0");
                    w.WriteElementString("xdr", "row", null, (ch.Row - 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteElementString("xdr", "rowOff", null, "0");
                    w.WriteEndElement();
                    w.WriteStartElement("xdr", "ext", null);
                    w.WriteAttributeString("cx", (ch.Def.Width * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("cy", (ch.Def.Height * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                    w.WriteStartElement("xdr", "graphicFrame", null);
                    w.WriteAttributeString("macro", "");
                    w.WriteStartElement("xdr", "nvGraphicFramePr", null);
                    w.WriteStartElement("xdr", "cNvPr", null);
                    w.WriteAttributeString("id", (imgs.Count + k2 + 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("name", Xml.Clean(string.IsNullOrEmpty(ch.Def.Name) ? "Chart " + k2 : ch.Def.Name));
                    w.WriteEndElement();
                    w.WriteElementString("xdr", "cNvGraphicFramePr", null, null);
                    w.WriteEndElement();
                    w.WriteStartElement("xdr", "xfrm", null);
                    w.WriteStartElement("a", "off", null); w.WriteAttributeString("x", "0"); w.WriteAttributeString("y", "0"); w.WriteEndElement();
                    w.WriteStartElement("a", "ext", null);
                    w.WriteAttributeString("cx", (ch.Def.Width * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("cy", (ch.Def.Height * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteStartElement("a", "graphic", null);
                    w.WriteStartElement("a", "graphicData", null);
                    w.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                    w.WriteStartElement("c", "chart", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                    w.WriteAttributeString("xmlns", "r", null, NsRel);
                    w.WriteAttributeString("id", NsRel, "rIdChart" + k2);
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement(); // graphicFrame
                    w.WriteElementString("xdr", "clientData", null, null);
                    w.WriteEndElement(); // anchor
                }
                w.WriteEndElement();
            });

            Part(zip, "xl/drawings/_rels/drawing" + info.Index + ".xml.rels", w =>
            {
                w.WriteStartElement("Relationships", NsPkgRel);
                int n = 0;
                foreach (var im in imgs)
                {
                    n++;
                    int mediaIdx = allMedia.IndexOf(im) + 1;
                    Rel(w, "rId" + n, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "../media/image" + mediaIdx + "." + Ext(im));
                }
                for (int k = 1; k <= info.Sheet.Charts.Count; k++)
                    Rel(w, "rIdChart" + k, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart" + (info.ChartBase + k) + ".xml");
                w.WriteEndElement();
            });
        }

        // ------------------------------------------------------------------ charts
        const string NsC = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const string NsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
        static readonly string[] DefaultColors = { "2E75B6", "C55A11", "548235", "7030A0", "BF9000", "C00000", "00A3A3", "7F7F7F" };

        static void WriteChartPart(ZipArchive zip, int number, RenderedChart ch)
        {
            Part(zip, "xl/charts/chart" + number + ".xml", w =>
            {
                w.WriteStartElement("c", "chartSpace", NsC);
                w.WriteAttributeString("xmlns", "a", null, NsA);
                w.WriteAttributeString("xmlns", "r", null, NsRel);
                w.WriteRaw(ChartXml(ch));
                w.WriteEndElement();
            });
        }

        static string ChartXml(RenderedChart ch)
        {
            var d = ch.Def; var type = (d.Type ?? "column").ToLowerInvariant();
            var sb = new StringBuilder();
            sb.Append("<c:roundedCorners val=\"0\"/><c:chart>");
            if (!string.IsNullOrWhiteSpace(d.Title))
                sb.Append("<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1400\" b=\"1\"/><a:t>").Append(Xml.Esc(d.Title)).Append("</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title><c:autoTitleDeleted val=\"0\"/>");
            else sb.Append("<c:autoTitleDeleted val=\"1\"/>");
            sb.Append("<c:plotArea><c:layout/>");

            bool pie = type == "pie";
            int count = pie ? 1 : ch.SeriesValues.Count;
            string grouping = d.Stacked ? "stacked" : "standard";
            string tag = type == "line" ? "lineChart" : type == "area" ? "areaChart" : pie ? "pieChart" : "barChart";
            sb.Append("<c:").Append(tag).Append('>');
            if (tag == "barChart") sb.Append("<c:barDir val=\"").Append(type == "bar" ? "bar" : "col").Append("\"/><c:grouping val=\"").Append(d.Stacked ? "stacked" : "clustered").Append("\"/>");
            else if (tag != "pieChart") sb.Append("<c:grouping val=\"").Append(grouping).Append("\"/>");
            sb.Append("<c:varyColors val=\"").Append(pie ? "1" : "0").Append("\"/>");

            for (int i = 0; i < count; i++)
            {
                var color = StyleBook.Rgb(ch.SeriesColors[i]);
                color = color != null ? color.Substring(2) : DefaultColors[i % DefaultColors.Length];
                sb.Append("<c:ser><c:idx val=\"").Append(i).Append("\"/><c:order val=\"").Append(i).Append("\"/><c:tx>");
                if (ch.SeriesNameIsRef[i]) sb.Append("<c:strRef><c:f>").Append(Xml.Esc(ch.SeriesNames[i])).Append("</c:f></c:strRef>");
                else sb.Append("<c:v>").Append(Xml.Esc(ch.SeriesNames[i])).Append("</c:v>");
                sb.Append("</c:tx>");
                if (!pie)
                {
                    if (tag == "lineChart") sb.Append("<c:spPr><a:ln w=\"28575\"><a:solidFill><a:srgbClr val=\"").Append(color).Append("\"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val=\"circle\"/><c:size val=\"5\"/></c:marker>");
                    else sb.Append("<c:spPr><a:solidFill><a:srgbClr val=\"").Append(color).Append("\"/></a:solidFill></c:spPr>");
                    if (tag == "barChart") sb.Append("<c:invertIfNegative val=\"0\"/>");
                }
                if (d.ShowDataLabels) sb.Append("<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>");
                if (ch.Categories != null) sb.Append("<c:cat><c:strRef><c:f>").Append(Xml.Esc(ch.Categories)).Append("</c:f></c:strRef></c:cat>");
                sb.Append("<c:val><c:numRef><c:f>").Append(Xml.Esc(ch.SeriesValues[i])).Append("</c:f></c:numRef></c:val>");
                if (tag == "lineChart") sb.Append("<c:smooth val=\"0\"/>");
                sb.Append("</c:ser>");
            }
            if (tag == "barChart") sb.Append("<c:gapWidth val=\"120\"/>").Append(d.Stacked ? "<c:overlap val=\"100\"/>" : "");
            if (tag == "lineChart") sb.Append("<c:marker val=\"1\"/>");
            if (tag == "pieChart") sb.Append("<c:firstSliceAng val=\"0\"/>");
            else sb.Append("<c:axId val=\"111111\"/><c:axId val=\"222222\"/>");
            sb.Append("</c:").Append(tag).Append('>');

            if (!pie)
            {
                bool horiz = type == "bar";
                sb.Append("<c:catAx><c:axId val=\"111111\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"").Append(horiz ? "l" : "b")
                  .Append("\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"222222\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/><c:noMultiLvlLbl val=\"0\"/></c:catAx>");
                sb.Append("<c:valAx><c:axId val=\"222222\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"").Append(horiz ? "b" : "l")
                  .Append("\"/><c:majorGridlines><c:spPr><a:ln w=\"6350\"><a:solidFill><a:srgbClr val=\"D9D9D9\"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"111111\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>");
            }
            sb.Append("</c:plotArea>");
            if (d.ShowLegend && (pie || count > 1)) sb.Append("<c:legend><c:legendPos val=\"b\"/><c:overlay val=\"0\"/></c:legend>");
            sb.Append("<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"gap\"/></c:chart>");
            return sb.ToString();
        }
    }

    internal static class Xml
    {
        /// <summary>Escapes element text (also strips characters that are illegal in XML).</summary>
        public static string Esc(string s)
        {
            s = Clean(s);
            if (s.IndexOfAny(EscChars) < 0) return s;
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        static readonly char[] EscChars = { '&', '<', '>' };

        /// <summary>Removes characters that are illegal in XML 1.0.</summary>
        public static string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return s ?? "";
            StringBuilder sb = null;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                bool ok = c == 0x9 || c == 0xA || c == 0xD || (c >= 0x20 && c <= 0xD7FF) || (c >= 0xE000 && c <= 0xFFFD)
                          || (char.IsHighSurrogate(c) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                          || (char.IsLowSurrogate(c) && i > 0 && char.IsHighSurrogate(s[i - 1]));
                if (!ok) { if (sb == null) sb = new StringBuilder(s, 0, i, s.Length); continue; }
                sb?.Append(c);
            }
            return sb == null ? s : sb.ToString();
        }
    }

    internal sealed class SharedStrings
    {
        readonly Dictionary<string, int> _idx = new Dictionary<string, int>(StringComparer.Ordinal);
        readonly List<string> _list = new List<string>();
        int _refs;

        public int Add(string s)
        {
            _refs++;
            if (s.Length > 32767) s = s.Substring(0, 32767);
            int i;
            if (_idx.TryGetValue(s, out i)) return i;
            i = _list.Count; _list.Add(s); _idx[s] = i;
            return i;
        }

        public void Write(XmlWriter w)
        {
            w.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            w.WriteAttributeString("count", _refs.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("uniqueCount", _list.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var s in _list)
            {
                w.WriteStartElement("si");
                w.WriteStartElement("t");
                var c = Xml.Clean(s);
                if (c.Length > 0 && (char.IsWhiteSpace(c[0]) || char.IsWhiteSpace(c[c.Length - 1]) || c.IndexOf('\n') >= 0))
                    w.WriteAttributeString("xml", "space", null, "preserve");
                w.WriteString(c);
                w.WriteEndElement(); w.WriteEndElement();
            }
            w.WriteEndElement();
        }
    }
}
