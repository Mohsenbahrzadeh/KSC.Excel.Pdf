using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using Newtonsoft.Json.Linq;

namespace KSC.Excel.Import
{
    public sealed class ImportResult
    {
        public ExcelTemplate Template { get; set; }
        public List<string> Warnings { get; } = new List<string>();
    }

    /// <summary>
    /// Builds a template from an existing .xlsx: sheets, values, formulas, styles (fonts, fills, borders, alignment, number formats,
    /// theme colors), merges, column widths, row heights, freeze panes, RTL, conditional formats, validations, hyperlinks,
    /// images, page setup. Charts, pivot tables and shared-formula children are not imported (reported as warnings).
    /// After importing, replace the cells that should come from data with {{Field}} placeholders in the designer.
    /// </summary>
    public static class XlsxImporter
    {
        const long MaxBytes = 50L * 1024 * 1024;
        const int MaxCells = 500000;
        static readonly XNamespace M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        static readonly XNamespace Pr = "http://schemas.openxmlformats.org/package/2006/relationships";

        public static ImportResult Import(byte[] xlsx, string templateName)
        {
            if (xlsx == null || xlsx.Length == 0) throw new InvalidDataException("The file is empty.");
            if (xlsx.Length > MaxBytes) throw new InvalidDataException("The file is larger than 50 MB.");
            var res = new ImportResult();
            using (var zip = new ZipArchive(new MemoryStream(xlsx), ZipArchiveMode.Read))
            {
                if (zip.Entries.Count > 4000) throw new InvalidDataException("Too many parts in the package.");
                var ctx = new Ctx { Zip = zip, Warn = res.Warnings };
                ctx.Load();
                var t = new ExcelTemplate { Name = templateName, Title = templateName, DefaultFont = ctx.DefaultFont, DefaultFontSize = ctx.DefaultSize };
                var styleIds = new Dictionary<string, string>();
                Func<StyleDef, string> sid = st =>
                {
                    if (st == null) return null;
                    var key = st.Key();
                    if (key == "{}") return null;
                    string id;
                    if (!styleIds.TryGetValue(key, out id)) { id = "s" + (styleIds.Count + 1); styleIds[key] = id; t.Styles[id] = st; }
                    return id;
                };
                int cells = 0;
                foreach (var sheet in ctx.Sheets)
                {
                    var sh = ImportSheet(ctx, sheet, sid, ref cells, res);
                    t.Sheets.Add(sh);
                    if (cells >= MaxCells) { res.Warnings.Add("Import stopped at " + MaxCells + " cells; the remaining sheets were skipped."); break; }
                }
                if (t.Sheets.Count == 0) t.Sheets.Add(new SheetTemplate { Name = "Sheet1" });
                res.Template = t;
            }
            return res;
        }

        // ------------------------------------------------------------------ context
        sealed class SheetRef { public string Name, Path; public int Index; public bool Hidden; }

        sealed class Ctx
        {
            public ZipArchive Zip; public List<string> Warn;
            public List<SheetRef> Sheets = new List<SheetRef>();
            public List<string> Strings = new List<string>();
            public List<StyleDef> Xfs = new List<StyleDef>();
            public List<StyleDef> Dxfs = new List<StyleDef>();
            public string DefaultFont = "Calibri"; public double DefaultSize = 11;
            public string[] Theme = new string[10];
            public Dictionary<int, string> Defined = new Dictionary<int, string>();   // localSheetId|name -> value (key built below)
            public List<Tuple<string, int, string>> Names = new List<Tuple<string, int, string>>();

            public XDocument Doc(string path)
            {
                var e = Zip.GetEntry(path.TrimStart('/'));
                if (e == null) return null;
                using (var s = e.Open())
                using (var xr = XmlReader.Create(s, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                    return XDocument.Load(xr);
            }

            public byte[] Bytes(string path)
            {
                var e = Zip.GetEntry(path.TrimStart('/'));
                if (e == null) return null;
                using (var s = e.Open()) using (var ms = new MemoryStream()) { s.CopyTo(ms); return ms.ToArray(); }
            }

            public void Load()
            {
                var wb = Doc("xl/workbook.xml");
                if (wb == null) throw new InvalidDataException("This is not an .xlsx file (xl/workbook.xml is missing).");
                var rels = Rels("xl/_rels/workbook.xml.rels");
                int i = 0;
                foreach (var s in wb.Descendants(M + "sheet"))
                {
                    var rid = (string)s.Attribute(R + "id");
                    string target;
                    if (rid == null || !rels.TryGetValue(rid, out target)) continue;
                    Sheets.Add(new SheetRef { Name = (string)s.Attribute("name"), Path = Resolve("xl/", target), Index = i++, Hidden = (string)s.Attribute("state") == "hidden" });
                }
                foreach (var dn in wb.Descendants(M + "definedName"))
                    Names.Add(Tuple.Create((string)dn.Attribute("name"), (int?)dn.Attribute("localSheetId") ?? -1, dn.Value));

                var sst = Doc("xl/sharedStrings.xml");
                if (sst != null)
                    foreach (var si in sst.Root.Elements(M + "si"))
                        Strings.Add(string.Concat(si.Descendants(M + "t").Where(t => t.Parent.Name != M + "rPh").Select(t => t.Value)));

                LoadTheme();
                LoadStyles();
            }

            public Dictionary<string, string> Rels(string path)
            {
                var d = new Dictionary<string, string>();
                var doc = Doc(path);
                if (doc != null) foreach (var r in doc.Root.Elements(Pr + "Relationship")) d[(string)r.Attribute("Id")] = ((string)r.Attribute("TargetMode") == "External" ? "!" : "") + (string)r.Attribute("Target");
                return d;
            }

            public static string Resolve(string baseDir, string target)
            {
                if (target.StartsWith("!", StringComparison.Ordinal)) return target;
                if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
                var parts = new List<string>((baseDir + target).Split('/'));
                for (int i = 0; i < parts.Count; i++)
                    if (parts[i] == "..") { parts.RemoveAt(i); if (i > 0) { parts.RemoveAt(i - 1); i--; } i--; }
                    else if (parts[i] == ".") { parts.RemoveAt(i); i--; }
                return string.Join("/", parts);
            }

            void LoadTheme()
            {
                var th = Doc("xl/theme/theme1.xml");
                // dk1 lt1 dk2 lt2 accent1..6 -> Excel theme index: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4..9 accents
                string[] order = { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6" };
                var map = new Dictionary<string, string>();
                if (th != null)
                {
                    var cs = th.Descendants(A + "clrScheme").FirstOrDefault();
                    if (cs != null)
                        foreach (var e in cs.Elements())
                        {
                            var c = e.Elements().FirstOrDefault(); if (c == null) continue;
                            map[e.Name.LocalName] = (string)c.Attribute("val") == null ? null : (c.Name.LocalName == "sysClr" ? ((string)c.Attribute("lastClr") ?? "000000") : (string)c.Attribute("val"));
                        }
                }
                string Get(string k, string def) { string v; return map.TryGetValue(k, out v) && v != null ? v : def; }
                Theme[0] = Get("lt1", "FFFFFF"); Theme[1] = Get("dk1", "000000"); Theme[2] = Get("lt2", "EEECE1"); Theme[3] = Get("dk2", "1F497D");
                string[] acc = { "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646" };
                for (int i = 0; i < 6; i++) Theme[4 + i] = Get("accent" + (i + 1), acc[i]);
            }

            public string Color(XElement c)
            {
                if (c == null) return null;
                var rgb = (string)c.Attribute("rgb");
                string hex = null;
                if (rgb != null) hex = rgb.Length == 8 ? rgb.Substring(2) : rgb;
                else if (c.Attribute("theme") != null)
                {
                    int t = (int)c.Attribute("theme");
                    if (t >= 0 && t < Theme.Length) hex = Theme[t];
                }
                else if (c.Attribute("indexed") != null) hex = Indexed((int)c.Attribute("indexed"));
                if (hex == null || hex.Length != 6) return null;
                double tint = c.Attribute("tint") != null ? (double)c.Attribute("tint") : 0;
                if (Math.Abs(tint) > 0.001) hex = Tint(hex, tint);
                return "#" + hex.ToUpperInvariant();
            }

            static string Indexed(int i)
            {
                string[] p = { "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF" };
                return i >= 0 && i < p.Length ? p[i] : i == 64 || i == 65 ? null : null;
            }

            static string Tint(string hex, double tint)
            {
                int r = Convert.ToInt32(hex.Substring(0, 2), 16), g = Convert.ToInt32(hex.Substring(2, 2), 16), b = Convert.ToInt32(hex.Substring(4, 2), 16);
                Func<int, int> f = v => (int)Math.Round(tint > 0 ? v * (1 - tint) + 255 * tint : v * (1 + tint));
                return f(r).ToString("X2") + f(g).ToString("X2") + f(b).ToString("X2");
            }

            void LoadStyles()
            {
                var st = Doc("xl/styles.xml");
                if (st == null) return;
                var root = st.Root;
                var fmts = new Dictionary<int, string>();
                foreach (var n in root.Descendants(M + "numFmt")) fmts[(int)n.Attribute("numFmtId")] = (string)n.Attribute("formatCode");

                var fonts = new List<Action<StyleDef>>();
                int fi = 0;
                foreach (var f in root.Element(M + "fonts")?.Elements(M + "font") ?? Enumerable.Empty<XElement>())
                {
                    var name = (string)f.Element(M + "name")?.Attribute("val"); var sz = (double?)f.Element(M + "sz")?.Attribute("val");
                    if (fi == 0) { if (name != null) DefaultFont = name; if (sz.HasValue) DefaultSize = sz.Value; }
                    bool b = f.Element(M + "b") != null && (string)f.Element(M + "b").Attribute("val") != "0";
                    bool it = f.Element(M + "i") != null && (string)f.Element(M + "i").Attribute("val") != "0";
                    bool u = f.Element(M + "u") != null && (string)f.Element(M + "u").Attribute("val") != "none";
                    bool sk = f.Element(M + "strike") != null && (string)f.Element(M + "strike").Attribute("val") != "0";
                    var col = Color(f.Element(M + "color"));
                    int idx = fi++;
                    fonts.Add(s =>
                    {
                        if (name != null && !name.Equals(DefaultFont, StringComparison.OrdinalIgnoreCase)) s.Font = name;
                        if (sz.HasValue && Math.Abs(sz.Value - DefaultSize) > 0.01) s.Size = sz;
                        s.Bold = b; s.Italic = it; s.Underline = u; s.Strike = sk;
                        if (col != null && col != "#000000") s.Color = col;
                    });
                }
                var fills = new List<string>();
                foreach (var f in root.Element(M + "fills")?.Elements(M + "fill") ?? Enumerable.Empty<XElement>())
                {
                    var pf = f.Element(M + "patternFill");
                    fills.Add(pf != null && ((string)pf.Attribute("patternType") ?? "none") == "solid" ? Color(pf.Element(M + "fgColor")) : null);
                }
                var borders = new List<BorderDef[]>();
                foreach (var b in root.Element(M + "borders")?.Elements(M + "border") ?? Enumerable.Empty<XElement>())
                {
                    Func<string, BorderDef> side = n =>
                    {
                        var e = b.Element(M + n); var style = (string)e?.Attribute("style");
                        return string.IsNullOrEmpty(style) ? null : new BorderDef { Style = style, Color = Color(e.Element(M + "color")) ?? "#000000" };
                    };
                    borders.Add(new[] { side("left"), side("right"), side("top"), side("bottom") });
                }
                foreach (var xf in root.Element(M + "cellXfs")?.Elements(M + "xf") ?? Enumerable.Empty<XElement>())
                {
                    var s = new StyleDef();
                    int fid = (int?)xf.Attribute("fontId") ?? 0, fl = (int?)xf.Attribute("fillId") ?? 0, bid = (int?)xf.Attribute("borderId") ?? 0, nf = (int?)xf.Attribute("numFmtId") ?? 0;
                    if (fid > 0 && fid < fonts.Count) fonts[fid](s);
                    else if (fid == 0 && fonts.Count > 0) { var tmp = new StyleDef(); fonts[0](tmp); s.Bold = tmp.Bold; s.Italic = tmp.Italic; s.Underline = tmp.Underline; s.Strike = tmp.Strike; s.Color = tmp.Color; }
                    if (fl > 1 && fl < fills.Count) s.Fill = fills[fl];
                    if (bid > 0 && bid < borders.Count) { var bb = borders[bid]; s.Left = bb[0]; s.Right = bb[1]; s.Top = bb[2]; s.Bottom = bb[3]; }
                    s.Format = FormatOf(nf, fmts);
                    var al = xf.Element(M + "alignment");
                    if (al != null)
                    {
                        var h = (string)al.Attribute("horizontal"); if (h != null && h != "general") s.H = h == "centerContinuous" ? "center" : h;
                        var v = (string)al.Attribute("vertical"); if (v != null && v != "bottom") s.V = v;
                        s.Wrap = (string)al.Attribute("wrapText") == "1";
                        s.Indent = (int?)al.Attribute("indent") ?? 0;
                        int rot = (int?)al.Attribute("textRotation") ?? 0; s.Rotation = rot > 90 && rot != 255 ? 90 - rot : rot;
                        s.Shrink = (string)al.Attribute("shrinkToFit") == "1";
                        var ro = (int?)al.Attribute("readingOrder"); if (ro == 2) s.Rtl = true; else if (ro == 1) s.Rtl = false;
                    }
                    Xfs.Add(s);
                }
                foreach (var dx in root.Element(M + "dxfs")?.Elements(M + "dxf") ?? Enumerable.Empty<XElement>())
                {
                    var s = new StyleDef();
                    var f = dx.Element(M + "font");
                    if (f != null)
                    {
                        s.Bold = f.Element(M + "b") != null; s.Italic = f.Element(M + "i") != null; s.Strike = f.Element(M + "strike") != null; s.Underline = f.Element(M + "u") != null;
                        s.Color = Color(f.Element(M + "color"));
                    }
                    var fill = dx.Element(M + "fill")?.Element(M + "patternFill");
                    if (fill != null) s.Fill = Color(fill.Element(M + "bgColor")) ?? Color(fill.Element(M + "fgColor"));
                    var nm = dx.Element(M + "numFmt"); if (nm != null) s.Format = (string)nm.Attribute("formatCode");
                    Dxfs.Add(s);
                }
            }

            static string FormatOf(int id, Dictionary<int, string> custom)
            {
                string c;
                if (custom.TryGetValue(id, out c)) return c == "General" ? null : c;
                switch (id)
                {
                    case 1: return "0"; case 2: return "0.00"; case 3: return "#,##0"; case 4: return "#,##0.00";
                    case 9: return "0%"; case 10: return "0.00%"; case 14: return "yyyy/mm/dd"; case 15: return "d-mmm-yy"; case 16: return "d-mmm";
                    case 17: return "mmm-yy"; case 18: return "h:mm AM/PM"; case 20: return "h:mm"; case 21: return "h:mm:ss"; case 22: return "yyyy/mm/dd hh:mm";
                    case 37: return "#,##0 ;(#,##0)"; case 38: return "#,##0 ;[Red](#,##0)"; case 39: return "#,##0.00;(#,##0.00)"; case 40: return "#,##0.00;[Red](#,##0.00)";
                    case 45: return "mm:ss"; case 46: return "[h]:mm:ss"; case 49: return "@";
                    default: return null;
                }
            }
        }

        // ------------------------------------------------------------------ sheet
        static SheetTemplate ImportSheet(Ctx ctx, SheetRef sref, Func<StyleDef, string> sid, ref int cellCount, ImportResult res)
        {
            var sh = new SheetTemplate { Name = sref.Name, ShowGridLines = true };
            var doc = ctx.Doc(sref.Path);
            if (doc == null) { res.Warnings.Add("Sheet '" + sref.Name + "' could not be read."); return sh; }
            var ws = doc.Root;
            var rels = ctx.Rels(Path.GetDirectoryName(sref.Path).Replace('\\', '/') + "/_rels/" + Path.GetFileName(sref.Path) + ".rels");

            var pr = ws.Element(M + "sheetPr");
            if (pr != null)
            {
                sh.TabColor = ctx.Color(pr.Element(M + "tabColor"));
                if ((string)pr.Element(M + "pageSetUpPr")?.Attribute("fitToPage") == "1") sh.Page.FitToWidth = true;
            }
            var view = ws.Element(M + "sheetViews")?.Element(M + "sheetView");
            if (view != null)
            {
                sh.Rtl = (string)view.Attribute("rightToLeft") == "1";
                sh.ShowGridLines = (string)view.Attribute("showGridLines") != "0";
                sh.Zoom = (int?)view.Attribute("zoomScale") ?? 100;
                var pane = view.Element(M + "pane");
                if (pane != null && (string)pane.Attribute("state") == "frozen") { sh.FreezeCols = (int?)(double?)pane.Attribute("xSplit") ?? 0; sh.FreezeRows = (int?)(double?)pane.Attribute("ySplit") ?? 0; }
            }
            var fmt = ws.Element(M + "sheetFormatPr");
            if (fmt != null)
            {
                var dcw = (double?)fmt.Attribute("defaultColWidth"); if (dcw.HasValue) sh.DefaultColWidth = Math.Max(2, Math.Round((dcw.Value * 7 - 5) / 7, 2));
                var drh = (double?)fmt.Attribute("defaultRowHeight"); if (drh.HasValue) sh.DefaultRowHeight = drh.Value;
            }
            foreach (var c in ws.Element(M + "cols")?.Elements(M + "col") ?? Enumerable.Empty<XElement>())
            {
                int min = (int?)c.Attribute("min") ?? 0, max = (int?)c.Attribute("max") ?? 0; var w = (double?)c.Attribute("width");
                bool hidden = (string)c.Attribute("hidden") == "1";
                if (min < 1 || max - min > 512 || (!w.HasValue && !hidden)) continue;
                for (int i = min; i <= Math.Min(max, 16384); i++)
                    sh.Columns.Add(new ColumnDef { Index = i, Width = w.HasValue ? (double?)Math.Round(Math.Max(1, (w.Value * 7 - 5) / 7), 2) : null, Hidden = hidden });
            }

            // cells
            var sd = ws.Element(M + "sheetData");
            var links = new Dictionary<string, string>();
            foreach (var h in ws.Element(M + "hyperlinks")?.Elements(M + "hyperlink") ?? Enumerable.Empty<XElement>())
            {
                var rid = (string)h.Attribute(R + "id"); string tgt;
                var loc = (string)h.Attribute("location");
                var cellRef = (string)h.Attribute("ref");
                if (cellRef == null) continue;
                if (rid != null && rels.TryGetValue(rid, out tgt) && tgt.StartsWith("!", StringComparison.Ordinal)) links[cellRef.Split(':')[0]] = tgt.Substring(1);
                else if (loc != null) links[cellRef.Split(':')[0]] = "#" + loc;
            }
            bool sharedWarn = false;
            foreach (var row in sd?.Elements(M + "row") ?? Enumerable.Empty<XElement>())
            {
                int rIdx = (int?)row.Attribute("r") ?? 0; if (rIdx < 1) continue;
                var ht = (double?)row.Attribute("ht");
                bool rh = (string)row.Attribute("hidden") == "1";
                if ((ht.HasValue && (string)row.Attribute("customHeight") != "0") || rh) sh.Rows.Add(new RowDef { Index = rIdx, Height = ht.HasValue ? ht : null, Hidden = rh });
                foreach (var c in row.Elements(M + "c"))
                {
                    if (cellCount >= MaxCells) break;
                    var addr = (string)c.Attribute("r"); int rr, cc;
                    if (!CellRef.TryParseCell(addr, out rr, out cc)) continue;
                    var cd = new CellDef { R = rr, C = cc };
                    int sIdx = (int?)c.Attribute("s") ?? 0;
                    if (sIdx > 0 && sIdx < ctx.Xfs.Count) cd.S = sid(ctx.Xfs[sIdx].Clone());
                    var t = (string)c.Attribute("t");
                    var v = (string)c.Element(M + "v");
                    var f = c.Element(M + "f");
                    if (f != null && !string.IsNullOrWhiteSpace(f.Value) && (string)f.Attribute("t") != "array") cd.V = "=" + f.Value;
                    else
                    {
                        if (f != null && (string)f.Attribute("t") == "shared" && !sharedWarn) { sharedWarn = true; res.Warnings.Add("Sheet '" + sh.Name + "': shared formulas were imported as their current values."); }
                        switch (t)
                        {
                            case "s": { int i; cd.V = int.TryParse(v, out i) && i >= 0 && i < ctx.Strings.Count ? ctx.Strings[i] : null; MarkText(cd); break; }
                            case "inlineStr": cd.V = string.Concat(c.Descendants(M + "t").Select(x => x.Value)); MarkText(cd); break;
                            case "str": cd.V = v; MarkText(cd); break;
                            case "b": cd.V = v == "1" ? "TRUE" : "FALSE"; cd.T = "text"; break;
                            case "e": cd.V = v; cd.T = "text"; break;
                            default: cd.V = v; break;   // numbers (and dates as serial numbers with a date format)
                        }
                    }
                    string link;
                    if (links.TryGetValue(addr, out link)) cd.L = link;
                    if (cd.V == null && cd.S == null && cd.L == null) continue;
                    sh.Cells.Add(cd); cellCount++;
                }
            }

            foreach (var m in ws.Element(M + "mergeCells")?.Elements(M + "mergeCell") ?? Enumerable.Empty<XElement>()) { var rf = (string)m.Attribute("ref"); if (rf != null) sh.Merges.Add(rf); }
            sh.AutoFilter = (string)ws.Element(M + "autoFilter")?.Attribute("ref");

            ImportConditionalFormats(ctx, ws, sh);
            ImportValidations(ws, sh);
            ImportPage(ctx, ws, sh, sref);
            ImportDrawing(ctx, ws, rels, sref, sh, res);
            if (ws.Elements(M + "tableParts").Any()) res.Warnings.Add("Sheet '" + sh.Name + "': Excel tables were imported as plain cells.");
            return sh;
        }

        static void MarkText(CellDef cd)
        {
            double d;
            if (cd.V != null && (cd.V.Length > 0 && (double.TryParse(cd.V, NumberStyles.Float, CultureInfo.InvariantCulture, out d) || cd.V.TrimStart().StartsWith("=", StringComparison.Ordinal) || cd.V.Contains("{{")))) cd.T = "text";
        }

        static void ImportConditionalFormats(Ctx ctx, XElement ws, SheetTemplate sh)
        {
            foreach (var cf in ws.Elements(M + "conditionalFormatting"))
            {
                var range = ((string)cf.Attribute("sqref") ?? "").Split(' ')[0];
                if (string.IsNullOrEmpty(range)) continue;
                foreach (var rule in cf.Elements(M + "cfRule"))
                {
                    var type = (string)rule.Attribute("type");
                    int dxf = (int?)rule.Attribute("dxfId") ?? -1;
                    StyleDef style = dxf >= 0 && dxf < ctx.Dxfs.Count ? ctx.Dxfs[dxf].Clone() : null;
                    var f1 = rule.Elements(M + "formula").FirstOrDefault()?.Value; var f2 = rule.Elements(M + "formula").Skip(1).FirstOrDefault()?.Value;
                    CondFormatDef def = null;
                    switch (type)
                    {
                        case "cellIs": def = new CondFormatDef { Type = "cellIs", Operator = (string)rule.Attribute("operator"), Formula1 = f1, Formula2 = f2, Style = style }; break;
                        case "expression": def = new CondFormatDef { Type = "expression", Formula1 = f1, Style = style }; break;
                        case "containsText": def = new CondFormatDef { Type = "containsText", Text = (string)rule.Attribute("text"), Style = style }; break;
                        case "duplicateValues": def = new CondFormatDef { Type = "duplicates", Style = style }; break;
                        case "top10": def = new CondFormatDef { Type = "top10", Rank = (int?)rule.Attribute("rank") ?? 10, Bottom = (string)rule.Attribute("bottom") == "1", Style = style }; break;
                        case "colorScale": def = new CondFormatDef { Type = "colorScale", Colors = rule.Descendants(M + "color").Select(ctx.Color).Where(x => x != null).ToList() }; break;
                        case "dataBar": def = new CondFormatDef { Type = "dataBar", BarColor = ctx.Color(rule.Descendants(M + "color").FirstOrDefault()) }; break;
                    }
                    if (def == null) continue;
                    def.Range = range;
                    sh.ConditionalFormats.Add(def);
                }
            }
        }

        static void ImportValidations(XElement ws, SheetTemplate sh)
        {
            foreach (var dv in ws.Element(M + "dataValidations")?.Elements(M + "dataValidation") ?? Enumerable.Empty<XElement>())
            {
                var type = (string)dv.Attribute("type"); var range = ((string)dv.Attribute("sqref") ?? "").Split(' ')[0];
                if (string.IsNullOrEmpty(type) || string.IsNullOrEmpty(range)) continue;
                var f1 = (string)dv.Element(M + "formula1");
                if (type == "list" && f1 != null && f1.Length > 1 && f1[0] == '"' && f1[f1.Length - 1] == '"') f1 = f1.Substring(1, f1.Length - 2).Replace("\"\"", "\"");
                else if (type == "list" && f1 != null) f1 = "=" + f1;
                sh.Validations.Add(new ValidationDef
                {
                    Range = range, Type = type == "textLength" ? "textLength" : type, Operator = (string)dv.Attribute("operator") ?? "between", Formula1 = f1, Formula2 = (string)dv.Element(M + "formula2"),
                    AllowBlank = (string)dv.Attribute("allowBlank") == "1", ErrorTitle = (string)dv.Attribute("errorTitle"), Error = (string)dv.Attribute("error"),
                    PromptTitle = (string)dv.Attribute("promptTitle"), Prompt = (string)dv.Attribute("prompt")
                });
            }
        }

        static void ImportPage(Ctx ctx, XElement ws, SheetTemplate sh, SheetRef sref)
        {
            var pg = sh.Page;
            var ps = ws.Element(M + "pageSetup");
            if (ps != null)
            {
                if ((string)ps.Attribute("orientation") == "landscape") pg.Orientation = "landscape";
                pg.PaperSize = (int?)ps.Attribute("paperSize") ?? 9;
                var scale = (int?)ps.Attribute("scale"); if (scale.HasValue && !pg.FitToWidth) pg.Scale = scale.Value;
            }
            var pm = ws.Element(M + "pageMargins");
            if (pm != null)
            {
                pg.MarginLeft = (double?)pm.Attribute("left") ?? 0.5; pg.MarginRight = (double?)pm.Attribute("right") ?? 0.5;
                pg.MarginTop = (double?)pm.Attribute("top") ?? 0.75; pg.MarginBottom = (double?)pm.Attribute("bottom") ?? 0.75;
            }
            pg.CenterHorizontally = (string)ws.Element(M + "printOptions")?.Attribute("horizontalCentered") == "1";
            var hf = ws.Element(M + "headerFooter");
            if (hf != null) { pg.Header = (string)hf.Element(M + "oddHeader"); pg.Footer = (string)hf.Element(M + "oddFooter"); }
            foreach (var n in ctx.Names.Where(x => x.Item2 == sref.Index))
            {
                var val = n.Item3; int bang = val.LastIndexOf('!'); if (bang >= 0) val = val.Substring(bang + 1);
                val = val.Replace("$", "");
                if (n.Item1 == "_xlnm.Print_Titles" && System.Text.RegularExpressions.Regex.IsMatch(val, @"^\d+:\d+")) pg.PrintTitleRows = val.Split(',')[0];
                else if (n.Item1 == "_xlnm.Print_Area") pg.PrintArea = val.Split(',')[0];
            }
        }

        static void ImportDrawing(Ctx ctx, XElement ws, Dictionary<string, string> sheetRels, SheetRef sref, SheetTemplate sh, ImportResult res)
        {
            var d = ws.Element(M + "drawing"); var rid = (string)d?.Attribute(R + "id");
            string target;
            if (rid == null || !sheetRels.TryGetValue(rid, out target)) return;
            var dpath = Ctx.Resolve(Path.GetDirectoryName(sref.Path).Replace('\\', '/') + "/", target);
            var drels = ctx.Rels(Path.GetDirectoryName(dpath).Replace('\\', '/') + "/_rels/" + Path.GetFileName(dpath) + ".rels");
            var doc = ctx.Doc(dpath); if (doc == null) return;
            const long Emu = 9525;
            int n = 0;
            foreach (var anchor in doc.Root.Elements())
            {
                var pic = anchor.Element(Xdr + "pic");
                if (anchor.Descendants(Xdr + "graphicFrame").Any()) { res.Warnings.Add("Sheet '" + sh.Name + "': a chart/object was not imported (charts are not read yet)."); continue; }
                if (pic == null) continue;
                var emb = (string)pic.Descendants(A + "blip").FirstOrDefault()?.Attribute(R + "embed");
                string mt;
                if (emb == null || !drels.TryGetValue(emb, out mt)) continue;
                var bytes = ctx.Bytes(Ctx.Resolve(Path.GetDirectoryName(dpath).Replace('\\', '/') + "/", mt));
                if (bytes == null || bytes.Length > 8 * 1024 * 1024) continue;
                var from = anchor.Element(Xdr + "from"); if (from == null) continue;
                int col = (int?)from.Element(Xdr + "col") ?? 0, row = (int?)from.Element(Xdr + "row") ?? 0;
                long cx = 0, cy = 0;
                var ext = anchor.Element(Xdr + "ext");
                if (ext != null) { cx = (long?)ext.Attribute("cx") ?? 0; cy = (long?)ext.Attribute("cy") ?? 0; }
                else
                {
                    var xfrm = pic.Descendants(A + "ext").FirstOrDefault();
                    if (xfrm != null) { cx = (long?)xfrm.Attribute("cx") ?? 0; cy = (long?)xfrm.Attribute("cy") ?? 0; }
                }
                if (cx <= 0 || cy <= 0) { cx = 120 * Emu; cy = 60 * Emu; }
                var extn = Path.GetExtension(mt).TrimStart('.').ToLowerInvariant();
                sh.Images.Add(new ImageDef { Name = "Image " + (++n), Row = row + 1, Col = col + 1, Width = (int)(cx / Emu), Height = (int)(cy / Emu), Data = Convert.ToBase64String(bytes), Ext = extn == "jpg" ? "jpeg" : (extn == "png" || extn == "gif" || extn == "jpeg" ? extn : "png") });
            }
        }
    }
}
