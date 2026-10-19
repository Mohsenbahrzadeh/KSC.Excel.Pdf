using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using KSC.Excel.Model;

namespace KSC.Excel.Engine
{
    /// <summary>Binds an <see cref="ExcelTemplate"/> to a data object and produces a <see cref="RenderedWorkbook"/>.</summary>
    /// <summary>One group produced by BandDef.GroupBy.</summary>
    public sealed class GroupItem
    {
        public object Key { get; set; }
        public List<object> Items { get; } = new List<object>();
        public int Count { get { return Items.Count; } }
    }

    public sealed class TemplateRenderer
    {
        readonly RenderOptions _opt;
        public TemplateRenderer(RenderOptions options = null) { _opt = options ?? new RenderOptions(); }

        public RenderedWorkbook Render(ExcelTemplate template, object data) { return RenderCore(template, data, null); }

        /// <summary>Renders while delivering rows to <paramref name="sink"/>; rows are not kept in memory. Check <see cref="StreamAnalyzer"/> first.</summary>
        public RenderedWorkbook RenderStreaming(ExcelTemplate template, object data, IStreamSink sink)
        {
            if (sink == null) throw new ArgumentNullException(nameof(sink));
            return RenderCore(template, data, sink);
        }

        RenderedWorkbook RenderCore(ExcelTemplate template, object data, IStreamSink sink)
        {
            if (template == null) throw new ArgumentNullException(nameof(template));
            var wb = new RenderedWorkbook { Template = template };
            var warn = new Action<string>(w => { if (!wb.Warnings.Contains(w)) wb.Warnings.Add(w); });
            var rootScope = new DataScope(DataAccess.Normalize(data));
            rootScope.Params = BuildParameters(template, _opt.Parameters);
            BuildDataSets(template, rootScope, warn);
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var sheet in template.Sheets)
            {
                if (!string.IsNullOrWhiteSpace(sheet.RepeatFor))
                {
                    var items = DataAccess.Items(DataAccess.Resolve(rootScope, sheet.RepeatFor)).ToList();
                    int i = 0;
                    foreach (var it in items)
                    {
                        var sc = new DataScope(it, rootScope) { Index = ++i, Count = items.Count };
                        wb.Sheets.Add(new SheetRenderer(template, sheet, _opt, warn, usedNames, sink).Render(sc));
                    }
                }
                else wb.Sheets.Add(new SheetRenderer(template, sheet, _opt, warn, usedNames, sink).Render(rootScope));
            }
            if (wb.Sheets.Count == 0)
                wb.Sheets.Add(new SheetRenderer(template, new SheetTemplate(), _opt, warn, usedNames, sink).Render(rootScope));
            return wb;
        }

        static Dictionary<string, object> BuildParameters(ExcelTemplate t, IDictionary<string, object> given)
        {
            var p = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in t.Parameters)
            {
                if (string.IsNullOrWhiteSpace(d.Name)) continue;
                object v = d.Default;
                if (d.Default != null)
                {
                    double n; DateTime dt; bool b;
                    switch ((d.Type ?? "text").ToLowerInvariant())
                    {
                        case "number": if (double.TryParse(d.Default, NumberStyles.Float, CultureInfo.InvariantCulture, out n)) v = n; break;
                        case "date": if (DateTime.TryParse(d.Default, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) v = dt; break;
                        case "bool": if (bool.TryParse(d.Default, out b)) v = b; break;
                        case "jdate": if (Jalali.TryParse(d.Default, out dt)) v = dt; break;
                    }
                }
                p[d.Name] = v;
            }
            if (given != null) foreach (var kv in given) p[kv.Key] = DataAccess.Normalize(kv.Value);
            foreach (var d in t.Parameters)   // a Jalali date parameter may be passed as text ("1404/03/15")
            {
                object cur; DateTime jd;
                if (!string.IsNullOrWhiteSpace(d.Name) && string.Equals(d.Type, "jdate", StringComparison.OrdinalIgnoreCase) && p.TryGetValue(d.Name, out cur) && cur is string && Jalali.TryParse((string)cur, out jd)) p[d.Name] = jd;
            }
            return p;
        }

        void BuildDataSets(ExcelTemplate t, DataScope root, Action<string> warn)
        {
            if (t.DataSets == null || t.DataSets.Count == 0) return;
            var x = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            root.Extras = x;   // visible while building, so a pivot can use an earlier SQL data set
            foreach (var d in t.DataSets)
            {
                if (string.IsNullOrWhiteSpace(d.Name)) continue;
                if (string.Equals(d.Type, "sql", StringComparison.OrdinalIgnoreCase))
                {
                    if (_opt.SkipSql) { warn("SQL data set '" + d.Name + "' was not executed (preview); sample data is used if available."); continue; }
                    if (!_opt.AllowSql) throw new ExcelRenderException("The template uses the SQL data set '" + d.Name + "' but SQL data sets are disabled. Set KscExcelOptions.AllowSql = true and register a connection.");
                    if (_opt.ConnectionFactory == null) throw new ExcelRenderException("No connection factory is configured for SQL data sets.");
                    var con = _opt.ConnectionFactory(d.Connection);
                    if (con == null) throw new ExcelRenderException("SQL data set '" + d.Name + "': connection '" + d.Connection + "' is not registered (KscExcelOptions.AddConnection).");
                    if (_opt.LazySql)
                    {
                        con.Dispose();
                        var cname = d.Connection;
                        x[d.Name] = SqlDataSource.QueryLazy(() => _opt.ConnectionFactory(cname), d.Sql, root.Params, _opt.SqlTimeoutSeconds, _opt.SqlMaxRows, warn);
                    }
                    else x[d.Name] = SqlDataSource.Query(con, d.Sql, root.Params, _opt.SqlTimeoutSeconds, _opt.SqlMaxRows, warn);
                    continue;
                }
                if (string.Equals(d.Type, "json", StringComparison.OrdinalIgnoreCase) || string.Equals(d.Type, "provider", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        var loaded = ExternalData.Load(d, _opt, root.Params, warn);
                        if (loaded != null) x[d.Name] = loaded;
                    }
                    catch (ExcelRenderException ex) when (_opt.Preview) { warn(ex.Message + " (preview: sample data is used if available)"); }
                    continue;
                }
                if (!string.Equals(d.Type ?? "pivot", "pivot", StringComparison.OrdinalIgnoreCase)) { warn("Unknown data set type '" + d.Type + "'."); continue; }
                if (string.IsNullOrWhiteSpace(d.Source) || string.IsNullOrWhiteSpace(d.RowPath) || string.IsNullOrWhiteSpace(d.ColPath)) { warn("Data set '" + d.Name + "' needs Source, RowPath and ColPath."); continue; }
                x[d.Name] = Pivot.Build(DataAccess.Resolve(root, d.Source), d.RowPath, d.ColPath, d.ValuePath, d.Agg);
            }
        }

        // ------------------------------------------------------------------
        sealed class SheetRenderer
        {
            // [Band.Field] = column of the cell bound to {{Field}}; [Band:E] = column letter E; [Field] = same-row cell
            static readonly Regex RowToken = new Regex(@"\[row:([A-Z]{1,3}):([A-Z]{1,3})\]", RegexOptions.Compiled);
            static readonly Regex FormulaToken = new Regex(@"\[([A-Za-z_][\w]*)(?:([.:])([A-Za-z_][\w\.]*))?\]", RegexOptions.Compiled);

            readonly ExcelTemplate _tpl;
            SheetTemplate _sh;
            readonly RenderOptions _opt;
            readonly Action<string> _warn;
            readonly HashSet<string> _usedNames;
            readonly TextEvaluator _text;

            readonly RenderedSheet _out = new RenderedSheet();
            readonly Dictionary<int, List<CellDef>> _cellsByRow = new Dictionary<int, List<CellDef>>();
            readonly Dictionary<int, RowDef> _rowDefs = new Dictionary<int, RowDef>();
            readonly Dictionary<int, List<int[]>> _mergesByStart = new Dictionary<int, List<int[]>>();
            readonly Dictionary<BandDef, List<BandDef>> _children = new Dictionary<BandDef, List<BandDef>>();
            readonly List<BandDef> _rootBands = new List<BandDef>();
            readonly Dictionary<string, BandDef> _bandByName = new Dictionary<string, BandDef>(StringComparer.OrdinalIgnoreCase);
            readonly List<Tuple<RenderedCell, int, DataScope>> _pendingFormulas = new List<Tuple<RenderedCell, int, DataScope>>();
            readonly Dictionary<StyleDef, Dictionary<string, StyleDef>> _derived = new Dictionary<StyleDef, Dictionary<string, StyleDef>>();

            int _outRow;
            bool _needRowMap;          // only formulas with [Field] tokens need template-row -> output-row maps
            long _cellCount;
            int _maxTplRow;
            int[] _first, _last;

            readonly IStreamSink _sink;

            public SheetRenderer(ExcelTemplate tpl, SheetTemplate sh, RenderOptions opt, Action<string> warn, HashSet<string> usedNames, IStreamSink sink = null)
            {
                _sink = sink;
                _tpl = tpl; _sh = sh; _opt = opt; _warn = warn; _usedNames = usedNames;
                _text = new TextEvaluator(warn, opt.Filters, opt.Functions);
            }

            public RenderedSheet Render(DataScope scope)
            {
                _sh = ColumnExpander.Expand(_sh, scope, _warn);
                Prepare();
                _out.Name = UniqueName(Sanitize(_text.EvalText(_sh.Name, scope)));
                _out.Rtl = _sh.Rtl; _out.ShowGridLines = _sh.ShowGridLines; _out.TabColor = _sh.TabColor; _out.Zoom = _sh.Zoom;
                _out.DefaultColWidth = _sh.DefaultColWidth; _out.DefaultRowHeight = _sh.DefaultRowHeight;
                _out.Protect = _sh.Protect; _out.Page = _sh.Page ?? new PageSetupDef();
                _out.Columns = _sh.Columns.Where(c => c.Index >= 1).ToList();
                if (_sink != null) { _out.FreezeRows = _sh.FreezeRows; _sink.BeginSheet(_out); }

                EmitRange(1, _maxTplRow, _rootBands, scope);

                ResolveFormulas();
                if (scope.HasRuns) foreach (var kv in scope.Runs) _out.BandRuns[kv.Key] = new[] { kv.Value[0], kv.Value[1] };
                MapDecorations(scope);

                foreach (var r in _out.Rows) if (r.Cells.Count > 0) _out.MaxCol = Math.Max(_out.MaxCol, r.Cells[r.Cells.Count - 1].Col);
                foreach (var c in _out.Columns) _out.MaxCol = Math.Max(_out.MaxCol, c.Index);

                if (_opt.ComputeFormulaValues && _sink == null)
                {
                    var ev = new FormulaEvaluator(_out, _opt.Functions);
                    foreach (var row in _out.Rows)
                        foreach (var c in row.Cells)
                            if (c.Formula != null) { object v; if (ev.TryEvaluate(c, out v)) c.Cached = v; }
                }
                if (_sink != null) _sink.EndSheet(_out);
                return _out;
            }

            // ---------- preparation ----------
            void Prepare()
            {
                foreach (var g in _sh.Cells.Where(c => c.R >= 1 && c.C >= 1).GroupBy(c => c.R))
                {
                    _cellsByRow[g.Key] = g.OrderBy(c => c.C).ToList();
                    _maxTplRow = Math.Max(_maxTplRow, g.Key);
                }
                foreach (var r in _sh.Rows) { _rowDefs[r.Index] = r; _maxTplRow = Math.Max(_maxTplRow, r.Index); }

                foreach (var m in _sh.Merges)
                {
                    int r1, c1, r2, c2;
                    if (!CellRef.TryParseRange(m, out r1, out c1, out r2, out c2) || (r1 == r2 && c1 == c2)) continue;
                    List<int[]> l;
                    if (!_mergesByStart.TryGetValue(r1, out l)) _mergesByStart[r1] = l = new List<int[]>();
                    l.Add(new[] { r1, c1, r2, c2 });
                    _maxTplRow = Math.Max(_maxTplRow, r2);
                }

                var bands = _sh.Bands.Where(b => b.FirstRow >= 1 && b.LastRow >= b.FirstRow && !string.IsNullOrWhiteSpace(b.Source)).ToList();
                foreach (var b in bands)
                {
                    if (string.IsNullOrWhiteSpace(b.Name)) b.Name = b.Source.Split('.').Last();
                    _bandByName[b.Name] = b;
                    _maxTplRow = Math.Max(_maxTplRow, b.LastRow);
                }
                foreach (var b in bands)
                {
                    // parent = smallest band that contains b
                    BandDef parent = null;
                    foreach (var p in bands)
                    {
                        if (p == b) continue;
                        bool contains = p.FirstRow <= b.FirstRow && p.LastRow >= b.LastRow && !(p.FirstRow == b.FirstRow && p.LastRow == b.LastRow);
                        bool overlaps = p.FirstRow <= b.LastRow && b.FirstRow <= p.LastRow;
                        if (overlaps && !contains && !(b.FirstRow <= p.FirstRow && b.LastRow >= p.LastRow))
                            _warn("Bands '" + p.Name + "' and '" + b.Name + "' overlap partially; fix their rows.");
                        if (contains && (parent == null || (p.LastRow - p.FirstRow) < (parent.LastRow - parent.FirstRow))) parent = p;
                    }
                    if (parent == null) _rootBands.Add(b);
                    else
                    {
                        List<BandDef> l;
                        if (!_children.TryGetValue(parent, out l)) _children[parent] = l = new List<BandDef>();
                        l.Add(b);
                    }
                }
                _needRowMap = _sh.Cells.Any(c => c.V != null && c.V.Length > 1 && c.V[0] == '=' && c.V.IndexOf('[') >= 0);
                _first = new int[_maxTplRow + 2];
                _last = new int[_maxTplRow + 2];
            }

            // ---------- emission ----------
            void EmitRange(int from, int to, List<BandDef> bandsHere, DataScope scope)
            {
                int r = from;
                while (r <= to)
                {
                    var b = bandsHere?.FirstOrDefault(x => x.FirstRow == r);
                    if (b != null) { EmitBand(b, scope); r = b.LastRow + 1; continue; }
                    EmitRow(r, scope);
                    r++;
                }
            }

            void EmitBand(BandDef b, DataScope scope)
            {
                var raw = DataAccess.Resolve(scope, b.Source);
                // streaming: never materialize a plain collection (sorting/grouping need it in memory anyway)
                bool lazy = _sink != null && string.IsNullOrEmpty(b.SortBy) && string.IsNullOrEmpty(b.GroupBy) && string.IsNullOrEmpty(b.Filter) && b.Top <= 0 && b.MinRows <= 0;
                IEnumerable<object> seq; int count = 0;
                if (lazy) seq = DataAccess.ItemsLazy(raw);
                else
                {
                    var items = DataAccess.Items(raw).ToList();
                    if (!string.IsNullOrWhiteSpace(b.Filter)) items = items.Where(it => Truthy(_text.Eval("{{" + b.Filter + "}}", new DataScope(it, scope)).Value)).ToList();
                    if (!string.IsNullOrEmpty(b.SortBy)) items = SortItems(items, b.SortBy, b.SortDesc, scope);
                    if (!string.IsNullOrEmpty(b.GroupBy)) items = GroupItems(items, b.GroupBy, scope);
                    if (b.Top > 0)
                    {
                        if (!string.IsNullOrEmpty(b.TopBy)) items = SortItems(items, b.TopBy, !b.TopAscending, scope);
                        if (items.Count > b.Top) items = items.Take(b.Top).ToList();
                    }
                    if (items.Count == 0 && b.KeepWhenEmpty) items.Add(null);
                    for (int pad = items.Count; pad < b.MinRows && b.MinRows <= 2000; pad++) items.Add(new PadItem());
                    count = items.Count; seq = items;
                }
                List<BandDef> kids;
                _children.TryGetValue(b, out kids);

                int runFirst = 0, runLast = 0, n = 0;
                var all = lazy ? null : (IList<object>)seq;
                foreach (var it in seq) { n++; EmitOccurrence(b, kids, scope, it, n, count, ref runFirst, ref runLast, all); }
                if (lazy && n == 0 && b.KeepWhenEmpty) EmitOccurrence(b, kids, scope, null, 1, 1, ref runFirst, ref runLast, null);
                if (runFirst > 0)
                {
                    scope.Runs[b.Name] = new[] { runFirst, runLast };
                    // outer occurrences (and the root) also see the rows of this band: a group total can cover its nested detail rows
                    for (var up = scope.Parent; up != null; up = up.Parent)
                    {
                        int[] cur;
                        if (up.Runs.TryGetValue(b.Name, out cur)) { cur[0] = Math.Min(cur[0], runFirst); cur[1] = Math.Max(cur[1], runLast); }
                        else up.Runs[b.Name] = new[] { runFirst, runLast };
                    }
                }
            }

            void EmitOccurrence(BandDef b, List<BandDef> kids, DataScope scope, object item, int n, int count, ref int runFirst, ref int runLast, IList<object> all)
            {
                var child = new DataScope(item, scope) { Index = n, Count = count, Band = b, Siblings = all };
                if (b.PageBreak && n > 1 && _outRow > 0 && !_out.RowBreaks.Contains(_outRow)) _out.RowBreaks.Add(_outRow);
                int start = _outRow + 1;
                EmitRange(b.FirstRow, b.LastRow, kids, child);
                if (runFirst == 0) runFirst = start;
                runLast = _outRow;
            }

            // zebra fill: innermost band whose occurrence is even wins (decided per cell, so it works while streaming)
            static string AltFill(DataScope scope)
            {
                for (var s = scope; s != null; s = s.Parent)
                    if (s.Band != null && !string.IsNullOrEmpty(s.Band.AlternateFill) && s.Index % 2 == 0) return s.Band.AlternateFill;
                return null;
            }

            static bool Truthy(object v)
            {
                if (v is bool) return (bool)v;
                if (v is string) { var s = ((string)v).Trim().ToLowerInvariant(); return s == "true" || s == "1" || s == "yes"; }
                double d; return DataAccess.TryNumber(v, out d) && d != 0;
            }

            // a path ("Date") or, when it starts with "=", an expression ("=JYM(Date)", "=IF(Qty>10,'big','small')")
            object KeyOf(string path, object it, DataScope scope)
            {
                if (path.Length > 1 && path[0] == '=') return _text.Eval("{{" + path.Substring(1) + "}}", new DataScope(it, scope)).Value;
                return DataAccess.Resolve(new DataScope(it), path);
            }

            List<object> SortItems(List<object> items, string path, bool desc, DataScope scope)
            {
                Func<object, object> key = it => KeyOf(path, it, scope);
                var q = desc ? items.OrderByDescending(key, SortKeyComparer.Instance) : items.OrderBy(key, SortKeyComparer.Instance);
                return q.ToList();
            }

            List<object> GroupItems(List<object> items, string path, DataScope scope)
            {
                var order = new List<GroupItem>();
                var map = new Dictionary<string, GroupItem>(StringComparer.Ordinal);
                foreach (var it in items)
                {
                    var k = KeyOf(path, it, scope);
                    var ks = TextEvaluator.ToDisplay(k);
                    GroupItem g;
                    if (!map.TryGetValue(ks, out g)) { g = new GroupItem { Key = k }; map[ks] = g; order.Add(g); }
                    g.Items.Add(it);
                }
                return order.Cast<object>().ToList();
            }

            static readonly StyleDef NullStyle = new StyleDef();

            // keyed by the base style's reference (never serialize per cell: that is very slow on big reports)
            StyleDef Derive(StyleDef baseStyle, string tag, Action<StyleDef> mutate)
            {
                var key = baseStyle ?? NullStyle;
                Dictionary<string, StyleDef> byTag;
                if (!_derived.TryGetValue(key, out byTag)) _derived[key] = byTag = new Dictionary<string, StyleDef>();
                StyleDef d;
                if (byTag.TryGetValue(tag, out d)) return d;
                d = baseStyle == null ? new StyleDef() : baseStyle.Clone();
                mutate(d);
                byTag[tag] = d;
                return d;
            }

            void EmitRow(int tplRow, DataScope scope)
            {
                _outRow++;
                if (_outRow > _opt.MaxRows) throw new ExcelRenderException("The report has more than " + _opt.MaxRows + " rows (sheet '" + _sh.Name + "'). Filter the data or split it into several sheets.");
                if (_needRowMap) scope.RowOf[tplRow] = _outRow;   // before the cells: [Field] formulas may refer to this row

                var row = new RenderedRow { Index = _outRow };
                List<CellDef> cells;
                if (_cellsByRow.TryGetValue(tplRow, out cells)) row.Cells = new List<RenderedCell>(cells.Count);
                RowDef rd;
                if (_rowDefs.TryGetValue(tplRow, out rd)) { row.Height = rd.Height; row.Hidden = rd.Hidden; }
                if (_sink == null) _out.Rows.Add(row);

                if (cells != null)
                {
                    _cellCount += cells.Count;
                    if (_sink == null && _cellCount > _opt.MaxCells) throw new ExcelRenderException("The report has more than " + _opt.MaxCells + " cells. Filter the data, or raise KscExcelOptions.MaxCells if you really need more (the whole report is built in memory).");
                    foreach (var cd in cells) row.Cells.Add(BuildCell(cd, scope, tplRow));
                    if (row.Cells.Count > 0) _out.MaxCol = Math.Max(_out.MaxCol, row.Cells[row.Cells.Count - 1].Col);
                }

                if (_first[tplRow] == 0) _first[tplRow] = _outRow;
                _last[tplRow] = _outRow;

                List<int[]> merges;
                if (_mergesByStart.TryGetValue(tplRow, out merges))
                    foreach (var m in merges)
                        if (!StraddlesBand(m[0], m[2]))
                            _out.Merges.Add(new[] { _outRow, m[1], _outRow + (m[2] - m[0]), m[3] });

                if (_sink != null) _sink.Row(row);
            }

            bool StraddlesBand(int r1, int r2)
            {
                foreach (var b in _bandByName.Values)
                {
                    bool inside = b.FirstRow <= r1 && b.LastRow >= r2;
                    bool outside = b.LastRow < r1 || b.FirstRow > r2;
                    if (!inside && !outside) { _warn("A merged range spans the edge of band '" + b.Name + "' and was ignored."); return true; }
                }
                return false;
            }

            RenderedCell BuildCell(CellDef cd, DataScope scope, int tplRow)
            {
                StyleDef style = null;
                if (!string.IsNullOrEmpty(cd.S)) _tpl.Styles.TryGetValue(cd.S, out style);
                var alt = AltFill(scope);
                if (alt != null && (style == null || string.IsNullOrEmpty(style.Fill))) style = Derive(style, "alt:" + alt, x => x.Fill = alt);
                var cell = new RenderedCell { Col = cd.C, Style = style, Row = _outRow };
                if (!string.IsNullOrEmpty(cd.L)) cell.Link = Services.ReportLinks.Resolve(_text.EvalText(cd.L, scope), _opt.ReportLinkBase);

                var v = cd.V;
                if (string.IsNullOrEmpty(v)) return cell;
                if (!string.IsNullOrEmpty(cd.Bc)) { AddBarcode(cd, v, scope, tplRow); return cell; }

                if (v[0] == '=' && v.Length > 1)
                {
                    cell.Formula = _text.EvalFormula(v.Substring(1), scope);
                    if (_sink != null) cell.Formula = ResolveFormula(cell, tplRow, scope);
                    else _pendingFormulas.Add(Tuple.Create(cell, tplRow, scope));
                    return cell;
                }

                var r = _text.Eval(v, scope);
                var val = r.Value;
                if (!r.HasPlaceholder && cd.T != "text" && val is string s0) val = TryLiteralNumber(s0);
                SetValue(cell, val);
                return cell;
            }

            // ---------- barcodes ----------
            void AddBarcode(CellDef cd, string expr, DataScope scope, int tplRow)
            {
                if (!Barcodes.IsSupported(cd.Bc.Split(':')[0])) { _warn("Unknown barcode type '" + cd.Bc + "'."); return; }
                string text = _text.EvalText(expr, scope);
                if (string.IsNullOrEmpty(text)) return;
                // cell size in pixels (merged range included)
                int rows = 1, cols = 1;
                List<int[]> ms;
                if (_mergesByStart.TryGetValue(tplRow, out ms))
                    foreach (var m in ms) if (m[1] == cd.C) { rows = m[2] - m[0] + 1; cols = m[3] - m[1] + 1; break; }
                double w = 0, h = 0;
                for (int c = cd.C; c < cd.C + cols; c++)
                {
                    var cdef = _sh.Columns.FirstOrDefault(x => x.Index == c);
                    w += ((cdef != null && cdef.Width.HasValue ? cdef.Width.Value : _sh.DefaultColWidth) * 7 + 5);
                }
                for (int r = tplRow; r < tplRow + rows; r++)
                {
                    RowDef rd;
                    h += ((_rowDefs.TryGetValue(r, out rd) && rd.Height.HasValue ? rd.Height.Value : _sh.DefaultRowHeight) / 0.75);
                }
                int maxW = Math.Max(20, (int)w - 4), maxH = Math.Max(14, (int)h - 4);
                try
                {
                    var mx = Barcodes.Encode(cd.Bc, text);
                    int pw, ph;
                    var png = Barcodes.ToPng(mx, maxW, maxH, out pw, out ph);
                    // 2-D codes stay square; 1-D codes use the full width of the cell
                    int dw = mx.Is2D ? Math.Min(maxW, maxH) : maxW, dh = mx.Is2D ? dw : maxH;
                    _out.Images.Add(new RenderedImage { Name = "barcode", Row = _outRow, Col = cd.C, Width = dw, Height = dh, Bytes = png, Ext = "png", BarType = cd.Bc, BarText = text });
                }
                catch (ArgumentException ex) { _warn("Barcode in " + CellRef.A1(tplRow, cd.C) + ": " + ex.Message); }
            }

            void SetValue(RenderedCell cell, object val)
            {
                double d;
                if (val == null) return;
                if (val is DateTime dt)
                {
                    cell.Value = dt;
                    var fmt = cell.Style?.Format;
                    if (string.IsNullOrEmpty(fmt) || fmt == "General")
                        cell.Style = Derive(cell.Style, "date" + (dt.TimeOfDay == TimeSpan.Zero ? "d" : "t"),
                            x => x.Format = dt.TimeOfDay == TimeSpan.Zero ? "yyyy/mm/dd" : "yyyy/mm/dd hh:mm");
                }
                else if (val is bool) cell.Value = val;
                else if (DataAccess.TryNumber(val, out d)) cell.Value = d;
                else if (val is byte[]) cell.Value = Convert.ToBase64String((byte[])val);
                else cell.Value = Convert.ToString(val, CultureInfo.InvariantCulture);
            }

            static object TryLiteralNumber(string s)
            {
                if (s.Length == 0 || s.Length > 15) return s;
                var t = s.Trim();
                if (t.Length == 0 || t[0] == '+' || (t.Length > 1 && t[0] == '0' && t[1] != '.') || t.IndexOf(',') >= 0) return s;
                double d;
                if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out d) && !double.IsInfinity(d) && !double.IsNaN(d)) return d;
                return s;
            }

            // ---------- formulas ----------
            void ResolveFormulas()
            {
                foreach (var pf in _pendingFormulas) pf.Item1.Formula = ResolveFormula(pf.Item1, pf.Item2, pf.Item3);
            }

            string ResolveFormula(RenderedCell cell, int tplRow, DataScope scope)
            {
                var f = cell.Formula;
                var rowNo = cell.Row;
                if (f.IndexOf("[row:", StringComparison.Ordinal) >= 0)
                    f = RowToken.Replace(f, m => "$" + m.Groups[1].Value + rowNo + ":$" + m.Groups[2].Value + rowNo);
                return FormulaToken.Replace(f, m => ResolveToken(m, cell, tplRow, scope));
            }

            string ResolveToken(Match m, RenderedCell cell, int tplRow, DataScope scope)
            {
                var g1 = m.Groups[1].Value;
                var field = m.Groups[3].Success ? m.Groups[3].Value : null;
                bool byLetter = m.Groups[2].Value == ":";
                if (field != null)
                {
                    // [Band.Field] / [Band:E] -> whole column range of the band
                    BandDef band;
                    if (!_bandByName.TryGetValue(g1, out band)) { _warn("Formula references unknown band '" + g1 + "'."); return "#REF!"; }
                    int col;
                    if (byLetter)
                    {
                        try { col = CellRef.ColIndex(field); } catch (FormatException) { col = 0; }
                        if (col == 0 || col > CellRef.MaxCols) { _warn("Bad column '" + field + "' in [" + g1 + ":" + field + "]."); return "#REF!"; }
                    }
                    else
                    {
                        col = FindFieldColumn(band, field);
                        if (col == 0) { _warn("Band '" + g1 + "' has no cell bound to {{" + field + "}} (used in a formula)."); return "#REF!"; }
                    }
                    int[] run = null;
                    for (var s = scope; s != null && run == null; s = s.Parent) if (s.HasRuns) s.Runs.TryGetValue(band.Name, out run);
                    if (run == null) return "0"; // empty band
                    return "$" + CellRef.ColName(col) + "$" + run[0] + ":$" + CellRef.ColName(col) + "$" + run[1];
                }

                // [Field] -> the same occurrence's cell
                for (var s = scope; s != null; s = s.Parent)
                {
                    if (s.Band == null) continue;
                    int col = FindFieldColumn(s.Band, g1, out int tRow);
                    if (col == 0) continue;
                    int outRow;
                    if (s.RowOf.TryGetValue(tRow, out outRow)) return CellRef.ColName(col) + outRow;
                }
                _warn("Formula token [" + g1 + "] is not a field of the enclosing band.");
                return "#REF!";
            }

            int FindFieldColumn(BandDef band, string field) { int t; return FindFieldColumn(band, field, out t); }

            int FindFieldColumn(BandDef band, string field, out int tplRow)
            {
                var rx = new Regex(@"^\s*\{\{\s*" + Regex.Escape(field) + @"\s*(\|.*)?\}\}\s*$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                for (int r = band.FirstRow; r <= band.LastRow; r++)
                {
                    List<CellDef> cells;
                    if (!_cellsByRow.TryGetValue(r, out cells)) continue;
                    foreach (var c in cells)
                        if (c.V != null && rx.IsMatch(c.V)) { tplRow = r; return c.C; }
                }
                tplRow = 0;
                return 0;
            }

            // ---------- row-mapped decorations ----------
            int MapFirst(int tplRow)
            {
                for (int r = tplRow; r <= _maxTplRow; r++) if (_first[r] != 0) return _first[r];
                return 0;
            }

            int MapLast(int tplRow)
            {
                for (int r = Math.Min(tplRow, _maxTplRow); r >= 1; r--) if (_last[r] != 0) return _last[r];
                return 0;
            }

            bool MapRange(string range, out int r1, out int c1, out int r2, out int c2, out int delta)
            {
                delta = 0; r1 = r2 = 0;
                int a, b;
                if (!CellRef.TryParseRange(range, out a, out c1, out b, out c2)) return false;
                if (a > _maxTplRow) { r1 = a; r2 = b; return true; } // beyond template: keep as is
                r1 = MapFirst(a); r2 = MapLast(Math.Min(b, _maxTplRow));
                if (b > _maxTplRow) r2 = Math.Max(r2, _outRow + (b - _maxTplRow));
                if (r1 == 0 || r2 == 0 || r2 < r1) return false;
                delta = r1 - a;
                return true;
            }

            void MapDecorations(DataScope scope)
            {
                int r1, c1, r2, c2, delta;

                foreach (var cf in _sh.ConditionalFormats)
                {
                    if (string.IsNullOrWhiteSpace(cf.Range) || !MapRange(cf.Range, out r1, out c1, out r2, out c2, out delta)) continue;
                    var copy = new CondFormatDef
                    {
                        Range = CellRef.Range(r1, c1, r2, c2), Type = cf.Type, Operator = cf.Operator, Text = cf.Text, Rank = cf.Rank,
                        Bottom = cf.Bottom, Style = cf.Style, Colors = cf.Colors, BarColor = cf.BarColor,
                        Formula1 = CellRef.ShiftRows(cf.Formula1, delta), Formula2 = CellRef.ShiftRows(cf.Formula2, delta)
                    };
                    _out.ConditionalFormats.Add(copy);
                }
                foreach (var dv in _sh.Validations)
                {
                    if (string.IsNullOrWhiteSpace(dv.Range) || !MapRange(dv.Range, out r1, out c1, out r2, out c2, out delta)) continue;
                    _out.Validations.Add(new ValidationDef
                    {
                        Range = CellRef.Range(r1, c1, r2, c2), Type = dv.Type, Operator = dv.Operator, Formula1 = dv.Formula1, Formula2 = dv.Formula2,
                        AllowBlank = dv.AllowBlank, ErrorTitle = dv.ErrorTitle, Error = dv.Error, PromptTitle = dv.PromptTitle, Prompt = dv.Prompt
                    });
                }
                if (!string.IsNullOrWhiteSpace(_sh.AutoFilter) && MapRange(_sh.AutoFilter, out r1, out c1, out r2, out c2, out delta))
                    _out.AutoFilter = CellRef.Range(r1, c1, r2, c2);

                _out.FreezeCols = _sh.FreezeCols;
                if (_sh.FreezeRows > 0) _out.FreezeRows = Math.Min(MapLast(_sh.FreezeRows), Math.Max(_outRow, 0));

                foreach (var img in _sh.Images)
                {
                    byte[] bytes = null;
                    if (!string.IsNullOrWhiteSpace(img.Binding))
                    {
                        var bv = DataAccess.Resolve(scope, img.Binding);
                        if (bv is byte[] bb) bytes = bb;
                        else if (bv is string bs && bs.Length > 0) bytes = TryBase64(bs);
                    }
                    if (bytes == null && !string.IsNullOrEmpty(img.Data)) bytes = TryBase64(img.Data);
                    if (bytes == null) continue;
                    int row = img.Row <= _maxTplRow ? MapFirst(img.Row) : img.Row;
                    if (row == 0) continue;
                    _out.Images.Add(new RenderedImage { Name = img.Name, Row = row, Col = Math.Max(1, img.Col), Width = img.Width, Height = img.Height, Bytes = bytes, Ext = NormalizeExt(img.Ext, bytes) });
                }

                foreach (var ch in _sh.Charts) MapChart(ch);

                var pg = _sh.Page ?? new PageSetupDef();
                if (!string.IsNullOrWhiteSpace(pg.PrintTitleRows))
                {
                    var m = Regex.Match(pg.PrintTitleRows, @"^\s*\$?(\d+)\s*:\s*\$?(\d+)\s*$");
                    if (m.Success)
                    {
                        int a = MapFirst(int.Parse(m.Groups[1].Value)), b = MapLast(int.Parse(m.Groups[2].Value));
                        if (a > 0 && b >= a) _out.PrintTitleRows = a + ":" + b;
                    }
                }
                if (string.IsNullOrWhiteSpace(pg.PrintTitleRows))
                {
                    var ph = _sh.Sections == null ? null : _sh.Sections.FirstOrDefault(x => string.Equals(x.Kind, "pageHeader", StringComparison.OrdinalIgnoreCase) && x.FirstRow >= 1 && x.LastRow >= x.FirstRow);
                    if (ph != null) { int a = MapFirst(ph.FirstRow), b = MapLast(ph.LastRow); if (a > 0 && b >= a) _out.PrintTitleRows = a + ":" + b; }
                }
                if (!string.IsNullOrWhiteSpace(pg.PrintArea) && MapRange(pg.PrintArea, out r1, out c1, out r2, out c2, out delta))
                    _out.PrintArea = CellRef.Range(r1, c1, Math.Max(r2, _outRow), c2);
            }

            string AbsRef(int r1, int c1, int r2, int c2)
                => "'" + _out.Name.Replace("'", "''") + "'!$" + CellRef.ColName(c1) + "$" + r1 + ":$" + CellRef.ColName(c2) + "$" + r2;

            void MapChart(ChartDef ch)
            {
                int r1, c1, r2, c2, delta;
                var rc = new RenderedChart { Def = ch, Col = Math.Max(1, ch.Col), Row = ch.Row <= _maxTplRow ? MapFirst(ch.Row) : ch.Row };
                if (rc.Row == 0) rc.Row = 1;
                if (!string.IsNullOrWhiteSpace(ch.Categories))
                {
                    if (!MapRange(ch.Categories, out r1, out c1, out r2, out c2, out delta)) { _warn("Chart '" + ch.Name + "': categories range '" + ch.Categories + "' is invalid or empty."); return; }
                    rc.Categories = AbsRef(r1, c1, r2, c2);
                }
                foreach (var s in ch.Series)
                {
                    if (string.IsNullOrWhiteSpace(s.Values) || !MapRange(s.Values, out r1, out c1, out r2, out c2, out delta)) { _warn("Chart '" + ch.Name + "': a series has an invalid or empty range."); continue; }
                    rc.SeriesValues.Add(AbsRef(r1, c1, r2, c2));
                    rc.SeriesColors.Add(s.Color); rc.SeriesTypes.Add(s.Type); rc.SeriesSecondary.Add(s.Secondary);
                    int nr, nc, nr2, nc2;
                    if (!string.IsNullOrWhiteSpace(s.NameCell) && MapRange(s.NameCell, out nr, out nc, out nr2, out nc2, out delta))
                    { rc.SeriesNames.Add(AbsRef(nr, nc, nr, nc)); rc.SeriesNameIsRef.Add(true); }
                    else { rc.SeriesNames.Add(s.Name ?? ""); rc.SeriesNameIsRef.Add(false); }
                }
                if (rc.SeriesValues.Count == 0) { _warn("Chart '" + ch.Name + "' has no valid series."); return; }
                _out.Charts.Add(rc);
            }

            static byte[] TryBase64(string s)
            {
                try
                {
                    int comma = s.IndexOf("base64,", StringComparison.OrdinalIgnoreCase);
                    if (comma >= 0) s = s.Substring(comma + 7);
                    return Convert.FromBase64String(s.Trim());
                }
                catch (FormatException) { return null; }
            }

            static string NormalizeExt(string ext, byte[] b)
            {
                if (b.Length > 4 && b[0] == 0x89 && b[1] == 0x50) return "png";
                if (b.Length > 3 && b[0] == 0xFF && b[1] == 0xD8) return "jpeg";
                if (b.Length > 3 && b[0] == 0x47 && b[1] == 0x49) return "gif";
                ext = (ext ?? "png").ToLowerInvariant().TrimStart('.');
                return ext == "jpg" ? "jpeg" : ext;
            }

            // ---------- names ----------
            static string Sanitize(string name)
            {
                var sb = new StringBuilder();
                foreach (var ch in name ?? "")
                    sb.Append("[]:*?/\\".IndexOf(ch) >= 0 || char.IsControl(ch) ? '_' : ch);
                var s = sb.ToString().Trim().Trim('\'');
                if (s.Length == 0) s = "Sheet";
                return s.Length > 31 ? s.Substring(0, 31) : s;
            }

            string UniqueName(string name)
            {
                var n = name; int i = 1;
                while (!_usedNames.Add(n))
                {
                    var suffix = " (" + (++i) + ")";
                    n = (name.Length + suffix.Length > 31 ? name.Substring(0, 31 - suffix.Length) : name) + suffix;
                }
                return n;
            }
        }
    }
}
