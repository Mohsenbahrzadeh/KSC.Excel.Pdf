using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using KSC.Excel.Model;
using Newtonsoft.Json;

namespace KSC.Excel.Engine
{
    /// <summary>
    /// Horizontal repeat ("column bands", used for cross-tabs): before rows are rendered the sheet template is expanded so each
    /// column band becomes N copies of its columns. Placeholders in the copies are rewritten to index the collection
    /// ({{Value}} in copy 2 of band "Cells" becomes {{Cells[1].Value}}), merges/ranges/widths are remapped.
    /// </summary>
    internal static class ColumnExpander
    {
        static readonly Regex Ph = new Regex(@"\{\{(.*?)\}\}", RegexOptions.Compiled | RegexOptions.Singleline);
        static readonly Regex Tok = new Regex(@"\[([A-Za-z_][\w]*)(?:([.:])([A-Za-z_][\w\.]*))?\]", RegexOptions.Compiled);
        static readonly Regex FuncHead = new Regex(@"^(count|sum|avg|average|min|max)\s*\(", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        sealed class Plan
        {
            public BandDef Def; public int First, Last, W, N, AddBefore;
            public bool Contains(int c) { return c >= First && c <= Last; }
        }

        public static bool IsColumnBand(BandDef b) => string.Equals(b.Orientation, "columns", StringComparison.OrdinalIgnoreCase);

        public static SheetTemplate Expand(SheetTemplate sh, DataScope root, Action<string> warn)
        {
            var defs = sh.Bands.Where(b => IsColumnBand(b) && b.FirstCol >= 1 && b.LastCol >= b.FirstCol && !string.IsNullOrWhiteSpace(b.Source)).OrderBy(b => b.FirstCol).ToList();
            if (defs.Count == 0) return sh;

            var plans = new List<Plan>(); int add = 0;
            foreach (var d in defs)
            {
                if (plans.Count > 0 && d.FirstCol <= plans[plans.Count - 1].Last) { warn("Column bands '" + plans[plans.Count - 1].Def.Name + "' and '" + d.Name + "' overlap; the second was ignored."); continue; }
                if (string.IsNullOrWhiteSpace(d.Name)) d.Name = d.Source.Split('.').Last();
                var p = new Plan { Def = d, First = d.FirstCol, Last = d.LastCol, W = d.LastCol - d.FirstCol + 1, AddBefore = add };
                p.N = CountOf(sh, d, root);
                if (p.N == 0 && d.KeepWhenEmpty) p.N = 1;
                add += (p.N - 1) * p.W;
                plans.Add(p);
            }

            Func<int, bool, int> map = (c, asEnd) =>
            {
                int a = 0;
                foreach (var p in plans)
                {
                    if (c > p.Last) { a += (p.N - 1) * p.W; continue; }
                    if (c >= p.First)
                    {
                        int rep = asEnd ? Math.Max(p.N - 1, 0) : 0;
                        return p.First + a + rep * p.W + (c - p.First);
                    }
                    break;
                }
                return c + a;
            };

            var o = new SheetTemplate
            {
                Name = sh.Name, RepeatFor = sh.RepeatFor, Rtl = sh.Rtl, ShowGridLines = sh.ShowGridLines, TabColor = sh.TabColor, Zoom = sh.Zoom,
                DefaultColWidth = sh.DefaultColWidth, DefaultRowHeight = sh.DefaultRowHeight, Rows = sh.Rows, Page = ClonePage(sh.Page, map),
                Protect = sh.Protect, FreezeRows = sh.FreezeRows, Sections = sh.Sections
            };
            o.FreezeCols = sh.FreezeCols > 0 ? map(sh.FreezeCols, true) : 0;
            o.AutoFilter = RangeMap(sh.AutoFilter, map);
            o.Bands = sh.Bands.Where(b => !IsColumnBand(b)).ToList();

            // columns
            foreach (var c in sh.Columns)
            {
                var p = plans.FirstOrDefault(x => x.Contains(c.Index));
                if (p == null) { o.Columns.Add(new ColumnDef { Index = map(c.Index, false), Width = c.Width, Hidden = c.Hidden }); continue; }
                for (int i = 0; i < p.N; i++) o.Columns.Add(new ColumnDef { Index = p.First + p.AddBefore + i * p.W + (c.Index - p.First), Width = c.Width, Hidden = c.Hidden });
            }
            // band columns without an explicit definition keep the default width (nothing to add)

            // cells
            foreach (var cd in sh.Cells)
            {
                var p = plans.FirstOrDefault(x => x.Contains(cd.C));
                if (p == null) { o.Cells.Add(Copy(cd, map(cd.C, false), RewriteOutside(cd.V, plans, sh, map), cd.L)); continue; }
                for (int i = 0; i < p.N; i++)
                    o.Cells.Add(Copy(cd, p.First + p.AddBefore + i * p.W + (cd.C - p.First), RewriteInside(cd.V, p, i, plans, sh, map), cd.L == null ? null : RewritePh(cd.L, p, i)));
            }

            // merges
            foreach (var m in sh.Merges)
            {
                int r1, c1, r2, c2;
                if (!CellRef.TryParseRange(m, out r1, out c1, out r2, out c2)) continue;
                var p = plans.FirstOrDefault(x => x.Contains(c1) || x.Contains(c2));
                if (p == null)
                {
                    if (plans.Any(x => c1 < x.First && c2 > x.Last)) { warn("A merged range spans a column band and was ignored."); continue; }
                    o.Merges.Add(CellRef.Range(r1, map(c1, false), r2, map(c2, true)));
                }
                else if (p.Contains(c1) && p.Contains(c2))
                    for (int i = 0; i < p.N; i++) o.Merges.Add(CellRef.Range(r1, p.First + p.AddBefore + i * p.W + (c1 - p.First), r2, p.First + p.AddBefore + i * p.W + (c2 - p.First)));
                else warn("A merged range crosses the edge of column band '" + p.Def.Name + "' and was ignored.");
            }

            foreach (var cf in sh.ConditionalFormats) { var x = Clone(cf); x.Range = RangeMap(x.Range, map); o.ConditionalFormats.Add(x); }
            foreach (var v in sh.Validations) { var x = Clone(v); x.Range = RangeMap(x.Range, map); o.Validations.Add(x); }
            foreach (var im in sh.Images) { var x = Clone(im); x.Col = map(im.Col, false); o.Images.Add(x); }
            foreach (var ch in sh.Charts)
            {
                var x = Clone(ch); x.Col = map(ch.Col, false);
                x.Categories = RangeMap(ch.Categories, map);
                foreach (var s in x.Series) { s.Values = RangeMap(s.Values, map); s.NameCell = RangeMap(s.NameCell, map); }
                o.Charts.Add(x);
            }
            return o;
        }

        // ---------------------------------------------------------------- helpers
        static int CountOf(SheetTemplate sh, BandDef d, DataScope root)
        {
            var direct = DataAccess.Resolve(root, d.Source);
            if (direct != null) return DataAccess.Items(direct).Count();
            // relative source: look inside the first item of any row band's collection
            foreach (var rb in sh.Bands.Where(b => !IsColumnBand(b) && !string.IsNullOrWhiteSpace(b.Source)))
            {
                var first = DataAccess.Items(DataAccess.Resolve(root, rb.Source)).FirstOrDefault();
                if (first == null) continue;
                var v = DataAccess.Resolve(new DataScope(first, root), d.Source);
                if (v != null) return DataAccess.Items(v).Count();
            }
            return 0;
        }

        static CellDef Copy(CellDef c, int col, string v, string l) => new CellDef { R = c.R, C = col, V = v, S = c.S, T = c.T, L = l ?? c.L };

        static string RangeMap(string range, Func<int, bool, int> map)
        {
            if (string.IsNullOrWhiteSpace(range)) return range;
            int r1, c1, r2, c2;
            if (!CellRef.TryParseRange(range, out r1, out c1, out r2, out c2)) return range;
            return CellRef.Range(r1, map(c1, false), r2, map(c2, true));
        }

        static PageSetupDef ClonePage(PageSetupDef p, Func<int, bool, int> map)
        {
            if (p == null) return new PageSetupDef();
            var c = JsonConvert.DeserializeObject<PageSetupDef>(JsonConvert.SerializeObject(p));
            c.PrintArea = RangeMap(c.PrintArea, map);
            return c;
        }

        static T Clone<T>(T x) => JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(x));

        static string Letter(int c) => CellRef.ColName(c);

        // placeholders inside a replica: {{Field}} -> {{Source[i].Field}}
        static string RewritePh(string text, Plan p, int i)
        {
            if (text == null) return null;
            return Ph.Replace(text, m =>
            {
                var body = m.Groups[1].Value; int bar = body.IndexOf('|');
                var head = (bar < 0 ? body : body.Substring(0, bar)).Trim(); var tail = bar < 0 ? "" : body.Substring(bar);
                string nh;
                if (head.IndexOf("[$col]", StringComparison.Ordinal) >= 0) nh = head.Replace("[$col]", "[" + i + "]");   // explicit path, e.g. Cells[$col].Value
                else if (head == "$item" || head.StartsWith("$item.", StringComparison.Ordinal)) nh = p.Def.Source + "[" + i + "]" + head.Substring(5);
                else if (head == "$col") nh = i.ToString();
                else if (head == "$index") nh = (i + 1).ToString();
                else if (head == "$count") nh = p.N.ToString();
                else if (head.StartsWith("$", StringComparison.Ordinal) || head.StartsWith("'", StringComparison.Ordinal) || FuncHead.IsMatch(head) || IsNumber(head)) return m.Value;
                else nh = p.Def.Source + "[" + i + "]." + head;
                return "{{" + nh + tail + "}}";
            });
        }

        static bool IsNumber(string s) { double d; return double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out d); }

        static string RewriteInside(string v, Plan p, int i, List<Plan> plans, SheetTemplate sh, Func<int, bool, int> map)
        {
            if (string.IsNullOrEmpty(v)) return v;
            v = RewritePh(v, p, i);
            return v[0] == '=' ? Tokens(v, plans, sh, map, p, i) : v;
        }

        static string RewriteOutside(string v, List<Plan> plans, SheetTemplate sh, Func<int, bool, int> map)
        {
            if (string.IsNullOrEmpty(v) || v[0] != '=') return v;
            return Tokens(v, plans, sh, map, null, 0);
        }

        static int FieldColumn(SheetTemplate sh, BandDef rowBand, string field, int fromCol, int toCol)
        {
            // the field may be the last segment of a longer path, e.g. {{Cells[$col].Value}} matches "Value"
            var rx = new Regex(@"^\s*\{\{\s*([\w\.\[\]\$]*\.)?" + Regex.Escape(field) + @"\s*(\|.*)?\}\}\s*$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (var c in sh.Cells.OrderBy(x => x.C))
            {
                if (c.V == null || c.C < fromCol || c.C > toCol) continue;
                if (rowBand != null && (c.R < rowBand.FirstRow || c.R > rowBand.LastRow)) continue;
                if (rx.IsMatch(c.V)) return c.C;
            }
            return 0;
        }

        static string Tokens(string f, List<Plan> plans, SheetTemplate sh, Func<int, bool, int> map, Plan inBand, int i)
        {
            return Tok.Replace(f, m =>
            {
                string g1 = m.Groups[1].Value, sep = m.Groups[2].Value, g3 = m.Groups[3].Success ? m.Groups[3].Value : null;
                if (g3 == null) return m.Value;

                var cb = plans.FirstOrDefault(x => string.Equals(x.Def.Name, g1, StringComparison.OrdinalIgnoreCase));
                if (cb != null)
                {
                    // [ColBand.Field] -> horizontal range across the copies, in the formula's own row
                    int t = 0;
                    if (sep == ".") t = FieldColumn(sh, null, g3, cb.First, cb.Last);
                    else { try { t = CellRef.ColIndex(g3); } catch (FormatException) { t = 0; } if (!cb.Contains(t)) t = 0; }
                    if (t == 0) return m.Value;
                    if (cb.N == 0) return "0";
                    int off = t - cb.First;
                    return "[row:" + Letter(cb.First + cb.AddBefore + off) + ":" + Letter(cb.First + cb.AddBefore + (cb.N - 1) * cb.W + off) + "]";
                }

                if (sep == ".")
                {
                    var rb = sh.Bands.FirstOrDefault(b => !IsColumnBand(b) && string.Equals(b.Name ?? b.Source.Split('.').Last(), g1, StringComparison.OrdinalIgnoreCase));
                    if (rb == null) return m.Value;
                    foreach (var p in plans)
                    {
                        int t = FieldColumn(sh, rb, g3, p.First, p.Last);
                        if (t == 0) continue;
                        int off = t - p.First;
                        int copy = inBand == p ? i : 0;
                        return "[" + g1 + ":" + Letter(p.First + p.AddBefore + copy * p.W + off) + "]";
                    }
                    return m.Value;
                }

                // [Band:E] with a template column letter
                int col;
                try { col = CellRef.ColIndex(g3); } catch (FormatException) { return m.Value; }
                var bp = plans.FirstOrDefault(x => x.Contains(col));
                int newCol = bp != null && inBand == bp ? bp.First + bp.AddBefore + i * bp.W + (col - bp.First) : map(col, false);
                return "[" + g1 + ":" + Letter(newCol) + "]";
            });
        }
    }
}
