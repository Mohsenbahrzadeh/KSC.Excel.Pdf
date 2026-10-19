using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using KSC.Excel.Model;

namespace KSC.Excel.Engine
{
    /// <summary>What conditional formatting does to one cell.</summary>
    public sealed class CfResult
    {
        /// <summary>Format of a matching cellIs / expression / containsText / duplicates / top10 rule (fill, color, bold, italic...).</summary>
        public StyleDef Style;
        /// <summary>Background from a color scale (#RRGGBB).</summary>
        public string ScaleFill;
        /// <summary>Data bar: 0..1 of the cell width, and its color.</summary>
        public double? BarFraction;
        public string BarColor;
    }

    /// <summary>
    /// Evaluates a sheet's conditional formats for rendered cells. Shared by the HTML preview and the PDF renderer so both agree
    /// (cellIs, expression, containsText, duplicates, top10, colorScale, dataBar).
    /// </summary>
    public sealed class CondFormatEvaluator
    {
        sealed class Rule
        {
            public CondFormatDef Def; public string Type; public int R1, C1, R2, C2;
            public double Min = double.MaxValue, Max = double.MinValue; public bool HasRange;
            public double? Threshold;                       // top10
            public HashSet<string> Duplicates;              // duplicates
        }

        readonly RenderedSheet _sheet;
        readonly FormulaEvaluator _eval;
        readonly List<Rule> _rules = new List<Rule>();
        readonly Dictionary<string, object> _operands = new Dictionary<string, object>();

        public bool HasRules { get { return _rules.Count > 0; } }

        public CondFormatEvaluator(RenderedSheet sheet, FormulaEvaluator eval = null)
        {
            _sheet = sheet; _eval = eval ?? new FormulaEvaluator(sheet);
            foreach (var cf in sheet.ConditionalFormats)
            {
                int r1, c1, r2, c2;
                if (cf == null || !CellRef.TryParseRange(cf.Range, out r1, out c1, out r2, out c2)) continue;
                var rule = new Rule { Def = cf, Type = (cf.Type ?? "").ToLowerInvariant(), R1 = r1, C1 = c1, R2 = r2, C2 = c2 };
                Prepare(rule);
                _rules.Add(rule);
            }
        }

        object Value(RenderedCell c)
        {
            if (c == null) return null;
            if (c.Formula == null) return c.Value;
            if (c.Cached != null) return c.Cached;
            object v; return _eval.TryEvaluate(c, out v) ? v : null;
        }

        IEnumerable<RenderedCell> CellsIn(Rule r)
        {
            for (int row = r.R1; row <= Math.Min(r.R2, _sheet.Rows.Count); row++)
                foreach (var c in _sheet.Rows[row - 1].Cells)
                    if (c.Col >= r.C1 && c.Col <= r.C2) yield return c;
        }

        void Prepare(Rule rule)
        {
            switch (rule.Type)
            {
                case "colorscale":
                case "databar":
                    foreach (var c in CellsIn(rule))
                    {
                        double d;
                        if (DataAccess.TryNumber(Value(c), out d)) { rule.Min = Math.Min(rule.Min, d); rule.Max = Math.Max(rule.Max, d); rule.HasRange = true; }
                    }
                    break;
                case "top10":
                {
                    var nums = new List<double>();
                    foreach (var c in CellsIn(rule)) { double d; if (DataAccess.TryNumber(Value(c), out d)) nums.Add(d); }
                    if (nums.Count > 0)
                    {
                        nums.Sort();
                        int n = Math.Max(1, Math.Min(rule.Def.Rank, nums.Count));
                        rule.Threshold = rule.Def.Bottom ? nums[n - 1] : nums[nums.Count - n];
                    }
                    break;
                }
                case "duplicates":
                {
                    var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    foreach (var c in CellsIn(rule))
                    {
                        var v = Value(c); if (v == null) continue;
                        var k = Convert.ToString(v, CultureInfo.InvariantCulture); if (k.Length == 0) continue;
                        int n; seen.TryGetValue(k, out n); seen[k] = n + 1;
                    }
                    rule.Duplicates = new HashSet<string>(seen.Where(kv => kv.Value > 1).Select(kv => kv.Key), StringComparer.OrdinalIgnoreCase);
                    break;
                }
            }
        }

        /// <summary>Returns the effect on the cell at (row, col), or null when no rule applies. <paramref name="text"/> is the displayed text.</summary>
        public CfResult Evaluate(int row, int col, object value, string text)
        {
            CfResult res = null;
            foreach (var rule in _rules)
            {
                if (row < rule.R1 || row > rule.R2 || col < rule.C1 || col > rule.C2) continue;
                double d; bool num = DataAccess.TryNumber(value, out d);
                var cf = rule.Def;
                switch (rule.Type)
                {
                    case "cellis":
                        if (CellIs(cf, value, d, num)) { res = res ?? new CfResult(); res.Style = Merge(res.Style, cf.Style); }
                        break;
                    case "expression":
                        if (Truthy(Eval(cf.Formula1, row - rule.R1, col - rule.C1))) { res = res ?? new CfResult(); res.Style = Merge(res.Style, cf.Style); }
                        break;
                    case "containstext":
                        if (!string.IsNullOrEmpty(cf.Text) && (text ?? "").IndexOf(cf.Text, StringComparison.OrdinalIgnoreCase) >= 0) { res = res ?? new CfResult(); res.Style = Merge(res.Style, cf.Style); }
                        break;
                    case "duplicates":
                        if (value != null && rule.Duplicates != null && rule.Duplicates.Contains(Convert.ToString(value, CultureInfo.InvariantCulture))) { res = res ?? new CfResult(); res.Style = Merge(res.Style, cf.Style); }
                        break;
                    case "top10":
                        if (num && rule.Threshold.HasValue && (cf.Bottom ? d <= rule.Threshold.Value : d >= rule.Threshold.Value)) { res = res ?? new CfResult(); res.Style = Merge(res.Style, cf.Style); }
                        break;
                    case "colorscale":
                        if (num && rule.HasRange)
                        {
                            var cols = (cf.Colors ?? new List<string>()).Where(IsHex).ToList();
                            if (cols.Count < 2) cols = new List<string> { "#F8696B", "#FFEB84", "#63BE7B" };
                            double f = rule.Max > rule.Min ? (d - rule.Min) / (rule.Max - rule.Min) : 0;
                            res = res ?? new CfResult(); res.ScaleFill = Lerp(cols, f);
                        }
                        break;
                    case "databar":
                        if (num && rule.HasRange)
                        {
                            res = res ?? new CfResult();
                            res.BarFraction = rule.Max > rule.Min ? (d - rule.Min) / (rule.Max - rule.Min) : 1;
                            res.BarColor = IsHex(cf.BarColor) ? cf.BarColor : "#638EC6";
                        }
                        break;
                }
            }
            return res;
        }

        bool CellIs(CondFormatDef cf, object value, double d, bool num)
        {
            var op = (cf.Operator ?? "equal").ToLowerInvariant();
            var a = Operand(cf.Formula1); var b = Operand(cf.Formula2);
            double da, db;
            if (num && DataAccess.TryNumber(a, out da))
            {
                DataAccess.TryNumber(b, out db);
                switch (op)
                {
                    case "greaterthan": return d > da;
                    case "lessthan": return d < da;
                    case "greaterthanorequal": return d >= da;
                    case "lessthanorequal": return d <= da;
                    case "notequal": return d != da;
                    case "between": return d >= da && d <= db;
                    case "notbetween": return d < da || d > db;
                    default: return d == da;
                }
            }
            // text comparison
            var sv = value == null ? "" : Convert.ToString(value, CultureInfo.InvariantCulture);
            var sa = a == null ? "" : Convert.ToString(a, CultureInfo.InvariantCulture);
            int cmp = string.Compare(sv, sa, StringComparison.OrdinalIgnoreCase);
            switch (op)
            {
                case "notequal": return cmp != 0;
                case "equal": return cmp == 0;
                case "greaterthan": return cmp > 0;
                case "lessthan": return cmp < 0;
                case "greaterthanorequal": return cmp >= 0;
                case "lessthanorequal": return cmp <= 0;
                default: return false;
            }
        }

        // a literal ("text", 5) or a formula / reference
        object Operand(string f)
        {
            if (string.IsNullOrWhiteSpace(f)) return null;
            object cached;
            if (_operands.TryGetValue(f, out cached)) return cached;
            var t = f.Trim().TrimStart('=').Trim();
            object v; double d;
            if (t.Length >= 2 && t[0] == '"' && t[t.Length - 1] == '"') v = t.Substring(1, t.Length - 2).Replace("\"\"", "\"");
            else if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) v = d;
            else v = Eval(t, 0, 0);
            _operands[f] = v;
            return v;
        }

        object Eval(string formula, int dRow, int dCol)
        {
            if (string.IsNullOrWhiteSpace(formula)) return null;
            var f = CellRef.Shift(formula.Trim().TrimStart('='), dRow, dCol);
            object v;
            return _eval.TryEvaluate(new RenderedCell { Formula = f }, out v) ? v : null;
        }

        static bool Truthy(object v)
        {
            if (v is bool) return (bool)v;
            double d; return !(v is FormulaEvaluator.XlError) && DataAccess.TryNumber(v, out d) && d != 0;
        }

        static StyleDef Merge(StyleDef a, StyleDef b)
        {
            if (b == null) return a;
            if (a == null) return b;
            var m = a.Clone();
            if (!string.IsNullOrEmpty(b.Fill)) m.Fill = b.Fill;
            if (!string.IsNullOrEmpty(b.Color)) m.Color = b.Color;
            m.Bold |= b.Bold; m.Italic |= b.Italic; m.Underline |= b.Underline; m.Strike |= b.Strike;
            return m;
        }

        static bool IsHex(string h)
        {
            if (string.IsNullOrEmpty(h) || h.Length != 7 || h[0] != '#') return false;
            for (int i = 1; i < 7; i++) if (!Uri.IsHexDigit(h[i])) return false;
            return true;
        }

        /// <summary>Linear interpolation across 2-3 colors; returns #RRGGBB.</summary>
        public static string Lerp(List<string> cols, double f)
        {
            f = Math.Max(0, Math.Min(1, f));
            double pos = f * (cols.Count - 1);
            int i = Math.Min((int)pos, cols.Count - 2);
            double t = pos - i;
            Func<string, int[]> p = h => new[] { Convert.ToInt32(h.Substring(1, 2), 16), Convert.ToInt32(h.Substring(3, 2), 16), Convert.ToInt32(h.Substring(5, 2), 16) };
            var a = p(cols[i]); var b = p(cols[i + 1]);
            return "#" + ((int)(a[0] + (b[0] - a[0]) * t)).ToString("X2") + ((int)(a[1] + (b[1] - a[1]) * t)).ToString("X2") + ((int)(a[2] + (b[2] - a[2]) * t)).ToString("X2");
        }
    }
}
