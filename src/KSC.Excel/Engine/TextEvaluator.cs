using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace KSC.Excel.Engine
{
    /// <summary>Result of evaluating a cell text: a typed value when the cell is a single placeholder, text otherwise.</summary>
    public struct EvalResult
    {
        public object Value;
        public bool HasPlaceholder;
    }

    /// <summary>Evaluates "{{path | filter:arg}}" placeholders against a <see cref="DataScope"/>.</summary>
    public sealed class TextEvaluator
    {
        static readonly Regex Placeholder = new Regex(@"\{\{(.*?)\}\}", RegexOptions.Compiled | RegexOptions.Singleline);
        static readonly Regex FuncRx = new Regex(@"^(count|sum|avg|average|min|max)\s*\(\s*([^)]*?)\s*\)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static readonly PersianCalendar Pc = new PersianCalendar();

        readonly Action<string> _warn;
        readonly IDictionary<string, Func<object, string, object>> _filters;
        readonly ExprEnv _env;

        public TextEvaluator(Action<string> warn = null, IDictionary<string, Func<object, string, object>> filters = null, IDictionary<string, Func<object[], object>> functions = null)
        {
            _warn = warn ?? (_ => { });
            _filters = filters;
            _env = new ExprEnv { Aggregate = (fn, path, sc) => Aggregate(fn, path, sc), Functions = functions };
        }

        public static bool HasPlaceholders(string text) => text != null && text.IndexOf("{{", StringComparison.Ordinal) >= 0;

        sealed class Ph
        {
            public string Head, Fn, FnArg;
            public object Literal;
            public Expr Expr;
            public List<string[]> Filters = new List<string[]>();   // {name, arg}
        }

        sealed class Compiled
        {
            public List<object> Parts = new List<object>();          // string literal | Ph
            public bool Single;
        }

        readonly Dictionary<string, Compiled> _cache = new Dictionary<string, Compiled>(StringComparer.Ordinal);

        Compiled Compile(string text)
        {
            Compiled c;
            if (_cache.TryGetValue(text, out c)) return c;
            c = new Compiled();
            var matches = Placeholder.Matches(text);
            int last = 0;
            foreach (Match m in matches)
            {
                if (m.Index > last) c.Parts.Add(text.Substring(last, m.Index - last));
                c.Parts.Add(ParsePlaceholder(m.Groups[1].Value));
                last = m.Index + m.Length;
            }
            if (matches.Count > 0 && last < text.Length) c.Parts.Add(text.Substring(last));
            c.Single = matches.Count == 1 && matches[0].Index == 0 && matches[0].Length == text.Length;
            if (matches.Count == 0) c.Parts.Clear();
            _cache[text] = c;
            return c;
        }

        Ph ParsePlaceholder(string body)
        {
            var parts = SplitPipes(body);
            var ph = new Ph { Head = parts[0].Trim() };
            double lit;
            if (double.TryParse(ph.Head, NumberStyles.Float, CultureInfo.InvariantCulture, out lit)) ph.Literal = lit;
            else if (ph.Head.Length >= 2 && ph.Head[0] == '\'' && ph.Head[ph.Head.Length - 1] == '\'') ph.Literal = ph.Head.Substring(1, ph.Head.Length - 2);
            var fm = FuncRx.Match(ph.Head);
            if (!fm.Success && ph.Literal == null && Expr.LooksLikeExpression(ph.Head))
            {
                try { ph.Expr = Expr.Parse(ph.Head); }
                catch (FormatException ex) { _warn(ex.Message); }
            }
            if (fm.Success) { ph.Fn = fm.Groups[1].Value.ToLowerInvariant(); ph.FnArg = fm.Groups[2].Value.Trim(); }
            for (int i = 1; i < parts.Count; i++)
            {
                var f = parts[i].Trim(); string name = f, arg = null;
                int c = f.IndexOf(':');
                if (c >= 0) { name = f.Substring(0, c).Trim(); arg = f.Substring(c + 1); }
                ph.Filters.Add(new[] { name, arg });
            }
            return ph;
        }

        public EvalResult Eval(string text, DataScope scope)
        {
            if (string.IsNullOrEmpty(text)) return new EvalResult();
            var c = Compile(text);
            if (c.Parts.Count == 0) return new EvalResult { Value = text };

            // single placeholder filling the whole text -> keep the type (number/date/bool)
            if (c.Single) return new EvalResult { Value = EvalPlaceholder((Ph)c.Parts[0], scope), HasPlaceholder = true };

            var sb = new StringBuilder();
            foreach (var part in c.Parts)
            {
                var lit = part as string;
                if (lit != null) sb.Append(lit);
                else sb.Append(ToDisplay(EvalPlaceholder((Ph)part, scope)));
            }
            return new EvalResult { Value = sb.ToString(), HasPlaceholder = true };
        }

        /// <summary>Text-only evaluation (sheet names, hyperlinks...).</summary>
        public string EvalText(string text, DataScope scope)
        {
            var r = Eval(text, scope);
            return ToDisplay(r.Value);
        }

        /// <summary>Replaces placeholders inside a formula with literals (numbers raw, text quoted).</summary>
        public string EvalFormula(string formula, DataScope scope)
        {
            if (!HasPlaceholders(formula)) return formula;
            return Placeholder.Replace(formula, m =>
            {
                var v = EvalPlaceholder(ParsePlaceholder(m.Groups[1].Value), scope);
                double d;
                if (v == null) return "\"\"";
                if (v is bool b) return b ? "TRUE" : "FALSE";
                if (v is DateTime dt) return dt.ToOADate().ToString("R", CultureInfo.InvariantCulture);
                if (DataAccess.TryNumber(v, out d)) return d.ToString("R", CultureInfo.InvariantCulture);
                return "\"" + Convert.ToString(v, CultureInfo.InvariantCulture).Replace("\"", "\"\"") + "\"";
            });
        }

        object EvalExpr(Expr e, DataScope scope)
        {
            try { return e.Eval(scope, _env); }
            catch (InvalidOperationException ex) { _warn(ex.Message); return null; }
        }

        object EvalPlaceholder(Ph ph, DataScope scope)
        {
            object value = ph.Expr != null ? EvalExpr(ph.Expr, scope) : ph.Literal != null ? ph.Literal : ph.Fn != null ? Aggregate(ph.Fn, ph.FnArg, scope) : DataAccess.Resolve(scope, ph.Head);
            foreach (var f in ph.Filters) value = ApplyFilter(value, f[0], f[1]);
            return value;
        }

        static List<string> SplitPipes(string s)
        {
            var list = new List<string>();
            var sb = new StringBuilder();
            foreach (var ch in s)
            {
                if (ch == '|') { list.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(ch);
            }
            list.Add(sb.ToString());
            return list;
        }

        object Aggregate(string fn, string path, DataScope scope)
        {
            path = path.Trim();
            if (fn == "count")
            {
                var c = DataAccess.Resolve(scope, path);
                return (double)DataAccess.Items(c).Count();
            }
            int dot = path.LastIndexOf('.');
            if (dot < 0) { _warn("Aggregate '" + fn + "(" + path + ")' needs Collection.Field"); return null; }
            var coll = DataAccess.Items(DataAccess.Resolve(scope, path.Substring(0, dot)));
            var field = path.Substring(dot + 1);
            var nums = new List<double>();
            foreach (var it in coll)
            {
                object v;
                double d;
                if (DataAccess.TryMember(it, field, out v) && DataAccess.TryNumber(v, out d)) nums.Add(d);
            }
            if (nums.Count == 0) return fn == "sum" ? (object)0d : null;
            switch (fn)
            {
                case "sum": return nums.Sum();
                case "avg": case "average": return nums.Average();
                case "min": return nums.Min();
                default: return nums.Max();
            }
        }

        object ApplyFilter(object v, string name, string arg)
        {
            switch (name.ToLowerInvariant())
            {
                case "fa": case "persian": case "jalali":
                {
                    DateTime dt;
                    if (!TryDate(v, out dt)) return v;
                    return FormatJalali(dt, string.IsNullOrWhiteSpace(arg) ? "yyyy/MM/dd" : arg.Trim());
                }
                case "jyear": case "jmonth": case "jday": case "jmonthname": case "jym": case "jdayname": case "jquarter":
                {
                    DateTime dt; int jy, jm, jd;
                    if (!TryDate(v, out dt) || !Jalali.TryParts(dt, out jy, out jm, out jd)) return v;
                    switch (name.ToLowerInvariant())
                    {
                        case "jyear": return (double)jy;
                        case "jmonth": return (double)jm;
                        case "jday": return (double)jd;
                        case "jquarter": return (double)Jalali.Quarter(jm);
                        case "jym": return jy.ToString("0000") + "/" + jm.ToString("00");
                        case "jmonthname": return Jalali.MonthName(jm);
                        default: return Jalali.DayNames[(int)dt.DayOfWeek];
                    }
                }
                case "urlenc": case "urlencode": return Uri.EscapeDataString(ToDisplay(v));
                case "fadigits": case "fanum":
                    return ToPersianDigits(ToDisplay(v));
                case "format":
                {
                    if (v is IFormattable f && arg != null) return f.ToString(arg.Trim(), CultureInfo.InvariantCulture);
                    return v;
                }
                case "upper": return ToDisplay(v).ToUpperInvariant();
                case "lower": return ToDisplay(v).ToLowerInvariant();
                case "trim": return ToDisplay(v).Trim();
                case "default": return v == null || (v is string s && s.Length == 0) ? arg : v;
                case "yesno":
                {
                    var p = (arg ?? "Yes/No").Split('/');
                    var yes = p[0]; var no = p.Length > 1 ? p[1] : "";
                    return Truthy(v) ? yes : no;
                }
                case "round":
                {
                    double d; int n = 0;
                    int.TryParse(arg, out n);
                    return DataAccess.TryNumber(v, out d) ? (object)Math.Round(d, n, MidpointRounding.AwayFromZero) : v;
                }
                case "join":
                    return string.Join(arg ?? ", ", DataAccess.Items(v).Select(ToDisplay));
                case "truncate":
                {
                    int n; var s2 = ToDisplay(v);
                    return int.TryParse(arg, out n) && s2.Length > n ? s2.Substring(0, n) + "…" : s2;
                }
                case "text": return ToDisplay(v);
                default:
                    Func<object, string, object> custom;
                    if (_filters != null && _filters.TryGetValue(name, out custom)) return DataAccess.Normalize(custom(v, arg));
                    _warn("Unknown filter '" + name + "'");
                    return v;
            }
        }

        static bool Truthy(object v)
        {
            if (v == null) return false;
            if (v is bool b) return b;
            double d;
            if (DataAccess.TryNumber(v, out d)) return d != 0;
            var s = v.ToString().Trim().ToLowerInvariant();
            return !(s == "" || s == "false" || s == "0" || s == "no");
        }

        static bool TryDate(object v, out DateTime dt)
        {
            dt = default(DateTime);
            if (v is DateTime d) { dt = d; return true; }
            if (v is string s && DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) return true;
            return false;
        }

        public static string FormatJalali(DateTime dt, string fmt)
        {
            if (dt.Year < 622 || dt.Year > 9999) return dt.ToString(fmt, CultureInfo.InvariantCulture);
            int y, m, d;
            try { y = Pc.GetYear(dt); m = Pc.GetMonth(dt); d = Pc.GetDayOfMonth(dt); }
            catch (ArgumentOutOfRangeException) { return dt.ToString(fmt, CultureInfo.InvariantCulture); }
            return Regex.Replace(fmt, "yyyy|yy|MM|dd|HH|mm|ss", t =>
            {
                switch (t.Value)
                {
                    case "yyyy": return y.ToString("0000");
                    case "yy": return (y % 100).ToString("00");
                    case "MM": return m.ToString("00");
                    case "dd": return d.ToString("00");
                    case "HH": return dt.Hour.ToString("00");
                    case "mm": return dt.Minute.ToString("00");
                    default: return dt.Second.ToString("00");
                }
            });
        }

        public static string ToPersianDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var sb = new StringBuilder(s.Length);
            foreach (var ch in s) sb.Append(ch >= '0' && ch <= '9' ? (char)('۰' + (ch - '0')) : ch);
            return sb.ToString();
        }

        public static string ToDisplay(object v)
        {
            if (v == null) return "";
            if (v is string s) return s;
            if (v is bool b) return b ? "TRUE" : "FALSE";
            if (v is DateTime dt) return dt.TimeOfDay == TimeSpan.Zero ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : dt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
            if (v is IFormattable f) return f.ToString(null, CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }
    }
}
