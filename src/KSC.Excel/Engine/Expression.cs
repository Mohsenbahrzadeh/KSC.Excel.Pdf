using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace KSC.Excel.Engine
{
    /// <summary>Services an expression needs from the renderer (aggregates over collections and custom functions).</summary>
    internal sealed class ExprEnv
    {
        public Func<string, string, DataScope, object> Aggregate;
        public IDictionary<string, Func<object[], object>> Functions;
    }

    /// <summary>
    /// Calculated-field expressions inside placeholders: {{Qty * UnitPrice}}, {{IF(Total &gt; 1000, 'VIP', 'Normal')}},
    /// {{ROUND(Price * 1.09, 0)}}, {{Tax(Price)}} (custom function). Paths resolve against the current data scope.
    /// </summary>
    internal sealed class Expr
    {
        readonly Func<DataScope, ExprEnv, object> _root;
        Expr(Func<DataScope, ExprEnv, object> root) { _root = root; }

        public object Eval(DataScope scope, ExprEnv env) { return _root(scope, env); }

        static readonly HashSet<string> Aggregates = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "count", "sum", "avg", "average", "min", "max" };

        /// <summary>True when the placeholder head is more than a plain path / literal (operators, calls, spaces).</summary>
        public static bool LooksLikeExpression(string head)
        {
            foreach (var ch in head)
                if (!(char.IsLetterOrDigit(ch) || ch == '_' || ch == '.' || ch == '[' || ch == ']' || ch == '$')) return true;
            return false;
        }

        public static Expr Parse(string text)
        {
            var p = new Parser(text);
            var node = p.ParseAll();
            return new Expr(node);
        }

        // ------------------------------------------------------------------
        sealed class Parser
        {
            readonly string _s; int _p;
            public Parser(string s) { _s = s; }

            public Func<DataScope, ExprEnv, object> ParseAll()
            {
                var n = Compare();
                Ws();
                if (_p < _s.Length) throw new FormatException("Unexpected '" + _s[_p] + "' in expression: " + _s);
                return n;
            }

            void Ws() { while (_p < _s.Length && char.IsWhiteSpace(_s[_p])) _p++; }
            bool Eat(string t) { Ws(); if (string.CompareOrdinal(_s, _p, t, 0, t.Length) == 0) { _p += t.Length; return true; } return false; }

            Func<DataScope, ExprEnv, object> Compare()
            {
                var l = Concat();
                while (true)
                {
                    string op = null;
                    foreach (var o in new[] { "<=", ">=", "<>", "!=", "==", "=", "<", ">" }) { Ws(); if (string.CompareOrdinal(_s, _p, o, 0, o.Length) == 0) { op = o; break; } }
                    if (op == null) return l;
                    _p += op.Length;
                    var r = Concat(); var left = l; var oper = op;
                    l = (sc, env) => Cmp(left(sc, env), r(sc, env), oper);
                }
            }

            Func<DataScope, ExprEnv, object> Concat()
            {
                var l = Add();
                while (Eat("&")) { var r = Add(); var left = l; l = (sc, env) => Str(left(sc, env)) + Str(r(sc, env)); }
                return l;
            }

            Func<DataScope, ExprEnv, object> Add()
            {
                var l = Mul();
                while (true)
                {
                    Ws();
                    if (_p < _s.Length && (_s[_p] == '+' || _s[_p] == '-'))
                    {
                        char op = _s[_p++]; var r = Mul(); var left = l;
                        l = (sc, env) =>
                        {
                            var a = left(sc, env); var b = r(sc, env);
                            double x, y;
                            if (op == '+' && (a is string || b is string) && !(Num(a, out x) && Num(b, out y))) return Str(a) + Str(b);
                            if (!Num(a, out x) && a != null) return null; if (!Num(b, out y) && b != null) return null;
                            x = a == null ? 0 : x; y = b == null ? 0 : y;
                            return op == '+' ? x + y : x - y;
                        };
                    }
                    else return l;
                }
            }

            Func<DataScope, ExprEnv, object> Mul()
            {
                var l = Pow();
                while (true)
                {
                    Ws();
                    if (_p < _s.Length && (_s[_p] == '*' || _s[_p] == '/' || _s[_p] == '%'))
                    {
                        char op = _s[_p++]; var r = Pow(); var left = l;
                        l = (sc, env) =>
                        {
                            double x, y;
                            if (!Num(left(sc, env), out x) || !Num(r(sc, env), out y)) return null;
                            if (op == '*') return x * y;
                            if (y == 0) return null;
                            return op == '/' ? x / y : x % y;
                        };
                    }
                    else return l;
                }
            }

            Func<DataScope, ExprEnv, object> Pow()
            {
                var l = Unary();
                while (Eat("^")) { var r = Unary(); var left = l; l = (sc, env) => { double x, y; return Num(left(sc, env), out x) && Num(r(sc, env), out y) ? (object)Math.Pow(x, y) : null; }; }
                return l;
            }

            Func<DataScope, ExprEnv, object> Unary()
            {
                Ws();
                if (Eat("-")) { var u = Unary(); return (sc, env) => { double x; return Num(u(sc, env), out x) ? (object)(-x) : null; }; }
                if (Eat("+")) return Unary();
                return Primary();
            }

            Func<DataScope, ExprEnv, object> Primary()
            {
                Ws();
                if (_p >= _s.Length) throw new FormatException("Unexpected end of expression: " + _s);
                char ch = _s[_p];
                if (ch == '(') { _p++; var v = Compare(); if (!Eat(")")) throw new FormatException("Missing ) in: " + _s); return v; }
                if (ch == '\'' || ch == '"')
                {
                    var sb = new StringBuilder(); char q = ch; _p++;
                    while (_p < _s.Length)
                    {
                        if (_s[_p] == q) { if (_p + 1 < _s.Length && _s[_p + 1] == q) { sb.Append(q); _p += 2; continue; } _p++; break; }
                        sb.Append(_s[_p++]);
                    }
                    var str = sb.ToString();
                    return (sc, env) => str;
                }
                if (char.IsDigit(ch) || (ch == '.' && _p + 1 < _s.Length && char.IsDigit(_s[_p + 1])))
                {
                    int st = _p;
                    while (_p < _s.Length && (char.IsDigit(_s[_p]) || _s[_p] == '.')) _p++;
                    double d = double.Parse(_s.Substring(st, _p - st), CultureInfo.InvariantCulture);
                    return (sc, env) => d;
                }
                if (char.IsLetter(ch) || ch == '_' || ch == '$')
                {
                    int st = _p;
                    while (_p < _s.Length && (char.IsLetterOrDigit(_s[_p]) || _s[_p] == '_' || _s[_p] == '.' || _s[_p] == '$' || _s[_p] == '[' || _s[_p] == ']')) _p++;
                    var id = _s.Substring(st, _p - st);
                    Ws();
                    if (_p < _s.Length && _s[_p] == '(') { _p++; return Call(id); }
                    var low = id.ToLowerInvariant();
                    if (low == "true") return (sc, env) => true;
                    if (low == "false") return (sc, env) => false;
                    if (low == "null") return (sc, env) => null;
                    return (sc, env) => DataAccess.Resolve(sc, id);
                }
                throw new FormatException("Unexpected '" + ch + "' in expression: " + _s);
            }

            Func<DataScope, ExprEnv, object> Call(string name)
            {
                if (Aggregates.Contains(name))
                {
                    // raw path argument: Sum(Items.Qty) / Count(Items)
                    int depth = 1, st = _p;
                    while (_p < _s.Length && depth > 0) { if (_s[_p] == '(') depth++; else if (_s[_p] == ')') depth--; _p++; }
                    if (depth != 0) throw new FormatException("Missing ) in: " + _s);
                    var arg = _s.Substring(st, _p - 1 - st).Trim(); var fn = name.ToLowerInvariant();
                    return (sc, env) => env.Aggregate(fn, arg, sc);
                }
                var args = new List<Func<DataScope, ExprEnv, object>>();
                Ws();
                if (!Eat(")"))
                {
                    do { Ws(); args.Add(Compare()); } while (Eat(","));
                    if (!Eat(")")) throw new FormatException("Missing ) in: " + _s);
                }
                var key = name.ToUpperInvariant();
                return (sc, env) => Invoke(key, args, sc, env);
            }
        }

        // ------------------------------------------------------------------ evaluation helpers
        static object Invoke(string fn, List<Func<DataScope, ExprEnv, object>> a, DataScope sc, ExprEnv env)
        {
            Func<object[], object> custom;
            if (env != null && env.Functions != null && env.Functions.TryGetValue(fn, out custom))
                return DataAccess.Normalize(custom(a.Select(x => x(sc, env)).ToArray()));

            double x1, x2;
            switch (fn)
            {
                case "IF": return Truth(a[0](sc, env)) ? (a.Count > 1 ? a[1](sc, env) : true) : (a.Count > 2 ? a[2](sc, env) : false);
                case "IFERROR": case "IFNULL": case "COALESCE":
                    foreach (var arg in a) { var v = arg(sc, env); if (v != null && !(v is string s && s.Length == 0)) return v; }
                    return null;
                case "ROUND": return Num(a[0](sc, env), out x1) ? (object)Math.Round(x1, a.Count > 1 && Num(a[1](sc, env), out x2) ? (int)x2 : 0, MidpointRounding.AwayFromZero) : null;
                case "ABS": return Num(a[0](sc, env), out x1) ? (object)Math.Abs(x1) : null;
                case "MIN": case "MAX": case "SUM": case "AVG":
                {
                    var nums = new List<double>();
                    foreach (var arg in a) { double d; if (Num(arg(sc, env), out d)) nums.Add(d); }
                    if (nums.Count == 0) return null;
                    return fn == "MIN" ? nums.Min() : fn == "MAX" ? nums.Max() : fn == "SUM" ? nums.Sum() : nums.Average();
                }
                case "LEN": return (double)Str(a[0](sc, env)).Length;
                case "UPPER": return Str(a[0](sc, env)).ToUpperInvariant();
                case "LOWER": return Str(a[0](sc, env)).ToLowerInvariant();
                case "TRIM": return Str(a[0](sc, env)).Trim();
                case "CONCAT": case "CONCATENATE": return string.Concat(a.Select(x => Str(x(sc, env))));
                case "AND": return a.All(x => Truth(x(sc, env)));
                case "OR": return a.Any(x => Truth(x(sc, env)));
                case "NOT": return !Truth(a[0](sc, env));
                case "TODAY": return DateTime.Today;
                case "NOW": return DateTime.Now;
                case "YEAR": case "MONTH": case "DAY":
                {
                    var v = a[0](sc, env); DateTime dt;
                    if (v is DateTime) dt = (DateTime)v; else if (!(v is string && DateTime.TryParse((string)v, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))) return null;
                    return (double)(fn == "YEAR" ? dt.Year : fn == "MONTH" ? dt.Month : dt.Day);
                }
                case "FA":
                {
                    var v = a[0](sc, env); DateTime dt;
                    if (v is DateTime) dt = (DateTime)v; else if (!(v is string && DateTime.TryParse((string)v, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))) return null;
                    return TextEvaluator.FormatJalali(dt, a.Count > 1 ? Str(a[1](sc, env)) : "yyyy/MM/dd");
                }
                case "JYEAR": case "JMONTH": case "JDAY": case "JQUARTER": case "JYM": case "JMONTHNAME": case "JDAYNAME":
                {
                    var v = a[0](sc, env); DateTime dt; int jy, jm, jd;
                    if (fn == "JMONTHNAME" && Num(v, out x1) && !(v is DateTime) && x1 >= 1 && x1 <= 12) return Jalali.MonthName((int)x1);
                    if (!AsDate(v, out dt) || !Jalali.TryParts(dt, out jy, out jm, out jd)) return null;
                    switch (fn)
                    {
                        case "JYEAR": return (double)jy;
                        case "JMONTH": return (double)jm;
                        case "JDAY": return (double)jd;
                        case "JQUARTER": return (double)Jalali.Quarter(jm);
                        case "JYM": return jy.ToString("0000") + "/" + jm.ToString("00");
                        case "JMONTHNAME": return Jalali.MonthName(jm);
                        default: return Jalali.DayNames[(int)dt.DayOfWeek];
                    }
                }
                case "JDATE":
                {
                    double jy, jm, jd;
                    if (a.Count < 3 || !Num(a[0](sc, env), out jy) || !Num(a[1](sc, env), out jm) || !Num(a[2](sc, env), out jd)) return null;
                    var r = Jalali.ToDate((int)jy, (int)jm, (int)jd);
                    return r.HasValue ? (object)r.Value : null;
                }
                case "JPARSE":
                {
                    DateTime dt;
                    return Jalali.TryParse(Str(a[0](sc, env)), out dt) ? (object)dt : null;
                }
                case "RUNSUM": case "PCT": case "PREV": case "DIFF": case "RANK": case "MOVAVG":
                {
                    var sib = sc.Siblings;
                    if (sib == null || sib.Count == 0 || a.Count == 0) return null;
                    int idx = Math.Max(0, Math.Min(sib.Count - 1, sc.Index - 1));
                    Func<int, object> val = j => a[0](new DataScope(sib[j], sc.Parent) { Index = j + 1, Count = sib.Count, Band = sc.Band, Siblings = sib }, env);
                    double cur, o;
                    switch (fn)
                    {
                        case "RUNSUM": { double t = 0; for (int j = 0; j <= idx; j++) if (Num(val(j), out o)) t += o; return t; }
                        case "PCT": { double t = 0; for (int j = 0; j < sib.Count; j++) if (Num(val(j), out o)) t += o; return Num(val(idx), out cur) && t != 0 ? (object)(cur / t) : null; }
                        case "PREV": return idx == 0 ? null : val(idx - 1);
                        case "DIFF": { double p; return idx > 0 && Num(val(idx), out cur) && Num(val(idx - 1), out p) ? (object)(cur - p) : null; }
                        case "RANK": { if (!Num(val(idx), out cur)) return null; int rank = 1; for (int j = 0; j < sib.Count; j++) if (Num(val(j), out o) && o > cur) rank++; return (double)rank; }
                        default:
                        {
                            int n = a.Count > 1 && Num(a[1](sc, env), out x2) ? Math.Max(1, (int)x2) : 3;
                            double t = 0; int c = 0;
                            for (int j = Math.Max(0, idx - n + 1); j <= idx; j++) if (Num(val(j), out o)) { t += o; c++; }
                            return c == 0 ? null : (object)(t / c);
                        }
                    }
                }
                default: throw new InvalidOperationException("Unknown function '" + fn + "'. Register it with options.AddFunction(...).");
            }
        }

        static bool AsDate(object v, out DateTime dt)
        {
            dt = default(DateTime);
            if (v is DateTime) { dt = (DateTime)v; return true; }
            var s = v as string;
            return s != null && DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt);
        }

        static bool Num(object v, out double d)
        {
            d = 0;
            if (v == null) return false;
            if (DataAccess.TryNumber(v, out d)) return true;
            if (v is bool) { d = (bool)v ? 1 : 0; return true; }
            var s = v as string;
            return s != null && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d);
        }

        static bool Truth(object v)
        {
            if (v == null) return false;
            if (v is bool) return (bool)v;
            double d;
            if (DataAccess.TryNumber(v, out d)) return d != 0;
            var s = v.ToString().Trim().ToLowerInvariant();
            return !(s == "" || s == "false" || s == "no" || s == "0");
        }

        static string Str(object v) => TextEvaluator.ToDisplay(v);

        static object Cmp(object a, object b, string op)
        {
            int c;
            double x, y;
            if (a == null && b == null) c = 0;
            else if (a is DateTime && b is DateTime) c = ((DateTime)a).CompareTo((DateTime)b);
            else if (!(a is string && !Num(a, out x)) && !(b is string && !Num(b, out y)) && Num(a, out x) && Num(b, out y)) c = x.CompareTo(y);
            else c = string.Compare(Str(a), Str(b), StringComparison.CurrentCultureIgnoreCase);
            switch (op)
            {
                case "=": case "==": return c == 0;
                case "<>": case "!=": return c != 0;
                case "<": return c < 0;
                case ">": return c > 0;
                case "<=": return c <= 0;
                default: return c >= 0;
            }
        }
    }
}
