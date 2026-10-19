using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace KSC.Excel.Engine
{
    /// <summary>
    /// A small Excel formula evaluator used to store cached values in the file and to show real numbers in the
    /// HTML preview. It supports the common subset (see docs); anything else returns "unsupported" and the
    /// cell simply has no cached value (Excel recalculates on open anyway).
    /// </summary>
    public sealed class FormulaEvaluator
    {
        /// <summary>An Excel error value such as #DIV/0!.</summary>
        public sealed class XlError { public string Code; public XlError(string c) { Code = c; } }
        sealed class Unsupported : Exception { }
        sealed class Range { public int R1, C1, R2, C2; }

        readonly RenderedSheet _sheet;
        readonly Dictionary<RenderedCell, object> _memo = new Dictionary<RenderedCell, object>();
        readonly HashSet<RenderedCell> _busy = new HashSet<RenderedCell>();

        readonly IDictionary<string, Func<object[], object>> _custom;

        public FormulaEvaluator(RenderedSheet sheet, IDictionary<string, Func<object[], object>> customFunctions = null) { _sheet = sheet; _custom = customFunctions; }

        /// <summary>Evaluates a formula cell. Returns false when the formula uses something unsupported; the result may be an <see cref="XlError"/>.</summary>
        public bool TryEvaluate(RenderedCell cell, out object result)
        {
            result = null;
            try
            {
                // errors (#DIV/0! ...) are legitimate results and are cached like any other value
                result = CellValue(cell);
                return true;
            }
            catch (Unsupported) { return false; }
            catch (FormatException) { return false; }
            catch (InvalidCastException) { return false; }
        }

        object CellValue(RenderedCell c)
        {
            if (c == null) return null;
            if (c.Formula == null) return c.Value;
            object v;
            if (_memo.TryGetValue(c, out v)) return v;
            if (!_busy.Add(c)) return new XlError("#REF!");
            try
            {
                v = new Parser(this, c.Formula).ParseAll();
                _memo[c] = v;
                return v;
            }
            finally { _busy.Remove(c); }
        }

        // ---------- parser / evaluator ----------
        sealed class Parser
        {
            readonly FormulaEvaluator _ev;
            readonly string _s;
            int _p;

            public Parser(FormulaEvaluator ev, string s) { _ev = ev; _s = s; }

            public object ParseAll()
            {
                var v = Compare();
                Ws();
                if (_p < _s.Length) throw new Unsupported();
                return Scalar(v);
            }

            object Scalar(object v) => _ev.Scalar(v);

            void Ws() { while (_p < _s.Length && char.IsWhiteSpace(_s[_p])) _p++; }
            bool Eat(string t) { Ws(); if (string.CompareOrdinal(_s, _p, t, 0, t.Length) == 0) { _p += t.Length; return true; } return false; }

            object Compare()
            {
                var l = Concat();
                while (true)
                {
                    string op = null;
                    foreach (var o in new[] { "<=", ">=", "<>", "=", "<", ">" }) { Ws(); if (string.CompareOrdinal(_s, _p, o, 0, o.Length) == 0) { op = o; break; } }
                    if (op == null) return l;
                    _p += op.Length;
                    var r = Concat();
                    l = Cmp(Scalar(l), Scalar(r), op);
                }
            }

            object Concat()
            {
                var l = Add();
                while (Eat("&")) { var r = Add(); l = Str(Scalar(l)) + Str(Scalar(r)); }
                return l;
            }

            object Add()
            {
                var l = Mul();
                while (true)
                {
                    Ws();
                    if (_p < _s.Length && (_s[_p] == '+' || _s[_p] == '-'))
                    {
                        char op = _s[_p++];
                        var r = Mul();
                        var a = Num(Scalar(l)); var b = Num(Scalar(r));
                        if (a is XlError) l = a; else if (b is XlError) l = b;
                        else l = op == '+' ? (double)a + (double)b : (double)a - (double)b;
                    }
                    else return l;
                }
            }

            object Mul()
            {
                var l = Pow();
                while (true)
                {
                    Ws();
                    if (_p < _s.Length && (_s[_p] == '*' || _s[_p] == '/'))
                    {
                        char op = _s[_p++];
                        var r = Pow();
                        var a = Num(Scalar(l)); var b = Num(Scalar(r));
                        if (a is XlError) l = a; else if (b is XlError) l = b;
                        else if (op == '/') l = (double)b == 0 ? new XlError("#DIV/0!") : (object)((double)a / (double)b);
                        else l = (double)a * (double)b;
                    }
                    else return l;
                }
            }

            object Pow()
            {
                var l = Unary();
                while (Eat("^")) { var r = Unary(); var a = Num(Scalar(l)); var b = Num(Scalar(r)); l = a is XlError ? a : b is XlError ? b : Math.Pow((double)a, (double)b); }
                return l;
            }

            object Unary()
            {
                Ws();
                if (Eat("-")) { var v = Num(Scalar(Unary())); return v is XlError ? v : -(double)v; }
                if (Eat("+")) return Unary();
                var p = Primary();
                while (Eat("%")) { var n = Num(Scalar(p)); p = n is XlError ? n : (double)n / 100; }
                return p;
            }

            object Primary()
            {
                Ws();
                if (_p >= _s.Length) throw new Unsupported();
                char ch = _s[_p];
                if (ch == '(') { _p++; var v = Compare(); if (!Eat(")")) throw new Unsupported(); return v; }
                if (ch == '"')
                {
                    var sb = new StringBuilder(); _p++;
                    while (_p < _s.Length)
                    {
                        if (_s[_p] == '"') { if (_p + 1 < _s.Length && _s[_p + 1] == '"') { sb.Append('"'); _p += 2; continue; } _p++; break; }
                        sb.Append(_s[_p++]);
                    }
                    return sb.ToString();
                }
                if (char.IsDigit(ch) || ch == '.')
                {
                    int st = _p;
                    while (_p < _s.Length && (char.IsDigit(_s[_p]) || _s[_p] == '.')) _p++;
                    if (_p < _s.Length && (_s[_p] == 'E' || _s[_p] == 'e') && _p + 1 < _s.Length && (char.IsDigit(_s[_p + 1]) || _s[_p + 1] == '-' || _s[_p + 1] == '+'))
                    { _p += 2; while (_p < _s.Length && char.IsDigit(_s[_p])) _p++; }
                    return double.Parse(_s.Substring(st, _p - st), CultureInfo.InvariantCulture);
                }
                if (ch == '#') // error literal
                {
                    int st = _p; _p++;
                    while (_p < _s.Length && (char.IsLetterOrDigit(_s[_p]) || _s[_p] == '/' || _s[_p] == '!' || _s[_p] == '?')) _p++;
                    return new XlError(_s.Substring(st, _p - st));
                }
                if (ch == '$' || char.IsLetter(ch) || ch == '_')
                {
                    int st = _p;
                    while (_p < _s.Length && (char.IsLetterOrDigit(_s[_p]) || _s[_p] == '$' || _s[_p] == '_' || _s[_p] == '.')) _p++;
                    var id = _s.Substring(st, _p - st);
                    Ws();
                    if (_p < _s.Length && _s[_p] == '(') { _p++; return Call(id.ToUpperInvariant()); }
                    if (_p < _s.Length && _s[_p] == '!') throw new Unsupported(); // other sheets
                    if (id.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                    if (id.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return false;
                    int r1, c1;
                    if (!CellRef.TryParseCell(id, out r1, out c1)) throw new Unsupported();
                    if (_p < _s.Length && _s[_p] == ':')
                    {
                        _p++;
                        int st2 = _p;
                        while (_p < _s.Length && (char.IsLetterOrDigit(_s[_p]) || _s[_p] == '$')) _p++;
                        int r2, c2;
                        if (!CellRef.TryParseCell(_s.Substring(st2, _p - st2), out r2, out c2)) throw new Unsupported();
                        return new Range { R1 = Math.Min(r1, r2), C1 = Math.Min(c1, c2), R2 = Math.Max(r1, r2), C2 = Math.Max(c1, c2) };
                    }
                    return new Range { R1 = r1, C1 = c1, R2 = r1, C2 = c1 };
                }
                throw new Unsupported();
            }

            object Call(string fn)
            {
                var args = new List<object>();
                Ws();
                if (!Eat(")"))
                {
                    do { Ws(); args.Add(Compare()); } while (Eat(","));
                    if (!Eat(")")) throw new Unsupported();
                }
                return _ev.Function(fn, args);
            }
        }

        // ---------- value helpers ----------
        object Scalar(object v)
        {
            var r = v as Range;
            if (r == null) return v;
            if (r.R1 == r.R2 && r.C1 == r.C2) return Norm(CellValue(_sheet.Get(r.R1, r.C1)));
            throw new Unsupported(); // implicit intersection not supported
        }

        static object Norm(object v) => v is DateTime dt ? (object)dt.ToOADate() : v;

        static object Num(object v)
        {
            if (v is XlError) return v;
            if (v == null) return 0d;
            if (v is double) return v;
            if (v is bool b) return b ? 1d : 0d;
            double d;
            if (v is string s) { if (s.Length == 0) return 0d; if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d; return new XlError("#VALUE!"); }
            return new XlError("#VALUE!");
        }

        static string Str(object v)
        {
            if (v == null) return "";
            if (v is bool b) return b ? "TRUE" : "FALSE";
            if (v is double d) return d.ToString("R", CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        static object Cmp(object a, object b, string op)
        {
            if (a is XlError) return a;
            if (b is XlError) return b;
            int c;
            if (a is string || b is string) c = string.Compare(Str(a), Str(b), StringComparison.OrdinalIgnoreCase);
            else c = ((double)Num(a)).CompareTo((double)Num(b));
            switch (op)
            {
                case "=": return c == 0;
                case "<>": return c != 0;
                case "<": return c < 0;
                case ">": return c > 0;
                case "<=": return c <= 0;
                default: return c >= 0;
            }
        }

        static bool Truth(object v)
        {
            if (v is bool b) return b;
            if (v is double d) return d != 0;
            if (v is string s) return s.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
            return false;
        }

        IEnumerable<object> Flatten(object a)
        {
            var r = a as Range;
            if (r == null) { yield return a; yield break; }
            for (int row = r.R1; row <= r.R2 && row <= _sheet.Rows.Count; row++)
            {
                foreach (var c in _sheet.Rows[row - 1].Cells)
                {
                    if (c.Col < r.C1) continue;
                    if (c.Col > r.C2) break;
                    yield return Norm(CellValue(c));
                }
            }
        }

        // numbers inside ranges only count when numeric; direct args are coerced
        List<double> Numbers(List<object> args, out XlError err)
        {
            err = null;
            var list = new List<double>();
            foreach (var a in args)
            {
                bool isRange = a is Range;
                foreach (var v in Flatten(a))
                {
                    if (v is XlError e) { err = e; return list; }
                    if (v is double d) list.Add(d);
                    else if (!isRange)
                    {
                        var n = Num(v);
                        if (n is XlError e2) { err = e2; return list; }
                        list.Add((double)n);
                    }
                }
            }
            return list;
        }

        object Function(string fn, List<object> a)
        {
            XlError err;
            switch (fn)
            {
                case "SUM": { var n = Numbers(a, out err); return err ?? (object)n.Sum(); }
                case "AVERAGE": { var n = Numbers(a, out err); if (err != null) return err; return n.Count == 0 ? new XlError("#DIV/0!") : (object)n.Average(); }
                case "MIN": { var n = Numbers(a, out err); return err ?? (object)(n.Count == 0 ? 0 : n.Min()); }
                case "MAX": { var n = Numbers(a, out err); return err ?? (object)(n.Count == 0 ? 0 : n.Max()); }
                case "COUNT": { int cnt = 0; foreach (var x in a) foreach (var v in Flatten(x)) if (v is double) cnt++; return (double)cnt; }
                case "COUNTA": { int cnt = 0; foreach (var x in a) foreach (var v in Flatten(x)) if (v != null && !(v is string s && s.Length == 0)) cnt++; return (double)cnt; }
                case "ABS": { var n = Num(Scalar(a[0])); return n is XlError ? n : Math.Abs((double)n); }
                case "ROUND": { var n = Num(Scalar(a[0])); var d = Num(Scalar(a[1])); if (n is XlError) return n; if (d is XlError) return d; return Math.Round((double)n, (int)(double)d, MidpointRounding.AwayFromZero); }
                case "IF": { var c = Scalar(a[0]); if (c is XlError) return c; if (Truth(c)) return a.Count > 1 ? Scalar(a[1]) : true; return a.Count > 2 ? Scalar(a[2]) : false; }
                case "IFERROR": { var v = Safe(a[0]); return v is XlError ? Scalar(a[1]) : v; }
                case "AND": return a.SelectMany(Flatten).All(Truth);
                case "OR": return a.SelectMany(Flatten).Any(Truth);
                case "NOT": return !Truth(Scalar(a[0]));
                case "CONCATENATE": case "CONCAT": return string.Concat(a.SelectMany(Flatten).Select(Str));
                case "LEN": return (double)Str(Scalar(a[0])).Length;
                case "UPPER": return Str(Scalar(a[0])).ToUpperInvariant();
                case "LOWER": return Str(Scalar(a[0])).ToLowerInvariant();
                case "TRIM": return Str(Scalar(a[0])).Trim();
                case "SUBTOTAL":
                {
                    // like Excel: cells that are themselves SUBTOTALs are ignored, so nested group totals never double count
                    var code = Num(Scalar(a[0])); if (code is XlError) return code;
                    int fc = (int)(double)code % 100;
                    var nums = new List<double>(); int cnt = 0;
                    foreach (var x in a.Skip(1))
                    {
                        var r = x as Range;
                        if (r == null) { var v = Norm(Scalar(x)); if (v is double) nums.Add((double)v); if (v != null) cnt++; continue; }
                        for (int row = r.R1; row <= r.R2 && row <= _sheet.Rows.Count; row++)
                            foreach (var c in _sheet.Rows[row - 1].Cells)
                            {
                                if (c.Col < r.C1) continue;
                                if (c.Col > r.C2) break;
                                if (c.Formula != null && c.Formula.TrimStart().StartsWith("SUBTOTAL(", StringComparison.OrdinalIgnoreCase)) continue;
                                var v = Norm(CellValue(c));
                                if (v is XlError) return v;
                                if (v is double) nums.Add((double)v);
                                if (v != null && !(v is string s0 && s0.Length == 0)) cnt++;
                            }
                    }
                    switch (fc)
                    {
                        case 1: return nums.Count == 0 ? new XlError("#DIV/0!") : (object)nums.Average();
                        case 2: return (double)nums.Count;
                        case 3: return (double)cnt;
                        case 4: return nums.Count == 0 ? 0 : nums.Max();
                        case 5: return nums.Count == 0 ? 0 : nums.Min();
                        case 9: return nums.Sum();
                        default: throw new Unsupported();
                    }
                }
                case "SUMIF": case "COUNTIF": return CondAgg(fn, a);
                default:
                {
                    Func<object[], object> f;
                    if (_custom == null || !_custom.TryGetValue(fn, out f)) throw new Unsupported();
                    var vals = a.Select(x => { var rg = x as Range; return rg != null && (rg.R1 != rg.R2 || rg.C1 != rg.C2) ? (object)Flatten(x).ToArray() : Scalar(x); }).ToArray();
                    var res = DataAccess.Normalize(f(vals));
                    if (res is DateTime) return ((DateTime)res).ToOADate();
                    double d;
                    if (DataAccess.TryNumber(res, out d)) return d;
                    return res;
                }
            }
        }

        object Safe(object v) { try { return Scalar(v); } catch (Unsupported) { throw; } }

        object CondAgg(string fn, List<object> a)
        {
            var range = a[0] as Range; if (range == null) throw new Unsupported();
            var crit = Scalar(a[1]);
            var sumRange = fn == "SUMIF" ? (a.Count > 2 ? a[2] as Range : range) : null;
            if (fn == "SUMIF" && sumRange == null) throw new Unsupported();

            string op = "="; string cs = Str(crit); object cval = crit;
            foreach (var o in new[] { "<=", ">=", "<>", "<", ">", "=" })
                if (cs.StartsWith(o, StringComparison.Ordinal)) { op = o; cs = cs.Substring(o.Length); double dd; cval = double.TryParse(cs, NumberStyles.Float, CultureInfo.InvariantCulture, out dd) ? (object)dd : cs; break; }

            double total = 0; int count = 0;
            for (int row = range.R1; row <= range.R2; row++)
                for (int col = range.C1; col <= range.C2; col++)
                {
                    var v = Norm(CellValue(_sheet.Get(row, col)));
                    var m = Cmp(v ?? "", cval, op);
                    if (m is bool mb && mb)
                    {
                        count++;
                        if (fn == "SUMIF")
                        {
                            var sv = Norm(CellValue(_sheet.Get(sumRange.R1 + (row - range.R1), sumRange.C1 + (col - range.C1))));
                            if (sv is double sd) total += sd;
                        }
                    }
                }
            return fn == "SUMIF" ? total : (double)count;
        }
    }
}
