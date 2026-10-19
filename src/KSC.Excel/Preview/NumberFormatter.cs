using System;
using System.Globalization;
using System.Text;

namespace KSC.Excel.Preview
{
    /// <summary>A pragmatic subset of Excel number formats, enough to show values in the HTML preview.</summary>
    public static class NumberFormatter
    {
        public static string Format(object value, string format)
        {
            if (value == null) return "";
            if (value is string s) return s;
            if (value is bool b) return b ? "TRUE" : "FALSE";

            var ci = CultureInfo.InvariantCulture;
            if (value is DateTime dt)
                return FormatDate(dt, string.IsNullOrEmpty(format) || format == "General" ? "yyyy/mm/dd" : format);

            double d;
            if (value is double) d = (double)value; else d = Convert.ToDouble(value, ci);

            if (string.IsNullOrEmpty(format) || format == "General")
                return d.ToString("0.##########", ci);
            if (format == "@") return d.ToString("0.##########", ci);

            // choose a section: positive;negative;zero
            var sections = SplitSections(format);
            string sec = sections[0];
            bool negSection = false;
            if (d < 0 && sections.Length > 1) { sec = sections[1]; negSection = true; }
            else if (d == 0 && sections.Length > 2) sec = sections[2];

            if (LooksLikeDate(sec))
            {
                try { return FormatDate(DateTime.FromOADate(d), sec); } catch (ArgumentException) { return d.ToString(ci); }
            }

            // literal prefix/suffix and numeric core
            var prefix = new StringBuilder(); var core = new StringBuilder(); var suffix = new StringBuilder();
            bool inQuote = false, seenCore = false, afterCore = false;
            for (int i = 0; i < sec.Length; i++)
            {
                char c = sec[i];
                if (c == '"') { inQuote = !inQuote; continue; }
                if (c == '\\' && i + 1 < sec.Length) { (afterCore ? suffix : prefix).Append(sec[++i]); continue; }
                if (c == '_' || c == '*') { i++; continue; }
                bool numeric = !inQuote && (c == '0' || c == '#' || c == '?' || c == '.' || c == ',' || c == '%');
                if (numeric && !afterCore) { core.Append(c); seenCore = true; }
                else if (seenCore) { afterCore = true; suffix.Append(c); }
                else prefix.Append(c);
            }

            var pattern = core.ToString();
            bool percent = pattern.IndexOf('%') >= 0;
            pattern = pattern.Replace("%", "");
            if (percent) d *= 100;

            int dot = pattern.IndexOf('.');
            int decimals = 0, optional = 0;
            if (dot >= 0) for (int i = dot + 1; i < pattern.Length; i++) { if (pattern[i] == '0') decimals++; else if (pattern[i] == '#' || pattern[i] == '?') optional++; }
            bool thousands = pattern.IndexOf(',') >= 0;
            int minInt = 0;
            for (int i = 0; i < (dot >= 0 ? dot : pattern.Length); i++) if (pattern[i] == '0') minInt++;

            double abs = negSection ? Math.Abs(d) : d;
            // '0' decimals are required, '#' decimals optional (trailing zeros dropped; Excel keeps the dot when none remain)
            var body = Math.Abs(abs).ToString((thousands ? "#,##" : "") + new string('0', Math.Max(minInt, 1)) + (decimals + optional > 0 ? "." + new string('0', decimals) + new string('#', optional) : ""), ci);
            if (optional > 0 && decimals == 0 && body.IndexOf('.') < 0) body += ".";
            if (minInt == 0 && body.StartsWith("0", StringComparison.Ordinal) && body.Length > 1 && body[1] == '.') body = body.Substring(1);
            if (percent) suffix.Insert(0, '%');
            var sign = abs < 0 && !negSection ? "-" : "";
            return prefix + sign + body + suffix;
        }

        static string[] SplitSections(string f)
        {
            var list = new System.Collections.Generic.List<string>();
            var sb = new StringBuilder(); bool q = false;
            foreach (var c in f)
            {
                if (c == '"') q = !q;
                if (c == ';' && !q) { list.Add(sb.ToString()); sb.Clear(); } else sb.Append(c);
            }
            list.Add(sb.ToString());
            return list.ToArray();
        }

        static bool LooksLikeDate(string f)
        {
            bool q = false;
            foreach (var c in f)
            {
                if (c == '"') q = !q;
                if (q) continue;
                char l = char.ToLowerInvariant(c);
                if (l == 'y' || l == 'd' || l == 'h' || l == 's' || (l == 'm' && f.IndexOf('0') < 0)) return true;
            }
            return false;
        }

        static string FormatDate(DateTime dt, string fmt)
        {
            var sb = new StringBuilder();
            int i = 0;
            while (i < fmt.Length)
            {
                char c = fmt[i];
                if (c == '"') { int j = fmt.IndexOf('"', i + 1); if (j < 0) j = fmt.Length - 1; sb.Append(fmt, i + 1, j - i - 1); i = j + 1; continue; }
                char l = char.ToLowerInvariant(c);
                if (l == 'y' || l == 'd' || l == 'h' || l == 'm' || l == 's')
                {
                    int n = 1; while (i + n < fmt.Length && char.ToLowerInvariant(fmt[i + n]) == l) n++;
                    switch (l)
                    {
                        case 'y': sb.Append(n <= 2 ? (dt.Year % 100).ToString("00") : dt.Year.ToString("0000")); break;
                        case 'd': sb.Append(n <= 1 ? dt.Day.ToString() : n == 2 ? dt.Day.ToString("00") : dt.ToString(n == 3 ? "ddd" : "dddd", CultureInfo.InvariantCulture)); break;
                        case 'h': sb.Append(n == 1 ? dt.Hour.ToString() : dt.Hour.ToString("00")); break;
                        case 's': sb.Append(n == 1 ? dt.Second.ToString() : dt.Second.ToString("00")); break;
                        default:
                        {
                            // 'm' directly after h or before s is minutes; otherwise month
                            bool minutes = PrevIsHour(sb, fmt, i) || NextIsSecond(fmt, i + n);
                            if (minutes) sb.Append(n == 1 ? dt.Minute.ToString() : dt.Minute.ToString("00"));
                            else sb.Append(n == 1 ? dt.Month.ToString() : n == 2 ? dt.Month.ToString("00") : dt.ToString(n == 3 ? "MMM" : "MMMM", CultureInfo.InvariantCulture));
                            break;
                        }
                    }
                    i += n; continue;
                }
                sb.Append(c); i++;
            }
            return sb.ToString();
        }

        static bool PrevIsHour(StringBuilder sb, string fmt, int i)
        {
            for (int k = i - 1; k >= 0; k--)
            {
                char c = char.ToLowerInvariant(fmt[k]);
                if (c == 'h') return true;
                if (c == 'y' || c == 'd' || c == 'm' || c == 's') return false;
            }
            return false;
        }

        static bool NextIsSecond(string fmt, int from)
        {
            for (int k = from; k < fmt.Length; k++)
            {
                char c = char.ToLowerInvariant(fmt[k]);
                if (c == 's') return true;
                if (c == 'y' || c == 'd' || c == 'm' || c == 'h') return false;
            }
            return false;
        }
    }
}
