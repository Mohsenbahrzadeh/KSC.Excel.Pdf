using System;
using System.Text;
using System.Text.RegularExpressions;

namespace KSC.Excel.Engine
{
    /// <summary>A1-notation helpers.</summary>
    public static class CellRef
    {
        public const int MaxRows = 1048576;
        public const int MaxCols = 16384;

        static readonly string[] ColCache = new string[1025];

        public static string ColName(int col)
        {
            if (col < 1) throw new ArgumentOutOfRangeException(nameof(col));
            if (col < ColCache.Length) return ColCache[col] ?? (ColCache[col] = BuildColName(col));
            return BuildColName(col);
        }

        static string BuildColName(int col)
        {
            var sb = new StringBuilder();
            while (col > 0)
            {
                int m = (col - 1) % 26;
                sb.Insert(0, (char)('A' + m));
                col = (col - 1) / 26;
            }
            return sb.ToString();
        }

        public static int ColIndex(string name)
        {
            int n = 0;
            foreach (char ch in name.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') throw new FormatException("Bad column name: " + name);
                n = n * 26 + (ch - 'A' + 1);
            }
            return n;
        }

        public static string A1(int row, int col) => ColName(col) + row;

        static readonly Regex CellRx = new Regex(@"^\$?([A-Za-z]{1,3})\$?(\d+)$", RegexOptions.Compiled);

        public static bool TryParseCell(string s, out int row, out int col)
        {
            row = col = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            var m = CellRx.Match(s.Trim());
            if (!m.Success) return false;
            col = ColIndex(m.Groups[1].Value);
            row = int.Parse(m.Groups[2].Value);
            return row >= 1 && row <= MaxRows && col >= 1 && col <= MaxCols;
        }

        /// <summary>Parses "A1" or "A1:C5" (rows/cols normalized so r1&lt;=r2, c1&lt;=c2).</summary>
        public static bool TryParseRange(string s, out int r1, out int c1, out int r2, out int c2)
        {
            r1 = c1 = r2 = c2 = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            var parts = s.Split(':');
            if (parts.Length == 1)
            {
                if (!TryParseCell(parts[0], out r1, out c1)) return false;
                r2 = r1; c2 = c1; return true;
            }
            if (parts.Length != 2) return false;
            if (!TryParseCell(parts[0], out r1, out c1) || !TryParseCell(parts[1], out r2, out c2)) return false;
            if (r1 > r2) { var t = r1; r1 = r2; r2 = t; }
            if (c1 > c2) { var t = c1; c1 = c2; c2 = t; }
            return true;
        }

        public static string Range(int r1, int c1, int r2, int c2)
            => (r1 == r2 && c1 == c2) ? A1(r1, c1) : A1(r1, c1) + ":" + A1(r2, c2);

        static readonly Regex RefInFormula = new Regex(@"(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])", RegexOptions.Compiled);

        /// <summary>
        /// Shifts the row of every relative A1 reference in a formula by <paramref name="delta"/>
        /// (used to rebase conditional-format formulas when bands move rows).
        /// </summary>
        public static string ShiftRows(string formula, int delta) => Shift(formula, delta, 0);

        /// <summary>Shifts relative row/column references of an A1 formula (absolute parts marked with $ stay).</summary>
        public static string Shift(string formula, int dRow, int dCol)
        {
            int delta = dRow;
            if (string.IsNullOrEmpty(formula) || (dRow == 0 && dCol == 0)) return formula;
            var sb = new StringBuilder();
            bool inStr = false;
            int i = 0;
            while (i < formula.Length)
            {
                if (formula[i] == '"') { inStr = !inStr; sb.Append('"'); i++; continue; }
                if (inStr) { sb.Append(formula[i]); i++; continue; }
                // copy a non-string run, then transform refs in it
                int j = i;
                while (j < formula.Length && formula[j] != '"') j++;
                var run = formula.Substring(i, j - i);
                sb.Append(RefInFormula.Replace(run, m =>
                {
                    // skip things like LOG10( or part of identifiers / sheet names
                    if (m.Index > 0 && (char.IsLetterOrDigit(run[m.Index - 1]) || run[m.Index - 1] == '_' || run[m.Index - 1] == '.')) return m.Value;
                    int row = int.Parse(m.Groups[4].Value);
                    if (m.Groups[3].Value != "$") row = Math.Max(1, row + delta);
                    string col = m.Groups[2].Value;
                    if (m.Groups[1].Value != "$" && dCol != 0)
                    {
                        int ci = ColIndex(col) + dCol;
                        if (ci < 1) ci = 1;
                        if (ci > MaxCols) ci = MaxCols;
                        col = ColName(ci);
                    }
                    return m.Groups[1].Value + col + m.Groups[3].Value + row;
                }));
                i = j;
            }
            return sb.ToString();
        }
    }
}
