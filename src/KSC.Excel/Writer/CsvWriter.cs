using System;
using System.Globalization;
using System.Text;
using KSC.Excel.Engine;
using KSC.Excel.Preview;

namespace KSC.Excel.Writer
{
    /// <summary>Writes one sheet of a rendered report as CSV (UTF-8 with BOM so Excel shows Persian correctly).</summary>
    public static class CsvWriter
    {
        public static byte[] ToBytes(RenderedWorkbook wb, int sheetIndex = 0, char delimiter = ',', bool bom = true)
        {
            var text = ToString(wb, sheetIndex, delimiter);
            var body = new UTF8Encoding(false).GetBytes(text);
            if (!bom) return body;
            var res = new byte[body.Length + 3];
            res[0] = 0xEF; res[1] = 0xBB; res[2] = 0xBF;
            Buffer.BlockCopy(body, 0, res, 3, body.Length);
            return res;
        }

        public static string ToString(RenderedWorkbook wb, int sheetIndex = 0, char delimiter = ',')
        {
            if (sheetIndex < 0 || sheetIndex >= wb.Sheets.Count) throw new ArgumentOutOfRangeException("sheetIndex", "The report has " + wb.Sheets.Count + " sheet(s).");
            var sh = wb.Sheets[sheetIndex];
            var ev = new FormulaEvaluator(sh);
            var sb = new StringBuilder();
            int cols = Math.Max(1, sh.MaxCol);
            var merged = new System.Collections.Generic.HashSet<long>();
            foreach (var m in sh.Merges)
                for (int r = m[0]; r <= m[2]; r++) for (int c = m[1]; c <= m[3]; c++) if (r != m[0] || c != m[1]) merged.Add((long)r * 100000 + c);

            for (int r = 1; r <= sh.Rows.Count; r++)
            {
                var row = sh.Rows[r - 1];
                if (row.Hidden) continue;
                for (int c = 1; c <= cols; c++)
                {
                    if (c > 1) sb.Append(delimiter);
                    if (merged.Contains((long)r * 100000 + c)) continue;
                    var cell = sh.Get(r, c);
                    if (cell == null) continue;
                    object v = cell.Value;
                    if (cell.Formula != null) { v = cell.Cached; object e; if (v == null && ev.TryEvaluate(cell, out e)) v = e; }
                    sb.Append(Field(v, delimiter));
                }
                sb.Append("\r\n");
            }
            return sb.ToString();
        }

        static string Field(object v, char delimiter)
        {
            string s;
            var err = v as FormulaEvaluator.XlError;
            if (err != null) s = err.Code;
            else if (v == null) return "";
            else if (v is double) s = ((double)v).ToString("R", CultureInfo.InvariantCulture);
            else if (v is DateTime) { var d = (DateTime)v; s = d.TimeOfDay == TimeSpan.Zero ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture); }
            else if (v is bool) s = (bool)v ? "TRUE" : "FALSE";
            else
            {
                s = Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
                // spreadsheet formula injection: text that Excel would treat as a formula is neutralised
                if (s.Length > 0 && (s[0] == '=' || s[0] == '+' || s[0] == '-' || s[0] == '@' || s[0] == '\t' || s[0] == '\r')) s = "'" + s;
            }
            if (s.IndexOf(delimiter) >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\n') >= 0 || s.IndexOf('\r') >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }
}
