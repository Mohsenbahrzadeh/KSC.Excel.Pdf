using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using KSC.Excel.Model;

namespace KSC.Excel.Engine
{
    /// <summary>
    /// Decides whether a template can be rendered by the streaming writer. Streaming writes each row as soon as it exists, so a
    /// row must never depend on data that is produced later in the report.
    /// </summary>
    public static class StreamAnalyzer
    {
        static readonly Regex Tok = new Regex(@"\[([A-Za-z_][\w]*)(?:([.:])([A-Za-z_][\w\.]*))?\]", RegexOptions.Compiled);

        /// <summary>Returns the reasons a template cannot be streamed (empty = streamable).</summary>
        public static IList<string> Check(ExcelTemplate t)
        {
            var reasons = new List<string>();
            foreach (var sh in t.Sheets)
            {
                var rowBands = sh.Bands.Where(b => !string.Equals(b.Orientation, "columns", StringComparison.OrdinalIgnoreCase) && b.FirstRow >= 1 && b.LastRow >= b.FirstRow).ToList();
                foreach (var b in rowBands) if (string.IsNullOrWhiteSpace(b.Name)) b.Name = (b.Source ?? "").Split('.').Last();

                foreach (var c in sh.Cells)
                {
                    if (!string.IsNullOrEmpty(c.Bc)) reasons.Add("Sheet '" + sh.Name + "': barcode cells are drawn as images and cannot be streamed.");
                    if (c.V == null) continue;
                    if (c.V.IndexOf("$count", StringComparison.OrdinalIgnoreCase) >= 0) reasons.Add("Sheet '" + sh.Name + "': {{$count}} needs the item count, which is unknown while streaming.");
                    if (c.V.Length < 2 || c.V[0] != '=') continue;

                    foreach (Match m in Tok.Matches(c.V))
                    {
                        string g1 = m.Groups[1].Value, g3 = m.Groups[3].Success ? m.Groups[3].Value : null;
                        if (g3 == null)
                        {
                            // [Field]: its cell must be in the same row or above inside the enclosing band
                            var band = rowBands.Where(b => c.R >= b.FirstRow && c.R <= b.LastRow).OrderBy(b => b.LastRow - b.FirstRow).FirstOrDefault();
                            if (band == null) continue;
                            var rx = new Regex(@"^\s*\{\{\s*" + Regex.Escape(g1) + @"\s*(\|.*)?\}\}\s*$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                            var later = sh.Cells.Any(x => x.R > c.R && x.R <= band.LastRow && x.V != null && rx.IsMatch(x.V))
                                        && !sh.Cells.Any(x => x.R <= c.R && x.R >= band.FirstRow && x.V != null && rx.IsMatch(x.V));
                            if (later) reasons.Add("Sheet '" + sh.Name + "' " + CellRef.A1(c.R, c.C) + ": [" + g1 + "] refers to a cell in a later row.");
                            continue;
                        }
                        // [Band.Field] / [Band:E]: the band must end above this cell
                        var colBand = sh.Bands.Any(b => string.Equals(b.Orientation, "columns", StringComparison.OrdinalIgnoreCase) && string.Equals(b.Name, g1, StringComparison.OrdinalIgnoreCase));
                        if (colBand) continue;
                        var target = rowBands.FirstOrDefault(b => string.Equals(b.Name, g1, StringComparison.OrdinalIgnoreCase));
                        if (target == null) continue;
                        if (target.LastRow >= c.R)
                            reasons.Add("Sheet '" + sh.Name + "' " + CellRef.A1(c.R, c.C) + ": formula uses band '" + g1 + "' that is not completely above it (totals placed above or inside the band need the full data).");
                    }
                }
                if (sh.FreezeRows > 0 && rowBands.Any(b => b.FirstRow <= sh.FreezeRows))
                    reasons.Add("Sheet '" + sh.Name + "': frozen rows must be above every band.");
            }
            return reasons.Distinct().ToList();
        }
    }
}
