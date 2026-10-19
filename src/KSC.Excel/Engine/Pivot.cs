using System;
using System.Collections.Generic;
using System.Linq;

namespace KSC.Excel.Engine
{
    public sealed class PivotColumn { public object Key { get; set; } public double? Total { get; set; } }
    public sealed class PivotCell { public double? Value { get; set; } public int Count { get; set; } }

    public sealed class PivotRow
    {
        public object Key { get; set; }
        /// <summary>One cell per <see cref="PivotTable.Columns"/> entry, in the same order (Value is null when no data).</summary>
        public List<PivotCell> Cells { get; } = new List<PivotCell>();
        public double? Total { get; set; }
    }

    /// <summary>
    /// Cross-tab result: bind <c>Columns</c> to a horizontal (column) band for the header and <c>Rows</c> to a row band whose
    /// <c>Cells</c> is a column band. Also available as a template DataSet, so no code is needed.
    /// </summary>
    public sealed class PivotTable
    {
        public List<PivotColumn> Columns { get; } = new List<PivotColumn>();
        public List<PivotRow> Rows { get; } = new List<PivotRow>();
        public double? GrandTotal { get; set; }
    }

    public static class Pivot
    {
        /// <summary>
        /// Builds a cross-tab from any collection (POCOs, dictionaries, JSON, DataTable...). Paths are relative to each item.
        /// <paramref name="agg"/>: sum, count, avg, min, max. Row/column keys are sorted naturally (numbers, dates, text).
        /// </summary>
        public static PivotTable Build(object source, string rowPath, string colPath, string valuePath, string agg = "sum")
        {
            var items = DataAccess.Items(source).ToList();
            agg = (agg ?? "sum").ToLowerInvariant();
            var rowKeys = new Dictionary<string, object>(StringComparer.Ordinal);
            var colKeys = new Dictionary<string, object>(StringComparer.Ordinal);
            var cells = new Dictionary<string, List<double>>(StringComparer.Ordinal);
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (var it in items)
            {
                var rk = DataAccess.Resolve(new DataScope(it), rowPath);
                var ck = DataAccess.Resolve(new DataScope(it), colPath);
                var rs = TextEvaluator.ToDisplay(rk); var cs = TextEvaluator.ToDisplay(ck);
                if (!rowKeys.ContainsKey(rs)) rowKeys[rs] = rk;
                if (!colKeys.ContainsKey(cs)) colKeys[cs] = ck;
                var key = rs + "\u0001" + cs;
                List<double> list;
                if (!cells.TryGetValue(key, out list)) cells[key] = list = new List<double>();
                int n; counts.TryGetValue(key, out n); counts[key] = n + 1;
                double d;
                if (agg == "count") list.Add(1);
                else if (DataAccess.TryNumber(DataAccess.Resolve(new DataScope(it), valuePath), out d)) list.Add(d);
            }

            var table = new PivotTable();
            var sortedCols = colKeys.OrderBy(k => k.Value, SortKeyComparer.Instance).ToList();
            var sortedRows = rowKeys.OrderBy(k => k.Value, SortKeyComparer.Instance).ToList();
            foreach (var c in sortedCols)
            {
                var all = sortedRows.SelectMany(r => Get(cells, r.Key, c.Key)).ToList();
                table.Columns.Add(new PivotColumn { Key = c.Value, Total = Aggregate(all, agg) });
            }
            foreach (var r in sortedRows)
            {
                var row = new PivotRow { Key = r.Value };
                var rowAll = new List<double>();
                foreach (var c in sortedCols)
                {
                    var vals = Get(cells, r.Key, c.Key);
                    int cnt; counts.TryGetValue(r.Key + "\u0001" + c.Key, out cnt);
                    row.Cells.Add(new PivotCell { Value = Aggregate(vals, agg), Count = cnt });
                    rowAll.AddRange(vals);
                }
                row.Total = Aggregate(rowAll, agg);
                table.Rows.Add(row);
            }
            table.GrandTotal = Aggregate(cells.Values.SelectMany(x => x).ToList(), agg);
            return table;
        }

        static List<double> Get(Dictionary<string, List<double>> cells, string r, string c)
        {
            List<double> l;
            return cells.TryGetValue(r + "\u0001" + c, out l) ? l : new List<double>();
        }

        static double? Aggregate(List<double> v, string agg)
        {
            if (v.Count == 0) return null;
            switch (agg)
            {
                case "count": return v.Count;
                case "avg": case "average": return v.Average();
                case "min": return v.Min();
                case "max": return v.Max();
                default: return v.Sum();
            }
        }
    }

    internal sealed class SortKeyComparer : IComparer<object>
    {
        public static readonly SortKeyComparer Instance = new SortKeyComparer();

        public int Compare(object a, object b)
        {
            if (a == null) return b == null ? 0 : -1;
            if (b == null) return 1;
            double x, y;
            if (DataAccess.TryNumber(a, out x) && DataAccess.TryNumber(b, out y)) return x.CompareTo(y);
            if (a is DateTime && b is DateTime) return ((DateTime)a).CompareTo((DateTime)b);
            return string.Compare(TextEvaluator.ToDisplay(a), TextEvaluator.ToDisplay(b), StringComparison.CurrentCultureIgnoreCase);
        }
    }
}
