using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace KSC.Excel.Engine
{
    /// <summary>
    /// Runs the read-only queries of SQL data sets. Values are always bound as parameters (@Name = report parameter),
    /// never concatenated. Defence in depth: single SELECT/WITH statement only; the real protection is a read-only DB login.
    /// </summary>
    public static class SqlDataSource
    {
        static readonly Regex ParamRx = new Regex(@"(?<![@\w])@([A-Za-z_][\w]*)", RegexOptions.Compiled);
        static readonly Regex Forbidden = new Regex(@"\b(insert|update|delete|drop|alter|create|truncate|exec|execute|merge|grant|revoke|deny|backup|restore|shutdown|kill|openrowset|opendatasource|bulk|into|xp_\w*|sp_\w*)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>Removes comments and string literals so keywords inside them do not matter.</summary>
        static string Strip(string sql)
        {
            var sb = new StringBuilder(); int i = 0;
            while (i < sql.Length)
            {
                char c = sql[i];
                if (c == '-' && i + 1 < sql.Length && sql[i + 1] == '-') { while (i < sql.Length && sql[i] != '\n') i++; sb.Append(' '); continue; }
                if (c == '/' && i + 1 < sql.Length && sql[i + 1] == '*') { int e = sql.IndexOf("*/", i + 2, StringComparison.Ordinal); i = e < 0 ? sql.Length : e + 2; sb.Append(' '); continue; }
                if (c == '\'') { i++; while (i < sql.Length) { if (sql[i] == '\'') { if (i + 1 < sql.Length && sql[i + 1] == '\'') { i += 2; continue; } break; } i++; } i++; sb.Append("''"); continue; }
                sb.Append(c); i++;
            }
            return sb.ToString();
        }

        /// <summary>Throws <see cref="ExcelRenderException"/> unless the text is one read-only SELECT/WITH statement.</summary>
        public static void Validate(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) throw new ExcelRenderException("The SQL text is empty.");
            var s = Strip(sql).Trim().TrimEnd(';').Trim();
            if (s.IndexOf(';') >= 0) throw new ExcelRenderException("Only a single SQL statement is allowed.");
            if (!Regex.IsMatch(s, @"^(select|with)\b", RegexOptions.IgnoreCase)) throw new ExcelRenderException("Only SELECT / WITH queries are allowed in a data set.");
            var m = Forbidden.Match(s);
            if (m.Success) throw new ExcelRenderException("The keyword '" + m.Value + "' is not allowed in a data set query (read-only).");
        }

        public static IList<string> ParameterNames(string sql)
        {
            return ParamRx.Matches(Strip(sql)).Cast<Match>().Select(m => m.Groups[1].Value).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        public static List<Dictionary<string, object>> Query(IDbConnection connection, string sql, IDictionary<string, object> parameters, int timeoutSeconds, int maxRows, Action<string> warn, bool truncateAtMax = false)
        {
            Validate(sql);
            var rows = new List<Dictionary<string, object>>();
            using (connection)
            {
                if (connection.State != ConnectionState.Open) connection.Open();
                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandTimeout = timeoutSeconds;
                    foreach (var name in ParameterNames(sql))
                    {
                        object v = null;
                        if (parameters == null || !TryGet(parameters, name, out v)) warn("SQL parameter @" + name + " has no report parameter; NULL was used.");
                        var p = cmd.CreateParameter();
                        p.ParameterName = "@" + name;
                        p.Value = v == null ? (object)DBNull.Value : v;
                        cmd.Parameters.Add(p);
                    }
                    using (var rd = cmd.ExecuteReader())
                    {
                        var cols = new string[rd.FieldCount];
                        for (int i = 0; i < cols.Length; i++) cols[i] = string.IsNullOrEmpty(rd.GetName(i)) ? "Column" + (i + 1) : rd.GetName(i);
                        while (rd.Read())
                        {
                            if (rows.Count >= maxRows && truncateAtMax) break;
                            if (rows.Count >= maxRows) throw new ExcelRenderException("The SQL data set returned more than " + maxRows + " rows. Add a filter or raise KscExcelOptions.SqlMaxRows.");
                            var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                            for (int i = 0; i < cols.Length; i++) row[cols[i]] = DataAccess.Normalize(rd.GetValue(i));
                            rows.Add(row);
                        }
                    }
                }
            }
            return rows;
        }

        /// <summary>Lazy variant: rows are read while the report is rendered; every enumeration re-runs the query.</summary>
        public static IEnumerable<object> QueryLazy(Func<IDbConnection> connection, string sql, IDictionary<string, object> parameters, int timeoutSeconds, int maxRows, Action<string> warn)
        {
            Validate(sql);
            return new LazyRows(connection, sql, parameters, timeoutSeconds, maxRows, warn);
        }

        sealed class LazyRows : IEnumerable<object>
        {
            readonly Func<IDbConnection> _c; readonly string _sql; readonly IDictionary<string, object> _p; readonly int _t, _max; readonly Action<string> _warn;
            public LazyRows(Func<IDbConnection> c, string sql, IDictionary<string, object> p, int t, int max, Action<string> warn) { _c = c; _sql = sql; _p = p; _t = t; _max = max; _warn = warn; }

            public IEnumerator<object> GetEnumerator()
            {
                using (var con = _c())
                {
                    if (con.State != ConnectionState.Open) con.Open();
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.CommandText = _sql; cmd.CommandType = CommandType.Text; cmd.CommandTimeout = _t;
                        foreach (var name in ParameterNames(_sql))
                        {
                            object v = null;
                            if (_p == null || !TryGet(_p, name, out v)) _warn("SQL parameter @" + name + " has no report parameter; NULL was used.");
                            var prm = cmd.CreateParameter(); prm.ParameterName = "@" + name; prm.Value = v == null ? (object)DBNull.Value : v; cmd.Parameters.Add(prm);
                        }
                        using (var rd = cmd.ExecuteReader())
                        {
                            var cols = new string[rd.FieldCount];
                            for (int i = 0; i < cols.Length; i++) cols[i] = string.IsNullOrEmpty(rd.GetName(i)) ? "Column" + (i + 1) : rd.GetName(i);
                            int n = 0;
                            while (rd.Read())
                            {
                                if (++n > _max) throw new ExcelRenderException("The SQL data set returned more than " + _max + " rows. Add a filter or raise KscExcelOptions.SqlMaxRows.");
                                var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                                for (int i = 0; i < cols.Length; i++) row[cols[i]] = DataAccess.Normalize(rd.GetValue(i));
                                yield return row;
                            }
                        }
                    }
                }
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        static bool TryGet(IDictionary<string, object> d, string name, out object v)
        {
            foreach (var kv in d) if (string.Equals(kv.Key, name, StringComparison.OrdinalIgnoreCase)) { v = kv.Value; return true; }
            v = null; return false;
        }

        /// <summary>Column names and .NET-ish types of the first rows (used by the designer "test query").</summary>
        public static string TypeName(object v)
        {
            if (v == null) return "text";
            if (v is DateTime) return "date";
            if (v is bool) return "bool";
            if (v is byte[]) return "image";
            double d;
            return DataAccess.TryNumber(v, out d) ? "number" : "text";
        }
    }
}
