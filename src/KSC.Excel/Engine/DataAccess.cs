using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Newtonsoft.Json.Linq;

namespace KSC.Excel.Engine
{
    /// <summary>
    /// One level of data context. The root scope wraps the data object; every band occurrence creates a child scope.
    /// </summary>
    /// <summary>Marks a padding row added by BandDef.MinRows.</summary>
    internal sealed class PadItem { }

    public sealed class DataScope
    {
        public object Item;
        public DataScope Parent;
        /// <summary>1-based position inside the band.</summary>
        public int Index = 1;
        public int Count = 1;
        public Model.BandDef Band;
        /// <summary>All items of the band this scope belongs to (null while streaming): used by RUNSUM / PCT / PREV / RANK...</summary>
        public IList<object> Siblings;
        /// <summary>Named virtual members (template DataSets such as pivots), visible from this scope and its children.</summary>
        public Dictionary<string, object> Extras;
        /// <summary>Report parameters (root scope).</summary>
        public Dictionary<string, object> Params;
        /// <summary>Band name -> first/last output row generated inside this scope.</summary>
        internal Dictionary<string, int[]> Runs { get { return _runs ?? (_runs = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase)); } }
        Dictionary<string, int[]> _runs;
        internal bool HasRuns => _runs != null;
        /// <summary>Template row -> output row emitted inside this scope.</summary>
        internal Dictionary<int, int> RowOf { get { return _rowOf ?? (_rowOf = new Dictionary<int, int>()); } }
        Dictionary<int, int> _rowOf;

        public DataScope(object item, DataScope parent = null) { Item = item; Parent = parent; }
    }

    /// <summary>Uniform member/collection access over POCOs, dictionaries, JSON, DataTable/DataRow and dynamic objects.</summary>
    public static class DataAccess
    {
        static readonly Func<Type, Dictionary<string, Func<object, object>>> AccessorFactory = BuildAccessors;
        static readonly ConcurrentDictionary<string, List<Seg>> PathCache = new ConcurrentDictionary<string, List<Seg>>();
        static readonly ConcurrentDictionary<Type, Dictionary<string, Func<object, object>>> Cache
            = new ConcurrentDictionary<Type, Dictionary<string, Func<object, object>>>();

        /// <summary>Converts JSON tokens / enums / DBNull to plain CLR values.</summary>
        public static object Normalize(object v)
        {
            if (v == null || v is DBNull) return null;
            if (v is JValue jv) return jv.Type == JTokenType.Null || jv.Type == JTokenType.Undefined ? null : Normalize(jv.Value);
            if (v is Enum e) return e.ToString();
            if (v is DateTimeOffset dto) return dto.DateTime;
            if (v is Guid || v is TimeSpan || v is Uri) return v.ToString();
            if (v is char c) return c.ToString();
            return v;
        }

        public static bool TryMember(object obj, string name, out object value)
        {
            value = null;
            if (obj == null) return false;

            if (obj is JObject jo)
            {
                JToken t;
                if (jo.TryGetValue(name, StringComparison.OrdinalIgnoreCase, out t)) { value = t is JValue ? Normalize(t) : t; return true; }
                return false;
            }
            if (obj is DataRow dr)
            {
                if (dr.Table.Columns.Contains(name)) { value = Normalize(dr[name]); return true; }
                return false;
            }
            if (obj is DataRowView drv)
            {
                if (drv.Row.Table.Columns.Contains(name)) { value = Normalize(drv[name]); return true; }
                return false;
            }
            if (obj is IDictionary dict)
            {
                foreach (DictionaryEntry en in dict)
                    if (string.Equals(Convert.ToString(en.Key), name, StringComparison.OrdinalIgnoreCase)) { value = Normalize(en.Value); return true; }
                return false;
            }
            if (obj is IDictionary<string, object> gd)
            {
                foreach (var kv in gd)
                    if (string.Equals(kv.Key, name, StringComparison.OrdinalIgnoreCase)) { value = Normalize(kv.Value); return true; }
                return false;
            }
            if (obj is string) return false;

            var map = Cache.GetOrAdd(obj.GetType(), AccessorFactory);
            Func<object, object> getter;
            if (map.TryGetValue(name, out getter)) { value = Normalize(getter(obj)); return true; }
            return false;
        }

        static Dictionary<string, Func<object, object>> BuildAccessors(Type t)
        {
            var d = new Dictionary<string, Func<object, object>>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in t.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (p.GetIndexParameters().Length > 0 || !p.CanRead) continue;
                d[p.Name] = Getter(p);
            }
            foreach (var f in t.GetFields(BindingFlags.Public | BindingFlags.Instance))
            {
                var fi = f;
                if (!d.ContainsKey(f.Name)) d[f.Name] = o => fi.GetValue(o);
            }
            return d;
        }

        // Compiled delegates are ~50x faster than PropertyInfo.GetValue on big reports.
        static Func<object, object> Getter(PropertyInfo p)
        {
            try
            {
                var o = Expression.Parameter(typeof(object), "o");
                var body = Expression.Convert(Expression.Property(Expression.Convert(o, p.DeclaringType), p), typeof(object));
                return Expression.Lambda<Func<object, object>>(body, o).Compile();
            }
            catch (Exception)
            {
                return x => p.GetValue(x, null);
            }
        }

        /// <summary>Enumerates a collection value (null and scalars give an empty sequence; a JObject is a single item).</summary>
        public static IEnumerable<object> Items(object v)
        {
            if (v == null || v is string) return Enumerable.Empty<object>();
            if (v is JArray ja) return ja.Select(x => x is JValue ? Normalize(x) : (object)x).ToList();
            if (v is JObject) return new[] { v };
            if (v is DataTable dt) return dt.Rows.Cast<DataRow>().ToList();
            if (v is DataView dv) return dv.Cast<DataRowView>().ToList();
            if (v is IDictionary || v is IDictionary<string, object>) return new[] { v };
            if (v is IEnumerable en) return en.Cast<object>().Select(Normalize).ToList();
            return Enumerable.Empty<object>();
        }

        /// <summary>Like <see cref="Items"/> but never materializes a sequence (used by the streaming renderer).</summary>
        public static IEnumerable<object> ItemsLazy(object v)
        {
            if (v == null || v is string) return Enumerable.Empty<object>();
            if (v is JArray || v is JObject || v is DataTable || v is DataView || v is IDictionary || v is IDictionary<string, object>) return Items(v);
            var en = v as IEnumerable;
            return en != null ? LazyNormalize(en) : Enumerable.Empty<object>();
        }

        static IEnumerable<object> LazyNormalize(IEnumerable en) { foreach (var x in en) yield return Normalize(x); }

        public static bool IsCollection(object v)
            => v != null && !(v is string) && !(v is JObject) && !(v is IDictionary) && !(v is IDictionary<string, object>) && (v is IEnumerable || v is DataTable);

        /// <summary>Resolves a dotted path such as "Order.Customer.Name" or "Items[0].Price" starting from a scope.</summary>
        public static object Resolve(DataScope scope, string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            path = path.Trim();

            List<Seg> segs;
            if (!PathCache.TryGetValue(path, out segs))
            {
                segs = SplitPath(path);
                if (PathCache.Count < 5000) PathCache[path] = segs;
            }
            if (segs.Count == 0) return null;
            var first = segs[0];

            object cur = null;
            bool found = false;
            if (first.Name.StartsWith("$", StringComparison.Ordinal))
            {
                cur = ResolveSpecial(scope, first.Name, segs);
                if (segs.Count == 1) return cur;
                found = true;
            }
            else
            {
                // first segment: look through the scope chain (current item first, then parents)
                for (var s = scope; s != null && !found; s = s.Parent)
                {
                    if (first.Name == "." || first.Name == "this") { cur = s.Item; found = true; break; }
                    if (s.Extras != null && s.Extras.TryGetValue(first.Name, out cur)) { found = true; break; }
                    found = TryMember(s.Item, first.Name, out cur);
                }
            }
            if (!found) return null;
            cur = ApplyIndexes(cur, first);
            for (int i = 1; i < segs.Count; i++)
            {
                if (cur == null) return null;
                object next;
                var seg = segs[i];
                if (!TryMember(cur, seg.Name, out next))
                {
                    if (seg.Name.Equals("Count", StringComparison.OrdinalIgnoreCase) && IsCollection(cur)) next = Items(cur).Count();
                    else if (seg.Name.Equals("Length", StringComparison.OrdinalIgnoreCase) && cur is string cs) next = cs.Length;
                    else return null;
                }
                cur = ApplyIndexes(next, seg);
            }
            return cur is JToken jt && jt is JValue ? Normalize(jt) : cur;
        }

        static object ResolveSpecial(DataScope scope, string name, List<Seg> segs)
        {
            switch (name.ToLowerInvariant())
            {
                case "$index": return (double)scope.Index;
                case "$count": return (double)scope.Count;
                case "$pad": return scope.Item is PadItem;
                case "$today": return DateTime.Today;
                case "$now": return DateTime.Now;
                case "$root": { var s = scope; while (s.Parent != null) s = s.Parent; return s.Item; }
                case "$parent": return scope.Parent?.Item;
                case "$item": return scope.Item;
                case "$p": case "$param": case "$params": { for (var s = scope; s != null; s = s.Parent) if (s.Params != null) return s.Params; return null; }
                default: return null;
            }
        }

        sealed class Seg { public string Name; public List<int> Indexes = new List<int>(); }

        static List<Seg> SplitPath(string path)
        {
            var list = new List<Seg>();
            foreach (var raw in path.Split('.'))
            {
                var p = raw.Trim();
                if (p.Length == 0) continue;
                var seg = new Seg();
                int b = p.IndexOf('[');
                if (b < 0) seg.Name = p;
                else
                {
                    seg.Name = p.Substring(0, b);
                    foreach (var part in p.Substring(b).Split(new[] { '[' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        int idx;
                        if (int.TryParse(part.TrimEnd(']'), out idx)) seg.Indexes.Add(idx);
                    }
                }
                list.Add(seg);
            }
            return list;
        }

        static object ApplyIndexes(object v, Seg seg)
        {
            foreach (var i in seg.Indexes)
            {
                var arr = Items(v).ToList();
                v = i >= 0 && i < arr.Count ? arr[i] : null;
            }
            return v;
        }

        /// <summary>Converts any value to a double when it is numeric (including numeric strings are NOT converted).</summary>
        public static bool TryNumber(object v, out double d)
        {
            d = 0;
            switch (v)
            {
                case null: return false;
                case double x: d = x; return true;
                case float f: d = f; return true;
                case decimal m: d = (double)m; return true;
                case int i: d = i; return true;
                case long l: d = l; return true;
                case short s: d = s; return true;
                case byte b: d = b; return true;
                case uint ui: d = ui; return true;
                case ulong ul: d = ul; return true;
                case ushort us: d = us; return true;
                case sbyte sb: d = sb; return true;
                default: return false;
            }
        }
    }
}
