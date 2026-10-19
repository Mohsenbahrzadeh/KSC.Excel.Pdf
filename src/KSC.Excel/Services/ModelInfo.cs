using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json.Linq;

namespace KSC.Excel.Services
{
    /// <summary>A node of the field tree the designer shows (text / number / date / bool / object / list).</summary>
    public sealed class FieldNode
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public List<FieldNode> Children { get; set; }
    }

    /// <summary>Shape + sample data of a data model that templates can be bound to.</summary>
    public sealed class ModelInfo
    {
        public string Name { get; set; }
        public List<FieldNode> Fields { get; set; } = new List<FieldNode>();
        public JToken Sample { get; set; }

        const int MaxDepth = 4;

        public static ModelInfo FromType(Type type, string name = null)
        {
            var info = new ModelInfo { Name = name ?? type.Name };
            var rootType = ElementTypeOrSelf(type);
            info.Fields = Describe(rootType, 0, new HashSet<Type>());
            var sample = SampleFor(type, "", 0, 1, new HashSet<Type>());
            info.Sample = sample;
            return info;
        }

        public static ModelInfo FromJson(string name, string json)
        {
            var token = JToken.Parse(json);
            return new ModelInfo { Name = name, Sample = token, Fields = Infer(token) };
        }

        /// <summary>Infers a field tree from sample JSON (used for models registered from JSON).</summary>
        public static List<FieldNode> Infer(JToken t)
        {
            var list = new List<FieldNode>();
            var obj = t as JObject ?? (t is JArray a ? a.FirstOrDefault() as JObject : null);
            if (obj == null) return list;
            foreach (var p in obj.Properties())
            {
                var v = p.Value;
                if (v is JArray arr) list.Add(new FieldNode { Name = p.Name, Type = "list", Children = Infer(arr) });
                else if (v is JObject) list.Add(new FieldNode { Name = p.Name, Type = "object", Children = Infer(v) });
                else list.Add(new FieldNode { Name = p.Name, Type = JsonType(v) });
            }
            return list;
        }

        static string JsonType(JToken v)
        {
            switch (v.Type)
            {
                case JTokenType.Integer: case JTokenType.Float: return "number";
                case JTokenType.Boolean: return "bool";
                case JTokenType.Date: return "date";
                case JTokenType.String:
                    DateTime d;
                    return DateTime.TryParse((string)v, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out d) && ((string)v).Length >= 8 && ((string)v).Contains("-") ? "date" : "text";
                default: return "text";
            }
        }

        static Type ElementTypeOrSelf(Type t)
        {
            var e = ElementType(t);
            return e ?? t;
        }

        static Type ElementType(Type t)
        {
            if (t == typeof(string) || t == typeof(byte[])) return null;
            if (t.IsArray) return t.GetElementType();
            var ie = t.IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>) ? t : t.GetInterfaces().FirstOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>));
            return ie?.GetGenericArguments()[0];
        }

        static string Kind(Type t)
        {
            t = Nullable.GetUnderlyingType(t) ?? t;
            if (t == typeof(string) || t == typeof(char) || t == typeof(Guid) || t.IsEnum) return "text";
            if (t == typeof(bool)) return "bool";
            if (t == typeof(DateTime) || t == typeof(DateTimeOffset)) return "date";
            if (t == typeof(byte[])) return "image";
            if (t.IsPrimitive || t == typeof(decimal)) return "number";
            return null;
        }

        static IEnumerable<PropertyInfo> Props(Type t)
            => t.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(p => p.CanRead && p.GetIndexParameters().Length == 0);

        static List<FieldNode> Describe(Type t, int depth, HashSet<Type> path)
        {
            var list = new List<FieldNode>();
            if (depth >= MaxDepth || !path.Add(t)) return list;
            foreach (var p in Props(t))
            {
                var kind = Kind(p.PropertyType);
                if (kind != null) { list.Add(new FieldNode { Name = p.Name, Type = kind }); continue; }
                var el = ElementType(p.PropertyType);
                if (el != null)
                {
                    var ek = Kind(el);
                    list.Add(new FieldNode { Name = p.Name, Type = "list", Children = ek != null ? new List<FieldNode>() : Describe(el, depth + 1, path) });
                }
                else list.Add(new FieldNode { Name = p.Name, Type = "object", Children = Describe(p.PropertyType, depth + 1, path) });
            }
            path.Remove(t);
            return list;
        }

        // ---------- sample data ----------
        static JToken SampleFor(Type t, string name, int depth, int seed, HashSet<Type> path)
        {
            var el = ElementType(t);
            if (el != null)
            {
                var arr = new JArray();
                for (int i = 1; i <= 3; i++) arr.Add(SampleFor(el, name, depth + 1, i, path));
                return arr;
            }
            var kind = Kind(t);
            if (kind != null) return Scalar(Nullable.GetUnderlyingType(t) ?? t, kind, name, seed);
            if (depth >= MaxDepth || !path.Add(t)) return JValue.CreateNull();
            var o = new JObject();
            foreach (var p in Props(t)) o[p.Name] = SampleFor(p.PropertyType, p.Name, depth + 1, seed, path);
            path.Remove(t);
            return o;
        }

        static JToken Scalar(Type t, string kind, string name, int seed)
        {
            switch (kind)
            {
                case "text": return t.IsEnum ? new JValue(Enum.GetNames(t).FirstOrDefault() ?? "") : new JValue((string.IsNullOrEmpty(name) ? "Text" : name) + " " + seed);
                case "bool": return new JValue(seed % 2 == 1);
                case "date": return new JValue(new DateTime(2025, 1, 1).AddDays(seed * 17));
                case "image": return JValue.CreateNull();
                default:
                    if (t == typeof(decimal) || t == typeof(double) || t == typeof(float)) return new JValue((double)(seed * 1250 + 100) + 0.5);
                    return new JValue(seed * 10);
            }
        }
    }
}
