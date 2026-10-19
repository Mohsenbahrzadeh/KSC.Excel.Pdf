using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using KSC.Excel.Model;

namespace KSC.Excel.Services
{
    public sealed class TemplateInfo
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public DateTime Modified { get; set; }
    }

    /// <summary>Where templates are stored. Implement it to keep templates in a database or blob storage.</summary>
    public interface ITemplateStore
    {
        IList<TemplateInfo> List();
        /// <summary>Returns null when the template does not exist.</summary>
        ExcelTemplate Get(string name);
        void Save(ExcelTemplate template);
        bool Delete(string name);
    }

    public static class TemplateNames
    {
        static readonly Regex Valid = new Regex(@"^[\p{L}\p{N}][\p{L}\p{N} _\-\.]{0,98}$", RegexOptions.Compiled);

        /// <summary>Letters (any language), digits, space, _ - . ; must start with a letter/digit; max 99 chars. Blocks path traversal.</summary>
        public static bool IsValid(string name) => !string.IsNullOrEmpty(name) && Valid.IsMatch(name) && !name.Contains("..") && !name.EndsWith(".", StringComparison.Ordinal) && !name.EndsWith(" ", StringComparison.Ordinal);

        public static void Ensure(string name)
        {
            if (!IsValid(name)) throw new ArgumentException("Invalid template name '" + name + "'. Use letters, digits, space, '_', '-' or '.'.", nameof(name));
        }
    }

    /// <summary>Stores each template as {name}.kxl.json in a folder.</summary>
    public sealed class FileTemplateStore : ITemplateStore
    {
        const string Suffix = ".kxl.json";
        readonly string _dir;
        readonly object _lock = new object();

        public FileTemplateStore(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("directory is required", nameof(directory));
            _dir = Path.GetFullPath(directory);
        }

        public string Directory => _dir;

        string PathOf(string name)
        {
            TemplateNames.Ensure(name);
            var p = Path.GetFullPath(Path.Combine(_dir, name + Suffix));
            if (!p.StartsWith(_dir, StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Invalid template name.");
            return p;
        }

        public IList<TemplateInfo> List()
        {
            if (!System.IO.Directory.Exists(_dir)) return new List<TemplateInfo>();
            var list = new List<TemplateInfo>();
            foreach (var f in System.IO.Directory.GetFiles(_dir, "*" + Suffix))
            {
                var name = Path.GetFileName(f);
                name = name.Substring(0, name.Length - Suffix.Length);
                if (!TemplateNames.IsValid(name)) continue;
                string title = name;
                try { title = ExcelTemplate.LoadFile(f).Title ?? name; } catch (Exception) { /* unreadable file: still list it */ }
                list.Add(new TemplateInfo { Name = name, Title = title, Modified = File.GetLastWriteTimeUtc(f) });
            }
            return list.OrderBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        public ExcelTemplate Get(string name)
        {
            var p = PathOf(name);
            lock (_lock)
            {
                if (!File.Exists(p)) return null;
                var t = ExcelTemplate.LoadFile(p);
                t.Name = name;
                return t;
            }
        }

        public void Save(ExcelTemplate template)
        {
            if (template == null) throw new ArgumentNullException(nameof(template));
            var p = PathOf(template.Name);
            lock (_lock)
            {
                System.IO.Directory.CreateDirectory(_dir);
                var tmp = p + ".tmp";
                File.WriteAllText(tmp, template.ToJson(true), new UTF8Encoding(false));
                if (File.Exists(p)) File.Delete(p);
                File.Move(tmp, p);
            }
        }

        public bool Delete(string name)
        {
            var p = PathOf(name);
            lock (_lock)
            {
                if (!File.Exists(p)) return false;
                File.Delete(p);
                return true;
            }
        }
    }

    /// <summary>In-memory store (tests, demos).</summary>
    public sealed class MemoryTemplateStore : ITemplateStore
    {
        readonly ConcurrentDictionary<string, Tuple<string, DateTime>> _items = new ConcurrentDictionary<string, Tuple<string, DateTime>>(StringComparer.OrdinalIgnoreCase);

        public IList<TemplateInfo> List() => _items.Select(kv => new TemplateInfo { Name = kv.Key, Title = ExcelTemplate.FromJson(kv.Value.Item1).Title ?? kv.Key, Modified = kv.Value.Item2 }).OrderBy(t => t.Name).ToList();

        public ExcelTemplate Get(string name)
        {
            Tuple<string, DateTime> v;
            return _items.TryGetValue(name ?? "", out v) ? ExcelTemplate.FromJson(v.Item1) : null;
        }

        public void Save(ExcelTemplate template)
        {
            TemplateNames.Ensure(template.Name);
            _items[template.Name] = Tuple.Create(template.ToJson(), DateTime.UtcNow);
        }

        public bool Delete(string name)
        {
            Tuple<string, DateTime> v;
            return _items.TryRemove(name ?? "", out v);
        }
    }
}
