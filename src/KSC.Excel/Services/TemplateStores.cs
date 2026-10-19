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

    /// <summary>One saved version of a template (the state BEFORE a later save replaced it).</summary>
    public sealed class TemplateVersionInfo
    {
        /// <summary>Opaque id (a UTC time stamp); pass it to <see cref="IVersionedTemplateStore.GetVersion"/>.</summary>
        public string Id { get; set; }
        public DateTime Saved { get; set; }
        public long Size { get; set; }
        public string Title { get; set; }
    }

    /// <summary>Optional capability of a store: keeps earlier versions so a template can be compared and restored.</summary>
    public interface IVersionedTemplateStore : ITemplateStore
    {
        /// <summary>Newest first.</summary>
        IList<TemplateVersionInfo> ListVersions(string name);
        /// <summary>Returns null when the version does not exist.</summary>
        ExcelTemplate GetVersion(string name, string versionId);
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
    public sealed class FileTemplateStore : IVersionedTemplateStore
    {
        /// <summary>How many earlier versions are kept per template (oldest are deleted). Set 0 to disable history.</summary>
        public int MaxVersions { get; set; } = 50;

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
                var json = template.ToJson(true);
                if (File.Exists(p) && MaxVersions > 0) KeepVersion(template.Name, p, json);
                var tmp = p + ".tmp";
                File.WriteAllText(tmp, json, new UTF8Encoding(false));
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
                if (MaxVersions > 0) KeepVersion(name, p, null);   // a deleted template can still be restored from its history
                File.Delete(p);
                return true;
            }
        }

        // ---- history: {dir}/.history/{name}/{yyyyMMddHHmmssfff}.kxl.json (never listed as a template: a dot-folder)
        static readonly Regex VersionId = new Regex(@"^\d{17}$", RegexOptions.Compiled);

        string HistoryDir(string name) { TemplateNames.Ensure(name); return Path.Combine(_dir, ".history", name); }

        void KeepVersion(string name, string currentPath, string newJson)
        {
            var old = File.ReadAllText(currentPath, Encoding.UTF8);
            if (newJson != null && old == newJson) return;      // nothing changed
            var dir = HistoryDir(name);
            System.IO.Directory.CreateDirectory(dir);
            var id = DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", System.Globalization.CultureInfo.InvariantCulture);
            var target = Path.Combine(dir, id + Suffix);
            while (File.Exists(target)) { id = (long.Parse(id) + 1).ToString(); target = Path.Combine(dir, id + Suffix); }
            File.WriteAllText(target, old, new UTF8Encoding(false));
            var files = System.IO.Directory.GetFiles(dir, "*" + Suffix).OrderByDescending(f => f, StringComparer.Ordinal).Skip(MaxVersions).ToList();
            foreach (var f in files) File.Delete(f);
        }

        public IList<TemplateVersionInfo> ListVersions(string name)
        {
            var dir = HistoryDir(name);
            var list = new List<TemplateVersionInfo>();
            lock (_lock)
            {
                if (!System.IO.Directory.Exists(dir)) return list;
                foreach (var f in System.IO.Directory.GetFiles(dir, "*" + Suffix).OrderByDescending(x => x, StringComparer.Ordinal))
                {
                    var id = Path.GetFileName(f); id = id.Substring(0, id.Length - Suffix.Length);
                    if (!VersionId.IsMatch(id)) continue;
                    string title = null;
                    try { title = ExcelTemplate.LoadFile(f).Title; } catch (Exception) { }
                    list.Add(new TemplateVersionInfo { Id = id, Saved = DateTime.ParseExact(id, "yyyyMMddHHmmssfff", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal), Size = new FileInfo(f).Length, Title = title });
                }
            }
            return list;
        }

        public ExcelTemplate GetVersion(string name, string versionId)
        {
            if (!VersionId.IsMatch(versionId ?? "")) throw new ArgumentException("Invalid version id.");
            var f = Path.Combine(HistoryDir(name), versionId + Suffix);
            lock (_lock)
            {
                if (!File.Exists(f)) return null;
                var t = ExcelTemplate.LoadFile(f); t.Name = name; return t;
            }
        }
    }

    /// <summary>In-memory store (tests, demos).</summary>
    public sealed class MemoryTemplateStore : IVersionedTemplateStore
    {
        readonly ConcurrentDictionary<string, List<Tuple<string, DateTime>>> _history = new ConcurrentDictionary<string, List<Tuple<string, DateTime>>>(StringComparer.OrdinalIgnoreCase);
        int _seq;

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
            var json = template.ToJson();
            Tuple<string, DateTime> prev;
            if (_items.TryGetValue(template.Name, out prev) && prev.Item1 != json) Keep(template.Name, prev);
            _items[template.Name] = Tuple.Create(json, DateTime.UtcNow);
        }

        void Keep(string name, Tuple<string, DateTime> v)
        {
            var list = _history.GetOrAdd(name, _ => new List<Tuple<string, DateTime>>());
            lock (list) { list.Add(v); if (list.Count > 50) list.RemoveAt(0); }
        }

        public IList<TemplateVersionInfo> ListVersions(string name)
        {
            List<Tuple<string, DateTime>> list;
            if (!_history.TryGetValue(name ?? "", out list)) return new List<TemplateVersionInfo>();
            lock (list)
                return list.Select((v, i) => new TemplateVersionInfo { Id = (i + 1).ToString("00000000000000000"), Saved = v.Item2, Size = v.Item1.Length, Title = ExcelTemplate.FromJson(v.Item1).Title }).Reverse().ToList();
        }

        public ExcelTemplate GetVersion(string name, string versionId)
        {
            List<Tuple<string, DateTime>> list; int n;
            if (!_history.TryGetValue(name ?? "", out list) || !int.TryParse(versionId, out n)) return null;
            lock (list) return n >= 1 && n <= list.Count ? ExcelTemplate.FromJson(list[n - 1].Item1) : null;
        }

        public bool Delete(string name)
        {
            Tuple<string, DateTime> v;
            return _items.TryRemove(name ?? "", out v);
        }
    }
}
