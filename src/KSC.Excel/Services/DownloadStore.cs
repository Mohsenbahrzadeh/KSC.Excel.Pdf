using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Security.Cryptography;

namespace KSC.Excel.Services
{
    /// <summary>Holds rendered files for short-lived, unguessable download links (bound to the signed-in user, if any).</summary>
    public sealed class DownloadStore
    {
        public sealed class Item
        {
            public byte[] Bytes { get; set; }
            public string FileName { get; set; }
            public string User { get; set; }
            /// <summary>MIME type sent to the browser (default: xlsx).</summary>
            public string ContentType { get; set; } = ExcelContent.Xlsx;
            /// <summary>True = shown in the browser (inline PDF viewer), false = download.</summary>
            public bool Inline { get; set; }
            public DateTime Expires { get; set; }
        }

        const int MaxItems = 64;
        readonly ConcurrentDictionary<string, Item> _items = new ConcurrentDictionary<string, Item>();

        public string Add(byte[] bytes, string fileName, string user, TimeSpan? ttl = null)
            => Add(bytes, fileName, user, ExcelContent.Xlsx, false, ttl);

        public string Add(byte[] bytes, string fileName, string user, string contentType, bool inline, TimeSpan? ttl = null)
        {
            Purge();
            if (_items.Count >= MaxItems)
            {
                var oldest = _items.OrderBy(kv => kv.Value.Expires).First().Key;
                Item removed;
                _items.TryRemove(oldest, out removed);
            }
            var buf = new byte[18];
            using (var rng = RandomNumberGenerator.Create()) rng.GetBytes(buf);
            var token = Convert.ToBase64String(buf).Replace('+', '-').Replace('/', '_');
            _items[token] = new Item { Bytes = bytes, FileName = fileName, User = user, ContentType = contentType ?? ExcelContent.Xlsx, Inline = inline, Expires = DateTime.UtcNow + (ttl ?? TimeSpan.FromMinutes(15)) };
            return token;
        }

        /// <summary>Returns the file when the token is valid and belongs to <paramref name="user"/>; links stay valid until they expire (re-downloads work).</summary>
        public Item Take(string token, string user)
        {
            Purge();
            Item item;
            if (string.IsNullOrEmpty(token) || !_items.TryGetValue(token, out item)) return null;
            if (item.User != null && !string.Equals(item.User, user, StringComparison.OrdinalIgnoreCase)) return null;
            return item;
        }

        void Purge()
        {
            var now = DateTime.UtcNow;
            foreach (var kv in _items)
                if (kv.Value.Expires < now) { Item removed; _items.TryRemove(kv.Key, out removed); }
        }
    }
}
