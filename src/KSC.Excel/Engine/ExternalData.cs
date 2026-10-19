using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using KSC.Excel.Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KSC.Excel.Engine
{
    /// <summary>A REST endpoint a developer allowed templates to read (host and credentials are fixed in code).</summary>
    public sealed class HttpEndpoint
    {
        public string BaseUrl;
        public IDictionary<string, string> Headers;
        public int TimeoutSeconds = 30;
        /// <summary>Largest response accepted (bytes).</summary>
        public long MaxBytes = 20 * 1024 * 1024;
    }

    /// <summary>Loads data sets of Type "json" (inline JSON or a registered REST endpoint) and "provider" (a registered function).</summary>
    internal static class ExternalData
    {
        public static object Load(DataSetDef d, RenderOptions opt, IDictionary<string, object> parameters, Action<string> warn)
        {
            if (string.Equals(d.Type, "provider", StringComparison.OrdinalIgnoreCase))
            {
                Func<IDictionary<string, object>, object> fn = null;
                if (opt.Providers == null || string.IsNullOrWhiteSpace(d.Connection) || !opt.Providers.TryGetValue(d.Connection, out fn))
                    throw new ExcelRenderException("Data set '" + d.Name + "': provider '" + d.Connection + "' is not registered (KscExcelOptions.AddProvider).");
                return DataAccess.Normalize(fn(parameters ?? new Dictionary<string, object>()));
            }

            JToken doc;
            if (!string.IsNullOrWhiteSpace(d.Json))
            {
                try { doc = ParseJson(d.Json); }
                catch (JsonException ex) { throw new ExcelRenderException("Data set '" + d.Name + "': the inline JSON is invalid (" + ex.Message + ")."); }
            }
            else
            {
                HttpEndpoint ep = null;
                if (opt.Endpoints == null || string.IsNullOrWhiteSpace(d.Connection) || !opt.Endpoints.TryGetValue(d.Connection, out ep))
                    throw new ExcelRenderException("Data set '" + d.Name + "': endpoint '" + d.Connection + "' is not registered (KscExcelOptions.AddHttpEndpoint), and no inline JSON is given.");
                doc = Fetch(d, ep, parameters);
            }

            if (!string.IsNullOrWhiteSpace(d.JsonPath))
            {
                JToken sel;
                try { sel = doc.SelectToken(d.JsonPath.Trim()); }
                catch (JsonException ex) { throw new ExcelRenderException("Data set '" + d.Name + "': invalid JSON path (" + ex.Message + ")."); }
                if (sel == null) { warn("Data set '" + d.Name + "': path '" + d.JsonPath + "' was not found in the JSON; the data set is empty."); return new JArray(); }
                doc = sel;
            }
            return DataAccess.Normalize(doc);
        }

        static JToken ParseJson(string json)
        {
            using (var r = new JsonTextReader(new StringReader(json)) { DateParseHandling = DateParseHandling.None })
                return JToken.Load(r);
        }

        static JToken Fetch(DataSetDef d, HttpEndpoint ep, IDictionary<string, object> parameters)
        {
            var rel = d.Url ?? "";
            if (parameters != null)
                foreach (var kv in parameters)
                    rel = rel.Replace("{" + kv.Key + "}", Uri.EscapeDataString(Convert.ToString(kv.Value, System.Globalization.CultureInfo.InvariantCulture) ?? ""));
            if (rel.Contains("{")) rel = System.Text.RegularExpressions.Regex.Replace(rel, @"\{[A-Za-z_]\w*\}", "");

            var baseUri = new Uri(ep.BaseUrl, UriKind.Absolute);
            Uri target;
            if (!Uri.TryCreate(baseUri, rel, out target)) throw new ExcelRenderException("Data set '" + d.Name + "': invalid URL path.");
            // the template may only choose the path: never another host / scheme (SSRF guard)
            if (!string.Equals(target.Scheme, baseUri.Scheme, StringComparison.OrdinalIgnoreCase) || !string.Equals(target.Authority, baseUri.Authority, StringComparison.OrdinalIgnoreCase))
                throw new ExcelRenderException("Data set '" + d.Name + "': the URL must stay on the registered endpoint host.");
            if (target.Scheme != Uri.UriSchemeHttps && target.Scheme != Uri.UriSchemeHttp) throw new ExcelRenderException("Data set '" + d.Name + "': only http/https endpoints are supported.");

            // run on the thread pool: no sync-context deadlock when called from classic ASP.NET
            return Task.Run(async () =>
            {
                using (var http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(1, ep.TimeoutSeconds)) })
                using (var req = new HttpRequestMessage(HttpMethod.Get, target))
                {
                    req.Headers.TryAddWithoutValidation("Accept", "application/json");
                    if (ep.Headers != null) foreach (var h in ep.Headers) req.Headers.TryAddWithoutValidation(h.Key, h.Value);
                    using (var res = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                    {
                        if (!res.IsSuccessStatusCode) throw new ExcelRenderException("Data set '" + d.Name + "': the endpoint answered " + (int)res.StatusCode + " " + res.ReasonPhrase + ".");
                        using (var stream = await res.Content.ReadAsStreamAsync().ConfigureAwait(false))
                        using (var ms = new MemoryStream())
                        {
                            var buf = new byte[81920]; int n;
                            while ((n = await stream.ReadAsync(buf, 0, buf.Length).ConfigureAwait(false)) > 0)
                            {
                                ms.Write(buf, 0, n);
                                if (ms.Length > ep.MaxBytes) throw new ExcelRenderException("Data set '" + d.Name + "': the response is larger than " + ep.MaxBytes + " bytes.");
                            }
                            return ParseJson(Encoding.UTF8.GetString(ms.ToArray()));
                        }
                    }
                }
            }).GetAwaiter().GetResult();
        }
    }
}
