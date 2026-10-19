using System;
using System.Collections.Generic;
using KSC.Excel.Preview;

namespace KSC.Excel.Services
{
    /// <summary>Configuration shared by the ASP.NET Core and .NET Framework hosts.</summary>
    public sealed class KscExcelOptions
    {
        /// <summary>Where templates live when no custom store is given. Relative paths are resolved by the host (content root / app base).</summary>
        public string TemplateDirectory { get; set; } = "ExcelTemplates";

        /// <summary>Custom template storage (database, blob...). Defaults to <see cref="FileTemplateStore"/>.</summary>
        public ITemplateStore Store { get; set; }

        /// <summary>Designer title shown in the header.</summary>
        public string DesignerTitle { get; set; } = "KSC Excel Designer";

        /// <summary>fa (Persian, RTL) or en. Persian is the default.</summary>
        public string Language { get; set; } = "fa";

        /// <summary>
        /// Hosts allow anonymous access to the designer only in Development unless this is true.
        /// The designer can read/write templates, so keep it behind your login in production.
        /// </summary>
        public bool AllowAnonymousDesigner { get; set; }

        /// <summary>URL prefix of the one-time download links created by the viewer component.</summary>
        public string DownloadPath { get; set; } = "/kxl-download";

        /// <summary>When true the designer can open/preview/download but cannot save or delete templates.</summary>
        public bool ReadOnlyDesigner { get; set; }

        /// <summary>Maximum rows an exported sheet may have (guards memory). Default: Excel's limit.</summary>
        public int MaxRows { get; set; } = 1048576;

        /// <summary>Maximum generated cells per sheet (the report is built in memory). Default 3,000,000.</summary>
        public long MaxCells { get; set; } = 3000000;

        /// <summary>Rows shown by HTML previews and the viewer component.</summary>
        public int PreviewMaxRows { get; set; } = 500;

        /// <summary>
        /// Allow templates to run their SQL data sets (read-only SELECT/WITH). Off by default: anyone who can save templates
        /// could otherwise query the connections you register. Use a read-only database login.
        /// </summary>
        public bool AllowSql { get; set; }
        public int SqlCommandTimeoutSeconds { get; set; } = 30;
        public int SqlMaxRows { get; set; } = 100000;

        readonly Dictionary<string, Func<System.Data.IDbConnection>> _connections = new Dictionary<string, Func<System.Data.IDbConnection>>(StringComparer.OrdinalIgnoreCase);
        internal IEnumerable<string> ConnectionNames { get { return _connections.Keys; } }

        /// <summary>Registers a named connection for SQL data sets (also sets nothing else: remember to enable AllowSql).</summary>
        public KscExcelOptions AddConnection(string name, Func<System.Data.IDbConnection> factory)
        {
            _connections[name] = factory; return this;
        }

        /// <summary>e.g. AddConnection("Sales", SqlClientFactory.Instance, connectionString).</summary>
        public KscExcelOptions AddConnection(string name, System.Data.Common.DbProviderFactory provider, string connectionString)
        {
            return AddConnection(name, () => { var c = provider.CreateConnection(); c.ConnectionString = connectionString; return c; });
        }

        internal System.Data.IDbConnection OpenFactory(string name)
        {
            Func<System.Data.IDbConnection> f;
            return name != null && _connections.TryGetValue(name, out f) ? f() : null;
        }

        // ---- external data (data sets of Type "provider" / "json") ----
        // Registered in CODE, never in templates: a template can only name what the developer allowed (like SQL connections).
        internal Dictionary<string, Func<IDictionary<string, object>, object>> Providers { get; } = new Dictionary<string, Func<IDictionary<string, object>, object>>(StringComparer.OrdinalIgnoreCase);
        internal Dictionary<string, KSC.Excel.Engine.HttpEndpoint> Endpoints { get; } = new Dictionary<string, KSC.Excel.Engine.HttpEndpoint>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Registers a named data provider usable by templates (data set Type "provider"): your code returns the data
        /// (EF query, Dapper, DataTable, JSON, any list) from the report parameters.
        /// <code>o.AddProvider("Orders", p =&gt; db.Orders.Where(x =&gt; x.Year == (int)p["Year"]).ToList());</code>
        /// </summary>
        public KscExcelOptions AddProvider(string name, Func<IDictionary<string, object>, object> provider)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("name");
            Providers[name] = provider ?? throw new ArgumentNullException(nameof(provider)); return this;
        }

        /// <summary>
        /// Registers a REST/JSON endpoint (data set Type "json"). Templates give only a relative path; the host is fixed here
        /// (no SSRF from templates). <paramref name="headers"/> may carry an API key.
        /// </summary>
        public KscExcelOptions AddHttpEndpoint(string name, string baseUrl, IDictionary<string, string> headers = null, int timeoutSeconds = 30)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("name");
            Endpoints[name] = new KSC.Excel.Engine.HttpEndpoint { BaseUrl = baseUrl, Headers = headers == null ? null : new Dictionary<string, string>(headers), TimeoutSeconds = timeoutSeconds };
            return this;
        }

        /// <summary>Custom placeholder filter: {{Price|tax:9}} -> AddFilter("tax", (value, arg) =&gt; ...).</summary>
        public Dictionary<string, Func<object, string, object>> Filters { get; } = new Dictionary<string, Func<object, string, object>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Custom function usable in placeholder expressions ({{Tax(Price)}}) and, for cached values, in Excel formulas (=TAX(B2)).</summary>
        public Dictionary<string, Func<object[], object>> Functions { get; } = new Dictionary<string, Func<object[], object>>(StringComparer.OrdinalIgnoreCase);

        public KscExcelOptions AddFilter(string name, Func<object, string, object> filter) { Filters[name] = filter; return this; }
        public KscExcelOptions AddFunction(string name, Func<object[], object> function) { Functions[name] = function; return this; }

        /// <summary>
        /// Turns a bound workbook into a PDF. Set by the KSC.Excel.Pdf package (<c>options.UsePdf()</c>); when set, the designer shows a PDF button.
        /// </summary>
        public Func<KSC.Excel.Engine.RenderedWorkbook, byte[]> PdfExporter { get; set; }

        /// <summary>Models whose shape the designer offers as a field tree (and from which sample data is generated).</summary>
        public Dictionary<string, ModelInfo> Models { get; } = new Dictionary<string, ModelInfo>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Registers a data model: the designer shows its fields (drag &amp; drop) and generates sample data for the preview.
        /// </summary>
        public KscExcelOptions AddModel<T>(string name = null) => AddModel(typeof(T), name);

        public KscExcelOptions AddModel(Type type, string name = null)
        {
            var info = ModelInfo.FromType(type, name);
            Models[info.Name] = info;
            return this;
        }

        /// <summary>Registers a model described only by sample JSON (the fields are inferred from it).</summary>
        public KscExcelOptions AddModelFromJson(string name, string sampleJson)
        {
            Models[name] = ModelInfo.FromJson(name, sampleJson);
            return this;
        }

        internal PreviewOptions Preview(bool headers) => new PreviewOptions { MaxRows = PreviewMaxRows, ShowHeaders = headers };
    }
}
