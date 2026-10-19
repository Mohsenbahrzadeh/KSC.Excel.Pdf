using System;
using System.IO;
using System.Web;
using KSC.Excel.Designer;
using KSC.Excel.Services;

namespace KSC.Excel.Mvc
{
    /// <summary>
    /// Entry point for .NET Framework apps (no DI container required).
    /// Call <see cref="Initialize"/> once, e.g. in Application_Start.
    /// </summary>
    public static class KscExcel
    {
        static readonly object Gate = new object();
        static IExcelReportService _service;
        static DesignerHost _host;

        public static KscExcelOptions Options { get; private set; }
        public static DownloadStore Downloads { get; } = new DownloadStore();

        /// <summary>
        /// Optional extra authorization for the designer, e.g. <c>ctx =&gt; ctx.User.IsInRole("Admin")</c>.
        /// Without it the designer needs an authenticated user (or &lt;compilation debug="true"&gt; / AllowAnonymousDesigner).
        /// </summary>
        public static Func<HttpContextBase, bool> Authorize { get; set; }

        public static IExcelReportService Service
        {
            get { EnsureInitialized(); return _service; }
        }

        internal static DesignerHost Host
        {
            get { EnsureInitialized(); return _host; }
        }

        public static void Initialize(Action<KscExcelOptions> configure = null)
        {
            lock (Gate)
            {
                var o = new KscExcelOptions();
                if (configure != null) configure(o);

                var store = o.Store;
                if (store == null)
                {
                    var dir = o.TemplateDirectory;
                    if (!Path.IsPathRooted(dir)) dir = HttpContext.Current != null ? HttpContext.Current.Server.MapPath("~/" + dir) : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, dir);
                    store = new FileTemplateStore(dir);
                }
                Options = o;
                _service = new ExcelReportService(o, store);
                _host = new DesignerHost(o, _service);
            }
        }

        static void EnsureInitialized()
        {
            if (_service == null) Initialize();
        }
    }
}
