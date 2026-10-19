using System;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Routing;
using KSC.Excel.Designer;
using KSC.Excel.Services;

namespace KSC.Excel.Mvc
{
    /// <summary>Serves the designer UI + API (System.Web). Register it with <see cref="RouteExtensions.MapKscExcelDesigner"/>.</summary>
    public sealed class DesignerHandler : IHttpHandler
    {
        readonly string _basePath;
        readonly string _path;

        internal DesignerHandler(string basePath, string path) { _basePath = basePath; _path = path; }

        public bool IsReusable => false;

        public void ProcessRequest(HttpContext context)
        {
            var ctx = new HttpContextWrapper(context);
            var host = KscExcel.Host;          // forces initialization
            var o = KscExcel.Options;

            bool allowed = o.AllowAnonymousDesigner || context.IsDebuggingEnabled || (context.User != null && context.User.Identity.IsAuthenticated);
            if (allowed && KscExcel.Authorize != null) allowed = KscExcel.Authorize(ctx);
            if (!allowed)
            {
                context.Response.StatusCode = (context.User != null && context.User.Identity.IsAuthenticated) ? 403 : 401;
                context.Response.Write("Not allowed to use the Excel designer.");
                return;
            }

            var req = new DesignerRequest
            {
                Method = context.Request.HttpMethod,
                Path = "/" + (_path ?? "").TrimStart('/'),
                BasePath = VirtualPathUtility.ToAbsolute("~/" + _basePath.Trim('/')).TrimEnd('/')
            };
            foreach (string key in context.Request.Headers.AllKeys) req.Headers[key] = context.Request.Headers[key];
            if (context.Request.HttpMethod != "GET" && context.Request.HttpMethod != "HEAD")
            {
                using (var ms = new MemoryStream())
                {
                    context.Request.InputStream.CopyTo(ms);
                    req.Body = ms.ToArray();
                }
            }

            var res = host.Handle(req);
            context.Response.StatusCode = res.Status;
            context.Response.ContentType = res.ContentType;
            context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            foreach (var h in res.Headers) context.Response.AddHeader(h.Key, h.Value);
            context.Response.BinaryWrite(res.Body);
        }
    }

    sealed class DesignerRouteHandler : IRouteHandler
    {
        readonly string _basePath;
        public DesignerRouteHandler(string basePath) { _basePath = basePath; }

        public IHttpHandler GetHttpHandler(RequestContext requestContext)
            => new DesignerHandler(_basePath, Convert.ToString(requestContext.RouteData.Values["path"]));
    }

    sealed class DownloadRouteHandler : IRouteHandler
    {
        public IHttpHandler GetHttpHandler(RequestContext requestContext) => new DownloadHandler(Convert.ToString(requestContext.RouteData.Values["token"]));
    }

    sealed class DownloadHandler : IHttpHandler
    {
        readonly string _token;
        public DownloadHandler(string token) { _token = token; }
        public bool IsReusable => false;

        public void ProcessRequest(HttpContext context)
        {
            var user = context.User != null && context.User.Identity.IsAuthenticated ? context.User.Identity.Name : null;
            var item = KscExcel.Downloads.Take(_token, user);
            if (item == null) { context.Response.StatusCode = 404; context.Response.Write("This download link has expired. Reload the page."); return; }
            context.Response.ContentType = item.ContentType;
            // inline items (PDF viewer) are shown by the browser; others are downloads
            context.Response.AddHeader("Content-Disposition", (item.Inline ? "inline" : "attachment") + "; filename=\"report\"; filename*=UTF-8''" + Uri.EscapeDataString(item.FileName));
            context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            context.Response.BinaryWrite(item.Bytes);
        }
    }

    sealed class ReportsRouteHandler : IRouteHandler
    {
        readonly string _basePath;
        public ReportsRouteHandler(string basePath) { _basePath = basePath; }
        public IHttpHandler GetHttpHandler(RequestContext requestContext) => new ReportsHandler(_basePath, Convert.ToString(requestContext.RouteData.Values["name"]));
    }

    /// <summary>Serves the report server (System.Web). Map it with <c>routes.MapKscExcelReports()</c>.</summary>
    sealed class ReportsHandler : IHttpHandler
    {
        readonly string _basePath, _name;
        public ReportsHandler(string basePath, string name) { _basePath = basePath; _name = name; }
        public bool IsReusable => false;

        public void ProcessRequest(HttpContext context)
        {
            var ctx = new HttpContextWrapper(context);
            if (KscExcel.AuthorizeReports != null && !KscExcel.AuthorizeReports(ctx))
            {
                context.Response.StatusCode = (context.User != null && context.User.Identity.IsAuthenticated) ? 403 : 401;
                context.Response.Write("Not allowed to use the reports.");
                return;
            }
            var req = new KSC.Excel.Services.ReportRequest { User = context.User, Name = string.IsNullOrEmpty(_name) ? null : _name, BasePath = VirtualPathUtility.ToAbsolute("~/" + _basePath.Trim('/')).TrimEnd('/') };
            foreach (string key in context.Request.QueryString.AllKeys) if (key != null) req.Query[key] = context.Request.QueryString[key];
            var res = KscExcel.Reports.Handle(req);
            context.Response.StatusCode = res.Status;
            context.Response.ContentType = res.ContentType;
            context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            foreach (var h in res.Headers) context.Response.AddHeader(h.Key, h.Value);
            context.Response.BinaryWrite(res.Body);
        }
    }

    public static class RouteExtensions
    {
        /// <summary>
        /// Maps the report server: <c>/reports</c> (catalog) and <c>/reports/{name}</c> for reports registered with <c>options.AddReport</c>.
        /// Restrict access with <c>KscExcel.AuthorizeReports</c> (reports expose your data).
        /// </summary>
        public static void MapKscExcelReports(this RouteCollection routes, string url = "reports")
        {
            url = url.Trim('/');
            if (KscExcel.Options != null) KscExcel.Options.ReportsPath = "/" + url;
            routes.Add("KscExcelReport", new Route(url + "/{name}", new RouteValueDictionary { { "name", "" } }, new ReportsRouteHandler(url)));
        }

        /// <summary>
        /// Maps the designer, e.g. <c>routes.MapKscExcelDesigner();</c> -&gt; /excel-designer.
        /// Call before the default MVC route. Needs <c>&lt;modules runAllManagedModulesForAllRequests="true"/&gt;</c> or the route to be handled by the UrlRoutingModule (default in MVC apps).
        /// </summary>
        public static void MapKscExcelDesigner(this RouteCollection routes, string url = "excel-designer")
        {
            url = url.Trim('/');
            routes.Add("KscExcelDesigner", new Route(url + "/{*path}", new RouteValueDictionary { { "path", "" } }, new DesignerRouteHandler(url)));
        }

        /// <summary>Maps the one-time download links used by <c>Html.KscExcelViewer</c> (required when you use the viewer).</summary>
        public static void MapKscExcelDownloads(this RouteCollection routes)
        {
            var url = KscExcel.Options != null ? KscExcel.Options.DownloadPath.Trim('/') : "kxl-download";
            routes.Add("KscExcelDownload", new Route(url + "/{token}", new DownloadRouteHandler()));
        }
    }
}
