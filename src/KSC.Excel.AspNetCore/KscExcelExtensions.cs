using KSC.Excel.Designer;
using KSC.Excel.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ApplicationParts;
using Microsoft.AspNetCore.Routing;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Hosting;

namespace KSC.Excel.AspNetCore;

public static class KscExcelExtensions
{
    /// <summary>
    /// Registers KSC.Excel: <see cref="IExcelReportService"/>, the template store (files under
    /// <c>{ContentRoot}/ExcelTemplates</c> by default) and the viewer ViewComponent.
    /// </summary>
    public static IServiceCollection AddKscExcel(this IServiceCollection services, Action<KscExcelOptions>? configure = null)
    {
        var options = new KscExcelOptions();
        configure?.Invoke(options);
        services.TryAddSingleton(options);

        services.TryAddSingleton<ITemplateStore>(sp =>
        {
            var o = sp.GetRequiredService<KscExcelOptions>();
            if (o.Store is not null) return o.Store;
            var dir = o.TemplateDirectory;
            if (!Path.IsPathRooted(dir)) dir = Path.Combine(sp.GetRequiredService<IHostEnvironment>().ContentRootPath, dir);
            return new FileTemplateStore(dir);
        });
        services.TryAddSingleton<IExcelReportService>(sp => new ExcelReportService(sp.GetRequiredService<KscExcelOptions>(), sp.GetRequiredService<ITemplateStore>()));
        services.TryAddSingleton<DesignerHost>(sp => new DesignerHost(sp.GetRequiredService<KscExcelOptions>(), sp.GetRequiredService<IExcelReportService>()));
        services.TryAddSingleton<DownloadStore>();

        // make the ViewComponent discoverable even if the app only references this package indirectly
        services.AddMvcCore().ConfigureApplicationPartManager(m =>
        {
            var asm = typeof(KscExcelViewerViewComponent).Assembly;
            if (!m.ApplicationParts.OfType<AssemblyPart>().Any(p => p.Assembly == asm)) m.ApplicationParts.Add(new AssemblyPart(asm));
        });
        return services;
    }

    /// <summary>
    /// Maps the visual designer (UI + API) under <paramref name="pattern"/>. Chain <c>.RequireAuthorization(...)</c> to protect it.
    /// Outside Development, anonymous users are rejected unless <see cref="KscExcelOptions.AllowAnonymousDesigner"/> is set.
    /// </summary>
    public static IEndpointConventionBuilder MapKscExcelDesigner(this IEndpointRouteBuilder endpoints, string pattern = "/excel-designer")
    {
        pattern = "/" + pattern.Trim('/');
        var host = endpoints.ServiceProvider.GetRequiredService<DesignerHost>();
        var options = endpoints.ServiceProvider.GetRequiredService<KscExcelOptions>();
        var env = endpoints.ServiceProvider.GetRequiredService<IHostEnvironment>();

        async Task Handle(HttpContext ctx)
        {
            if (!options.AllowAnonymousDesigner && !env.IsDevelopment() && ctx.User.Identity?.IsAuthenticated != true)
            {
                ctx.Response.StatusCode = StatusCodes.Status401Unauthorized;
                await ctx.Response.WriteAsync("Sign in to use the Excel designer.");
                return;
            }

            var rel = (string?)ctx.Request.RouteValues["path"] ?? "";
            if (rel.Length == 0 && HttpMethods.IsGet(ctx.Request.Method) && !ctx.Request.Path.Value!.EndsWith('/'))
            {
                // relative URLs in the UI need the trailing slash
                ctx.Response.Redirect(ctx.Request.PathBase + pattern + "/" + ctx.Request.QueryString);
                return;
            }

            var req = new DesignerRequest
            {
                Method = ctx.Request.Method,
                Path = "/" + rel.TrimStart('/'),
                BasePath = ctx.Request.PathBase + pattern
            };
            foreach (var h in ctx.Request.Headers) req.Headers[h.Key] = h.Value.ToString();
            if (ctx.Request.ContentLength is > 0 or null && !HttpMethods.IsGet(ctx.Request.Method))
            {
                using var ms = new MemoryStream();
                await ctx.Request.Body.CopyToAsync(ms, ctx.RequestAborted);
                req.Body = ms.ToArray();
            }

            var res = host.Handle(req);
            ctx.Response.StatusCode = res.Status;
            ctx.Response.ContentType = res.ContentType;
            ctx.Response.Headers["Cache-Control"] = "no-store";
            foreach (var h in res.Headers) ctx.Response.Headers[h.Key] = h.Value;
            await ctx.Response.Body.WriteAsync(res.Body, ctx.RequestAborted);
        }

        endpoints.MapMethods(pattern, new[] { "GET" }, Handle).ExcludeFromDescription();

        return endpoints.MapMethods(pattern + "/{**path}", new[] { "GET", "POST", "PUT", "DELETE" }, Handle).ExcludeFromDescription();
    }

    /// <summary>Maps the one-time download links created by <c>&lt;vc:ksc-excel-viewer&gt;</c> (required when you use the viewer).</summary>
    public static IEndpointConventionBuilder MapKscExcelDownloads(this IEndpointRouteBuilder endpoints)
    {
        var options = endpoints.ServiceProvider.GetRequiredService<KscExcelOptions>();
        var pattern = "/" + options.DownloadPath.Trim('/');
        return endpoints.MapGet(pattern + "/{token}", (string token, HttpContext ctx, DownloadStore store) =>
        {
            var item = store.Take(token, ctx.User.Identity?.Name);
            if (item is null) return Results.NotFound("This download link has expired. Reload the page.");
            ctx.Response.Headers["Cache-Control"] = "no-store";
            // inline items (PDF viewer) are shown by the browser; others are downloads
            return item.Inline
                ? Results.File(item.Bytes, item.ContentType)
                : Results.File(item.Bytes, item.ContentType, item.FileName);
        }).ExcludeFromDescription();
    }
}
