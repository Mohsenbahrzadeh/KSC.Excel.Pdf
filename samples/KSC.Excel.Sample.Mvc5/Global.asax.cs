using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using KSC.Excel.Mvc;
#if KXL_PDF
using KSC.Excel.Pdf;
#endif
using KSC.Excel.Sample.Mvc5.Models;

namespace KSC.Excel.Sample.Mvc5
{
    public class MvcApplication : HttpApplication
    {
        protected void Application_Start()
        {
            KscExcel.Initialize(o =>
            {
                o.DesignerTitle = "طراح قالب اکسل";
                o.TemplateDirectory = "ExcelTemplates";   // ~/ExcelTemplates
#if KXL_PDF
                o.UsePdf();                                // PDF button in the designer + this.Pdf / Html.KscPdfViewer
#endif
                o.AddModel<Invoice>("Invoice");
            });
            // The designer can write templates: in a real application require a login / role here.
            // (Without it the designer is open only while <compilation debug="true"> is set.)
            // KscExcel.Authorize = ctx => ctx.User.IsInRole("Admin");

            var routes = RouteTable.Routes;
            routes.MapKscExcelDesigner();      // /excel-designer
            routes.MapKscExcelDownloads();     // links used by the viewers
            routes.MapRoute("Default", "{controller}/{action}/{id}", new { controller = "Home", action = "Index", id = UrlParameter.Optional });
        }
    }
}
