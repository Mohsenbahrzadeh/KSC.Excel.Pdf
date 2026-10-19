using System.Web.Mvc;
using KSC.Excel.Mvc;
using KSC.Excel.Sample.Mvc5.Models;

namespace KSC.Excel.Sample.Mvc5.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index() { return View(); }

        /// <summary>The report inside the page (HTML) with a download button.</summary>
        public ActionResult Viewer() { return View(Invoice.Demo()); }

#if KXL_PDF
        /// <summary>The report as a PDF inside the page (browser PDF reader + Download / Print).</summary>
        public ActionResult PdfViewer() { return View(Invoice.Demo()); }

        public ActionResult Download() { return this.Excel("Invoice", Invoice.Demo(), "invoice.xlsx"); }

        public ActionResult Pdf() { return this.Pdf("Invoice", Invoice.Demo(), "invoice.pdf"); }
#endif
    }
}
