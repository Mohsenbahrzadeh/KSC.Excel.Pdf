using KSC.Excel.AspNetCore;
#if KXL_PDF
using KSC.Excel.Pdf;
#endif
using KSC.Excel.Services;
using KSC.Excel.Sample.Web.Models;
using Microsoft.AspNetCore.Mvc;

namespace KSC.Excel.Sample.Web.Controllers;

public class HomeController : Controller
{
    public IActionResult Index() => View();

    /// <summary>Show the report inside a page (ViewComponent).</summary>
    public IActionResult Viewer() => View(Invoice.Demo());

    /// <summary>Download mode: one line of code, no spreadsheet library.</summary>
    public IActionResult Download() => this.ExcelFile("Invoice", Invoice.Demo(), "invoice.xlsx");

#if KXL_PDF
    /// <summary>PDF viewer inside a page (browser PDF reader + Download / Print buttons).</summary>
    public IActionResult PdfViewer() => View(Invoice.Demo());

    /// <summary>PDF from the same template (Persian shaping + RTL).</summary>
    public IActionResult Pdf([FromServices] IExcelReportService svc)
        => File(svc.RenderPdf("Invoice", Invoice.Demo()), "application/pdf", "invoice.pdf");
#endif
}
