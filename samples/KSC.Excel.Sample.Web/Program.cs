using KSC.Excel.AspNetCore;
#if KXL_PDF
using KSC.Excel.Pdf;
#endif
using KSC.Excel.Sample.Web.Models;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews();
builder.Services.AddKscExcel(o =>
{
    o.DesignerTitle = "طراح قالب اکسل";
#if KXL_PDF
    o.UsePdf();                              // adds the PDF button to the designer
#endif
    o.AddModel<Invoice>("Invoice");          // fields + generated sample data for the designer
    o.AddReport("Invoice", p => Invoice.Demo(), title: "فاکتور نمونه");   // published by the report server (/reports)
});

var app = builder.Build();

app.UseStaticFiles();
app.UseRouting();

// The designer: protect it with your own login in production (.RequireAuthorization()).
app.MapKscExcelDesigner("/excel-designer");
app.MapKscExcelDownloads();
app.MapKscExcelReports("/reports");      // report server: catalog, parameter forms, pdf/xlsx/docx/csv/html/png (add .RequireAuthorization() in real apps)
app.MapDefaultControllerRoute();

app.Run();
