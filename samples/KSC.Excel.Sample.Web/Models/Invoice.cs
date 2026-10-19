namespace KSC.Excel.Sample.Web.Models;

public class Invoice
{
    public string Number { get; set; } = "";
    public DateTime Date { get; set; }
    public Customer Customer { get; set; } = new();
    public List<InvoiceLine> Lines { get; set; } = new();
    public string Notes { get; set; } = "";

    public static Invoice Demo() => new()
    {
        Number = "INV-1403-021",
        Date = new DateTime(2025, 3, 1),
        Customer = new Customer { Name = "شرکت نمونه", Phone = "021-12345678" },
        Notes = "پرداخت ظرف ۳۰ روز",
        Lines =
        {
            new InvoiceLine { Product = "ورق فولادی ۲ میلی‌متر", Qty = 12, UnitPrice = 4_500_000m },
            new InvoiceLine { Product = "سیم جوش", Qty = 30, UnitPrice = 380_000m },
            new InvoiceLine { Product = "حمل و نقل", Qty = 1, UnitPrice = 2_500_000m },
        }
    };
}

public class Customer
{
    public string Name { get; set; } = "";
    public string Phone { get; set; } = "";
}

public class InvoiceLine
{
    public string Product { get; set; } = "";
    public int Qty { get; set; }
    public decimal UnitPrice { get; set; }
}
