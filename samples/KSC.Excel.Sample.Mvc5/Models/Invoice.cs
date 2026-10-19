using System;
using System.Collections.Generic;

namespace KSC.Excel.Sample.Mvc5.Models
{
    public class Invoice
    {
        public string Number { get; set; }
        public DateTime Date { get; set; }
        public Customer Customer { get; set; }
        public List<InvoiceLine> Lines { get; set; }
        public string Notes { get; set; }

        public static Invoice Demo()
        {
            return new Invoice
            {
                Number = "INV-1403-021",
                Date = new DateTime(2025, 3, 1),
                Customer = new Customer { Name = "شرکت نمونه", Phone = "021-12345678" },
                Notes = "پرداخت ظرف ۳۰ روز",
                Lines = new List<InvoiceLine>
                {
                    new InvoiceLine { Product = "ورق فولادی ۲ میلی‌متر", Qty = 12, UnitPrice = 4500000m },
                    new InvoiceLine { Product = "سیم جوش", Qty = 30, UnitPrice = 380000m },
                    new InvoiceLine { Product = "حمل و نقل", Qty = 1, UnitPrice = 2500000m }
                }
            };
        }
    }

    public class Customer { public string Name { get; set; } public string Phone { get; set; } }
    public class InvoiceLine { public string Product { get; set; } public int Qty { get; set; } public decimal UnitPrice { get; set; } }
}
