using System;
using System.Collections.Generic;
using System.Linq;
using KSC.Excel.Engine;

namespace KSC.Excel.Model
{
    /// <summary>
    /// Code-first alternative to the visual designer (handy for tests and for seeding default templates).
    /// </summary>
    public sealed class TemplateBuilder
    {
        readonly ExcelTemplate _t;
        public TemplateBuilder(string name, string title = null)
        {
            _t = new ExcelTemplate { Name = name, Title = title ?? name };
        }

        public TemplateBuilder Font(string name, double size) { _t.DefaultFont = name; _t.DefaultFontSize = size; return this; }

        /// <summary>Registers a style under an id and returns the id for use in Cell(...).</summary>
        public string Style(string id, StyleDef style) { _t.Styles[id] = style; return id; }

        public TemplateBuilder Sample(string json) { _t.SampleData = Newtonsoft.Json.Linq.JToken.Parse(json); return this; }

        public TemplateBuilder Sheet(string name, Action<SheetBuilder> build, bool rtl = false)
        {
            var s = new SheetTemplate { Name = name, Rtl = rtl };
            build(new SheetBuilder(s));
            _t.Sheets.Add(s);
            return this;
        }

        public ExcelTemplate Build() => _t;
    }

    public sealed class SheetBuilder
    {
        readonly SheetTemplate _s;
        internal SheetBuilder(SheetTemplate s) { _s = s; }

        public SheetBuilder Cell(string a1, string value, string styleId = null)
        {
            int r, c;
            if (!CellRef.TryParseCell(a1, out r, out c)) throw new ArgumentException("Bad cell address: " + a1);
            var existing = _s.Cells.FirstOrDefault(x => x.R == r && x.C == c);
            if (existing != null) { existing.V = value; existing.S = styleId ?? existing.S; }
            else _s.Cells.Add(new CellDef { R = r, C = c, V = value, S = styleId });
            return this;
        }

        /// <summary>Applies a style to every cell of a range (creating empty cells as needed).</summary>
        public SheetBuilder StyleRange(string range, string styleId)
        {
            int r1, c1, r2, c2;
            if (!CellRef.TryParseRange(range, out r1, out c1, out r2, out c2)) throw new ArgumentException("Bad range: " + range);
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                {
                    var existing = _s.Cells.FirstOrDefault(x => x.R == r && x.C == c);
                    if (existing != null) existing.S = styleId;
                    else _s.Cells.Add(new CellDef { R = r, C = c, S = styleId });
                }
            return this;
        }

        public SheetBuilder Merge(string range) { _s.Merges.Add(range); return this; }
        public SheetBuilder Col(int index, double width) { _s.Columns.Add(new ColumnDef { Index = index, Width = width }); return this; }
        public SheetBuilder Row(int index, double height) { _s.Rows.Add(new RowDef { Index = index, Height = height }); return this; }
        public SheetBuilder Freeze(int rows, int cols = 0) { _s.FreezeRows = rows; _s.FreezeCols = cols; return this; }
        public SheetBuilder Band(string name, string source, int firstRow, int lastRow, string alternateFill = null)
        {
            _s.Bands.Add(new BandDef { Name = name, Source = source, FirstRow = firstRow, LastRow = lastRow, AlternateFill = alternateFill });
            return this;
        }
        public SheetBuilder Landscape() { _s.Page.Orientation = "landscape"; _s.Page.FitToWidth = true; return this; }
        public SheetBuilder NoGrid() { _s.ShowGridLines = false; return this; }
        public SheetBuilder AutoFilter(string range) { _s.AutoFilter = range; return this; }
        public SheetBuilder ConditionalFormat(CondFormatDef cf) { _s.ConditionalFormats.Add(cf); return this; }
        public SheetBuilder Validation(ValidationDef v) { _s.Validations.Add(v); return this; }
        public SheetBuilder Raw(Action<SheetTemplate> edit) { edit(_s); return this; }
    }

    /// <summary>Ready-made starting points offered by the designer ("New from sample").</summary>
    public static class StarterTemplates
    {
        static BorderDef Thin(string color = "#C9D1D9") => new BorderDef { Style = "thin", Color = color };

        static StyleDef Box(StyleDef s, string color = "#C9D1D9")
        {
            s.Left = Thin(color); s.Right = Thin(color); s.Top = Thin(color); s.Bottom = Thin(color);
            return s;
        }

        public static IList<KeyValuePair<string, ExcelTemplate>> All()
            => new List<KeyValuePair<string, ExcelTemplate>>
            {
                new KeyValuePair<string, ExcelTemplate>("table", SimpleTable()),
                new KeyValuePair<string, ExcelTemplate>("invoice", Invoice())
            };

        /// <summary>Title, header row, repeating detail band and a totals row.</summary>
        public static ExcelTemplate SimpleTable()
        {
            var b = new TemplateBuilder("SimpleTable", "Simple table report").Font("Tahoma", 10);
            var title = b.Style("title", new StyleDef { Size = 16, Bold = true, Color = "#FFFFFF", Fill = "#1F4E79", H = "center", V = "center" });
            var sub = b.Style("sub", new StyleDef { Color = "#555555", H = "center", Italic = true });
            var head = b.Style("head", Box(new StyleDef { Bold = true, Color = "#FFFFFF", Fill = "#2E75B6", H = "center", V = "center", Wrap = true }, "#1F4E79"));
            var text = b.Style("text", Box(new StyleDef { V = "center" }));
            var num = b.Style("num", Box(new StyleDef { Format = "#,##0", V = "center" }));
            var center = b.Style("center", Box(new StyleDef { H = "center", V = "center" }));
            var totalL = b.Style("totalL", Box(new StyleDef { Bold = true, Fill = "#DDEBF7", H = "left" }, "#1F4E79"));
            var totalN = b.Style("totalN", Box(new StyleDef { Bold = true, Fill = "#DDEBF7", Format = "#,##0" }, "#1F4E79"));

            b.Sample("{\"Title\":\"Sales report\",\"Period\":\"Q1\",\"Items\":[{\"Name\":\"Item A\",\"Category\":\"Hardware\",\"Qty\":4,\"Price\":125000,\"Total\":500000},{\"Name\":\"Item B\",\"Category\":\"Software\",\"Qty\":2,\"Price\":720000,\"Total\":1440000},{\"Name\":\"Item C\",\"Category\":\"Service\",\"Qty\":10,\"Price\":35000,\"Total\":350000}]}");
            b.Sheet("Report", s =>
            {
                s.Col(1, 6).Col(2, 28).Col(3, 18).Col(4, 10).Col(5, 16).Col(6, 18).Row(1, 34).Row(4, 24).Row(5, 22);
                s.Cell("A1", "{{Title}}", title).StyleRange("A1:F1", title).Merge("A1:F1");
                s.Cell("A2", "{{Period}}", sub).StyleRange("A2:F2", sub).Merge("A2:F2");
                s.Cell("A4", "#", head).Cell("B4", "Name", head).Cell("C4", "Category", head).Cell("D4", "Qty", head).Cell("E4", "Price", head).Cell("F4", "Total", head);
                s.Cell("A5", "{{$index}}", center).Cell("B5", "{{Name}}", text).Cell("C5", "{{Category}}", text).Cell("D5", "{{Qty}}", num).Cell("E5", "{{Price}}", num).Cell("F5", "{{Total}}", num);
                s.Band("Items", "Items", 5, 5, "#F2F7FC");
                s.Cell("A6", "Total", totalL).StyleRange("A6:C6", totalL).Merge("A6:C6").Cell("D6", "=SUM([Items.Qty])", totalN).Cell("E6", "", totalN).Cell("F6", "=SUM([Items.Total])", totalN);
                s.Freeze(4).AutoFilter("A4:F4").Landscape();
            });
            var t = b.Build();
            return t;
        }

        /// <summary>A formal invoice: header block, line items with formulas, totals.</summary>
        public static ExcelTemplate Invoice()
        {
            var b = new TemplateBuilder("Invoice", "Invoice").Font("Tahoma", 10);
            var title = b.Style("title", new StyleDef { Size = 20, Bold = true, Color = "#1F4E79", H = "right", V = "center" });
            var label = b.Style("label", new StyleDef { Bold = true, Color = "#555555", H = "right" });
            var value = b.Style("value", new StyleDef { Color = "#111111", H = "right" });
            var head = b.Style("head", Box(new StyleDef { Bold = true, Color = "#FFFFFF", Fill = "#1F4E79", H = "center", V = "center" }, "#1F4E79"));
            var text = b.Style("text", Box(new StyleDef { V = "center" }));
            var num = b.Style("num", Box(new StyleDef { Format = "#,##0", V = "center" }));
            var center = b.Style("center", Box(new StyleDef { H = "center", V = "center" }));
            var totL = b.Style("totL", Box(new StyleDef { Bold = true, Fill = "#DDEBF7", H = "right" }, "#1F4E79"));
            var totN = b.Style("totN", Box(new StyleDef { Bold = true, Fill = "#DDEBF7", Format = "#,##0" }, "#1F4E79"));

            b.Sample("{\"Number\":\"INV-1403-021\",\"Date\":\"2025-03-01\",\"Customer\":{\"Name\":\"Acme Co.\",\"Phone\":\"021-1234567\"},\"Lines\":[{\"Product\":\"Steel sheet 2mm\",\"Qty\":12,\"UnitPrice\":4500000},{\"Product\":\"Welding wire\",\"Qty\":30,\"UnitPrice\":380000},{\"Product\":\"Transport\",\"Qty\":1,\"UnitPrice\":2500000}]}");
            b.Sheet("Invoice", s =>
            {
                s.Col(1, 6).Col(2, 34).Col(3, 10).Col(4, 18).Col(5, 20).Row(1, 36);
                s.Cell("A1", "Invoice", title).StyleRange("A1:E1", title).Merge("A1:E1");
                s.Cell("A3", "Number", label).Merge("A3:B3").Cell("C3", "{{Number}}", value).Merge("C3:E3");
                s.Cell("A4", "Date", label).Merge("A4:B4").Cell("C4", "{{Date|format:yyyy-MM-dd}}", value).Merge("C4:E4");
                s.Cell("A5", "Customer", label).Merge("A5:B5").Cell("C5", "{{Customer.Name}}", value).Merge("C5:E5");
                s.Cell("A7", "#", head).Cell("B7", "Product", head).Cell("C7", "Qty", head).Cell("D7", "Unit price", head).Cell("E7", "Amount", head).Row(7, 24);
                s.Cell("A8", "{{$index}}", center).Cell("B8", "{{Product}}", text).Cell("C8", "{{Qty}}", num).Cell("D8", "{{UnitPrice}}", num).Cell("E8", "=[Qty]*[UnitPrice]", num);
                s.Band("Lines", "Lines", 8, 8, "#F7F9FB");
                s.Cell("A9", "Grand total", totL).StyleRange("A9:D9", totL).Merge("A9:D9").Cell("E9", "=SUM([Lines:E])", totN);
                s.NoGrid();
            }, rtl: false);
            return b.Build();
        }
    }
}
