using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace KSC.Excel.Model
{
    /// <summary>
    /// A complete, serializable Excel template - what the designer edits and what the engine renders.
    /// Persisted as plain JSON (see docs/KSC.Excel-Design.md for the schema).
    /// </summary>
    public class ExcelTemplate
    {
        public const int CurrentSchemaVersion = 1;

        public int SchemaVersion { get; set; } = CurrentSchemaVersion;
        /// <summary>Unique key used by the store and by IExcelReportService.Render(name, data).</summary>
        public string Name { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Author { get; set; }
        /// <summary>Optional: name of a model registered with options.AddModel&lt;T&gt;(); drives the field list.</summary>
        public string ModelName { get; set; }
        public string DefaultFont { get; set; } = "Tahoma";
        public double DefaultFontSize { get; set; } = 10;
        public List<SheetTemplate> Sheets { get; set; } = new List<SheetTemplate>();
        /// <summary>Style table: id -> style. Cells reference styles by id.</summary>
        public Dictionary<string, StyleDef> Styles { get; set; } = new Dictionary<string, StyleDef>();
        /// <summary>Sample data used by the designer preview (never used when rendering at runtime).</summary>
        public JToken SampleData { get; set; }
        public List<DataSetDef> DataSets { get; set; } = new List<DataSetDef>();
        public List<ParameterDef> Parameters { get; set; } = new List<ParameterDef>();

        /// <summary>
        /// Nulls are omitted everywhere. A value-type property is omitted when it equals the type's default only if the
        /// class initializer leaves it at that default too (so "ShowGridLines = true" or "Zoom = 100" are always written;
        /// omitting a false/0 there would silently change meaning on reload).
        /// </summary>
        sealed class TemplateContractResolver : CamelCasePropertyNamesContractResolver
        {
            protected override JsonProperty CreateProperty(System.Reflection.MemberInfo member, MemberSerialization memberSerialization)
            {
                var p = base.CreateProperty(member, memberSerialization);
                if (p.PropertyType != null && p.PropertyType.IsValueType && Nullable.GetUnderlyingType(p.PropertyType) == null && member.DeclaringType.GetConstructor(Type.EmptyTypes) != null)
                {
                    try
                    {
                        var inst = Activator.CreateInstance(member.DeclaringType);
                        var initial = p.ValueProvider.GetValue(inst);
                        if (Equals(initial, Activator.CreateInstance(p.PropertyType))) p.DefaultValueHandling = DefaultValueHandling.Ignore;
                    }
                    catch (Exception) { /* keep the property always serialized */ }
                }
                return p;
            }
        }

        static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            ContractResolver = new TemplateContractResolver(),
            NullValueHandling = NullValueHandling.Ignore,
            DefaultValueHandling = DefaultValueHandling.Include,
            DateParseHandling = DateParseHandling.None
        };

        public static ExcelTemplate FromJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) throw new ArgumentException("Template json is empty.", nameof(json));
            var t = JsonConvert.DeserializeObject<ExcelTemplate>(json, Settings);
            if (t == null) throw new InvalidDataException("Invalid template json.");
            if (t.SchemaVersion > CurrentSchemaVersion)
                throw new InvalidDataException("Template schema " + t.SchemaVersion + " is newer than this library supports (" + CurrentSchemaVersion + "). Upgrade KSC.Excel.");
            if (t.Sheets == null) t.Sheets = new List<SheetTemplate>();
            if (t.Styles == null) t.Styles = new Dictionary<string, StyleDef>();
            if (t.DataSets == null) t.DataSets = new List<DataSetDef>();
            if (t.Parameters == null) t.Parameters = new List<ParameterDef>();
            foreach (var d in t.DataSets) if (d.Columns == null) d.Columns = new List<DataSetColumn>();
            foreach (var s in t.Sheets) s.Normalize();
            return t;
        }

        public string ToJson(bool indented = false)
            => JsonConvert.SerializeObject(this, indented ? Formatting.Indented : Formatting.None, Settings);

        public static ExcelTemplate LoadFile(string path) => FromJson(File.ReadAllText(path, Encoding.UTF8));

        public void SaveFile(string path) => File.WriteAllText(path, ToJson(true), new UTF8Encoding(false));

        public ExcelTemplate Clone() => FromJson(ToJson());

        /// <summary>A blank template with one empty sheet.</summary>
        public static ExcelTemplate CreateBlank(string name, bool rtl = true)
        {
            var t = new ExcelTemplate { Name = name, Title = name };
            t.Sheets.Add(new SheetTemplate { Name = "Sheet1", Rtl = rtl });
            return t;
        }
    }

    public class SheetTemplate
    {
        public string Name { get; set; } = "Sheet1";
        /// <summary>Optional path (e.g. "Branches"): the sheet is rendered once per item (name may use {{placeholders}}).</summary>
        public string RepeatFor { get; set; }
        public bool Rtl { get; set; }
        public bool ShowGridLines { get; set; } = true;
        public string TabColor { get; set; }
        public int Zoom { get; set; } = 100;
        /// <summary>Default column width in characters.</summary>
        public double DefaultColWidth { get; set; } = 12;
        /// <summary>Default row height in points.</summary>
        public double DefaultRowHeight { get; set; } = 18;
        public List<ColumnDef> Columns { get; set; } = new List<ColumnDef>();
        public List<RowDef> Rows { get; set; } = new List<RowDef>();
        public List<CellDef> Cells { get; set; } = new List<CellDef>();
        /// <summary>Merged ranges in A1 notation ("A1:C1").</summary>
        public List<string> Merges { get; set; } = new List<string>();
        public List<BandDef> Bands { get; set; } = new List<BandDef>();
        public int FreezeRows { get; set; }
        public int FreezeCols { get; set; }
        public List<CondFormatDef> ConditionalFormats { get; set; } = new List<CondFormatDef>();
        public List<ValidationDef> Validations { get; set; } = new List<ValidationDef>();
        public List<ImageDef> Images { get; set; } = new List<ImageDef>();
        public List<ChartDef> Charts { get; set; } = new List<ChartDef>();
        /// <summary>Report sections (Stimulsoft-style labels over rows): reportHeader, pageHeader, reportFooter.</summary>
        public List<SectionDef> Sections { get; set; } = new List<SectionDef>();
        public PageSetupDef Page { get; set; } = new PageSetupDef();
        /// <summary>Header row range for AutoFilter, e.g. "A5:F5".</summary>
        public string AutoFilter { get; set; }
        /// <summary>Protect the sheet (no password; stops accidental edits).</summary>
        public bool Protect { get; set; }

        internal void Normalize()
        {
            if (Columns == null) Columns = new List<ColumnDef>();
            if (Rows == null) Rows = new List<RowDef>();
            if (Cells == null) Cells = new List<CellDef>();
            if (Merges == null) Merges = new List<string>();
            if (Bands == null) Bands = new List<BandDef>();
            if (ConditionalFormats == null) ConditionalFormats = new List<CondFormatDef>();
            if (Validations == null) Validations = new List<ValidationDef>();
            if (Images == null) Images = new List<ImageDef>();
            if (Charts == null) Charts = new List<ChartDef>();
            if (Sections == null) Sections = new List<SectionDef>();
            if (Page == null) Page = new PageSetupDef();
            if (Zoom <= 0) Zoom = 100;
        }
    }

    public class ColumnDef
    {
        /// <summary>1-based column index.</summary>
        public int Index { get; set; }
        public double? Width { get; set; }
        public bool Hidden { get; set; }
    }

    public class RowDef
    {
        /// <summary>1-based row index.</summary>
        public int Index { get; set; }
        public double? Height { get; set; }
        public bool Hidden { get; set; }
    }

    public class CellDef
    {
        /// <summary>1-based row.</summary>
        public int R { get; set; }
        /// <summary>1-based column.</summary>
        public int C { get; set; }
        /// <summary>Text, number, "{{Field}}" placeholder text, or a formula starting with "=".</summary>
        public string V { get; set; }
        /// <summary>Style id (key of <see cref="ExcelTemplate.Styles"/>).</summary>
        public string S { get; set; }
        /// <summary>Value type hint: null = auto (numbers become numbers), "text" = force text.</summary>
        public string T { get; set; }
        /// <summary>Hyperlink: URL (may contain placeholders) or "#Sheet!A1" for in-workbook links.</summary>
        public string L { get; set; }
        /// <summary>Barcode: "qr" (or "qr:L|M|Q|H") or "code128". The cell text ({{Field}}) is encoded and drawn as an image that fills the cell.</summary>
        public string Bc { get; set; }
    }

    /// <summary>
    /// A band repeats a block of template rows for every item of a collection (the detail rows of a report).
    /// Bands may be nested: a band fully inside another band's rows is its child and is bound relative to the parent item.
    /// </summary>
    public class BandDef
    {
        /// <summary>Identifier used inside formulas: =SUM([Items.Price]).</summary>
        public string Name { get; set; }
        /// <summary>Path of the collection, e.g. "Items" or "Order.Lines".</summary>
        public string Source { get; set; }
        public int FirstRow { get; set; }
        public int LastRow { get; set; }
        /// <summary>When the collection is empty: false = the rows vanish, true = keep one blank block.</summary>
        public bool KeepWhenEmpty { get; set; }
        /// <summary>Zebra striping: fill color applied to every even occurrence (#RRGGBB).</summary>
        public string AlternateFill { get; set; }
        /// <summary>Sort the items before repeating (path relative to the item, e.g. "Date").</summary>
        public string SortBy { get; set; }
        public bool SortDesc { get; set; }
        /// <summary>
        /// Group the items by this path. Each repeat then represents one group whose fields are
        /// Key, Count and Items; put a nested band bound to "Items" inside to list the members,
        /// and use =SUM([Inner.Field]) or {{Sum(Items.Field)}} in the group footer row.
        /// </summary>
        public string GroupBy { get; set; }
        /// <summary>"rows" (default): repeats rows. "columns": repeats columns (cross-tab) using FirstCol/LastCol.</summary>
        public string Orientation { get; set; }
        public int FirstCol { get; set; }
        public int LastCol { get; set; }
        /// <summary>Start every item (except the first) on a new printed page / PDF page (page per group, one invoice per page...).</summary>
        public bool PageBreak { get; set; }
    }

    /// <summary>A computed collection available to the template (no code needed), e.g. a cross-tab.</summary>
    public class DataSetDef
    {
        /// <summary>Name used in templates: {{Matrix.GrandTotal}}, band source "Matrix.Rows".</summary>
        public string Name { get; set; }
        /// <summary>"pivot" (cross-tab of a collection) or "sql" (read-only query).</summary>
        public string Type { get; set; } = "pivot";
        /// <summary>Path of the source collection in the data.</summary>
        public string Source { get; set; }
        public string RowPath { get; set; }
        public string ColPath { get; set; }
        public string ValuePath { get; set; }
        /// <summary>sum | count | avg | min | max</summary>
        public string Agg { get; set; } = "sum";
        /// <summary>For Type "sql": name of a connection registered with KscExcelOptions.AddConnection.</summary>
        public string Connection { get; set; }
        /// <summary>For Type "sql": one SELECT/WITH statement; @Name binds the report parameter Name.</summary>
        public string Sql { get; set; }
        /// <summary>
        /// For Type "json": inline JSON text (lookup tables, constants) - or leave empty and set <see cref="Connection"/> to a
        /// REST endpoint registered with KscExcelOptions.AddHttpEndpoint and <see cref="Url"/> to the path.
        /// For Type "provider": <see cref="Connection"/> names a function registered with KscExcelOptions.AddProvider.
        /// </summary>
        public string Json { get; set; }
        /// <summary>For Type "json" + endpoint: path/query appended to the endpoint's base URL; {Name} is replaced by the report parameter Name (URL-encoded).</summary>
        public string Url { get; set; }
        /// <summary>For Type "json": JSONPath-like selector of the array inside the document (e.g. "data.items"); empty = the document itself.</summary>
        public string JsonPath { get; set; }
        /// <summary>Designer metadata: columns discovered by "test query" (used for the field tree).</summary>
        public List<DataSetColumn> Columns { get; set; } = new List<DataSetColumn>();
    }

    public class DataSetColumn
    {
        public string Name { get; set; }
        public string Type { get; set; } = "text";
    }

    /// <summary>A report parameter. Available as {{$p.Name}}; callers pass values, otherwise Default is used.</summary>
    public class ParameterDef
    {
        public string Name { get; set; }
        public string Label { get; set; }
        /// <summary>text | number | date | bool</summary>
        public string Type { get; set; } = "text";
        public string Default { get; set; }
        /// <summary>Choices for a drop-down in the parameter form: "A|B|C" or "value:Label|value2:Label2".</summary>
        public string Options { get; set; }
    }

    public class CondFormatDef
    {
        public string Range { get; set; }
        /// <summary>cellIs, expression, containsText, duplicates, top10, colorScale, dataBar.</summary>
        public string Type { get; set; }
        /// <summary>For cellIs: equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, between, notBetween.</summary>
        public string Operator { get; set; }
        public string Formula1 { get; set; }
        public string Formula2 { get; set; }
        public string Text { get; set; }
        public int Rank { get; set; } = 10;
        public bool Bottom { get; set; }
        /// <summary>Format applied when the rule matches (cellIs/expression/containsText/duplicates/top10).</summary>
        public StyleDef Style { get; set; }
        /// <summary>colorScale: 2 or 3 colors (#RRGGBB) from low to high.</summary>
        public List<string> Colors { get; set; }
        /// <summary>dataBar color.</summary>
        public string BarColor { get; set; }
    }

    public class ValidationDef
    {
        public string Range { get; set; }
        /// <summary>list, whole, decimal, date, textLength, custom.</summary>
        public string Type { get; set; }
        public string Operator { get; set; } = "between";
        /// <summary>For lists: comma separated literal ("Yes,No") or a range ("=$H$1:$H$5").</summary>
        public string Formula1 { get; set; }
        public string Formula2 { get; set; }
        public bool AllowBlank { get; set; } = true;
        public string ErrorTitle { get; set; }
        public string Error { get; set; }
        public string PromptTitle { get; set; }
        public string Prompt { get; set; }
    }

    public class ImageDef
    {
        public string Name { get; set; }
        /// <summary>1-based anchor row/column (top-left corner).</summary>
        public int Row { get; set; } = 1;
        public int Col { get; set; } = 1;
        /// <summary>Size in pixels.</summary>
        public int Width { get; set; } = 120;
        public int Height { get; set; } = 60;
        /// <summary>Static image: base64 (no data: prefix).</summary>
        public string Data { get; set; }
        /// <summary>png, jpeg or gif.</summary>
        public string Ext { get; set; } = "png";
        /// <summary>Dynamic image: path to a byte[] or base64 string in the data (overrides Data when it resolves).</summary>
        public string Binding { get; set; }
    }

    /// <summary>
    /// A labelled block of rows. "pageHeader" rows repeat on every printed page / PDF page (unless PrintTitleRows is set);
    /// "reportHeader" and "reportFooter" are visual labels in the designer (the rows are rendered once, normally).
    /// </summary>
    public class SectionDef
    {
        /// <summary>reportHeader | pageHeader | reportFooter</summary>
        public string Kind { get; set; }
        public string Name { get; set; }
        public int FirstRow { get; set; }
        public int LastRow { get; set; }
    }

    public class ChartSeriesDef
    {
        /// <summary>Literal series name (or leave empty and use NameCell).</summary>
        public string Name { get; set; }
        /// <summary>Optional cell holding the series name, e.g. "D8".</summary>
        public string NameCell { get; set; }
        /// <summary>Values range in template coordinates, e.g. "D9:D9" (a range covering a band expands to all its rows).</summary>
        public string Values { get; set; }
        /// <summary>#RRGGBB (optional).</summary>
        public string Color { get; set; }
        /// <summary>Combo charts: column, line or area for this series (empty = the chart's type). Ignored for pie and bar (horizontal) charts.</summary>
        public string Type { get; set; }
        /// <summary>Plot this series against a second value axis on the right (e.g. a % line over amount columns).</summary>
        public bool Secondary { get; set; }
    }

    /// <summary>A chart bound to cell ranges. Ranges that cover a band grow with the data.</summary>
    public class ChartDef
    {
        public string Name { get; set; }
        /// <summary>column, bar, line, area or pie.</summary>
        public string Type { get; set; } = "column";
        public string Title { get; set; }
        /// <summary>1-based anchor cell (top-left) and size in pixels.</summary>
        public int Row { get; set; } = 1;
        public int Col { get; set; } = 1;
        public int Width { get; set; } = 480;
        public int Height { get; set; } = 288;
        /// <summary>Category labels range, e.g. "B9:B9".</summary>
        public string Categories { get; set; }
        public List<ChartSeriesDef> Series { get; set; } = new List<ChartSeriesDef>();
        public bool Stacked { get; set; }
        public bool ShowLegend { get; set; } = true;
        public bool ShowDataLabels { get; set; }
    }

    public class PageSetupDef
    {
        /// <summary>portrait | landscape</summary>
        public string Orientation { get; set; } = "portrait";
        /// <summary>Excel paper size: 9 = A4, 8 = A3, 1 = Letter, 5 = Legal.</summary>
        public int PaperSize { get; set; } = 9;
        /// <summary>Fit all columns on one page wide (height unlimited).</summary>
        public bool FitToWidth { get; set; }
        public int Scale { get; set; } = 100;
        /// <summary>Margins in inches.</summary>
        public double MarginTop { get; set; } = 0.75;
        public double MarginBottom { get; set; } = 0.75;
        public double MarginLeft { get; set; } = 0.5;
        public double MarginRight { get; set; } = 0.5;
        public bool CenterHorizontally { get; set; }
        /// <summary>Excel header/footer codes allowed: &amp;P page, &amp;N pages, &amp;D date, &amp;L/&amp;C/&amp;R sections.</summary>
        public string Header { get; set; }
        public string Footer { get; set; }
        /// <summary>Rows repeated on every printed page, e.g. "1:3".</summary>
        public string PrintTitleRows { get; set; }
        public string PrintArea { get; set; }
        /// <summary>PDF only: diagonal text drawn behind every page (e.g. "DRAFT", "محرمانه").</summary>
        public string Watermark { get; set; }
    }
}
