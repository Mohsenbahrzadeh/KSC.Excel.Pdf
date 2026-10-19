using System;
using System.Collections.Generic;
using KSC.Excel.Model;

namespace KSC.Excel.Engine
{
    /// <summary>The result of binding a template to data - what the writers (xlsx / html preview) consume.</summary>
    public sealed class RenderedWorkbook
    {
        public ExcelTemplate Template;
        public List<RenderedSheet> Sheets = new List<RenderedSheet>();
        public List<string> Warnings = new List<string>();
        /// <summary>True when the preview limit cut rows (never set by the normal renderer).</summary>
        public bool Truncated;
    }

    public sealed class RenderedSheet
    {
        public string Name;
        public bool Rtl, ShowGridLines = true, Protect;
        public string TabColor;
        public int Zoom = 100;
        public double DefaultColWidth = 12, DefaultRowHeight = 18;
        public List<ColumnDef> Columns = new List<ColumnDef>();
        /// <summary>Contiguous rows: Rows[i].Index == i + 1.</summary>
        public List<RenderedRow> Rows = new List<RenderedRow>();
        public List<int[]> Merges = new List<int[]>();   // r1,c1,r2,c2
        public int MaxCol;
        public int FreezeRows, FreezeCols;
        public List<CondFormatDef> ConditionalFormats = new List<CondFormatDef>();
        public List<ValidationDef> Validations = new List<ValidationDef>();
        public List<RenderedImage> Images = new List<RenderedImage>();
        public List<RenderedChart> Charts = new List<RenderedChart>();
        public PageSetupDef Page = new PageSetupDef();
        public string AutoFilter;
        public string PrintTitleRows;
        public string PrintArea;
        /// <summary>Rows after which a manual page break occurs (1-based, output coordinates).</summary>
        public List<int> RowBreaks = new List<int>();
        /// <summary>Band name -> first/last output row of all its occurrences (used for page-by-page carried totals).</summary>
        public Dictionary<string, int[]> BandRuns = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);

        public RenderedCell Get(int row, int col)
        {
            if (row < 1 || row > Rows.Count) return null;
            var cells = Rows[row - 1].Cells;
            // cells are sorted by column; binary search
            int lo = 0, hi = cells.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                int c = cells[mid].Col;
                if (c == col) return cells[mid];
                if (c < col) lo = mid + 1; else hi = mid - 1;
            }
            return null;
        }
    }

    public sealed class RenderedRow
    {
        public int Index;
        public double? Height;
        public bool Hidden;
        public List<RenderedCell> Cells = new List<RenderedCell>();
    }

    public sealed class RenderedCell
    {
        public int Col;
        /// <summary>null, string, double, bool or DateTime.</summary>
        public object Value;
        /// <summary>Formula text without the leading "=".</summary>
        public string Formula;
        /// <summary>Value computed by the engine for formula cells (cached value in the file), null if unknown.</summary>
        public object Cached;
        public StyleDef Style;
        public string Link;
        internal int Row;
    }

    public sealed class RenderedChart
    {
        public ChartDef Def;
        public int Row, Col;
        /// <summary>Absolute references with sheet name, e.g. 'Sheet'!$B$9:$B$20.</summary>
        public string Categories;
        public List<string> SeriesValues = new List<string>();
        public List<string> SeriesNames = new List<string>();   // literal text or a reference
        public List<bool> SeriesNameIsRef = new List<bool>();
        public List<string> SeriesColors = new List<string>();
        public List<string> SeriesTypes = new List<string>();
        public List<bool> SeriesSecondary = new List<bool>();
    }

    public sealed class RenderedImage
    {
        public string Name;
        public int Row, Col, Width, Height;
        public byte[] Bytes;
        public string Ext;
        /// <summary>Set for generated barcodes: the PDF renderer redraws them as crisp vector shapes.</summary>
        public string BarType, BarText;
    }

    /// <summary>Thrown for invalid templates or data that cannot be rendered (e.g. more rows than Excel allows).</summary>
    public class ExcelRenderException : Exception
    {
        public ExcelRenderException(string message) : base(message) { }
    }

    /// <summary>Receives sheets/rows as they are produced (streaming output).</summary>
    public interface IStreamSink
    {
        /// <summary>Sheet metadata (name, columns, view settings) is ready; rows follow.</summary>
        void BeginSheet(RenderedSheet meta);
        void Row(RenderedRow row);
        /// <summary>All rows were delivered; the sheet now also carries merges, conditional formats, images, charts...</summary>
        void EndSheet(RenderedSheet sheet);
    }

    public sealed class RenderOptions
    {
        /// <summary>Hard cap of generated rows per sheet (Excel allows 1,048,576).</summary>
        public int MaxRows = CellRef.MaxRows;
        /// <summary>Hard cap of generated cells per workbook (guards memory: the whole report is built in memory).</summary>
        public long MaxCells = 3000000;
        /// <summary>Compute formula results so viewers that do not recalculate still show numbers.</summary>
        public bool ComputeFormulaValues = true;
        /// <summary>Values for {{$p.Name}}; they override the defaults declared in the template.</summary>
        public System.Collections.Generic.IDictionary<string, object> Parameters;
        /// <summary>Creates a new connection for a SQL data set connection name (null = SQL data sets unavailable).</summary>
        public Func<string, System.Data.IDbConnection> ConnectionFactory;
        /// <summary>False = a template with SQL data sets throws. True = allowed.</summary>
        public bool AllowSql;
        /// <summary>True (designer preview): SQL data sets are skipped with a warning instead of failing, so sample data can be used.</summary>
        public bool SkipSql;
        /// <summary>Designer preview: failures of external data sets (REST/provider) become warnings so sample data can be used.</summary>
        public bool Preview;
        /// <summary>Base URL for cell links of the form <c>report:Name?Param=value</c> (null = such links are left as they are).</summary>
        public string ReportLinkBase;
        public int SqlTimeoutSeconds = 30;
        public int SqlMaxRows = 100000;
        /// <summary>SQL data sets are read lazily (one pass per use) instead of being loaded into memory. Used by streaming.</summary>
        public bool LazySql;
        public System.Collections.Generic.IDictionary<string, Func<object, string, object>> Filters;
        public System.Collections.Generic.IDictionary<string, Func<object[], object>> Functions;
        /// <summary>Named data providers for data sets of Type "provider": parameters in, any data (list, DataTable, JSON...) out.</summary>
        public System.Collections.Generic.IDictionary<string, Func<System.Collections.Generic.IDictionary<string, object>, object>> Providers;
        /// <summary>Named REST endpoints for data sets of Type "json".</summary>
        public System.Collections.Generic.IDictionary<string, HttpEndpoint> Endpoints;
    }
}
