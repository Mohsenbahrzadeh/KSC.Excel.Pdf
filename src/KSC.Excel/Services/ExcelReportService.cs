using System;
using System.Collections.Generic;
using System.IO;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Preview;
using KSC.Excel.Writer;

namespace KSC.Excel.Services
{
    public static class ExcelContent
    {
        public const string Xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public const string Pdf = "application/pdf";
    }

    /// <summary>The only API most developers need: template name + data in, .xlsx bytes out.</summary>
    public interface IExcelReportService
    {
        /// <summary>Renders a stored template with the given data and returns the .xlsx file.</summary>
        byte[] Render(string templateName, object data);
        /// <summary>Same, with report parameters available as {{$p.Name}}.</summary>
        byte[] Render(string templateName, object data, IDictionary<string, object> parameters);
        RenderedWorkbook Bind(ExcelTemplate template, object data, IDictionary<string, object> parameters);
        /// <summary>For the designer: like Bind, but SQL data sets are skipped (with a warning) when SQL is not allowed.</summary>
        RenderedWorkbook BindPreview(ExcelTemplate template, object data, IDictionary<string, object> parameters);
        byte[] Render(ExcelTemplate template, object data);
        void Render(ExcelTemplate template, object data, Stream output);

        /// <summary>Renders the same report as an HTML table with sheet tabs (for on-screen viewing).</summary>
        string RenderHtml(string templateName, object data, bool showHeaders = false);
        string RenderHtml(ExcelTemplate template, object data, bool showHeaders = false);

        /// <summary>Binds data without writing a file (warnings, diagnostics).</summary>
        RenderedWorkbook Bind(ExcelTemplate template, object data);

        /// <summary>
        /// Writes the .xlsx while the report is produced, so memory stays flat even for millions of rows (rows are not kept).
        /// Returns true when streamed; false when the template cannot be streamed (see <see cref="StreamAnalyzer.Check"/>), in which case
        /// the normal buffered render was used. Streamed files have no cached formula values (Excel calculates on open).
        /// An exception while streaming leaves a partial file in <paramref name="output"/>: write to a temp file or a buffer you can discard.
        /// </summary>
        bool RenderStreaming(ExcelTemplate template, object data, Stream output, IDictionary<string, object> parameters = null);
        bool RenderStreaming(string templateName, object data, Stream output, IDictionary<string, object> parameters = null);

        ExcelTemplate GetTemplate(string name);
        ITemplateStore Store { get; }
    }

    public sealed class ExcelReportService : IExcelReportService
    {
        readonly KscExcelOptions _opt;

        public ExcelReportService(KscExcelOptions options, ITemplateStore store)
        {
            _opt = options ?? new KscExcelOptions();
            Store = store ?? throw new ArgumentNullException(nameof(store));
        }

        public ITemplateStore Store { get; }

        public ExcelTemplate GetTemplate(string name)
        {
            var t = Store.Get(name);
            if (t == null) throw new FileNotFoundException("Excel template '" + name + "' was not found. Create it in the designer first.");
            return t;
        }

        public RenderedWorkbook Bind(ExcelTemplate template, object data) => Bind(template, data, null);

        public RenderedWorkbook Bind(ExcelTemplate template, object data, IDictionary<string, object> parameters)
            => new TemplateRenderer(Options(parameters, false)).Render(template, data);

        public RenderedWorkbook BindPreview(ExcelTemplate template, object data, IDictionary<string, object> parameters)
            => new TemplateRenderer(Options(parameters, true)).Render(template, data);

        public bool RenderStreaming(string templateName, object data, Stream output, IDictionary<string, object> parameters = null)
            => RenderStreaming(GetTemplate(templateName), data, output, parameters);

        public bool RenderStreaming(ExcelTemplate template, object data, Stream output, IDictionary<string, object> parameters = null)
        {
            if (template == null) throw new ArgumentNullException(nameof(template));
            if (output == null) throw new ArgumentNullException(nameof(output));
            if (StreamAnalyzer.Check(template).Count > 0)
            {
                XlsxWriter.Write(Bind(template, data, parameters), output);
                return false;
            }
            var opt = Options(parameters, false);
            opt.LazySql = true;
            using (var sink = new XlsxStreamSink(output, template))
            {
                new TemplateRenderer(opt).RenderStreaming(template, data, sink);
                sink.Complete();
            }
            return true;
        }

        RenderOptions Options(IDictionary<string, object> parameters, bool preview) => new RenderOptions
        {
            MaxRows = _opt.MaxRows, MaxCells = _opt.MaxCells, Parameters = parameters,
            AllowSql = _opt.AllowSql, SkipSql = preview && !_opt.AllowSql, Preview = preview, ConnectionFactory = _opt.OpenFactory,
            SqlTimeoutSeconds = _opt.SqlCommandTimeoutSeconds, SqlMaxRows = _opt.SqlMaxRows,
            Filters = _opt.Filters, Functions = _opt.Functions, Providers = _opt.Providers, Endpoints = _opt.Endpoints
        };

        public byte[] Render(string templateName, object data, IDictionary<string, object> parameters)
            => XlsxWriter.ToBytes(Bind(GetTemplate(templateName), data, parameters));

        public byte[] Render(string templateName, object data) => Render(GetTemplate(templateName), data);

        public byte[] Render(ExcelTemplate template, object data) => XlsxWriter.ToBytes(Bind(template, data));

        public void Render(ExcelTemplate template, object data, Stream output) => XlsxWriter.Write(Bind(template, data), output);

        public string RenderHtml(string templateName, object data, bool showHeaders = false) => RenderHtml(GetTemplate(templateName), data, showHeaders);

        public string RenderHtml(ExcelTemplate template, object data, bool showHeaders = false)
            => HtmlPreviewRenderer.Render(Bind(template, data), _opt.Preview(showHeaders));
    }
}
