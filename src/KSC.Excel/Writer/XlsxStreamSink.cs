using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using KSC.Excel.Engine;
using KSC.Excel.Model;

namespace KSC.Excel.Writer
{
    /// <summary>
    /// Streaming .xlsx writer: rows go straight into the zip while the report is rendered, so memory stays flat for huge reports.
    /// Text is written as inline strings and formulas have no cached values (Excel calculates on open).
    /// Use through <c>IExcelReportService.RenderStreaming</c>.
    /// </summary>
    public sealed class XlsxStreamSink : IStreamSink, IDisposable
    {
        readonly ZipArchive _zip;
        readonly ExcelTemplate _tpl;
        readonly StyleBook _styles;
        readonly List<XlsxWriter.SheetInfo> _infos = new List<XlsxWriter.SheetInfo>();
        readonly List<RenderedImage> _media = new List<RenderedImage>();
        int _chartCounter;

        Stream _entry; XmlWriter _w;
        XlsxWriter.SheetInfo _cur;
        Dictionary<StyleDef, int> _xf;
        List<KeyValuePair<string, string>> _links;

        public XlsxStreamSink(Stream output, ExcelTemplate template)
        {
            _tpl = template;
            _zip = new ZipArchive(output, ZipArchiveMode.Create, true, new UTF8Encoding(false));
            _styles = new StyleBook(template.DefaultFont, template.DefaultFontSize);
        }

        public void BeginSheet(RenderedSheet meta)
        {
            _cur = new XlsxWriter.SheetInfo { Index = _infos.Count + 1, Sheet = meta, ChartBase = _chartCounter };
            _xf = new Dictionary<StyleDef, int>();
            _links = new List<KeyValuePair<string, string>>();
            _entry = _zip.CreateEntry("xl/worksheets/sheet" + _cur.Index + ".xml", CompressionLevel.Optimal).Open();
            _w = XmlWriter.Create(_entry, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false, CheckCharacters = false });
            _w.WriteStartDocument(true);
            XlsxWriter.WriteSheetHead(_w, _cur, true);
            _w.WriteStartElement("sheetData");
        }

        public void Row(RenderedRow row) { XlsxWriter.WriteRow(_w, row, _styles, null, _xf, _links); }

        public void EndSheet(RenderedSheet sheet)
        {
            _w.WriteEndElement();                                  // sheetData
            XlsxWriter.WriteSheetTail(_w, _cur, _styles, _links);
            _w.WriteEndDocument();
            _w.Flush(); _w.Dispose(); _entry.Dispose();
            _w = null; _entry = null;

            _chartCounter += sheet.Charts.Count;
            foreach (var im in sheet.Images) _media.Add(im);
            XlsxWriter.WriteSheetExtras(_zip, _cur, _links, _media);
            _infos.Add(_cur);
        }

        /// <summary>Writes the workbook-level parts and closes the package. Call once, after rendering.</summary>
        public void Complete()
        {
            var wb = new RenderedWorkbook { Template = _tpl };
            wb.Sheets.AddRange(_infos.Select(i => i.Sheet));
            for (int i = 0; i < _media.Count; i++)
                using (var s = _zip.CreateEntry("xl/media/image" + (i + 1) + "." + XlsxWriter.ImageExt(_media[i]), CompressionLevel.NoCompression).Open())
                    s.Write(_media[i].Bytes, 0, _media[i].Bytes.Length);

            XlsxWriter.WriteContentTypes(_zip, wb, _media);
            XlsxWriter.WriteDocParts(_zip, _tpl);
            XlsxWriter.WriteWorkbook(_zip, _infos);
            XlsxWriter.Part(_zip, "xl/styles.xml", _styles.Write);
            XlsxWriter.Part(_zip, "xl/sharedStrings.xml", new SharedStrings().Write);
            _zip.Dispose();
        }

        public void Dispose() { if (_w != null) { _w.Dispose(); _entry.Dispose(); } }
    }
}
