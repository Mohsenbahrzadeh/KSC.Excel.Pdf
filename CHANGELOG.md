# Changelog

All notable changes are documented here. The project follows [Semantic Versioning](https://semver.org/);
template JSON compatibility is governed by `schemaVersion`.

## [1.0.0] - first public release

### Added
- Visual designer (web, embedded in your app; Windows desktop app for .NET 10 **and** .NET Framework 4.8).
- Engine: bands, nested bands, sort/group with totals, cross-tab, parameters, calculated fields, custom filters/functions, report sections, sheet repeat.
- Data: objects/dictionaries/JSON/DataTable, read-only parameterised **SQL**, **REST/JSON endpoints**, **code providers**, inline JSON.
- `.xlsx` writer (live formulas, charts, conditional formats, validation, images, print setup, page breaks) + **streaming writer** for millions of rows; **xlsx importer**.
- **PDF** (`netstandard2.0`: .NET Framework 4.6.1+ and .NET 8/10): Persian/Arabic shaping and RTL, conditional formats, rotated text, text overflow, charts, **watermark**, page break per group, header/footer, repeating title rows.
- **Barcodes & QR codes** (QR byte mode v1-10, Code 128) in `.xlsx` and as vector shapes in PDF.
- Viewers: HTML (`KscExcelViewer`) and **in-page PDF viewer** (`KscPdfViewer` / `Html.KscPdfViewer`).
- Hosts: ASP.NET Core 8/10, ASP.NET MVC 5 (.NET Framework 4.7.2+/4.8, verified on IIS Express), Windows desktop.
- Multi-level grouping wizard with `SUBTOTAL` totals, group/sort by expression, calculated columns (`RUNSUM`, `PCT`, `PREV`, `DIFF`, `RANK`, `MOVAVG`), Jalali calendar functions/filters and `jdate` parameters.
- PDF password protection and permissions (AES-256).
- Outputs: CSV, Word (docx), stand-alone HTML, PNG; `ReportFile` / `Report` one-liners.
- **Report server** (`AddReport`, `MapKscExcelReports`): parameter forms (text/number/date/jdate/bool/drop-down), HTML/PDF view, downloads, drill-down links (`report:Name?Id={{Id|urlenc}}`).
- **Scheduler** (`ReportScheduler`): daily/weekly/monthly/Jalali-monthly/every-N-minutes, e-mail (SMTP) and folder delivery.
- Template **history** (restore earlier versions), **combo charts** (column + line, secondary axis), EAN-13 / EAN-8 / Code 39 barcodes.
- Complete English and Persian guides.
