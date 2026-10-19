# Contributing

Thanks for helping! Bug reports, docs fixes and features are welcome.

## Build & test

```bash
dotnet build                                   # needs the .NET 10 SDK; Windows for net48 / WPF / MVC 5 projects
dotnet test tests/KSC.Excel.Tests              # net10.0 and net48
dotnet test tests/KSC.Excel.Pdf.Tests          # net10.0 and net48
```

Run the samples:

```bash
dotnet run --project samples/KSC.Excel.Sample.Web --urls http://localhost:5199 -- --environment Development   # ASP.NET Core
# .NET Framework 4.8 + MVC 5 (Windows, IIS Express):
dotnet build samples/KSC.Excel.Sample.Mvc5
"C:\Program Files\IIS Express\iisexpress.exe" /path:samples\KSC.Excel.Sample.Mvc5 /port:5290
```

## Guidelines

- A behaviour change needs a test. The engine tests are plain xUnit (`tests/KSC.Excel.Tests`); PDF output is checked by rendering to PNG and inspecting pixels (`tests/KSC.Excel.Pdf.Tests`). Set `KXL_PDF_PNG_DIR` to dump PNGs for a visual check.
- An incompatible change of the **template JSON** must bump `ExcelTemplate.CurrentSchemaVersion` and keep reading old files.
- The core (`KSC.Excel`) targets `netstandard2.0` and stays **dependency-free** except `Newtonsoft.Json`; the designer UI is plain JavaScript embedded in the assembly (no build step, no CDN).
- Update the docs in `docs/` (English + Persian guides) when you change user-visible behaviour. Rebuild the guide PDFs as described in `docs/guide/README.md`.
- Keep changes focused; run the full test-suite before opening a pull request.

## Reporting bugs

Please include the template JSON (remove private data), the data shape, the output you got and the output you expected. For rendering problems a screenshot of the PDF/Excel result helps a lot.
