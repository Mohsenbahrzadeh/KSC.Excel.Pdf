namespace KSC.Excel.Pdf;

public sealed class PdfOptions
{
    /// <summary>Font tried after the cell's own font when it lacks glyphs (Persian needs Tahoma/Segoe UI/Arial or any Arabic-capable font).</summary>
    public string FallbackFontFamily { get; set; } = "Tahoma";

    /// <summary>Optional path to a .ttf/.otf with Persian glyphs (e.g. Vazirmatn). Highest priority when it covers the text. Recommended on Linux servers.</summary>
    public string? FontFile { get; set; }

    /// <summary>Stop after this many pages (guards against accidental million-row PDFs).</summary>
    public int MaxPages { get; set; } = 2000;

    /// <summary>Document title metadata (defaults to the template title).</summary>
    public string? Title { get; set; }

    /// <summary>Draw light gridlines like the on-screen sheet (Excel prints without them by default).</summary>
    public bool PrintGridLines { get; set; }

    /// <summary>Diagonal text behind every page (overrides the template's own watermark), e.g. "DRAFT".</summary>
    public string? Watermark { get; set; }

    /// <summary>Password protection / permissions (AES-256). Null = an ordinary PDF. Example: <c>new PdfOptions { Security = new PdfSecurity { UserPassword = "1234", AllowCopy = false } }</c>.</summary>
    public PdfSecurity? Security { get; set; }

    /// <summary>Show digits as Persian digits on every page (overrides the template's own setting).</summary>
    public bool PersianDigits { get; set; }
}
