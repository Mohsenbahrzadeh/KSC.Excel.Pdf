using System.Globalization;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Preview;
using SkiaSharp;

namespace KSC.Excel.Pdf;

/// <summary>Renders a bound workbook (template + data) to PDF - same engine as the .xlsx, with Persian/Arabic shaping and RTL layout.</summary>
public static class ExcelPdf
{
    public static byte[] Render(RenderedWorkbook wb, PdfOptions? options = null)
    {
        using var ms = new MemoryStream();
        Render(wb, ms, options);
        return ms.ToArray();
    }

    public static void Render(RenderedWorkbook wb, Stream output, PdfOptions? options = null)
    {
        options ??= new PdfOptions();
        if (options.Security != null)
        {
            // draw normally, then encrypt the finished file
            using var plain = new MemoryStream();
            Render(wb, plain, new PdfOptions { FallbackFontFamily = options.FallbackFontFamily, FontFile = options.FontFile, MaxPages = options.MaxPages, Title = options.Title, PrintGridLines = options.PrintGridLines, Watermark = options.Watermark, PersianDigits = options.PersianDigits });
            var enc = PdfEncryptor.Encrypt(plain.ToArray(), options.Security);
            output.Write(enc, 0, enc.Length);
            return;
        }
        using var job = new PdfJob(wb, options);
        var meta = new SKDocumentPdfMetadata { Title = options.Title ?? wb.Template?.Title ?? "Report", Creator = "KSC.Excel", Producer = "KSC.Excel" };
        using var stream = new SKManagedWStream(output);
        using var doc = SKDocument.CreatePdf(stream, meta);
        foreach (var page in job.Pages)
        {
            using var canvas = doc.BeginPage(page.W, page.H);
            job.Draw(canvas, page);
            doc.EndPage();
        }
        doc.Close();
    }

    /// <summary>Renders one page to PNG (visual checks, thumbnails). <paramref name="scale"/> is pixels per point.</summary>
    public static byte[] RenderPng(RenderedWorkbook wb, int pageNumber = 1, float scale = 1.5f, PdfOptions? options = null)
    {
        options ??= new PdfOptions();
        using var job = new PdfJob(wb, options);
        if (pageNumber < 1 || pageNumber > job.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The report has " + job.Pages.Count + " page(s).");
        var page = job.Pages[pageNumber - 1];
        using var bmp = new SKBitmap((int)(page.W * scale), (int)(page.H * scale));
        using (var canvas = new SKCanvas(bmp))
        {
            canvas.Clear(SKColors.White);
            canvas.Scale(scale);
            job.Draw(canvas, page);
        }
        using var img = SKImage.FromBitmap(bmp);
        return img.Encode(SKEncodedImageFormat.Png, 100).ToArray();
    }

    public static int CountPages(RenderedWorkbook wb, PdfOptions? options = null)
    {
        using var job = new PdfJob(wb, options ?? new PdfOptions());
        return job.Pages.Count;
    }
}

internal sealed class PageSpec
{
    public int Sheet, Number, Total;
    public float W, H, Scale;
    public int ColFrom, ColTo;
    public List<int> Rows = new();
    /// <summary>How many leading entries of <see cref="Rows"/> are repeated title rows.</summary>
    public int TitleCount;
    /// <summary>Carried-forward totals (one per carry column) shown at the top / bottom of the page, or null.</summary>
    public double[]? CarryIn, CarryOut;
    public int[] CarryCols = Array.Empty<int>();
    public float CarryH;
    public int[]? CarryRun;
}

internal sealed class SheetGeo
{
    public RenderedSheet S = null!;
    public float[] ColW = Array.Empty<float>();     // index = column (1-based)
    public float[] ColOff = Array.Empty<float>();   // cumulative start offset of column c
    public float[] RowH = Array.Empty<float>();
    public int Cols, Rows;
    public Dictionary<(int, int), int[]> Anchors = new();   // (r,c) -> r2,c2
    public HashSet<(int, int)> Covered = new();
    public FormulaEvaluator Eval = null!;
    public CondFormatEvaluator Cf = null!;
}

internal sealed class PdfJob : IDisposable
{
    const float PxToPt = 0.75f;
    readonly RenderedWorkbook _wb;
    readonly PdfOptions _opt;
    readonly TextPainter _text;
    readonly List<SheetGeo> _geo = new();
    public List<PageSpec> Pages { get; } = new();
    readonly Dictionary<string, SKBitmap?> _images = new();

    public PdfJob(RenderedWorkbook wb, PdfOptions opt)
    {
        _wb = wb; _opt = opt; _text = new TextPainter(opt);
        for (int i = 0; i < wb.Sheets.Count; i++) { _geo.Add(Geometry(wb.Sheets[i])); Paginate(i); }
        for (int i = 0; i < Pages.Count; i++) { Pages[i].Number = i + 1; Pages[i].Total = Pages.Count; }
    }

    public void Dispose() { _text.Dispose(); foreach (var b in _images.Values) b?.Dispose(); }

    // ------------------------------------------------------------ geometry
    SheetGeo Geometry(RenderedSheet s)
    {
        var g = new SheetGeo { S = s, Rows = s.Rows.Count, Cols = Math.Max(s.MaxCol, 1), Eval = new FormulaEvaluator(s) };
        g.Cf = new CondFormatEvaluator(s, g.Eval);
        g.ColW = new float[g.Cols + 2]; g.ColOff = new float[g.Cols + 3];
        for (int c = 1; c <= g.Cols; c++) g.ColW[c] = ColPt(s.DefaultColWidth);
        foreach (var cd in s.Columns)
            if (cd.Index >= 1 && cd.Index <= g.Cols) g.ColW[cd.Index] = cd.Hidden ? 0 : ColPt(cd.Width ?? s.DefaultColWidth);
        for (int c = 1; c <= g.Cols + 1; c++) g.ColOff[c + 1] = g.ColOff[c] + (c <= g.Cols ? g.ColW[c] : 0);
        // charts/images may sit below the last data row: add blank virtual rows so they have a place (and a page)
        int needed = g.Rows;
        foreach (var im in s.Images) needed = Math.Max(needed, im.Row - 1 + (int)Math.Ceiling(im.Height * PxToPt / s.DefaultRowHeight));
        foreach (var ch in s.Charts) needed = Math.Max(needed, ch.Row - 1 + (int)Math.Ceiling(ch.Def.Height * PxToPt / s.DefaultRowHeight));
        int dataRows = g.Rows; g.Rows = Math.Min(Math.Max(needed, g.Rows), 5000);
        g.RowH = new float[g.Rows + 2];
        for (int r = 1; r <= g.Rows; r++)
        {
            if (r > dataRows) { g.RowH[r] = (float)s.DefaultRowHeight; continue; }
            var row = s.Rows[r - 1]; g.RowH[r] = row.Hidden ? 0 : (float)(row.Height ?? s.DefaultRowHeight);
        }
        foreach (var m in s.Merges)
        {
            g.Anchors[(m[0], m[1])] = new[] { m[2], m[3] };
            for (int r = m[0]; r <= m[2]; r++) for (int c = m[1]; c <= m[3]; c++) if (r != m[0] || c != m[1]) g.Covered.Add((r, c));
        }
        AutoRowHeights(g, dataRows);
        return g;
    }

    /// <summary>Rows without an explicit height grow to fit wrapped text (like Excel's auto-fit).</summary>
    void AutoRowHeights(SheetGeo g, int dataRows)
    {
        var s = g.S; var tpl = _wb.Template;
        for (int r = 1; r <= dataRows && r <= s.Rows.Count; r++)
        {
            var row = s.Rows[r - 1];
            if (row.Height.HasValue || row.Hidden) continue;
            float need = g.RowH[r];
            foreach (var cell in row.Cells)
            {
                var st = cell.Style;
                if (st == null || !st.Wrap || st.Rotation != 0) continue;
                object? val = cell.Formula != null ? cell.Cached : cell.Value;
                if (cell.Formula != null && val == null && g.Eval.TryEvaluate(cell, out var ev)) val = ev;
                string text = val is FormulaEvaluator.XlError xe ? xe.Code : NumberFormatter.Format(val, st.Format);
                if (text.Length == 0) continue;
                if (_opt.PersianDigits || (s.Page?.PersianDigits ?? false)) text = TextEvaluator.ToPersianDigits(text);
                int c1 = cell.Col, c2 = cell.Col;
                if (g.Covered.Contains((r, c1))) continue;
                if (g.Anchors.TryGetValue((r, c1), out var span)) { if (span[0] != r) continue; c2 = Math.Min(span[1], g.Cols); }
                float width = 0; for (int c = c1; c <= c2; c++) width += g.ColW[c];
                float size = (float)(st.Size ?? tpl?.DefaultFontSize ?? 10);
                var face = _text.Face(st.Font ?? tpl?.DefaultFont, st.Bold, st.Italic, text);
                using var font = new SKFont(face, size) { Subpixel = true };
                bool paraRtl = st.Rtl ?? (TextPainter.HasRtl(text) || (s.Rtl && !(val is double || val is DateTime)));
                int lines = _text.Wrap(text, Math.Max(1, width - 2f - 2f - st.Indent * 6.75f), font, paraRtl).Count;
                if (lines > 1) need = Math.Max(need, lines * size * 1.2f + 3f);
            }
            g.RowH[r] = need;
        }
    }

    static float ColPt(double chars) => (float)Math.Round((chars * 7 + 5) * PxToPt, 2);

    static (float w, float h) Paper(PageSetupDef p)
    {
        float w, h;
        switch (p.PaperSize) { case 8: w = 842; h = 1191; break; case 11: w = 420; h = 595; break; case 1: w = 612; h = 792; break; case 5: w = 612; h = 1008; break; default: w = 595; h = 842; break; }
        return string.Equals(p.Orientation, "landscape", StringComparison.OrdinalIgnoreCase) ? (h, w) : (w, h);
    }

    void Paginate(int si)
    {
        var g = _geo[si]; var s = g.S; var pg = s.Page ?? new PageSetupDef();
        var (pw, ph) = Paper(pg);
        float ml = (float)pg.MarginLeft * 72, mr = (float)pg.MarginRight * 72, mt = (float)pg.MarginTop * 72, mb = (float)pg.MarginBottom * 72;
        float cw = pw - ml - mr, ch = ph - mt - mb;
        float total = g.ColOff[g.Cols + 1];
        float scale = pg.FitToWidth ? Math.Min(1f, total > 0 ? cw / total : 1f) : Math.Max(0.1f, (pg.Scale > 0 ? pg.Scale : 100) / 100f);

        // column pages
        var colPages = new List<(int a, int b)>(); int start = 1; float acc = 0;
        for (int c = 1; c <= g.Cols; c++)
        {
            if (acc + g.ColW[c] * scale > cw + 0.01f && c > start) { colPages.Add((start, c - 1)); start = c; acc = 0; }
            acc += g.ColW[c] * scale;
        }
        colPages.Add((start, g.Cols));

        // row pages (print titles repeat)
        var titles = new List<int>();
        if (!string.IsNullOrEmpty(s.PrintTitleRows)) { var p = s.PrintTitleRows.Split(':'); int a, b; if (p.Length == 2 && int.TryParse(p[0], out a) && int.TryParse(p[1], out b)) for (int r = a; r <= Math.Min(b, g.Rows); r++) titles.Add(r); }
        float titleH = titles.Sum(r => g.RowH[r]) * scale;
        // carried-forward totals reserve two extra rows on every page
        var carryCols = new List<int>(); int[]? carryRun = null;
        if (!string.IsNullOrWhiteSpace(pg.CarryBand) && !string.IsNullOrWhiteSpace(pg.CarryColumns) && s.BandRuns.TryGetValue(pg.CarryBand!.Trim(), out var br))
        {
            foreach (var part in pg.CarryColumns!.Split(',')) { var nm = part.Trim().ToUpperInvariant(); if (nm.Length > 0 && nm.Length <= 3 && nm.All(ch => ch >= 'A' && ch <= 'Z')) { int ci = CellRef.ColIndex(nm); if (ci >= 1 && ci <= g.Cols) carryCols.Add(ci); } }
            if (carryCols.Count > 0) carryRun = br;
        }
        float carryH = (float)s.DefaultRowHeight, reserve = carryRun != null ? 2 * carryH * scale : 0f;
        var rowPages = new List<List<int>>(); var cur = new List<int>(); float used = 0;
        bool first = true;
        for (int r = 1; r <= g.Rows; r++)
        {
            float h = g.RowH[r] * scale;
            if (h == 0) { cur.Add(r); continue; }
            float limit = ch - (first ? 0 : titleH) - reserve;
            bool forced = s.RowBreaks != null && s.RowBreaks.Contains(r - 1) && cur.Any(x => g.RowH[x] > 0);
            if (forced || (used + h > limit + 0.01f && cur.Any(x => g.RowH[x] > 0))) { rowPages.Add(cur); cur = new List<int>(); used = 0; first = false; }
            cur.Add(r); used += h;
        }
        if (cur.Count > 0 || rowPages.Count == 0) rowPages.Add(cur);

        // page-by-page sums of the carry columns over the detail band (formula cells such as group subtotals are not counted)
        var cumBefore = new double[rowPages.Count][]; var cumAfter = new double[rowPages.Count][]; var hasBand = new bool[rowPages.Count];
        if (carryRun != null)
        {
            var running = new double[carryCols.Count];
            for (int i = 0; i < rowPages.Count; i++)
            {
                cumBefore[i] = (double[])running.Clone();
                foreach (var r in rowPages[i])
                {
                    if (r < carryRun[0] || r > carryRun[1] || r > s.Rows.Count) continue;
                    hasBand[i] = true;
                    for (int k = 0; k < carryCols.Count; k++)
                    {
                        var cell = s.Get(r, carryCols[k]);
                        if (cell != null && cell.Formula == null && cell.Value is double d) running[k] += d;
                    }
                }
                cumAfter[i] = (double[])running.Clone();
            }
        }

        foreach (var cp in colPages)
            for (int i = 0; i < rowPages.Count; i++)
            {
                if (Pages.Count >= _opt.MaxPages) throw new ExcelRenderException("The PDF would exceed " + _opt.MaxPages + " pages. Filter the data or raise PdfOptions.MaxPages.");
                var lead = i == 0 ? new List<int>() : titles.Where(t => !rowPages[i].Contains(t)).ToList();
                var rows = lead.Concat(rowPages[i]).ToList();
                var spec = new PageSpec { Sheet = si, W = pw, H = ph, Scale = scale, ColFrom = cp.a, ColTo = cp.b, Rows = rows, TitleCount = lead.Count };
                if (carryRun != null && hasBand[i])
                {
                    bool before = hasBand.Take(i).Any(x => x), after = hasBand.Skip(i + 1).Any(x => x);
                    if (before) spec.CarryIn = cumBefore[i];
                    if (after) spec.CarryOut = cumAfter[i];
                    spec.CarryCols = carryCols.ToArray(); spec.CarryH = carryH; spec.CarryRun = carryRun;
                }
                Pages.Add(spec);
            }
    }

    // ------------------------------------------------------------ drawing
    public void Draw(SKCanvas canvas, PageSpec page)
    {
        var g = _geo[page.Sheet]; var s = g.S; var pg = s.Page ?? new PageSetupDef();
        float ml = (float)pg.MarginLeft * 72, mt = (float)pg.MarginTop * 72, mr = (float)pg.MarginRight * 72;
        float cw = page.W - ml - mr;
        bool rtl = s.Rtl;

        DrawWatermark(canvas, page, string.IsNullOrWhiteSpace(_opt.Watermark) ? pg.Watermark : _opt.Watermark);
        if (pg.PageFrame)
        {
            using var fp = new SKPaint { Color = Color(pg.FrameColor, new SKColor(0xBF, 0xBF, 0xBF)), Style = SKPaintStyle.Stroke, StrokeWidth = 1.4f, IsAntialias = true };
            canvas.DrawRect(new SKRect(ml - 2, mt - 2, page.W - mr + 2, page.H - (float)pg.MarginBottom * 72 + 2), fp);
        }
        DrawHeaderFooter(canvas, page, s, pg, ml, mr, mt);

        canvas.Save();
        canvas.Translate(ml, mt);
        canvas.Scale(page.Scale);

        float blockW = g.ColOff[page.ColTo + 1] - g.ColOff[page.ColFrom];
        float x0 = rtl ? cw / page.Scale - blockW : 0;
        var yOf = new Dictionary<int, float>(); float y = 0;
        float carryInY = -1, carryOutY = -1; int idx = 0;
        foreach (var r in page.Rows)
        {
            if (idx++ == page.TitleCount && page.CarryIn != null) { carryInY = y; y += page.CarryH; }
            yOf[r] = y; y += g.RowH[r];
        }
        if (page.CarryIn != null && carryInY < 0) { carryInY = y; y += page.CarryH; }
        if (page.CarryOut != null) { carryOutY = y; y += page.CarryH; }

        float Left(int c1, int c2)
        {
            float rel = g.ColOff[c1] - g.ColOff[page.ColFrom], w = g.ColOff[c2 + 1] - g.ColOff[c1];
            return rtl ? x0 + blockW - rel - w : x0 + rel;
        }

        var fill = new SKPaint { Style = SKPaintStyle.Fill, IsAntialias = false };
        foreach (var r in page.Rows)
        {
            if (g.RowH[r] == 0) continue;
            if (r > s.Rows.Count) continue;
            var row = s.Rows[r - 1];
            foreach (var cell in row.Cells)
            {
                int c = cell.Col; if (c < page.ColFrom || c > page.ColTo || g.ColW[c] == 0) continue;
                if (g.Covered.Contains((r, c))) continue;
                int r2 = r, c2 = c;
                if (g.Anchors.TryGetValue((r, c), out var span)) { r2 = span[0]; c2 = Math.Min(span[1], g.Cols); }
                float left = Left(c, c2), width = g.ColOff[c2 + 1] - g.ColOff[c];
                float top = yOf[r], height = 0;
                for (int rr = r; rr <= r2 && rr <= g.Rows; rr++) height += g.RowH[rr];
                var rect = new SKRect(left, top, left + width, top + height);
                DrawCell(canvas, g, cell, rect, fill, rtl, r, page, c, c2);
            }
        }
        if (carryInY >= 0) DrawCarryRow(canvas, g, page, carryInY, pg.CarryLabelIn ?? (rtl ? "انتقال از صفحه قبل" : "Brought forward"), page.CarryIn!, Left, blockW, x0, rtl);
        if (carryOutY >= 0) DrawCarryRow(canvas, g, page, carryOutY, pg.CarryLabelOut ?? (rtl ? "انتقال به صفحه بعد" : "Carried forward"), page.CarryOut!, Left, blockW, x0, rtl);
        if (_opt.PrintGridLines) DrawGrid(canvas, g, page, yOf, Left);

        // images & charts anchored on this page
        foreach (var im in s.Images)
        {
            if (!yOf.TryGetValue(im.Row, out var iy) || im.Col < page.ColFrom || im.Col > page.ColTo) continue;
            float w = im.Width * PxToPt, h = im.Height * PxToPt, left = rtl ? Left(im.Col, im.Col) + (g.ColOff[im.Col + 1] - g.ColOff[im.Col]) - w : Left(im.Col, im.Col);
            if (im.BarType != null && DrawBarcode(canvas, im, new SKRect(left, iy, left + w, iy + h))) continue;   // crisp vector barcode
            var bmp = Bitmap(im);
            if (bmp == null) continue;
            canvas.DrawBitmap(bmp, new SKRect(left, iy, left + w, iy + h));
        }
        foreach (var ch in s.Charts)
        {
            if (!yOf.TryGetValue(ch.Row, out var cy) || ch.Col < page.ColFrom || ch.Col > page.ColTo) continue;
            float w = ch.Def.Width * PxToPt, h = ch.Def.Height * PxToPt, left = rtl ? Left(ch.Col, ch.Col) + (g.ColOff[ch.Col + 1] - g.ColOff[ch.Col]) - w : Left(ch.Col, ch.Col);
            new ChartPainter(_wb, _text, g.S.Rtl).Draw(canvas, ch, new SKRect(left, cy, left + w, cy + h));
        }
        fill.Dispose();
        canvas.Restore();
    }

    /// <summary>A full-width "brought forward / carried forward" row: label at the start, running totals under the carried columns.</summary>
    void DrawCarryRow(SKCanvas canvas, SheetGeo g, PageSpec page, float y, string label, double[] values, Func<int, int, float> left, float blockW, float x0, bool rtl)
    {
        var rect = new SKRect(x0, y, x0 + blockW, y + page.CarryH);
        using (var bg = new SKPaint { Color = new SKColor(0xF2, 0xF2, 0xF2), Style = SKPaintStyle.Fill }) canvas.DrawRect(rect, bg);
        using (var ln = new SKPaint { Color = new SKColor(0x7F, 0x7F, 0x7F), StrokeWidth = 0.6f, Style = SKPaintStyle.Stroke })
        { canvas.DrawLine(rect.Left, rect.Top, rect.Right, rect.Top, ln); canvas.DrawLine(rect.Left, rect.Bottom, rect.Right, rect.Bottom, ln); }

        float size = (float)(_wb.Template?.DefaultFontSize ?? 10);
        var face = _text.Face(_wb.Template?.DefaultFont, true, false, label + "0123456789");
        using var font = new SKFont(face, size) { Subpixel = true };
        using var paint = new SKPaint { Color = new SKColor(0x26, 0x32, 0x38), IsAntialias = true };
        var m = font.Metrics; float baseline = y + (page.CarryH - (m.Descent - m.Ascent)) / 2 - m.Ascent;
        bool lblRtl = TextPainter.HasRtl(label) || rtl;
        float lw = _text.Measure(label, font, lblRtl);
        _text.Draw(canvas, label, rtl ? rect.Right - 4 - lw : rect.Left + 4, baseline, font, paint, lblRtl);

        for (int k = 0; k < page.CarryCols.Length; k++)
        {
            int c = page.CarryCols[k];
            if (c < page.ColFrom || c > page.ColTo || g.ColW[c] == 0) continue;
            string fmt = "#,##0";
            if (page.CarryRun != null)
                for (int r = page.CarryRun[0]; r <= page.CarryRun[1] && r <= g.S.Rows.Count; r++)
                { var cell = g.S.Get(r, c); if (cell?.Style != null && !string.IsNullOrEmpty(cell.Style.Format) && cell.Value is double) { fmt = cell.Style.Format!; break; } }
            var text = NumberFormatter.Format(values[k], fmt);
            float w = _text.Measure(text, font, false);
            float cl = left(c, c), cwid = g.ColOff[c + 1] - g.ColOff[c];
            _text.Draw(canvas, text, cl + cwid - 2f - w, baseline, font, paint, false);
        }
    }

    static bool DrawBarcode(SKCanvas canvas, RenderedImage im, SKRect box)
    {
        try
        {
            var m = Barcodes.Encode(im.BarType!, im.BarText!);
            int totalCols = m.Cols + 2 * m.Quiet, totalRows = m.Is2D ? m.Rows + 2 * m.Quiet : 1;
            float mod = Math.Min(box.Width / totalCols, m.Is2D ? box.Height / totalRows : float.MaxValue);
            float x0 = box.Left + (box.Width - mod * totalCols) / 2, y0 = box.Top + (m.Is2D ? (box.Height - mod * totalRows) / 2 : 0);
            using var paint = new SKPaint { Color = SKColors.Black, IsAntialias = false, Style = SKPaintStyle.Fill };
            for (int r = 0; r < (m.Is2D ? m.Rows : 1); r++)
            {
                int c = 0;
                while (c < m.Cols)
                {
                    if (!m.Modules[r, c]) { c++; continue; }
                    int start = c;
                    while (c < m.Cols && m.Modules[r, c]) c++;     // merge horizontal runs: no hairline seams
                    float x = x0 + (start + m.Quiet) * mod;
                    float y = m.Is2D ? y0 + (r + m.Quiet) * mod : y0;
                    canvas.DrawRect(x, y, (c - start) * mod, m.Is2D ? mod : box.Height, paint);
                }
            }
            return true;
        }
        catch (ArgumentException) { return false; }
    }

    void DrawWatermark(SKCanvas canvas, PageSpec page, string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return;
        text = text!.Trim();
        var face = _text.Face(_wb.Template?.DefaultFont, true, false, text);
        using var font = new SKFont(face, 100) { Subpixel = true };
        bool rtl = TextPainter.HasRtl(text);
        float w = _text.Measure(text, font, rtl);
        float diag = (float)Math.Sqrt(page.W * page.W + page.H * page.H);
        font.Size = Math.Max(12f, Math.Min(140f, 100f * diag * 0.55f / Math.Max(1f, w)));
        w = _text.Measure(text, font, rtl);
        using var paint = new SKPaint { Color = new SKColor(0x80, 0x80, 0x80, 55), IsAntialias = true };
        canvas.Save();
        canvas.Translate(page.W / 2, page.H / 2);
        canvas.RotateDegrees(-(float)(Math.Atan2(page.H, page.W) * 180 / Math.PI) * 0.8f);
        _text.Draw(canvas, text, -w / 2, font.Size * 0.3f, font, paint, rtl);
        canvas.Restore();
    }

    void DrawGrid(SKCanvas canvas, SheetGeo g, PageSpec page, Dictionary<int, float> yOf, Func<int, int, float> left)
    {
        using var p = new SKPaint { Color = new SKColor(0xD9, 0xD9, 0xD9), StrokeWidth = 0.4f, Style = SKPaintStyle.Stroke };
        foreach (var r in page.Rows)
            for (int c = page.ColFrom; c <= page.ColTo; c++)
            {
                if (g.ColW[c] == 0 || g.RowH[r] == 0) continue;
                float l = left(c, c), t = yOf[r];
                canvas.DrawRect(l, t, g.ColW[c], g.RowH[r], p);
            }
    }

    SKBitmap? Bitmap(RenderedImage im)
    {
        var key = im.GetHashCode().ToString();
        if (!_images.TryGetValue(key, out var b)) _images[key] = b = SKBitmap.Decode(im.Bytes);
        return b;
    }

    static SKColor Color(string? hex, SKColor fallback)
    {
        if (string.IsNullOrWhiteSpace(hex)) return fallback;
        var h = hex!.Trim().TrimStart('#');
        if (h.Length == 3) h = "" + h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
        return h.Length == 6 && uint.TryParse(h, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v)
            ? new SKColor((byte)(v >> 16), (byte)(v >> 8), (byte)v) : fallback;
    }

    static bool EmptyCell(SheetGeo g, int r, int c)
    {
        if (c < 1 || c > g.Cols || g.Covered.Contains((r, c)) || g.Anchors.ContainsKey((r, c))) return false;
        var x = g.S.Get(r, c);
        return x == null || (x.Formula == null && (x.Value == null || (x.Value is string str && str.Length == 0)));
    }

    void DrawCell(SKCanvas canvas, SheetGeo g, RenderedCell cell, SKRect rect, SKPaint fill, bool sheetRtl, int row, PageSpec page, int c1, int c2)
    {
        var st = cell.Style;
        var tpl = _wb.Template;

        // value
        object? val = cell.Value;
        if (cell.Formula != null)
        {
            val = cell.Cached;
            if (val == null && g.Eval.TryEvaluate(cell, out var ev)) val = ev;
        }
        string text = val is FormulaEvaluator.XlError xe ? xe.Code : NumberFormatter.Format(val, st?.Format);
        bool persian = _opt.PersianDigits || (g.S.Page?.PersianDigits ?? false);
        if (persian) text = TextEvaluator.ToPersianDigits(text);

        // conditional formatting (same evaluator as the HTML preview)
        CfResult? cf = g.Cf.HasRules ? g.Cf.Evaluate(row, cell.Col, val, text) : null;
        if (cf?.Style != null)
        {
            var m = st?.Clone() ?? new StyleDef();
            if (!string.IsNullOrEmpty(cf.Style.Fill)) m.Fill = cf.Style.Fill;
            if (!string.IsNullOrEmpty(cf.Style.Color)) m.Color = cf.Style.Color;
            m.Bold |= cf.Style.Bold; m.Italic |= cf.Style.Italic; m.Underline |= cf.Style.Underline; m.Strike |= cf.Style.Strike;
            st = m;
        }
        string? bg = cf?.ScaleFill ?? (st != null && !string.IsNullOrEmpty(st.Fill) ? st.Fill : null);
        if (bg != null) { fill.Color = Color(bg, SKColors.White); canvas.DrawRect(rect, fill); }
        if (cf?.BarFraction != null)
        {
            float bw = rect.Width * (float)Math.Max(0.02, cf.BarFraction.Value);
            fill.Color = Color(cf.BarColor, new SKColor(0x63, 0x8E, 0xC6)).WithAlpha(190);
            var bar = sheetRtl ? new SKRect(rect.Right - bw, rect.Top + 1, rect.Right, rect.Bottom - 1) : new SKRect(rect.Left, rect.Top + 1, rect.Left + bw, rect.Bottom - 1);
            canvas.DrawRect(bar, fill);
        }

        if (text.Length > 0)
        {
            float size = (float)(st?.Size ?? tpl?.DefaultFontSize ?? 10);
            bool bold = st?.Bold ?? false, italic = st?.Italic ?? false;
            var face = _text.Face(st?.Font ?? tpl?.DefaultFont, bold, italic, text);
            using var font = new SKFont(face, size) { Subpixel = true };
            using var paint = new SKPaint { Color = Color(st?.Color, SKColors.Black), IsAntialias = true };
            bool paraRtl = st?.Rtl ?? (TextPainter.HasRtl(text) || (sheetRtl && !(val is double || val is DateTime)));
            bool numeric = val is double || val is DateTime;
            string h = (st?.H ?? "").ToLowerInvariant();
            float pad = 2f + (st?.Indent ?? 0) * 6.75f;
            float avail = Math.Max(1, rect.Width - pad - 2f);
            string v = (st?.V ?? "").ToLowerInvariant();
            int rot = st?.Rotation ?? 0;

            if (rot != 0)
            {
                DrawRotated(canvas, text, rect, font, paint, paraRtl, rot, h, v);
            }
            else
            {
                // shrink to fit (single line)
                if (st?.Shrink == true && st.Wrap != true)
                {
                    float w0 = _text.Measure(text, font, paraRtl);
                    if (w0 > avail) font.Size = Math.Max(4f, size * avail / w0);
                }
                var lines = st?.Wrap == true ? _text.Wrap(text, avail, font, paraRtl) : new List<string> { text.Replace("\n", " ") };
                var m = font.Metrics; float lineH = font.Size * 1.2f; float blockH = lines.Count * lineH;
                float top = v == "top" ? rect.Top + 1f : v == "center" || v == "middle" ? rect.Top + (rect.Height - blockH) / 2 : rect.Bottom - blockH - 1f;
                if (top < rect.Top) top = rect.Top;
                string align = h is "left" or "right" or "center" ? h : (h == "justify" || h == "distributed" ? (paraRtl ? "right" : "left") : numeric ? "right" : (paraRtl ? "right" : "left"));

                // text spills into empty neighbouring cells, like in Excel (single-line, left/right aligned, not numbers)
                var clip = rect;
                if (st?.Wrap != true && align != "center" && !numeric)
                {
                    float need = _text.Measure(lines[0], font, paraRtl) + pad + 2f - rect.Width;
                    if (need > 0)
                    {
                        bool toVisualRight = align == "left";
                        // the visual-right neighbour is the previous logical column on an RTL sheet, the next one otherwise
                        int step = (toVisualRight ^ sheetRtl) ? 1 : -1;
                        int cc = step > 0 ? c2 + 1 : c1 - 1; float extra = 0;
                        while (extra < need && cc >= page.ColFrom && cc <= page.ColTo && g.ColW[cc] > 0 && EmptyCell(g, row, cc)) { extra += g.ColW[cc]; cc += step; }
                        if (extra > 0) { if (toVisualRight) clip.Right += extra; else clip.Left -= extra; }
                    }
                }

                canvas.Save(); canvas.ClipRect(clip);
                float baseline = top + (lineH - (m.Descent - m.Ascent)) / 2 - m.Ascent;
                foreach (var line in lines)
                {
                    float w = _text.Measure(line, font, paraRtl);
                    float x;
                    if (align == "center") x = rect.Left + (rect.Width - w) / 2;
                    else if (align == "right") x = rect.Right - 2f - (st?.Indent ?? 0) * 6.75f - w;
                    else x = rect.Left + pad;
                    _text.Draw(canvas, line, x, baseline, font, paint, paraRtl);
                    if (st?.Underline == true) { using var up = new SKPaint { Color = paint.Color, StrokeWidth = font.Size / 14f }; canvas.DrawLine(x, baseline + 1.5f, x + w, baseline + 1.5f, up); }
                    if (st?.Strike == true) { using var sp = new SKPaint { Color = paint.Color, StrokeWidth = font.Size / 14f }; canvas.DrawLine(x, baseline - font.Size * 0.3f, x + w, baseline - font.Size * 0.3f, sp); }
                    baseline += lineH;
                }
                canvas.Restore();
            }
        }

        if (st != null) DrawBorders(canvas, st, rect);
        if (!string.IsNullOrEmpty(cell.Link) && (cell.Link.StartsWith("http", StringComparison.OrdinalIgnoreCase) || cell.Link.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)))
            canvas.DrawUrlAnnotation(rect, cell.Link);
    }

    /// <summary>Rotated text (Excel semantics: positive = counter-clockwise, 255 = vertical stack), clipped to the cell.</summary>
    void DrawRotated(SKCanvas canvas, string text, SKRect rect, SKFont font, SKPaint paint, bool paraRtl, int rot, string h, string v)
    {
        canvas.Save(); canvas.ClipRect(rect);
        var m = font.Metrics; float lineH = font.Size * 1.2f;
        text = text.Replace("\n", " ");
        if (rot == 255)
        {
            // characters stacked top to bottom
            var chars = text.Select(ch => ch.ToString()).ToList();
            float total = chars.Count * lineH;
            float y = v == "top" ? rect.Top + 1 : v == "bottom" ? rect.Bottom - total - 1 : rect.Top + (rect.Height - total) / 2;
            float baseY = Math.Max(rect.Top, y) + (lineH - (m.Descent - m.Ascent)) / 2 - m.Ascent;
            foreach (var chr in chars)
            {
                float w = _text.Measure(chr, font, paraRtl);
                _text.Draw(canvas, chr, rect.Left + (rect.Width - w) / 2, baseY, font, paint, paraRtl);
                baseY += lineH;
            }
        }
        else
        {
            float angle = Math.Max(-90, Math.Min(90, rot));
            float w = _text.Measure(text, font, paraRtl);
            float sin = Math.Abs((float)Math.Sin(angle * Math.PI / 180));
            float cx = h == "left" ? rect.Left + lineH / 2 + 1 : h == "right" ? rect.Right - lineH / 2 - 1 : rect.MidX;
            float cy = v == "top" ? rect.Top + w / 2 * sin + lineH / 2 : v == "bottom" ? rect.Bottom - w / 2 * sin - lineH / 2 : rect.MidY;
            canvas.Translate(cx, cy);
            canvas.RotateDegrees(-angle);
            float baseline = (lineH - (m.Descent - m.Ascent)) / 2 - m.Ascent - lineH / 2;
            _text.Draw(canvas, text, -w / 2, baseline, font, paint, paraRtl);
        }
        canvas.Restore();
    }

    static void DrawBorders(SKCanvas canvas, StyleDef st, SKRect r)
    {
        void Side(BorderDef? b, float x1, float y1, float x2, float y2)
        {
            if (b == null || string.IsNullOrEmpty(b.Style)) return;
            float w = b.Style switch { "medium" => 1.5f, "thick" => 2.25f, "hair" => 0.25f, "double" => 1.8f, _ => 0.75f };
            using var p = new SKPaint { Color = Color(b.Color, SKColors.Black), StrokeWidth = w, Style = SKPaintStyle.Stroke, IsAntialias = false };
            if (b.Style is "dashed" or "mediumdashed") p.PathEffect = SKPathEffect.CreateDash(new[] { 3f, 2f }, 0);
            else if (b.Style is "dotted" or "hair") p.PathEffect = SKPathEffect.CreateDash(new[] { 1f, 1.5f }, 0);
            canvas.DrawLine(x1, y1, x2, y2, p);
        }
        Side(st.Top, r.Left, r.Top, r.Right, r.Top);
        Side(st.Bottom, r.Left, r.Bottom, r.Right, r.Bottom);
        Side(st.Left, r.Left, r.Top, r.Left, r.Bottom);
        Side(st.Right, r.Right, r.Top, r.Right, r.Bottom);
    }

    // ------------------------------------------------------------ header / footer
    void DrawHeaderFooter(SKCanvas canvas, PageSpec page, RenderedSheet s, PageSetupDef pg, float ml, float mr, float mt)
    {
        void Draw(string? spec, bool header)
        {
            if (spec == null || string.IsNullOrWhiteSpace(spec)) return;
            string L = "", C = "", R = ""; string cur = "C";
            var sb = new System.Text.StringBuilder();
            void Flush() { var t = sb.ToString(); sb.Clear(); if (cur == "L") L += t; else if (cur == "R") R += t; else C += t; }
            for (int i = 0; i < spec.Length; i++)
            {
                if (spec[i] != '&' || i + 1 >= spec.Length) { sb.Append(spec[i]); continue; }
                char c = spec[++i];
                switch (char.ToUpperInvariant(c))
                {
                    case 'L': Flush(); cur = "L"; break;
                    case 'C': Flush(); cur = "C"; break;
                    case 'R': Flush(); cur = "R"; break;
                    case 'P': sb.Append(page.Number); break;
                    case 'N': sb.Append(page.Total); break;
                    case 'D': sb.Append(DateTime.Now.ToString("yyyy/MM/dd")); break;
                    case 'T': sb.Append(DateTime.Now.ToString("HH:mm")); break;
                    case 'A': sb.Append(s.Name); break;
                    case 'F': sb.Append(_wb.Template?.Title ?? ""); break;
                    case '&': sb.Append('&'); break;
                    case '"': { int e = spec.IndexOf('"', i + 1); i = e < 0 ? spec.Length : e; break; }
                    default: if (char.IsDigit(c)) { while (i + 1 < spec.Length && char.IsDigit(spec[i + 1])) i++; } break;
                }
            }
            Flush();
            if (_opt.PersianDigits || pg.PersianDigits) { L = TextEvaluator.ToPersianDigits(L); C = TextEvaluator.ToPersianDigits(C); R = TextEvaluator.ToPersianDigits(R); }
            float y = header ? Math.Max(14, mt / 2) : page.H - Math.Max(14, (float)pg.MarginBottom * 72 / 2) + 3;
            foreach (var (t, align) in new[] { (L, "L"), (C, "C"), (R, "R") })
            {
                if (t.Length == 0) continue;
                var face = _text.Face(_wb.Template?.DefaultFont, false, false, t);
                using var font = new SKFont(face, 8.5f); using var paint = new SKPaint { Color = new SKColor(0x55, 0x55, 0x55), IsAntialias = true };
                bool rtl = TextPainter.HasRtl(t);
                float w = _text.Measure(t, font, rtl);
                float x = align == "L" ? ml : align == "R" ? page.W - mr - w : (page.W - w) / 2;
                _text.Draw(canvas, t, x, y, font, paint, rtl);
            }
        }
        Draw(pg.Header, true);
        Draw(pg.Footer, false);
    }
}
