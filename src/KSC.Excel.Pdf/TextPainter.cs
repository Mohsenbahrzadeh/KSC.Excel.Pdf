using System.Text;
using HarfBuzzSharp;
using SkiaSharp;
using SkiaSharp.HarfBuzz;

namespace KSC.Excel.Pdf;

/// <summary>
/// Draws one line of mixed Persian/Latin/digit text correctly: Arabic-script runs are shaped with HarfBuzz (joining forms),
/// Latin runs and digits are drawn left-to-right, and runs are ordered visually for the paragraph direction.
/// </summary>
internal sealed class TextPainter : IDisposable
{
    enum Cls { R, L, N }

    sealed record Run(string Text, bool Rtl);

    readonly Dictionary<(string, bool, bool), SKTypeface> _faces = new();
    readonly Dictionary<SKTypeface, SKShaper> _shapers = new();
    readonly PdfOptions _opt;

    public TextPainter(PdfOptions opt) { _opt = opt; }

    // ------------------------------------------------------------ fonts
    public SKTypeface Face(string? family, bool bold, bool italic, string sample)
    {
        var key = (family ?? "", bold, italic);
        if (_faces.TryGetValue(key, out var f) && Covers(f, sample)) return f;

        var candidates = new List<string>();
        if (!string.IsNullOrWhiteSpace(family)) candidates.Add(family!);
        candidates.Add(_opt.FallbackFontFamily);
        candidates.AddRange(new[] { "Tahoma", "Segoe UI", "Arial", "Noto Naskh Arabic", "Noto Sans Arabic", "DejaVu Sans", "Liberation Sans" });

        SKTypeface? best = null;
        if (!string.IsNullOrEmpty(_opt.FontFile) && File.Exists(_opt.FontFile)) best = SKTypeface.FromFile(_opt.FontFile);
        if (best == null || !Covers(best, sample))
        {
            best = null;
            foreach (var c in candidates)
            {
                var t = SKTypeface.FromFamilyName(c, bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal, SKFontStyleWidth.Normal, italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);
                if (t != null && Covers(t, sample)) { best = t; break; }
            }
        }
        if (best == null)
        {
            // ask the OS font manager for any font with the first uncovered character
            foreach (var ch in sample.Where(c => !char.IsWhiteSpace(c)))
            {
                var m = SKFontManager.Default.MatchCharacter(ch);
                if (m != null) { best = m; break; }
            }
        }
        best ??= SKTypeface.Default;
        _faces[key] = best;
        return best;
    }

    static bool Covers(SKTypeface t, string sample)
    {
        foreach (var ch in sample)
        {
            if (char.IsWhiteSpace(ch) || char.IsControl(ch)) continue;
            if (t.GetGlyphs(ch.ToString())[0] == 0) return false;
        }
        return true;
    }

    SKShaper Shaper(SKTypeface t)
    {
        if (!_shapers.TryGetValue(t, out var s)) _shapers[t] = s = new SKShaper(t);
        return s;
    }

    // ------------------------------------------------------------ bidi
    static Cls Classify(char c)
    {
        if ((c >= 0x0590 && c <= 0x08FF && !(c >= 0x0660 && c <= 0x0669) && !(c >= 0x06F0 && c <= 0x06F9)) || (c >= 0xFB1D && c <= 0xFDFF) || (c >= 0xFE70 && c <= 0xFEFF)) return Cls.R;
        if (char.IsLetterOrDigit(c)) return Cls.L;      // Latin letters and all digits (incl. Persian digits) run left-to-right
        return Cls.N;
    }

    public static bool HasRtl(string s) => s.Any(c => Classify(c) == Cls.R);

    static List<Run> Runs(string text, bool paragraphRtl)
    {
        var cls = text.Select(Classify).ToArray();
        for (int i = 0; i < cls.Length; i++)
        {
            if (cls[i] != Cls.N) continue;
            int p = i - 1; while (p >= 0 && cls[p] == Cls.N) p--;
            int n = i + 1; while (n < cls.Length && cls[n] == Cls.N) n++;
            var before = p >= 0 ? cls[p] : (Cls?)null; var after = n < cls.Length ? cls[n] : (Cls?)null;
            cls[i] = before.HasValue && after.HasValue && before == after ? before.Value : (paragraphRtl ? Cls.R : Cls.L);
        }
        var runs = new List<Run>(); var sb = new StringBuilder(); Cls? cur = null;
        for (int i = 0; i < text.Length; i++)
        {
            if (cur != null && cls[i] != cur) { runs.Add(new Run(sb.ToString(), cur == Cls.R)); sb.Clear(); }
            cur = cls[i]; sb.Append(text[i]);
        }
        if (sb.Length > 0) runs.Add(new Run(sb.ToString(), cur == Cls.R));
        if (paragraphRtl) runs.Reverse();   // visual order: logical first run sits at the right
        return runs;
    }

    // ------------------------------------------------------------ measure / draw
    public float Measure(string text, SKFont font, bool paragraphRtl)
    {
        float w = 0;
        foreach (var r in Runs(text, paragraphRtl)) w += RunWidth(r, font);
        return w;
    }

    float RunWidth(Run r, SKFont font)
    {
        if (r.Rtl) return Shaper(font.Typeface).Shape(r.Text, font).Width;
        return font.MeasureText(r.Text);
    }

    /// <summary>Draws the line with its left edge at <paramref name="x"/> (baseline y).</summary>
    public void Draw(SKCanvas canvas, string text, float x, float y, SKFont font, SKPaint paint, bool paragraphRtl)
    {
        foreach (var r in Runs(text, paragraphRtl))
        {
            if (r.Rtl)
            {
                var shaper = Shaper(font.Typeface);
                canvas.DrawShapedText(shaper, r.Text, x, y, font, paint);
                x += shaper.Shape(r.Text, font).Width;
            }
            else
            {
                canvas.DrawText(r.Text, x, y, font, paint);
                x += font.MeasureText(r.Text);
            }
        }
    }

    public List<string> Wrap(string text, float maxWidth, SKFont font, bool rtl)
    {
        var lines = new List<string>();
        foreach (var para in text.Replace("\r", "").Split('\n'))
        {
            var words = para.Split(' ');
            var cur = "";
            foreach (var w in words)
            {
                var trial = cur.Length == 0 ? w : cur + " " + w;
                if (cur.Length > 0 && Measure(trial, font, rtl) > maxWidth) { lines.Add(cur); cur = w; } else cur = trial;
            }
            lines.Add(cur);
        }
        return lines;
    }

    public void Dispose() { foreach (var s in _shapers.Values) s.Dispose(); }
}
