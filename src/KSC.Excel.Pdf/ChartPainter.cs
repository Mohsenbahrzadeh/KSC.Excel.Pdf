using System.Text.RegularExpressions;
using KSC.Excel.Engine;
using SkiaSharp;

namespace KSC.Excel.Pdf;

/// <summary>Simple, clean charts (column, bar, line, area, pie) read from the rendered cells the chart refers to.</summary>
internal sealed class ChartPainter
{
    static readonly string[] Palette = { "2E75B6", "C55A11", "548235", "7030A0", "BF9000", "C00000", "00A3A3", "7F7F7F" };
    static readonly Regex RefRx = new(@"^'?(.*?)'?!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$", RegexOptions.Compiled);

    readonly RenderedWorkbook _wb; readonly TextPainter _text; readonly bool _rtl;
    readonly Dictionary<string, FormulaEvaluator> _eval = new();

    public ChartPainter(RenderedWorkbook wb, TextPainter text, bool rtl) { _wb = wb; _text = text; _rtl = rtl; }

    List<object?> Read(string reference)
    {
        var res = new List<object?>();
        var m = RefRx.Match(reference);
        if (!m.Success) return res;
        var name = m.Groups[1].Value.Replace("''", "'");
        var sh = _wb.Sheets.FirstOrDefault(x => x.Name == name);
        if (sh == null) return res;
        int c1 = CellRef.ColIndex(m.Groups[2].Value), r1 = int.Parse(m.Groups[3].Value);
        int c2 = m.Groups[4].Success ? CellRef.ColIndex(m.Groups[4].Value) : c1, r2 = m.Groups[5].Success ? int.Parse(m.Groups[5].Value) : r1;
        if (!_eval.TryGetValue(name, out var ev)) _eval[name] = ev = new FormulaEvaluator(sh);
        for (int r = r1; r <= r2; r++)
            for (int c = c1; c <= c2; c++)
            {
                var cell = sh.Get(r, c); object? v = cell?.Value;
                if (cell?.Formula != null) { v = cell.Cached; if (v == null && ev.TryEvaluate(cell, out var x)) v = x; }
                res.Add(v);
            }
        return res;
    }

    static double Num(object? v) => v switch { double d => d, DateTime dt => dt.ToOADate(), bool b => b ? 1 : 0, _ => double.TryParse(v?.ToString(), out var d2) ? d2 : 0 };
    static string Label(object? v) => v is DateTime dt ? dt.ToString("yyyy/MM/dd") : v is double d ? d.ToString("0.##") : v?.ToString() ?? "";

    public void Draw(SKCanvas canvas, RenderedChart ch, SKRect box)
    {
        using var bg = new SKPaint { Color = SKColors.White, Style = SKPaintStyle.Fill };
        using var border = new SKPaint { Color = new SKColor(0xBF, 0xBF, 0xBF), Style = SKPaintStyle.Stroke, StrokeWidth = 0.6f };
        canvas.DrawRect(box, bg); canvas.DrawRect(box, border);

        var cats = ch.Categories != null ? Read(ch.Categories).Select(Label).ToList() : new List<string>();
        var series = ch.SeriesValues.Select((v, i) => new Ser
        {
            Name = ch.SeriesNameIsRef[i] ? string.Concat(Read(ch.SeriesNames[i]).Select(Label)) : ch.SeriesNames[i],
            Vals = Read(v).Select(Num).ToList(),
            Color = ParseColor(ch.SeriesColors[i], i),
            Kind = SeriesKind(ch, i),
            Sec = i < ch.SeriesSecondary.Count && ch.SeriesSecondary[i] && SeriesKind(ch, i) != "pie"
        }).ToList();
        int n = series.Max(s => s.Vals.Count);
        if (n == 0) return;
        while (cats.Count < n) cats.Add((cats.Count + 1).ToString());

        var type = (ch.Def.Type ?? "column").ToLowerInvariant();
        float top = box.Top + 6, bottom = box.Bottom - 6, left = box.Left + 6, right = box.Right - 6;
        using (var tf = new SKFont(_text.Face(_wb.Template?.DefaultFont, true, false, ch.Def.Title ?? "A"), 10))
        using (var tp = new SKPaint { Color = new SKColor(0x26, 0x32, 0x38), IsAntialias = true })
            if (!string.IsNullOrWhiteSpace(ch.Def.Title))
            {
                var title = ch.Def.Title ?? "";
                bool rtl = TextPainter.HasRtl(title);
                float w = _text.Measure(title, tf, rtl);
                _text.Draw(canvas, title, box.MidX - w / 2, top + 10, tf, tp, rtl);
                top += 20;
            }

        bool legend = ch.Def.ShowLegend && (type == "pie" || series.Count > 1);
        if (legend) { DrawLegend(canvas, type == "pie" ? cats.Take(n).Select((c, i) => (c, ParseColor(null, i))).ToList() : series.Select(s => (s.Name, s.Color)).ToList(), box, ref bottom); }

        if (type == "pie") { DrawPie(canvas, series[0].Vals, new SKRect(left, top, right, bottom)); return; }

        // value axes (primary; optional secondary on the right for combo charts)
        bool horiz = type == "bar";
        var prim = series.Where(s => !s.Sec).ToList(); if (prim.Count == 0) prim = series;
        var secs = horiz || prim == series ? new List<Ser>() : series.Where(s => s.Sec).ToList();
        bool stackPrim = ch.Def.Stacked;
        double max = Math.Max(1e-9, stackPrim ? Enumerable.Range(0, n).Max(i => prim.Where(s => s.Kind != "line").Sum(s => i < s.Vals.Count ? Math.Max(0, s.Vals[i]) : 0)) : prim.Max(s => s.Vals.DefaultIfEmpty(0).Max()));
        if (stackPrim) max = Math.Max(max, prim.Max(s => s.Vals.DefaultIfEmpty(0).Max()));
        double min = Math.Min(0, prim.Min(s => s.Vals.DefaultIfEmpty(0).Min()));
        var (step, top2) = Nice(min, max);
        double max2 = secs.Count > 0 ? Math.Max(1e-9, secs.Max(s => s.Vals.DefaultIfEmpty(0).Max())) : 1, min2 = secs.Count > 0 ? Math.Min(0, (double)secs.Min(s => s.Vals.DefaultIfEmpty(0).Min())) : 0;
        var (step2, top22) = Nice(min2, max2);
        float labelW = 38, labelH = 18;
        var plot = new SKRect(left + (horiz ? 56 : labelW), top + 4, right - 4 - (secs.Count > 0 ? 34 : 0), bottom - labelH);
        if (plot.Width < 20 || plot.Height < 20) return;

        using var gridP = new SKPaint { Color = new SKColor(0xE0, 0xE0, 0xE0), StrokeWidth = 0.4f, Style = SKPaintStyle.Stroke };
        using var axisFont = new SKFont(_text.Face(_wb.Template?.DefaultFont, false, false, "0123456789"), 7.5f);
        using var axisPaint = new SKPaint { Color = new SKColor(0x59, 0x59, 0x59), IsAntialias = true };
        double span = top2 - min;
        float V(double v) => horiz ? plot.Left + (float)((v - min) / span) * plot.Width : plot.Bottom - (float)((v - min) / span) * plot.Height;
        double span2 = Math.Max(1e-9, top22 - min2);
        float V2(double v) => plot.Bottom - (float)((v - min2) / span2) * plot.Height;
        if (secs.Count > 0)
            for (double t = min2; t <= top22 + step2 / 2; t += step2)
                canvas.DrawText(t.ToString("#,##0.##"), plot.Right + 4, V2(t) + 3, axisFont, axisPaint);
        for (double t = min; t <= top2 + step / 2; t += step)
        {
            float p = V(t);
            if (horiz) canvas.DrawLine(p, plot.Top, p, plot.Bottom, gridP); else canvas.DrawLine(plot.Left, p, plot.Right, p, gridP);
            var lab = t.ToString("#,##0.##");
            if (horiz) canvas.DrawText(lab, p - axisFont.MeasureText(lab) / 2, plot.Bottom + 10, axisFont, axisPaint);
            else canvas.DrawText(lab, plot.Left - 4 - axisFont.MeasureText(lab), p + 3, axisFont, axisPaint);
        }

        float slot = (horiz ? plot.Height : plot.Width) / n;
        float C(int i) => (horiz ? plot.Top : plot.Left) + slot * (i + 0.5f);
        int skip = (int)Math.Ceiling(n / Math.Max(1.0, (horiz ? plot.Height / 12 : plot.Width / 44)));
        for (int i = 0; i < n; i += Math.Max(1, skip))
        {
            var lab = cats[i].Length > 12 ? cats[i].Substring(0, 11) + "…" : cats[i];
            bool rtl = TextPainter.HasRtl(lab);
            float w = _text.Measure(lab, axisFont, rtl);
            if (horiz) _text.Draw(canvas, lab, plot.Left - 4 - w, C(i) + 3, axisFont, axisPaint, rtl);
            else _text.Draw(canvas, lab, C(i) - w / 2, plot.Bottom + 11, axisFont, axisPaint, rtl);
        }

        double baseV = Math.Max(0, min);
        var colSeries = series.Where(s => s.Kind is "column" or "bar").ToList();
        if (colSeries.Count > 0)
        {
            int k = ch.Def.Stacked ? 1 : colSeries.Count;
            float bw = Math.Max(1, slot * 0.72f / k);
            for (int i = 0; i < n; i++)
            {
                double acc = 0;
                for (int s = 0; s < colSeries.Count; s++)
                {
                    double v = i < colSeries[s].Vals.Count ? colSeries[s].Vals[i] : 0;
                    using var fp = new SKPaint { Color = colSeries[s].Color, Style = SKPaintStyle.Fill, IsAntialias = true };
                    double from = ch.Def.Stacked ? acc : baseV, to = ch.Def.Stacked ? acc + v : v; acc += v;
                    float off = ch.Def.Stacked ? -bw / 2 : (s - k / 2f) * bw;
                    SKRect r = horiz ? new SKRect(V(Math.Min(from, to)), C(i) + off, V(Math.Max(from, to)), C(i) + off + bw)
                                     : new SKRect(C(i) + off, V(Math.Max(from, to)), C(i) + off + bw, V(Math.Min(from, to)));
                    canvas.DrawRect(r, fp);
                    if (ch.Def.ShowDataLabels) canvas.DrawText(v.ToString("#,##0.##"), horiz ? r.Right + 2 : r.Left, horiz ? r.MidY + 3 : r.Top - 2, axisFont, axisPaint);
                }
            }
        }
        {
            foreach (var s in series.Where(x => x.Kind is "line" or "area"))
            {
                Func<double, float> VV = s.Sec ? V2 : V;
                double baseLine = s.Sec ? Math.Max(0, min2) : baseV;
                using var lp = new SKPaint { Color = s.Color, StrokeWidth = 1.6f, Style = SKPaintStyle.Stroke, IsAntialias = true };
                var path = new SKPath();
                for (int i = 0; i < s.Vals.Count; i++) { var pt = new SKPoint(C(i), VV(s.Vals[i])); if (i == 0) path.MoveTo(pt); else path.LineTo(pt); }
                if (s.Kind == "area")
                {
                    var area = new SKPath(path); area.LineTo(C(s.Vals.Count - 1), VV(baseLine)); area.LineTo(C(0), VV(baseLine)); area.Close();
                    using var ap = new SKPaint { Color = s.Color.WithAlpha(110), Style = SKPaintStyle.Fill, IsAntialias = true };
                    canvas.DrawPath(area, ap);
                }
                canvas.DrawPath(path, lp);
                using var dot = new SKPaint { Color = s.Color, Style = SKPaintStyle.Fill, IsAntialias = true };
                for (int i = 0; i < s.Vals.Count; i++)
                {
                    canvas.DrawCircle(C(i), VV(s.Vals[i]), 2.2f, dot);
                    if (ch.Def.ShowDataLabels) canvas.DrawText(s.Vals[i].ToString("#,##0.##"), C(i) - 6, VV(s.Vals[i]) - 4, axisFont, axisPaint);
                }
            }
        }
    }

    sealed class Ser { public string Name = ""; public List<double> Vals = new(); public SKColor Color; public string Kind = "column"; public bool Sec; }

    /// <summary>The drawing kind of series <paramref name="i"/>: its own Type in a combo chart, otherwise the chart's type.</summary>
    static string SeriesKind(RenderedChart ch, int i)
    {
        var type = (ch.Def.Type ?? "column").ToLowerInvariant();
        if (type == "pie" || type == "bar") return type;
        var t = i < ch.SeriesTypes.Count ? (ch.SeriesTypes[i] ?? "").ToLowerInvariant() : "";
        return t is "column" or "line" or "area" ? t : type;
    }

    static (double step, double top) Nice(double min, double max)
    {
        double range = Math.Max(1e-9, max - Math.Min(0, min));
        double raw = range / 5, mag = Math.Pow(10, Math.Floor(Math.Log10(raw)));
        double norm = raw / mag; double step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * mag;
        return (step, Math.Ceiling(max / step) * step);
    }

    static SKColor ParseColor(string? hex, int i)
    {
        var h = (hex ?? "").Trim().TrimStart('#');
        if (h.Length != 6 || !uint.TryParse(h, System.Globalization.NumberStyles.HexNumber, null, out var v)) h = Palette[i % Palette.Length];
        v = Convert.ToUInt32(h, 16);
        return new SKColor((byte)(v >> 16), (byte)(v >> 8), (byte)v);
    }

    void DrawLegend(SKCanvas canvas, List<(string name, SKColor color)> items, SKRect box, ref float bottom)
    {
        using var font = new SKFont(_text.Face(_wb.Template?.DefaultFont, false, false, string.Concat(items.Select(i => i.name))), 8);
        using var tp = new SKPaint { Color = new SKColor(0x33, 0x33, 0x33), IsAntialias = true };
        float total = items.Sum(i => 14 + _text.Measure(i.name, font, TextPainter.HasRtl(i.name)) + 10);
        float x = box.MidX - Math.Min(total, box.Width - 12) / 2, y = box.Bottom - 8;
        foreach (var (name, color) in items)
        {
            using var sw = new SKPaint { Color = color, Style = SKPaintStyle.Fill };
            canvas.DrawRect(x, y - 6, 8, 8, sw);
            bool rtl = TextPainter.HasRtl(name);
            _text.Draw(canvas, name, x + 11, y + 1, font, tp, rtl);
            x += 14 + _text.Measure(name, font, rtl) + 10;
            if (x > box.Right - 20) break;
        }
        bottom -= 16;
    }

    static void DrawPie(SKCanvas canvas, List<double> vals, SKRect area)
    {
        double total = vals.Where(v => v > 0).Sum();
        if (total <= 0) return;
        float d = Math.Min(area.Width, area.Height) - 8; if (d < 10) return;
        var oval = new SKRect(area.MidX - d / 2, area.MidY - d / 2, area.MidX + d / 2, area.MidY + d / 2);
        float start = -90;
        for (int i = 0; i < vals.Count; i++)
        {
            if (vals[i] <= 0) continue;
            float sweep = (float)(vals[i] / total * 360);
            using var p = new SKPaint { Color = ParseColor(null, i), Style = SKPaintStyle.Fill, IsAntialias = true };
            canvas.DrawArc(oval, start, sweep, true, p);
            using var edge = new SKPaint { Color = SKColors.White, Style = SKPaintStyle.Stroke, StrokeWidth = 0.8f, IsAntialias = true };
            canvas.DrawArc(oval, start, sweep, true, edge);
            start += sweep;
        }
    }
}
