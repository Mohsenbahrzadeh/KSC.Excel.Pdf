using Newtonsoft.Json;

namespace KSC.Excel.Model
{
    /// <summary>One side of a cell border.</summary>
    public class BorderDef
    {
        /// <summary>thin, medium, thick, dashed, dotted, double, hair.</summary>
        public string Style { get; set; }
        /// <summary>#RRGGBB</summary>
        public string Color { get; set; }

        public BorderDef Clone() => new BorderDef { Style = Style, Color = Color };
    }

    /// <summary>
    /// A flat, designer-friendly cell style. Every property is optional: null/false means "inherit the default".
    /// </summary>
    public class StyleDef
    {
        public string Font { get; set; }
        public double? Size { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strike { get; set; }
        /// <summary>Font color, #RRGGBB.</summary>
        public string Color { get; set; }
        /// <summary>Background color, #RRGGBB.</summary>
        public string Fill { get; set; }
        public BorderDef Left { get; set; }
        public BorderDef Right { get; set; }
        public BorderDef Top { get; set; }
        public BorderDef Bottom { get; set; }
        /// <summary>left, center, right, justify, fill, distributed (null = general).</summary>
        public string H { get; set; }
        /// <summary>top, center, bottom.</summary>
        public string V { get; set; }
        public bool Wrap { get; set; }
        public int Indent { get; set; }
        /// <summary>Text rotation in degrees (-90..90) or 255 for vertical text.</summary>
        public int Rotation { get; set; }
        public bool Shrink { get; set; }
        /// <summary>Excel number format, e.g. #,##0 or yyyy/mm/dd.</summary>
        public string Format { get; set; }
        /// <summary>Reading order: true = right-to-left, false = left-to-right, null = context.</summary>
        public bool? Rtl { get; set; }

        public StyleDef Clone() => new StyleDef
        {
            Font = Font, Size = Size, Bold = Bold, Italic = Italic, Underline = Underline, Strike = Strike,
            Color = Color, Fill = Fill,
            Left = Left?.Clone(), Right = Right?.Clone(), Top = Top?.Clone(), Bottom = Bottom?.Clone(),
            H = H, V = V, Wrap = Wrap, Indent = Indent, Rotation = Rotation, Shrink = Shrink, Format = Format, Rtl = Rtl
        };

        static readonly JsonSerializerSettings KeySettings = new JsonSerializerSettings
        {
            NullValueHandling = NullValueHandling.Ignore,
            DefaultValueHandling = DefaultValueHandling.Ignore
        };

        /// <summary>Canonical identity used to de-duplicate styles.</summary>
        public string Key() => JsonConvert.SerializeObject(this, KeySettings);
    }
}
