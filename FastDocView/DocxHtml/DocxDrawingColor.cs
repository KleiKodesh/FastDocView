using System.Globalization;
using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// A DrawingML colour (ECMA-376 Part 1 20.1.2.3: a:srgbClr / a:schemeClr / a:sysClr / a:prstClr) with its transforms
/// applied in document order, ported from LibreOffice oox/source/drawingml/color.cxx (checked online 2026-11-03):
/// lumMod / lumOff work on HSL luminance (:784-796, with black / white losing saturation), tint / shade on the gamma-2.3
/// linear RGB ("CRGB", :798-821, toCrgb :1043, toRgb :978), percentages in 100000ths (MAX_PERCENT,
/// include/oox/drawingml/drawingmltypes.hxx:50-51). The elements come without prefixes (DocxDrawingStripper.ShapeTree).
/// a:alpha is ignored (opacity is not drawn); a:prstClr is passed to the browser as a CSS colour name (both use the
/// HTML / SVG names; not checked name by name).
/// </summary>
internal static class DocxDrawingColor
{
    private const int MaxPercent = 100000, MaxDegree = 360 * 60000, PerDegree = 60000;
    private const double DecGamma = 2.3, IncGamma = 1 / DecGamma;

    /// <summary>The CSS colour of the first colour child of <paramref name="parent"/>; null when it has none or it cannot be resolved.</summary>
    public static string? Css(XElement? parent, DocxTheme? theme, string? placeholderHex = null)
    {
        if (parent is null) return null;
        foreach (XElement color in parent.Elements())
        {
            string? baseHex;
            switch (color.Name.LocalName)
            {
                case "srgbClr": baseHex = (string?)color.Attribute("val"); break;
                case "sysClr": baseHex = (string?)color.Attribute("lastClr"); break;
                case "schemeClr":
                    string? scheme = (string?)color.Attribute("val");
                    baseHex = scheme == "phClr" ? placeholderHex : theme?.DrawingSchemeColorHex(scheme);
                    break;
                case "prstClr":
                    string? name = (string?)color.Attribute("val");
                    return name is { Length: > 0 } && name.All(char.IsLetter) ? name.ToLowerInvariant() : null;
                default: continue;
            }
            if (baseHex is not { Length: 6 } || !int.TryParse(baseHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb)) return null;
            return "#" + Transformed(rgb, color).ToString("X6", CultureInfo.InvariantCulture);
        }
        return null;
    }

    /// <summary>The RRGGBB (no '#') of the first colour child, for use as a placeholder colour (phClr); null when none.</summary>
    public static string? Hex(XElement? parent, DocxTheme? theme) => Css(parent, theme) is { Length: 7 } css && css[0] == '#' ? css.Substring(1) : null;

    private enum Mode { Rgb, Crgb, Hsl }

    private static int Transformed(int rgb, XElement color)
    {
        Mode mode = Mode.Rgb;
        int c1 = (rgb >> 16) & 0xFF, c2 = (rgb >> 8) & 0xFF, c3 = rgb & 0xFF;
        foreach (XElement transform in color.Elements())
        {
            if (!int.TryParse((string?)transform.Attribute("val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)) continue;
            switch (transform.Name.LocalName)
            {
                case "lumMod":
                    ToHsl(ref mode, ref c1, ref c2, ref c3);
                    c3 = (int)Math.Max(0, Math.Min(MaxPercent, (double)c3 * value / MaxPercent));
                    if (c3 == 0 || c3 == MaxPercent) c2 = 0;
                    break;
                case "lumOff":
                    ToHsl(ref mode, ref c1, ref c2, ref c3);
                    c3 = Math.Max(0, Math.Min(MaxPercent, c3 + value));
                    if (c3 == 0 || c3 == MaxPercent) c2 = 0;
                    break;
                case "shade" when value is >= 0 and <= MaxPercent:
                {
                    ToCrgb(ref mode, ref c1, ref c2, ref c3);
                    double factor = (double)value / MaxPercent;
                    c1 = (int)(c1 * factor); c2 = (int)(c2 * factor); c3 = (int)(c3 * factor);
                    break;
                }
                case "tint" when value is >= 0 and <= MaxPercent:
                {
                    ToCrgb(ref mode, ref c1, ref c2, ref c3);
                    double factor = (double)value / MaxPercent;
                    c1 = (int)(MaxPercent - (MaxPercent - c1) * factor); c2 = (int)(MaxPercent - (MaxPercent - c2) * factor); c3 = (int)(MaxPercent - (MaxPercent - c3) * factor);
                    break;
                }
            }
        }
        ToRgb(ref mode, ref c1, ref c2, ref c3);
        return (c1 << 16) | (c2 << 8) | c3;
    }

    private static int Gamma(int component, double gamma) => (int)(Math.Pow((double)component / MaxPercent, gamma) * MaxPercent + 0.5);

    private static void ToCrgb(ref Mode mode, ref int c1, ref int c2, ref int c3)
    {
        if (mode == Mode.Hsl) ToRgb(ref mode, ref c1, ref c2, ref c3);
        if (mode != Mode.Rgb) return;
        mode = Mode.Crgb;
        c1 = Gamma(c1 * MaxPercent / 255, DecGamma); c2 = Gamma(c2 * MaxPercent / 255, DecGamma); c3 = Gamma(c3 * MaxPercent / 255, DecGamma);
    }

    private static void ToRgb(ref Mode mode, ref int c1, ref int c2, ref int c3)
    {
        if (mode == Mode.Crgb)
        {
            mode = Mode.Rgb;
            c1 = Gamma(c1, IncGamma) * 255 / MaxPercent; c2 = Gamma(c2, IncGamma) * 255 / MaxPercent; c3 = Gamma(c3, IncGamma) * 255 / MaxPercent;
            return;
        }
        if (mode != Mode.Hsl) return;
        mode = Mode.Rgb;
        double r = 0, g = 0, b = 0;
        if (c2 == 0 || c3 == MaxPercent) r = g = b = (double)c3 / MaxPercent;
        else if (c3 > 0)
        {
            double hue = (double)c1 / MaxDegree * 6.0;
            if (hue <= 1) { r = 1; g = hue; }
            else if (hue <= 2) { r = 2 - hue; g = 1; }
            else if (hue <= 3) { g = 1; b = hue - 2; }
            else if (hue <= 4) { g = 4 - hue; b = 1; }
            else if (hue <= 5) { r = hue - 4; b = 1; }
            else { r = 1; b = 6 - hue; }
            double saturation = (double)c2 / MaxPercent;
            r = (r - 0.5) * saturation + 0.5; g = (g - 0.5) * saturation + 0.5; b = (b - 0.5) * saturation + 0.5;
            double luminance = 2.0 * c3 / MaxPercent - 1.0;
            if (luminance < 0) { double shade = luminance + 1; r *= shade; g *= shade; b *= shade; }
            else if (luminance > 0) { double tint = 1 - luminance; r = 1 - (1 - r) * tint; g = 1 - (1 - g) * tint; b = 1 - (1 - b) * tint; }
        }
        c1 = (int)(r * 255 + 0.5); c2 = (int)(g * 255 + 0.5); c3 = (int)(b * 255 + 0.5);
    }

    private static void ToHsl(ref Mode mode, ref int c1, ref int c2, ref int c3)
    {
        if (mode == Mode.Crgb) ToRgb(ref mode, ref c1, ref c2, ref c3);
        if (mode != Mode.Rgb) return;
        mode = Mode.Hsl;
        double r = c1 / 255.0, g = c2 / 255.0, b = c3 / 255.0;
        double min = Math.Min(Math.Min(r, g), b), max = Math.Max(Math.Max(r, g), b), delta = max - min;
        int hue;
        if (delta == 0) hue = 0;
        else if (max == r) hue = (int)(((g - b) / delta * 60 + 360) * PerDegree + 0.5) % MaxDegree;
        else if (max == g) hue = (int)(((b - r) / delta * 60 + 120) * PerDegree + 0.5);
        else hue = (int)(((r - g) / delta * 60 + 240) * PerDegree + 0.5);
        int luminance = (int)((min + max) / 2 * MaxPercent + 0.5);
        int saturation = luminance == 0 || luminance == MaxPercent ? 0
                       : luminance <= 50 * 1000 ? (int)(delta / (min + max) * MaxPercent + 0.5)
                       : (int)(delta / (2 - max - min) * MaxPercent + 0.5);
        c1 = hue; c2 = saturation; c3 = luminance;
    }
}
