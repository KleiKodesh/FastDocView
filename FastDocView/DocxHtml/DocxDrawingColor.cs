using System.Globalization;
using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>A resolved DrawingML colour: CSS "#RRGGBB" and its opacity (1 = opaque).</summary>
internal readonly record struct DocxResolvedColor(string Css, double Opacity);

/// <summary>
/// A DrawingML colour (ECMA-376 Part 1 20.1.2.3: a:srgbClr / a:scrgbClr / a:hslClr / a:schemeClr / a:sysClr / a:prstClr) with
/// its transforms applied in document order, ported from LibreOffice oox/source/drawingml/color.cxx (checked online
/// 2026-11-03): base colours (setSrgbClr :360, setScrgbClr :367 = linear "CRGB", setHslClr :378, the preset colour table
/// constDmlColors :39-108, scheme names :211-238), every transform of Color::getColor :756-874 (red / green / blue (Mod / Off)
/// on CRGB, hue / sat / lum (Mod / Off) on HSL with black / white losing saturation, shade / tint on CRGB, gray, comp, inv,
/// gamma, invGamma), and alpha / alphaMod / alphaOff, which LibreOffice applies as it reads them (:467-477 with lclSetValue /
/// lclModValue / lclOffValue :193-210); opacity = alpha / 100000 (getTransparency :904-907). Percentages in 100000ths, angles
/// in 60000ths of a degree. The elements come without prefixes (DocxDrawingStripper.ShapeTree) or, from the theme, without
/// them too (DocxTheme strips them): names are compared by local name.
/// </summary>
internal static class DocxDrawingColor
{
    private const int MaxPercent = 100000, MaxDegree = 360 * 60000, PerDegree = 60000;
    private const double DecGamma = 2.3, IncGamma = 1 / DecGamma;

    /// <summary>The CSS colour of the first colour child of <paramref name="parent"/>; null when it has none or it cannot be resolved.</summary>
    public static string? Css(XElement? parent, DocxTheme? theme, string? placeholderHex = null) => Resolve(parent, theme, placeholderHex)?.Css;

    /// <summary>The RRGGBB (no '#') of the first colour child, for use as a placeholder colour (phClr); null when none.</summary>
    public static string? Hex(XElement? parent, DocxTheme? theme) => Resolve(parent, theme)?.Css.Substring(1);

    /// <summary>The first colour child of <paramref name="parent"/> with its transforms and alpha; null when none or unresolved.</summary>
    public static DocxResolvedColor? Resolve(XElement? parent, DocxTheme? theme, string? placeholderHex = null)
    {
        if (parent is null) return null;
        foreach (XElement color in parent.Elements())
        {
            Mode mode = Mode.Rgb;
            int c1, c2, c3;
            switch (color.Name.LocalName)
            {
                case "srgbClr":
                    if (!TryHex((string?)color.Attribute("val"), out int srgb)) return null;
                    (c1, c2, c3) = Split(srgb);
                    break;
                case "sysClr":
                    if (!TryHex((string?)color.Attribute("lastClr"), out int system)) return null;
                    (c1, c2, c3) = Split(system);
                    break;
                case "schemeClr":
                {
                    string? scheme = (string?)color.Attribute("val");
                    if (!TryHex(scheme == "phClr" ? placeholderHex : theme?.DrawingSchemeColorHex(scheme), out int schemeRgb)) return null;
                    (c1, c2, c3) = Split(schemeRgb);
                    break;
                }
                case "prstClr":
                    if (!PresetColors.TryGetValue((string?)color.Attribute("val") ?? "", out string? preset) || !TryHex(preset, out int presetRgb)) return null;
                    (c1, c2, c3) = Split(presetRgb);
                    break;
                case "scrgbClr":
                    mode = Mode.Crgb;
                    c1 = Limit(Integer(color, "r"), 0, MaxPercent); c2 = Limit(Integer(color, "g"), 0, MaxPercent); c3 = Limit(Integer(color, "b"), 0, MaxPercent);
                    break;
                case "hslClr":
                    mode = Mode.Hsl;
                    c1 = Limit(Integer(color, "hue"), 0, MaxDegree); c2 = Limit(Integer(color, "sat"), 0, MaxPercent); c3 = Limit(Integer(color, "lum"), 0, MaxPercent);
                    break;
                default: continue;
            }
            int alpha = MaxPercent;
            Transform(color, ref mode, ref c1, ref c2, ref c3, ref alpha);
            ToRgb(ref mode, ref c1, ref c2, ref c3);
            int rgb = (Limit(c1, 0, 255) << 16) | (Limit(c2, 0, 255) << 8) | Limit(c3, 0, 255);
            return new DocxResolvedColor("#" + rgb.ToString("X6", CultureInfo.InvariantCulture), (double)alpha / MaxPercent);
        }
        return null;
    }

    private static (int, int, int) Split(int rgb) => ((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);

    private static bool TryHex(string? hex, out int rgb)
    {
        rgb = 0;
        return hex is { Length: 6 } && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out rgb);
    }

    private static int Integer(XElement element, string name)
        => int.TryParse((string?)element.Attribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : 0;

    private static int Limit(int value, int min, int max) => value < min ? min : value > max ? max : value;

    // color.cxx lclSetValue / lclModValue / lclOffValue (:193-210)
    private static void SetValue(ref int target, int value, int max = MaxPercent) { if (value >= 0 && value <= max) target = value; }
    private static void ModValue(ref int target, int value, int max = MaxPercent) => target = (int)Math.Max(0, Math.Min(max, (double)target * value / MaxPercent));
    private static void OffValue(ref int target, int value, int max = MaxPercent) => target = Limit(target + value, 0, max);

    private enum Mode { Rgb, Crgb, Hsl }

    private static void Transform(XElement color, ref Mode mode, ref int c1, ref int c2, ref int c3, ref int alpha)
    {
        foreach (XElement transform in color.Elements())
        {
            string name = transform.Name.LocalName;
            switch (name)
            {
                case "gray":
                    ToRgb(ref mode, ref c1, ref c2, ref c3);
                    c1 = c2 = c3 = (c1 * 22 + c2 * 72 + c3 * 6) / 100;
                    continue;
                case "comp":
                    ToHsl(ref mode, ref c1, ref c2, ref c3);
                    c1 = (c1 + 180 * PerDegree) % MaxDegree;
                    continue;
                case "inv":
                    ToCrgb(ref mode, ref c1, ref c2, ref c3);
                    c1 = MaxPercent - c1; c2 = MaxPercent - c2; c3 = MaxPercent - c3;
                    continue;
                case "gamma":
                    ToCrgb(ref mode, ref c1, ref c2, ref c3);
                    c1 = Gamma(c1, IncGamma); c2 = Gamma(c2, IncGamma); c3 = Gamma(c3, IncGamma);
                    continue;
                case "invGamma":
                    ToCrgb(ref mode, ref c1, ref c2, ref c3);
                    c1 = Gamma(c1, DecGamma); c2 = Gamma(c2, DecGamma); c3 = Gamma(c3, DecGamma);
                    continue;
            }
            if (!int.TryParse((string?)transform.Attribute("val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)) continue;
            switch (name)
            {
                case "alpha": SetValue(ref alpha, value); break;
                case "alphaMod": ModValue(ref alpha, value); break;
                case "alphaOff": OffValue(ref alpha, value); break;
                case "red": ToCrgb(ref mode, ref c1, ref c2, ref c3); SetValue(ref c1, value); break;
                case "redMod": ToCrgb(ref mode, ref c1, ref c2, ref c3); ModValue(ref c1, value); break;
                case "redOff": ToCrgb(ref mode, ref c1, ref c2, ref c3); OffValue(ref c1, value); break;
                case "green": ToCrgb(ref mode, ref c1, ref c2, ref c3); SetValue(ref c2, value); break;
                case "greenMod": ToCrgb(ref mode, ref c1, ref c2, ref c3); ModValue(ref c2, value); break;
                case "greenOff": ToCrgb(ref mode, ref c1, ref c2, ref c3); OffValue(ref c2, value); break;
                case "blue": ToCrgb(ref mode, ref c1, ref c2, ref c3); SetValue(ref c3, value); break;
                case "blueMod": ToCrgb(ref mode, ref c1, ref c2, ref c3); ModValue(ref c3, value); break;
                case "blueOff": ToCrgb(ref mode, ref c1, ref c2, ref c3); OffValue(ref c3, value); break;
                case "hue": ToHsl(ref mode, ref c1, ref c2, ref c3); SetValue(ref c1, value, MaxDegree); break;
                case "hueMod": ToHsl(ref mode, ref c1, ref c2, ref c3); ModValue(ref c1, value, MaxDegree); break;
                case "hueOff": ToHsl(ref mode, ref c1, ref c2, ref c3); OffValue(ref c1, value, MaxDegree); break;
                case "sat": ToHsl(ref mode, ref c1, ref c2, ref c3); SetValue(ref c2, value); break;
                case "satMod": ToHsl(ref mode, ref c1, ref c2, ref c3); ModValue(ref c2, value); break;
                case "satOff": ToHsl(ref mode, ref c1, ref c2, ref c3); OffValue(ref c2, value); break;
                case "lum":
                    ToHsl(ref mode, ref c1, ref c2, ref c3); SetValue(ref c3, value);
                    if (c3 == 0 || c3 == MaxPercent) c2 = 0;
                    break;
                case "lumMod":
                    ToHsl(ref mode, ref c1, ref c2, ref c3); ModValue(ref c3, value);
                    if (c3 == 0 || c3 == MaxPercent) c2 = 0;
                    break;
                case "lumOff":
                    ToHsl(ref mode, ref c1, ref c2, ref c3); OffValue(ref c3, value);
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

    /// <summary>ST_PresetColorVal -> RRGGBB: LibreOffice color.cxx constDmlColors (:39-108), as the XML writes the names.</summary>
    private static readonly Dictionary<string, string> PresetColors = new(StringComparer.Ordinal)
    {
        ["aliceBlue"] = "F0F8FF", ["antiqueWhite"] = "FAEBD7", ["aqua"] = "00FFFF", ["aquamarine"] = "7FFFD4", ["azure"] = "F0FFFF", ["beige"] = "F5F5DC",
        ["bisque"] = "FFE4C4", ["black"] = "000000", ["blanchedAlmond"] = "FFEBCD", ["blue"] = "0000FF", ["blueViolet"] = "8A2BE2", ["brown"] = "A52A2A",
        ["burlyWood"] = "DEB887", ["cadetBlue"] = "5F9EA0", ["chartreuse"] = "7FFF00", ["chocolate"] = "D2691E", ["coral"] = "FF7F50", ["cornflowerBlue"] = "6495ED",
        ["cornsilk"] = "FFF8DC", ["crimson"] = "DC143C", ["cyan"] = "00FFFF", ["deepPink"] = "FF1493", ["deepSkyBlue"] = "00BFFF", ["dimGray"] = "696969",
        ["dkBlue"] = "00008B", ["dkCyan"] = "008B8B", ["dkGoldenrod"] = "B8860B", ["dkGray"] = "A9A9A9", ["dkGreen"] = "006400", ["dkKhaki"] = "BDB76B",
        ["dkMagenta"] = "8B008B", ["dkOliveGreen"] = "556B2F", ["dkOrange"] = "FF8C00", ["dkOrchid"] = "9932CC", ["dkRed"] = "8B0000", ["dkSalmon"] = "E9967A",
        ["dkSeaGreen"] = "8FBC8B", ["dkSlateBlue"] = "483D8B", ["dkSlateGray"] = "2F4F4F", ["dkTurquoise"] = "00CED1", ["dkViolet"] = "9400D3",
        ["dodgerBlue"] = "1E90FF", ["firebrick"] = "B22222", ["floralWhite"] = "FFFAF0", ["forestGreen"] = "228B22", ["fuchsia"] = "FF00FF", ["gainsboro"] = "DCDCDC",
        ["ghostWhite"] = "F8F8FF", ["gold"] = "FFD700", ["goldenrod"] = "DAA520", ["gray"] = "808080", ["green"] = "008000", ["greenYellow"] = "ADFF2F",
        ["honeydew"] = "F0FFF0", ["hotPink"] = "FF69B4", ["indianRed"] = "CD5C5C", ["indigo"] = "4B0082", ["ivory"] = "FFFFF0", ["khaki"] = "F0E68C",
        ["lavender"] = "E6E6FA", ["lavenderBlush"] = "FFF0F5", ["lawnGreen"] = "7CFC00", ["lemonChiffon"] = "FFFACD", ["lime"] = "00FF00",
        ["limeGreen"] = "32CD32", ["linen"] = "FAF0E6", ["ltBlue"] = "ADD8E6", ["ltCoral"] = "F08080", ["ltCyan"] = "E0FFFF", ["ltGoldenrodYellow"] = "FAFA78",
        ["ltGray"] = "D3D3D3", ["ltGreen"] = "90EE90", ["ltPink"] = "FFB6C1", ["ltSalmon"] = "FFA07A", ["ltSeaGreen"] = "20B2AA", ["ltSkyBlue"] = "87CEFA",
        ["ltSlateGray"] = "778899", ["ltSteelBlue"] = "B0C4DE", ["ltYellow"] = "FFFFE0", ["magenta"] = "FF00FF", ["maroon"] = "800000",
        ["medAquamarine"] = "66CDAA", ["medBlue"] = "0000CD", ["medOrchid"] = "BA55D3", ["medPurple"] = "9370DB", ["medSeaGreen"] = "3CB371",
        ["medSlateBlue"] = "7B68EE", ["medSpringGreen"] = "00FA9A", ["medTurquoise"] = "48D1CC", ["medVioletRed"] = "C71585", ["midnightBlue"] = "191970",
        ["mintCream"] = "F5FFFA", ["mistyRose"] = "FFE4E1", ["moccasin"] = "FFE4B5", ["navajoWhite"] = "FFDEAD", ["navy"] = "000080", ["oldLace"] = "FDF5E6",
        ["olive"] = "808000", ["oliveDrab"] = "6B8E23", ["orange"] = "FFA500", ["orangeRed"] = "FF4500", ["orchid"] = "DA70D6", ["paleGoldenrod"] = "EEE8AA",
        ["paleGreen"] = "98FB98", ["paleTurquoise"] = "AFEEEE", ["paleVioletRed"] = "DB7093", ["papayaWhip"] = "FFEFD5", ["peachPuff"] = "FFDAB9", ["peru"] = "CD853F",
        ["pink"] = "FFC0CB", ["plum"] = "DDA0DD", ["powderBlue"] = "B0E0E6", ["purple"] = "800080", ["red"] = "FF0000", ["rosyBrown"] = "BC8F8F",
        ["royalBlue"] = "4169E1", ["saddleBrown"] = "8B4513", ["salmon"] = "FA8072", ["sandyBrown"] = "F4A460", ["seaGreen"] = "2E8B57", ["seaShell"] = "FFF5EE",
        ["sienna"] = "A0522D", ["silver"] = "C0C0C0", ["skyBlue"] = "87CEEB", ["slateBlue"] = "6A5ACD", ["slateGray"] = "708090", ["snow"] = "FFFAFA",
        ["springGreen"] = "00FF7F", ["steelBlue"] = "4682B4", ["tan"] = "D2B48C", ["teal"] = "008080", ["thistle"] = "D8BFD8", ["tomato"] = "FF6347",
        ["turquoise"] = "40E0D0", ["violet"] = "EE82EE", ["wheat"] = "F5DEB3", ["white"] = "FFFFFF", ["whiteSmoke"] = "F5F5F5", ["yellow"] = "FFFF00",
        ["yellowGreen"] = "9ACD32",
    };
}
