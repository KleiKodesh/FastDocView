using System.Globalization;

namespace FastDocView.Html;

/// <summary>
/// A document's explicit colour as CSS that reads in both themes, without making the HTML theme-dependent:
/// <c>light-dark(original, darkVariant)</c>. The page's <c>color-scheme</c> picks one (the app sets <c>dark</c> on its
/// root in a dark theme), so a theme switch needs no re-render. "Auto" colours are never written: they inherit the
/// page's text colour (currentColor).
/// <para>
/// The dark variant keeps hue and chroma and inverts lightness in CIELAB: L' = min(110 - L, 100), a and b unchanged.
/// Source: Chromium's dark mode, third_party/blink/renderer/platform/graphics/dark_mode_color_filter.cc
/// (LABColorFilter::InvertColor) — the same remap for text and backgrounds. Word's own dark mode likewise "shifts"
/// the document's colours instead of inverting them (support.microsoft.com, "Dark Mode in Word"); its exact formula
/// is not published, so Chromium's is the cited one. Black text becomes near-white, white shading near-black.
/// </para>
/// </summary>
public static class WordHtmlColor
{
    /// <summary>
    /// <paramref name="color"/> as theme-aware CSS. Accepts "RRGGBB", "#RRGGBB", "#RGB" and the basic named colours
    /// Word's highlights use; any other value (a gradient, "transparent", an unknown name) is returned unchanged.
    /// </summary>
    public static string Css(string color)
    {
        if (!TryParse(color, out int red, out int green, out int blue)) return color;
        // the light value is the document's own, as written (light mode output stays as before)
        string light = color.StartsWith("#", StringComparison.Ordinal) ? color : IsHexDigits(color) ? "#" + color : Hex(red, green, blue);
        (int darkRed, int darkGreen, int darkBlue) = DarkVariant(red, green, blue);
        return "light-dark(" + light + "," + Hex(darkRed, darkGreen, darkBlue) + ")";
    }

    /// <summary>
    /// <paramref name="html"/> with every bare hex colour inside its tags (style values, SVG fill / stroke / stop-color
    /// attributes, gradients, borders) written through <see cref="Css"/>, so a shape's fill, its line and the text in it
    /// are remapped together. Text content, values already in light-dark(), url(#id) and href="#id" are left as they are.
    /// Pictures are &lt;img&gt; elements and are never touched (bitmaps keep their colours, as in Word).
    /// One linear pass; a block without '#' is returned as is. Hosts apply it to block and note HTML at their transport
    /// (renderers keep writing plain colours, which the fidelity harness checks).
    /// </summary>
    public static string ThemeAwareMarkup(string html)
    {
        if (html.IndexOf('#') < 0) return html;
        System.Text.StringBuilder? result = null;
        int copied = 0;
        bool inTag = false;
        for (int i = 0; i < html.Length; i++)
        {
            char c = html[i];
            if (!inTag) { if (c == '<') inTag = true; continue; }
            if (c == '>') { inTag = false; continue; }
            if (c == 'l' && string.CompareOrdinal(html, i, "light-dark(", 0, 11) == 0)
            {
                int close = html.IndexOf(')', i);
                if (close < 0) break;
                i = close;
                continue;
            }
            if (char.IsLetter(c) && (i == 0 || !char.IsLetterOrDigit(html[i - 1]) && html[i - 1] != '-') && IsColorValueStart(html, i))
            {
                // a named colour (VML and highlight names: the 16 HTML colours, VML 2.2.4 / CSS 2.1 4.3.6) as a paint value
                int wordEnd = i;
                while (wordEnd < html.Length && char.IsLetter(html[wordEnd])) wordEnd++;
                string word = html.Substring(i, wordEnd - i);
                if (!word.All(Uri.IsHexDigit) && TryParse(word, out _, out _, out _))
                {
                    result ??= new System.Text.StringBuilder(html.Length + 64);
                    result.Append(html, copied, i - copied).Append(Css(word));
                    copied = wordEnd;
                }
                i = wordEnd - 1;
                continue;
            }
            if (c != '#' || (i > 0 && html[i - 1] == '&') || IsReference(html, i)) continue;   // &#160; is a character reference
            int end = i + 1;
            while (end < html.Length && Uri.IsHexDigit(html[end])) end++;
            int length = end - i - 1;
            if (length is not (3 or 6) || (end < html.Length && (char.IsLetterOrDigit(html[end]) || html[end] is '_' or '-'))) continue;
            result ??= new System.Text.StringBuilder(html.Length + 64);
            result.Append(html, copied, i - copied).Append(Css(html.Substring(i, end - i)));
            copied = end;
            i = end - 1;
        }
        if (result is null) return html;
        return result.Append(html, copied, html.Length - copied).ToString();
    }

    /// <summary>Position <paramref name="index"/> starts the value of a paint attribute or colour property.</summary>
    private static bool IsColorValueStart(string html, int index)
    {
        int start = index;
        if (start > 0 && html[start - 1] is '"' or '\'') start--;
        while (start > 0 && html[start - 1] == ' ') start--;
        foreach (string name in PaintNames)
            if (EndsWith(html, start, name)) return true;
        return false;
    }

    private static readonly string[] PaintNames =
        ["fill=", "stroke=", "stop-color=", "flood-color=", "fill:", "stroke:", "stop-color:", "color:", "background:", "background-color:"];

    /// <summary>'#' starting a fragment reference (url(#id), href="#id"), not a colour.</summary>
    private static bool IsReference(string html, int hashIndex)
    {
        int start = hashIndex;
        while (start > 0 && html[start - 1] is '"' or '\'' or ' ') start--;
        return EndsWith(html, start, "url(") || EndsWith(html, start, "href=");
    }

    private static bool EndsWith(string html, int end, string suffix)
        => end >= suffix.Length && string.CompareOrdinal(html, end - suffix.Length, suffix, 0, suffix.Length) == 0;

    /// <summary>The dark-theme variant: CIELAB lightness inverted (110 - L, at most 100), a and b kept.</summary>
    public static (int Red, int Green, int Blue) DarkVariant(int red, int green, int blue)
    {
        (double l, double a, double b) = ToLab(red, green, blue);
        return FromLab(Math.Min(110 - l, 100), a, b);
    }

    private static bool IsHexDigits(string value) => value.Length is 3 or 6 && value.All(Uri.IsHexDigit);

    private static string Hex(int red, int green, int blue) => "#" + red.ToString("x2") + green.ToString("x2") + blue.ToString("x2");

    private static bool TryParse(string color, out int red, out int green, out int blue)
    {
        red = green = blue = 0;
        string value = color.Trim().ToLowerInvariant();
        value = value switch
        {
            "black" => "000000", "white" => "ffffff", "red" => "ff0000", "lime" => "00ff00", "green" => "008000",
            "blue" => "0000ff", "yellow" => "ffff00", "cyan" or "aqua" => "00ffff", "magenta" or "fuchsia" => "ff00ff",
            "gray" or "grey" => "808080", "silver" => "c0c0c0", "maroon" => "800000", "navy" => "000080", "olive" => "808000",
            "purple" => "800080", "teal" => "008080",
            _ => value.StartsWith("#", StringComparison.Ordinal) ? value.Substring(1) : value,
        };
        if (value.Length == 3) value = new string([value[0], value[0], value[1], value[1], value[2], value[2]]);
        if (value.Length != 6 || !int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int packed)) return false;
        red = packed >> 16 & 0xff; green = packed >> 8 & 0xff; blue = packed & 0xff;
        return true;
    }

    // sRGB (D65) <-> CIELAB, the standard formulas (IEC 61966-2-1 transfer curve, CIE 1976 L*a*b*).
    private const double WhiteX = 0.95047, WhiteY = 1.0, WhiteZ = 1.08883;

    private static (double L, double A, double B) ToLab(int red, int green, int blue)
    {
        double r = ToLinear(red / 255.0), g = ToLinear(green / 255.0), b = ToLinear(blue / 255.0);
        double x = (0.4124564 * r + 0.3575761 * g + 0.1804375 * b) / WhiteX;
        double y = (0.2126729 * r + 0.7151522 * g + 0.0721750 * b) / WhiteY;
        double z = (0.0193339 * r + 0.1191920 * g + 0.9503041 * b) / WhiteZ;
        double fx = LabF(x), fy = LabF(y), fz = LabF(z);
        return (116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz));
    }

    private static (int Red, int Green, int Blue) FromLab(double l, double a, double b)
    {
        double fy = (l + 16) / 116, fx = fy + a / 500, fz = fy - b / 200;
        double x = LabFInverse(fx) * WhiteX, y = LabFInverse(fy) * WhiteY, z = LabFInverse(fz) * WhiteZ;
        double r = 3.2404542 * x - 1.5371385 * y - 0.4985314 * z;
        double g = -0.9692660 * x + 1.8760108 * y + 0.0415560 * z;
        double bl = 0.0556434 * x - 0.2040259 * y + 1.0572252 * z;
        return (ToByte(r), ToByte(g), ToByte(bl));
    }

    private static double ToLinear(double channel) => channel <= 0.04045 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);

    private static int ToByte(double linear)
    {
        double clamped = Math.Max(0, Math.Min(1, linear));
        double encoded = clamped <= 0.0031308 ? clamped * 12.92 : 1.055 * Math.Pow(clamped, 1 / 2.4) - 0.055;
        return (int)Math.Round(encoded * 255);
    }

    private const double Epsilon = 216.0 / 24389, Kappa = 24389.0 / 27;
    private static double LabF(double t) => t > Epsilon ? Math.Pow(t, 1.0 / 3) : (Kappa * t + 16) / 116;
    private static double LabFInverse(double f) { double cube = f * f * f; return cube > Epsilon ? cube : (116 * f - 16) / Kappa; }
}
