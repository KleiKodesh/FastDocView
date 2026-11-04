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
            "blue" => "0000ff", "yellow" => "ffff00", "cyan" => "00ffff", "magenta" => "ff00ff",
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
