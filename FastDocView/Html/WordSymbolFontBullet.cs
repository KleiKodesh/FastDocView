using System.Text;

namespace FastDocView.Html;

/// <summary>
/// A bullet written in a symbol font: Word stores the font's glyph code as U+F020..U+F0FF (the symbol-font cmap range) and
/// names the font in the level's run properties. Symbol is recoded to real Unicode with LibreOffice's Adobe-Symbol table
/// (unotools/source/misc/fontcvt.cxx:707 aAdobeSymbolToAppleSymbolTab, used for "symbol" at :1332). LibreOffice recodes
/// Wingdings / Wingdings 2 / Wingdings 3 / Webdings (fontcvt.cxx:185-482, :1324-1327) only into OpenSymbol's private-use
/// area (197 of Wingdings' 224 codes go to U+E4xx), which no browser can draw without OpenSymbol; those bullets keep their
/// code and their font, as Word itself draws them (the fonts ship with Windows). Any other font: a plain bullet.
/// </summary>
public static class WordSymbolFontBullet
{
    private const char Bullet = (char)0x2022;

    // fontcvt.cxx:707-781 aAdobeSymbolToAppleSymbolTab, F020..F0FF
    private const string AdobeSymbolToUnicode =
        " !∀#∃%&∍()∗+,−./" +
        "0123456789:;<=>?" +
        "≅ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟ" +
        "\u03A0\u0398\u03A1\u03A3\u03A4\u03A5\u03C2\u03A9\u039E\u03A8\u0396\u005B\u2234\u005D\u22A5\u005F" +
        "\uF8E5\u03B1\u03B2\u03C7\u03B4\u03B5\u03C6\u03B3\u03B7\u03B9\u03D5\u03BA\u03BB\u03BC\u03BD\u03BF" +
        "\u03C0\u03B8\u03C1\u03C3\u03C4\u03C5\u03D6\u03C9\u03BE\u03C8\u03B6\u007B\u007C\u007D\u223C\u007F" +
        "\u0080\u0081\u0082\u0083\u0084\u0085\u0086\u0087\u0088\u0089\u008A\u008B\u008C\u008D\u008E\u008F" +
        "\u0090\u0091\u0092\u0093\u0094\u0095\u0096\u0097\u0098\u0099\u009A\u009B\u009C\u009D\u009E\u009F" +
        "\u20AC\u03D2\u2032\u2264\u2044\u221E\u0192\u2663\u2666\u2665\u2660\u2194\u2190\u2191\u2192\u2193" +
        "\u00B0\u00B1\u2033\u2065\u00D7\u221D\u2202\u2022\u00F7\u2260\u2261\u2248\u2026\u23D0\u23AF\u21B5" +
        "\u2135\u2111\u211C\u2118\u2297\u2295\u2205\u2229\u222A\u2283\u2287\u2284\u2282\u2286\u2208\u2209" +
        "\u2220\u2207\u00AE\u00A9\u2122\u220F\u221A\u22C5\u00AC\u2227\u2228\u21D4\u21D0\u21D1\u21D2\u21D3" +
        "\u25CA\u3008\u00AE\u00A9\u2122\u2211\u239B\u239C\u239D\u23A1\u23A2\u23A3\u23A7\u23A8\u23A9\u23AA" +
        "\uF8FF\u3009\u222B\u2320\u23AE\u2321\u239E\u239F\u23A0\u23A4\u23A5\u23A6\u23AB\u23AC\u23AD\u00FF";

    /// <summary>
    /// The text to draw for a bullet's level text in <paramref name="fontName"/>, and whether the label keeps that font
    /// (true only for the Windows symbol fonts whose codes are left as they are).
    /// </summary>
    public static (string text, bool keepsFont) Resolve(string text, string? fontName)
    {
        if (text.Length == 0) return (Bullet.ToString(), false);
        bool hasSymbolCode = false;
        foreach (char c in text) if (IsSymbolCode(c)) { hasSymbolCode = true; break; }
        if (!hasSymbolCode) return (text, true);
        string font = Normalize(fontName);
        if (font is "wingdings" or "wingdings2" or "wingdings3" or "webdings") return (text, true);
        var resolved = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            if (!IsSymbolCode(c)) { resolved.Append(c); continue; }
            char mapped = font == "symbol" ? AdobeSymbolToUnicode[c - 0xF020] : '\0';
            resolved.Append(IsDrawable(mapped) ? mapped : Bullet);
        }
        return (resolved.ToString(), false);
    }

    /// <summary>The table's unassigned codes (C1 controls, Apple private-use) draw as a plain bullet.</summary>
    private static bool IsDrawable(char c) => c > ' ' && c is not ((>= (char)0x7F and <= (char)0x9F) or (>= (char)0xE000 and <= (char)0xF8FF));

    private static bool IsSymbolCode(char c) => c is >= (char)0xF020 and <= (char)0xF0FF;

    private static string Normalize(string? fontName)
    {
        if (fontName is null) return "";
        var name = new StringBuilder(fontName.Length);
        foreach (char c in fontName) if (c != ' ') name.Append(char.ToLowerInvariant(c));
        return name.ToString();
    }
}
