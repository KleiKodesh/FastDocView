using System.Globalization;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// word/theme/theme1.xml (ECMA-376 Part 1 §20.1.6: a:clrScheme, a:fontScheme) plus the two
/// settings.xml hooks that bend it: w:clrSchemeMapping (§17.15.1.20) and w:themeFontLang
/// (§17.15.1.88). A theme part is a few KB, so it is loaded whole once per document.
///
/// Fonts: rFonts asciiTheme / hAnsiTheme / cstheme name a major or minor face; per the spec the
/// theme attribute wins over the explicit one ("If the asciiTheme attribute is also specified, then
/// this attribute shall be ignored" — rFonts §17.3.2.26). For the complex-script slot LibreOffice
/// (writerfilter ThemeHandler::getFontNameForTheme) first looks up the supplemental a:font whose
/// script tag matches w:themeFontLang/@bidi (he-IL → "Hebr") and only then falls back to a:cs —
/// which Word leaves empty in the Office theme. The same order is used here.
///
/// Colors: w:themeColor goes through clrSchemeMapping (background1 → bg1 → light1 → a:lt1, …), the
/// scheme entry gives RRGGBB (a:srgbClr/@val or a:sysClr/@lastClr), then themeTint / themeShade
/// (a byte, 0xFF = unchanged) are applied per RGB component as the spec's worked examples do:
/// tint C' = 255 − (255 − C)·T/255, shade C' = C·S/255 (w:color §17.3.2.6 — accent1 C0504D with
/// tint 99 → D99694). LibreOffice instead scales HSL luminance by the same factor; both agree on the
/// spec example and differ by at most one unit per channel from Word's own rounding.
/// </summary>
public sealed class DocxTheme
{
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace W = DocxStyleCascade.W;

    private readonly ThemeFontSet _major;
    private readonly ThemeFontSet _minor;
    private readonly Dictionary<string, string> _schemeColors = new(StringComparer.Ordinal);       // dk1, lt1, …, accent6, hlink, folHlink → RRGGBB
    private readonly Dictionary<string, string> _schemeMapping = new(StringComparer.Ordinal);      // bg1 → light1, t1 → dark1, accent1 → accent1 …
    private readonly string? _bidiScript;        // ISO 15924 tag from w:themeFontLang/@bidi, e.g. "Hebr"
    private readonly string? _eastAsiaScript;

    public bool HasFonts => _major.Latin is not null || _minor.Latin is not null;
    public bool HasColors => _schemeColors.Count > 0;

    public DocxTheme(XDocument? themeDocument, XDocument? settingsDocument)
    {
        XElement? fontScheme = themeDocument?.Root?.Element(A + "themeElements")?.Element(A + "fontScheme");
        _major = ThemeFontSet.Read(fontScheme?.Element(A + "majorFont"));
        _minor = ThemeFontSet.Read(fontScheme?.Element(A + "minorFont"));

        XElement? colorScheme = themeDocument?.Root?.Element(A + "themeElements")?.Element(A + "clrScheme");
        if (colorScheme is not null)
            foreach (XElement entry in colorScheme.Elements())
            {
                string? hex = (string?)entry.Element(A + "srgbClr")?.Attribute("val") ?? (string?)entry.Element(A + "sysClr")?.Attribute("lastClr");
                if (hex is { Length: 6 } && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) _schemeColors[entry.Name.LocalName] = hex.ToUpperInvariant();
            }

        XElement? settings = settingsDocument?.Root;
        if (settings?.Element(W + "clrSchemeMapping") is { } mapping)
            foreach (XAttribute attribute in mapping.Attributes())
                if (attribute.Name.Namespace == W) _schemeMapping[attribute.Name.LocalName] = attribute.Value;
        if (settings?.Element(W + "themeFontLang") is { } fontLanguage)
        {
            _bidiScript = ScriptTagOfLanguage((string?)fontLanguage.Attribute(W + "bidi"));
            _eastAsiaScript = ScriptTagOfLanguage((string?)fontLanguage.Attribute(W + "eastAsia"));
        }
    }

    // ── fonts ──────────────────────────────────────────────────────────────────────────

    /// <summary>The typeface for an ST_Theme value (minorHAnsi, majorBidi, …); null when the theme has no usable face for it.</summary>
    public string? FontFor(string? themeValue)
    {
        if (themeValue is null) return null;
        ThemeFontSet set = themeValue.StartsWith("major", StringComparison.Ordinal) ? _major : _minor;
        string? face = themeValue switch
        {
            "majorAscii" or "majorHAnsi" or "minorAscii" or "minorHAnsi" => set.Latin,
            "majorBidi" or "minorBidi" => set.SupplementalFor(_bidiScript) ?? set.ComplexScript,
            "majorEastAsia" or "minorEastAsia" => set.SupplementalFor(_eastAsiaScript) ?? set.EastAsian,
            _ => null,
        };
        return string.IsNullOrEmpty(face) ? null : face;
    }

    private readonly struct ThemeFontSet
    {
        public readonly string? Latin;
        public readonly string? EastAsian;
        public readonly string? ComplexScript;
        private readonly Dictionary<string, string>? _byScript;

        private ThemeFontSet(string? latin, string? eastAsian, string? complexScript, Dictionary<string, string>? byScript)
        {
            Latin = latin; EastAsian = eastAsian; ComplexScript = complexScript; _byScript = byScript;
        }

        public static ThemeFontSet Read(XElement? fontSet)
        {
            if (fontSet is null) return default;
            Dictionary<string, string>? byScript = null;
            foreach (XElement font in fontSet.Elements(A + "font"))
            {
                string? script = (string?)font.Attribute("script");
                string? typeface = (string?)font.Attribute("typeface");
                if (script is null || string.IsNullOrEmpty(typeface)) continue;
                (byScript ??= new Dictionary<string, string>(StringComparer.Ordinal))[script] = typeface;
            }
            return new ThemeFontSet(
                (string?)fontSet.Element(A + "latin")?.Attribute("typeface"),
                (string?)fontSet.Element(A + "ea")?.Attribute("typeface"),
                (string?)fontSet.Element(A + "cs")?.Attribute("typeface"),
                byScript);
        }

        public string? SupplementalFor(string? script)
            => script is not null && _byScript is not null && _byScript.TryGetValue(script, out string? face) ? face : null;
    }

    /// <summary>
    /// BCP-47 language → the ISO 15924 script tag the theme's a:font entries use. Only the scripts a
    /// Word bidi / eastAsia language can name; the list mirrors LibreOffice ThemeHandler.cxx
    /// fromLCIDToScriptTag (which keys on LCIDs — settings.xml carries tags, so the primary subtag is matched).
    /// </summary>
    private static string? ScriptTagOfLanguage(string? languageTag)
    {
        if (string.IsNullOrEmpty(languageTag)) return null;
        int dash = languageTag.IndexOf('-');
        string primary = (dash > 0 ? languageTag.Substring(0, dash) : languageTag).ToLowerInvariant();
        return primary switch
        {
            "he" or "yi" => "Hebr",
            "ar" or "fa" or "ur" or "ps" or "prs" or "sd" or "ug" or "ks" or "tzm" => "Arab",
            "syr" => "Syrc",
            "th" => "Thai",
            "dv" => "Thaa",
            "hi" or "mr" or "sa" or "ne" or "kok" => "Deva",
            "bn" or "as" => "Beng",
            "gu" => "Gujr", "pa" => "Guru", "ta" => "Taml", "te" => "Telu", "kn" => "Knda", "ml" => "Mlym", "or" => "Orya", "si" => "Sinh",
            "km" => "Khmr", "lo" => "Laoo", "my" => "Mymr", "bo" => "Tibt", "mn" => "Mong", "am" or "ti" => "Ethi",
            "ja" => "Jpan", "ko" => "Hang",
            "zh" => languageTag.EndsWith("TW", StringComparison.OrdinalIgnoreCase) || languageTag.EndsWith("HK", StringComparison.OrdinalIgnoreCase)
                    || languageTag.EndsWith("MO", StringComparison.OrdinalIgnoreCase) || languageTag.Contains("Hant", StringComparison.OrdinalIgnoreCase) ? "Hant" : "Hans",
            _ => null,
        };
    }

    // ── colors ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// The RRGGBB of an ST_ThemeColor value with its tint / shade bytes (hex strings as written, null = none),
    /// or null when the theme cannot resolve it (no theme part, "none", unknown name).
    /// </summary>
    public string? ColorHex(string? themeColor, string? themeTint, string? themeShade)
    {
        if (themeColor is null || _schemeColors.Count == 0) return null;
        string? schemeEntry = SchemeEntryOf(themeColor);
        if (schemeEntry is null || !_schemeColors.TryGetValue(schemeEntry, out string? hex)) return null;
        int tint = ByteOf(themeTint), shade = ByteOf(themeShade);
        if (tint is < 0 or 255 && shade is < 0 or 255) return hex;
        int r = Convert.ToInt32(hex.Substring(0, 2), 16), g = Convert.ToInt32(hex.Substring(2, 2), 16), b = Convert.ToInt32(hex.Substring(4, 2), 16);
        if (tint is >= 0 and < 255) { r = Tint(r, tint); g = Tint(g, tint); b = Tint(b, tint); }
        if (shade is >= 0 and < 255) { r = Shade(r, shade); g = Shade(g, shade); b = Shade(b, shade); }
        return r.ToString("X2", CultureInfo.InvariantCulture) + g.ToString("X2", CultureInfo.InvariantCulture) + b.ToString("X2", CultureInfo.InvariantCulture);
    }

    private static int Tint(int component, int tint) => (int)Math.Round(255 - (255 - component) * (tint / 255.0), MidpointRounding.AwayFromZero);
    private static int Shade(int component, int shade) => (int)Math.Round(component * (shade / 255.0), MidpointRounding.AwayFromZero);

    private static int ByteOf(string? hexByte)
        => hexByte is { Length: 1 or 2 } && int.TryParse(hexByte, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int value) ? value : -1;

    /// <summary>ST_ThemeColor → (clrSchemeMapping) → ST_ColorSchemeIndex → the a:clrScheme child name.</summary>
    private string? SchemeEntryOf(string themeColor)
    {
        // The document-side names go through the mapping; the mapping's defaults are the spec's (bg1 → light1, t1 → dark1 …).
        string? mappingKey = themeColor switch
        {
            "background1" => "bg1", "text1" => "t1", "background2" => "bg2", "text2" => "t2",
            "accent1" or "accent2" or "accent3" or "accent4" or "accent5" or "accent6" or "hyperlink" or "followedHyperlink" => themeColor,
            _ => null,
        };
        string schemeIndex;
        if (mappingKey is not null)
        {
            if (!_schemeMapping.TryGetValue(mappingKey, out string? mapped))
                mapped = mappingKey switch { "bg1" => "light1", "t1" => "dark1", "bg2" => "light2", "t2" => "dark2", _ => mappingKey };
            schemeIndex = mapped;
        }
        else schemeIndex = themeColor;                        // dark1 / light1 / dark2 / light2 reference the scheme directly

        return schemeIndex switch
        {
            "dark1" => "dk1", "light1" => "lt1", "dark2" => "dk2", "light2" => "lt2",
            "hyperlink" => "hlink", "followedHyperlink" => "folHlink",
            "accent1" or "accent2" or "accent3" or "accent4" or "accent5" or "accent6" => schemeIndex,
            _ => null,
        };
    }

    /// <summary>
    /// The color of an element carrying ECMA-376's color attribute group (w:color/@val, w:u/@color,
    /// w:bdr/@color: <paramref name="valueAttribute"/>) plus themeColor / themeTint / themeShade. The
    /// theme wins when it resolves ("If the themeColor attribute is specified, then the val attribute
    /// is ignored", §17.3.2.6); otherwise the explicit value is used. "auto" and malformed values → null.
    /// </summary>
    public static string? ResolveColorHex(XElement element, string valueAttribute, DocxTheme? theme,
                                          string themeColorAttribute = "themeColor", string themeTintAttribute = "themeTint", string themeShadeAttribute = "themeShade")
    {
        string? fromTheme = theme?.ColorHex((string?)element.Attribute(W + themeColorAttribute),
                                            (string?)element.Attribute(W + themeTintAttribute),
                                            (string?)element.Attribute(W + themeShadeAttribute));
        if (fromTheme is not null) return fromTheme;
        string? explicitValue = (string?)element.Attribute(W + valueAttribute);
        return explicitValue is { Length: 6 } && !explicitValue.Equals("auto", StringComparison.OrdinalIgnoreCase) ? explicitValue : null;
    }
}
