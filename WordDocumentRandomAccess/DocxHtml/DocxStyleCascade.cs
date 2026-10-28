using System.Globalization;
using System.Xml.Linq;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// word/styles.xml resolver (ECMA-376 Part 1 §17.7): document defaults → the style chain through
/// w:basedOn → direct pPr/rPr. styles.xml is small, so it is loaded fully once per document; it
/// also tells which styles are headings (w:outlineLvl or built-in "heading N" names) and whether
/// the document reads right-to-left by default.
/// </summary>
public sealed class DocxStyleCascade
{
    public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly Dictionary<string, XElement> _stylesById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, WordParagraphFormatting> _resolvedParagraphStyles = new(StringComparer.Ordinal);
    private readonly string? _defaultParagraphStyleId;
    private readonly string? _defaultCharacterStyleId;
    private readonly DocxTheme? _theme;

    public WordParagraphFormatting DocumentDefaults { get; } = new();

    /// <summary>The document theme (fonts + colors) the cascade resolves theme references through; null = no theme part.</summary>
    public DocxTheme? Theme => _theme;

    public DocxStyleCascade(XDocument? stylesDocument, DocxTheme? theme = null)
    {
        _theme = theme;
        XElement? root = stylesDocument?.Root;
        if (root is null) return;

        XElement? documentDefaults = root.Element(W + "docDefaults");
        ApplyRunProperties(DocumentDefaults.RunFormatting, documentDefaults?.Element(W + "rPrDefault")?.Element(W + "rPr"), _theme, WordStyleLayer.DocumentDefaults);
        ResolveToggles(DocumentDefaults.RunFormatting);
        ApplyParagraphProperties(DocumentDefaults, documentDefaults?.Element(W + "pPrDefault")?.Element(W + "pPr"), _theme);

        foreach (XElement style in root.Elements(W + "style"))
        {
            string? id = (string?)style.Attribute(W + "styleId");
            if (id is null) continue;
            _stylesById[id] = style;
            string? type = (string?)style.Attribute(W + "type");
            // IsOn treats a missing attribute as "on" (right for <w:b/>); a missing w:default means NOT default.
            bool isDefault = style.Attribute(W + "default") is { } defaultAttribute && IsOn(defaultAttribute);
            if (isDefault && type == "paragraph") _defaultParagraphStyleId = id;
            if (isDefault && type == "character") _defaultCharacterStyleId = id;
        }
    }

    /// <summary>The paragraph style resolved through its chain (cached), without direct formatting.</summary>
    public WordParagraphFormatting ResolveParagraphStyle(string? styleId)
    {
        string key = styleId ?? _defaultParagraphStyleId ?? "";
        if (_resolvedParagraphStyles.TryGetValue(key, out WordParagraphFormatting? cached)) return cached;

        WordParagraphFormatting formatting = DocumentDefaults.Clone();
        foreach (XElement style in StyleChain(key))
        {
            ApplyParagraphProperties(formatting, style.Element(W + "pPr"), _theme);
            ApplyRunProperties(formatting.RunFormatting, style.Element(W + "rPr"), _theme, WordStyleLayer.ParagraphStyle);
        }
        ResolveToggles(formatting.RunFormatting);
        _resolvedParagraphStyles[key] = formatting;
        return formatting;
    }

    /// <summary>
    /// Inside a table, Word layers the table style's pPr/rPr between the document defaults and the
    /// paragraph style ([MS-OI29500] / ECMA-376 §17.7.2). Cached per (table style, paragraph style).
    /// </summary>
    public WordParagraphFormatting ResolveParagraphStyleInTable(string? tableStyleId, string? styleId)
    {
        if (tableStyleId is null) return ResolveParagraphStyle(styleId);
        string key = tableStyleId + "\u0001" + (styleId ?? _defaultParagraphStyleId ?? "");
        if (_resolvedParagraphStyles.TryGetValue(key, out WordParagraphFormatting? cached)) return cached;
        WordParagraphFormatting formatting = DocumentDefaults.Clone();
        foreach (XElement style in StyleChain(tableStyleId))
        {
            ApplyParagraphProperties(formatting, style.Element(W + "pPr"), _theme);
            ApplyRunProperties(formatting.RunFormatting, style.Element(W + "rPr"), _theme, WordStyleLayer.TableStyle);
        }
        foreach (XElement style in StyleChain(styleId ?? _defaultParagraphStyleId))
        {
            ApplyParagraphProperties(formatting, style.Element(W + "pPr"), _theme);
            ApplyRunProperties(formatting.RunFormatting, style.Element(W + "rPr"), _theme, WordStyleLayer.ParagraphStyle);
        }
        ResolveToggles(formatting.RunFormatting);
        _resolvedParagraphStyles[key] = formatting;
        return formatting;
    }

    /// <summary>A table style's own w:tblPr / w:tcPr children through its basedOn chain, most specific last (borders, cell margins, shading).</summary>
    public IEnumerable<XElement> TableStyleProperties(string? tableStyleId, string elementName)
    {
        foreach (XElement style in StyleChain(tableStyleId))
            if (style.Element(W + elementName) is { } properties) yield return properties;
    }

    /// <summary>Full paragraph formatting: style chain + direct pPr (+ its paragraph-mark rPr as the base run).</summary>
    public WordParagraphFormatting ResolveParagraph(string? styleId, XElement? directParagraphProperties, string? tableStyleId = null)
    {
        WordParagraphFormatting formatting = ResolveParagraphStyleInTable(tableStyleId, styleId).Clone();
        ApplyParagraphProperties(formatting, directParagraphProperties, _theme);
        ResolveToggles(formatting.RunFormatting);
        formatting.StyleKey = styleId ?? _defaultParagraphStyleId ?? "";
        return formatting;
    }

    /// <summary>Run formatting: paragraph base → character style chain → direct rPr, toggle properties combined per §17.7.3 (<see cref="WordToggleLayers"/>).</summary>
    public WordRunFormatting ResolveRun(WordRunFormatting paragraphBase, XElement? directRunProperties)
    {
        WordRunFormatting formatting = paragraphBase.Clone();
        string? characterStyleId = (string?)directRunProperties?.Element(W + "rStyle")?.Attribute(W + "val") ?? _defaultCharacterStyleId;
        foreach (XElement style in StyleChain(characterStyleId)) ApplyRunProperties(formatting, style.Element(W + "rPr"), _theme, WordStyleLayer.CharacterStyle);
        ApplyRunProperties(formatting, directRunProperties, _theme, WordStyleLayer.Direct);
        ResolveToggles(formatting);
        return formatting;
    }

    /// <summary>
    /// The toggle layers → the resolved flags. Bold / italic come from b / i or from bCs / iCs depending on
    /// whether the run is complex script (w:rtl): in Word 16, a run with w:rtl draws all its text with
    /// bCs / iCs / szCs / the cs font, and a run without it uses b / i / sz / ascii even for Hebrew letters
    /// (checked by PDF export of a synthetic document, 2026-10-28). LibreOffice keeps the pairs apart too
    /// (DomainMapper.cxx: b → CharWeight, bCs → CharWeightComplex; i → CharPosture, iCs → CharPostureComplex).
    /// </summary>
    public static void ResolveToggles(WordRunFormatting formatting)
    {
        bool isComplexScript = formatting.UsesComplexScriptProperties;
        WordToggleLayers toggles = formatting.Toggles;
        formatting.IsBold = toggles.IsOn(isComplexScript ? WordToggleProperty.BoldComplexScript : WordToggleProperty.Bold);
        formatting.IsItalic = toggles.IsOn(isComplexScript ? WordToggleProperty.ItalicComplexScript : WordToggleProperty.Italic);
        formatting.IsAllCaps = toggles.IsOn(WordToggleProperty.Caps);
        formatting.IsSmallCaps = toggles.IsOn(WordToggleProperty.SmallCaps);
        formatting.IsStrikethrough = toggles.IsOn(WordToggleProperty.Strike);
        formatting.IsEmboss = toggles.IsOn(WordToggleProperty.Emboss);
        formatting.IsImprint = toggles.IsOn(WordToggleProperty.Imprint);
        formatting.IsOutline = toggles.IsOn(WordToggleProperty.Outline);
        formatting.IsShadow = toggles.IsOn(WordToggleProperty.Shadow);
        formatting.IsHidden = toggles.IsOn(WordToggleProperty.Vanish) || formatting.IsWebHidden;
    }

    /// <summary>Heading level 1-9 for a style id, or 0 when the style is not a heading.</summary>
    public int HeadingLevelOf(string? styleId)
    {
        if (styleId is null) return 0;
        int outline = ResolveParagraphStyle(styleId).OutlineLevel;
        if (outline is >= 0 and <= 8) return outline + 1;
        if (!_stylesById.TryGetValue(styleId, out XElement? style)) return 0;
        string? name = ((string?)style.Element(W + "name")?.Attribute(W + "val"))?.ToLowerInvariant();
        if (name is not null && name.StartsWith("heading ", StringComparison.Ordinal)
            && int.TryParse(name.Substring(8), NumberStyles.Integer, CultureInfo.InvariantCulture, out int level) && level is >= 1 and <= 9)
            return level;
        return 0;
    }

    /// <summary>Numbering a paragraph inherits from its style (numId, ilvl); (-1, 0) when none.</summary>
    public (int numId, int level) NumberingOf(string? styleId)
    {
        WordParagraphFormatting formatting = ResolveParagraphStyle(styleId);
        return (formatting.NumberingId, formatting.NumberingLevel);
    }

    public bool IsDocumentRightToLeft => ResolveParagraphStyle(null).IsRightToLeft || DocumentDefaults.RunFormatting.IsRightToLeft;

    private List<XElement> StyleChain(string? styleId)
    {
        var chain = new List<XElement>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (styleId is not null && seen.Add(styleId) && _stylesById.TryGetValue(styleId, out XElement? style) && chain.Count < 32)
        {
            chain.Add(style);
            styleId = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
        }
        chain.Reverse();                                   // base first, most specific last
        return chain;
    }

    // ── property application (shared with the block renderer for direct formatting) ─────

    public static void ApplyParagraphProperties(WordParagraphFormatting formatting, XElement? pPr) => ApplyParagraphProperties(formatting, pPr, null);

    public static void ApplyParagraphProperties(WordParagraphFormatting formatting, XElement? pPr, DocxTheme? theme)
    {
        if (pPr is null) return;
        // w:bidi sets the paragraph's direction only: it does not make its runs complex script (Word 16 draws a bidi paragraph's
        // run without w:rtl with the ascii font / b / sz — PDF export check, 2026-10-28).
        if (pPr.Element(W + "bidi") is { } bidi) formatting.IsRightToLeft = IsOn(bidi.Attribute(W + "val"));
        if ((string?)pPr.Element(W + "jc")?.Attribute(W + "val") is { } jc)
            formatting.Alignment = jc switch
            {
                "center" => WordParagraphAlignment.Center,
                "right" or "end" => WordParagraphAlignment.End,
                "both" or "distribute" => WordParagraphAlignment.Justify,
                _ => WordParagraphAlignment.Start,
            };
        if (pPr.Element(W + "ind") is { } ind)
        {
            if ((TwipsAttribute(ind, "start") ?? TwipsAttribute(ind, "left")) is { } left) formatting.LeftIndentTwips = left;
            if ((TwipsAttribute(ind, "end") ?? TwipsAttribute(ind, "right")) is { } right) formatting.RightIndentTwips = right;
            if (TwipsAttribute(ind, "hanging") is { } hanging) formatting.FirstLineIndentTwips = -hanging;
            else if (TwipsAttribute(ind, "firstLine") is { } firstLine) formatting.FirstLineIndentTwips = firstLine;
        }
        if (pPr.Element(W + "spacing") is { } spacing)
        {
            if (TwipsAttribute(spacing, "before") is { } before) formatting.SpaceBeforeTwips = before;
            if (TwipsAttribute(spacing, "after") is { } after) formatting.SpaceAfterTwips = after;
            if (TwipsAttribute(spacing, "line") is { } line)
            {
                formatting.LineSpacingTwipsOrAuto = line;
                formatting.IsLineSpacingAuto = ((string?)spacing.Attribute(W + "lineRule") ?? "auto") == "auto";
            }
        }
        if (TwipsAttribute(pPr.Element(W + "outlineLvl"), "val") is { } outline) formatting.OutlineLevel = outline;
        if (pPr.Element(W + "pBdr") is { } borders)
        {
            formatting.BorderTopCss = DocxTableFormatting.BorderCss(borders.Element(W + "top"), theme) ?? formatting.BorderTopCss;
            formatting.BorderBottomCss = DocxTableFormatting.BorderCss(borders.Element(W + "bottom"), theme) ?? formatting.BorderBottomCss;
            formatting.BorderStartCss = DocxTableFormatting.BorderCss(borders.Element(W + "start") ?? borders.Element(W + "left"), theme) ?? formatting.BorderStartCss;
            formatting.BorderEndCss = DocxTableFormatting.BorderCss(borders.Element(W + "end") ?? borders.Element(W + "right"), theme) ?? formatting.BorderEndCss;
        }
        if (pPr.Element(W + "contextualSpacing") is { } contextual) formatting.IsContextualSpacing = IsOn(contextual.Attribute(W + "val"));
        if (pPr.Element(W + "shd") is { } shading) formatting.ShadingHex = FillHex(shading, theme);
        if (pPr.Element(W + "tabs") is { } tabs)
        {
            // Flow view has no ruler; the first explicit stop (relative to the indent) becomes tab-size.
            int? firstStop = tabs.Elements(W + "tab")
                .Where(tab => (string?)tab.Attribute(W + "val") is not ("clear" or "bar"))
                .Select(tab => TwipsAttribute(tab, "pos")).Where(position => position > 0).Min();
            if (firstStop is { } stop) formatting.FirstTabStopTwips = Math.Max(120, stop - formatting.LeftIndentTwips);
        }
        if (pPr.Element(W + "numPr") is { } numPr)
        {
            formatting.NumberingId = TwipsAttribute(numPr.Element(W + "numId"), "val") ?? -1;
            formatting.NumberingLevel = TwipsAttribute(numPr.Element(W + "ilvl"), "val") ?? 0;
        }
        ApplyRunProperties(formatting.RunFormatting, pPr.Element(W + "rPr"), theme, WordStyleLayer.Direct);
    }

    /// <param name="layer">Where <paramref name="rPr"/> sits in the style hierarchy: toggle properties combine per layer (<see cref="WordToggleLayers"/>);
    /// call <see cref="ResolveToggles"/> after the last layer.</param>
    public static void ApplyRunProperties(WordRunFormatting formatting, XElement? rPr, DocxTheme? theme, WordStyleLayer layer)
    {
        if (rPr is null) return;
        foreach (XElement property in rPr.Elements())
        {
            string name = property.Name.LocalName;
            switch (name)
            {
                case "b": formatting.Toggles.Set(layer, WordToggleProperty.Bold, IsOn(property.Attribute(W + "val"))); break;
                case "bCs": formatting.Toggles.Set(layer, WordToggleProperty.BoldComplexScript, IsOn(property.Attribute(W + "val"))); break;
                case "i": formatting.Toggles.Set(layer, WordToggleProperty.Italic, IsOn(property.Attribute(W + "val"))); break;
                case "iCs": formatting.Toggles.Set(layer, WordToggleProperty.ItalicComplexScript, IsOn(property.Attribute(W + "val"))); break;
                case "u": ApplyUnderline(formatting, property, theme); break;
                case "em":
                    // ST_Em (§17.18.24): dot / comma / circle above the glyphs, underDot below; CSS text-emphasis draws the same marks.
                    string? emphasis = (string?)property.Attribute(W + "val");
                    formatting.EmphasisMarkCss = emphasis switch { "dot" or "underDot" => "filled dot", "comma" => "filled sesame", "circle" => "open circle", _ => null };
                    formatting.IsEmphasisMarkBelow = emphasis == "underDot";
                    break;
                case "strike": formatting.Toggles.Set(layer, WordToggleProperty.Strike, IsOn(property.Attribute(W + "val"))); break;
                case "dstrike": formatting.IsDoubleStrikethrough = IsOn(property.Attribute(W + "val")); break;   // not a toggle property
                case "vanish": formatting.Toggles.Set(layer, WordToggleProperty.Vanish, IsOn(property.Attribute(W + "val"))); break;
                case "webHidden": formatting.IsWebHidden = IsOn(property.Attribute(W + "val")); break;
                case "smallCaps": formatting.Toggles.Set(layer, WordToggleProperty.SmallCaps, IsOn(property.Attribute(W + "val"))); break;
                case "position": formatting.PositionHalfPoints = TwipsAttribute(property, "val") ?? formatting.PositionHalfPoints; break;
                case "outline": formatting.Toggles.Set(layer, WordToggleProperty.Outline, property.Name.Namespace != W || IsOn(property.Attribute(W + "val"))); break;
                case "shadow": formatting.Toggles.Set(layer, WordToggleProperty.Shadow, property.Name.Namespace != W || IsOn(property.Attribute(W + "val"))); break;   // w14:shadow / w14:textOutline effects: drawn
                case "emboss": formatting.Toggles.Set(layer, WordToggleProperty.Emboss, IsOn(property.Attribute(W + "val"))); break;
                case "imprint": formatting.Toggles.Set(layer, WordToggleProperty.Imprint, IsOn(property.Attribute(W + "val"))); break;
                case "bdr": formatting.BorderCss = DocxTableFormatting.BorderCss(property, theme); break;
                case "textFill":
                    // Word 2010 text fill (w14): a solid fill is just a color; gradients fall back to their first stop.
                    string? fillColor = property.Descendants().FirstOrDefault(e => e.Name.LocalName == "srgbClr")?.Attributes()
                        .FirstOrDefault(a => a.Name.LocalName == "val")?.Value;
                    if (fillColor is { Length: 6 }) formatting.ColorHex = fillColor;
                    break;
                case "caps": formatting.Toggles.Set(layer, WordToggleProperty.Caps, IsOn(property.Attribute(W + "val"))); break;
                case "spacing": formatting.CharacterSpacingTwips = TwipsAttribute(property, "val") ?? formatting.CharacterSpacingTwips; break;
                case "rtl": formatting.IsRightToLeft = IsOn(property.Attribute(W + "val")); break;
                // w:cs (§17.3.2.7): use the complex-script properties for this run's text whatever its characters (Word 16: like w:rtl, without the direction).
                case "cs": formatting.IsComplexScript = IsOn(property.Attribute(W + "val")); break;
                case "vertAlign":
                    string? vertical = (string?)property.Attribute(W + "val");
                    formatting.IsSuperscript = vertical == "superscript";
                    formatting.IsSubscript = vertical == "subscript";
                    break;
                case "sz": formatting.FontSizeHalfPoints = TwipsAttribute(property, "val") ?? formatting.FontSizeHalfPoints; break;
                case "szCs": formatting.ComplexScriptFontSizeHalfPoints = TwipsAttribute(property, "val") ?? formatting.ComplexScriptFontSizeHalfPoints; break;
                case "color":
                    // themeColor (+ tint / shade) supersedes val when the theme resolves it (§17.3.2.6); val is Word's precomputed fallback.
                    formatting.ColorHex = DocxTheme.ResolveColorHex(property, "val", theme);
                    break;
                case "highlight": formatting.HighlightCssColor = WordHtmlText.HighlightCss((string?)property.Attribute(W + "val")); break;
                case "shd":
                    if (DocxTheme.ResolveColorHex(property, "fill", theme, "themeFill", "themeFillTint", "themeFillShade") is { } fill) formatting.HighlightCssColor = "#" + fill;
                    break;
                case "rFonts":
                    // The theme attribute wins over the explicit one (rFonts §17.3.2.26: "If the asciiTheme attribute is also
                    // specified, then this attribute shall be ignored"); a theme slot that resolves to nothing falls back to the explicit face.
                    formatting.FontName = theme?.FontFor((string?)property.Attribute(W + "asciiTheme")) ?? (string?)property.Attribute(W + "ascii")
                                       ?? theme?.FontFor((string?)property.Attribute(W + "hAnsiTheme")) ?? (string?)property.Attribute(W + "hAnsi") ?? formatting.FontName;
                    formatting.ComplexScriptFontName = theme?.FontFor((string?)property.Attribute(W + "cstheme")) ?? (string?)property.Attribute(W + "cs") ?? formatting.ComplexScriptFontName;
                    formatting.EastAsianFontName = theme?.FontFor((string?)property.Attribute(W + "eastAsiaTheme")) ?? (string?)property.Attribute(W + "eastAsia") ?? formatting.EastAsianFontName;
                    break;
            }
        }
    }

    /// <summary>
    /// w:u (§17.3.2.40). The pattern (ST_Underline §17.18.99) → CSS: double; dotted / dottedHeavy → dotted; the dash
    /// family (dash, dashLong, dotDash, dotDotDash and their Heavy forms) → dashed (CSS has one dash pattern; LibreOffice
    /// keeps DASH / LONGDASH / DASHDOT / DASHDOTDOT apart — DomainMapper::handleUnderlineType); wave / wavyHeavy /
    /// wavyDouble → wavy (CSS has no double wave); thick and every *Heavy → a thicker line; words → single, but not
    /// under spaces ("no underline beneath any space character (breaking or non-breaking)"); none → off.
    /// A w:u without w:val leaves the pattern as inherited — LibreOffice only sets the underline from the val attribute,
    /// and Word 16 draws no line for a bare &lt;w:u/&gt; (PDF export check, 2026-10-28). color / themeColor (+ tint / shade)
    /// → text-decoration-color; themeColor supersedes color; "auto" = the text color.
    /// </summary>
    private static void ApplyUnderline(WordRunFormatting formatting, XElement underline, DocxTheme? theme)
    {
        if ((string?)underline.Attribute(W + "val") is { } pattern)
        {
            formatting.IsUnderlined = pattern != "none";
            formatting.IsUnderlineWordsOnly = pattern == "words";
            formatting.UnderlineStyleCss = pattern switch
            {
                "double" => "double",
                "dotted" or "dottedHeavy" => "dotted",
                "dash" or "dashedHeavy" or "dashLong" or "dashLongHeavy" or "dotDash" or "dashDotHeavy" or "dotDotDash" or "dashDotDotHeavy" => "dashed",
                "wave" or "wavyHeavy" or "wavyDouble" => "wavy",
                _ => null,                                        // single, words, thick
            };
            formatting.IsUnderlineThick = pattern is "thick" or "dottedHeavy" or "dashedHeavy" or "dashLongHeavy" or "dashDotHeavy" or "dashDotDotHeavy" or "wavyHeavy";
        }
        if (underline.Attribute(W + "color") is not null || underline.Attribute(W + "themeColor") is not null)
            formatting.UnderlineColorHex = DocxTheme.ResolveColorHex(underline, "color", theme);
    }

    /// <summary>A w:shd's fill as "RRGGBB" (themeFill + tint / shade first, then fill), or null for auto / clear / white.</summary>
    public static string? FillHex(XElement shading) => FillHex(shading, null);

    public static string? FillHex(XElement shading, DocxTheme? theme)
    {
        string? fill = DocxTheme.ResolveColorHex(shading, "fill", theme, "themeFill", "themeFillTint", "themeFillShade");
        return fill is not null && !fill.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase) ? fill : null;
    }

    public static bool IsOn(XAttribute? value)
    {
        string? text = (string?)value;
        return text is null || text is "1" or "true" or "on";
    }

    public static int? TwipsAttribute(XElement? element, string attributeName)
    {
        string? text = (string?)element?.Attribute(W + attributeName);
        return text is not null && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : null;
    }
}
