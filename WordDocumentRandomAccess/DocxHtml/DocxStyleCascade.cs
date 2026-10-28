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

    public WordParagraphFormatting DocumentDefaults { get; } = new();

    public DocxStyleCascade(XDocument? stylesDocument)
    {
        XElement? root = stylesDocument?.Root;
        if (root is null) return;

        XElement? documentDefaults = root.Element(W + "docDefaults");
        ApplyRunProperties(DocumentDefaults.RunFormatting, documentDefaults?.Element(W + "rPrDefault")?.Element(W + "rPr"));
        ApplyParagraphProperties(DocumentDefaults, documentDefaults?.Element(W + "pPrDefault")?.Element(W + "pPr"));

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
            ApplyParagraphProperties(formatting, style.Element(W + "pPr"));
            ApplyRunProperties(formatting.RunFormatting, style.Element(W + "rPr"));
        }
        _resolvedParagraphStyles[key] = formatting;
        return formatting;
    }

    /// <summary>Full paragraph formatting: style chain + direct pPr (+ its paragraph-mark rPr as the base run).</summary>
    public WordParagraphFormatting ResolveParagraph(string? styleId, XElement? directParagraphProperties)
    {
        WordParagraphFormatting formatting = ResolveParagraphStyle(styleId).Clone();
        ApplyParagraphProperties(formatting, directParagraphProperties);
        return formatting;
    }

    /// <summary>Run formatting: paragraph base → character style chain → direct rPr.</summary>
    public WordRunFormatting ResolveRun(WordRunFormatting paragraphBase, XElement? directRunProperties)
    {
        WordRunFormatting formatting = paragraphBase.Clone();
        string? characterStyleId = (string?)directRunProperties?.Element(W + "rStyle")?.Attribute(W + "val") ?? _defaultCharacterStyleId;
        foreach (XElement style in StyleChain(characterStyleId)) ApplyRunProperties(formatting, style.Element(W + "rPr"));
        ApplyRunProperties(formatting, directRunProperties);
        return formatting;
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

    public static void ApplyParagraphProperties(WordParagraphFormatting formatting, XElement? pPr)
    {
        if (pPr is null) return;
        if (pPr.Element(W + "bidi") is { } bidi) { formatting.IsRightToLeft = IsOn(bidi.Attribute(W + "val")); formatting.RunFormatting.IsRightToLeft = formatting.IsRightToLeft; }
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
        if (pPr.Element(W + "numPr") is { } numPr)
        {
            formatting.NumberingId = TwipsAttribute(numPr.Element(W + "numId"), "val") ?? -1;
            formatting.NumberingLevel = TwipsAttribute(numPr.Element(W + "ilvl"), "val") ?? 0;
        }
        ApplyRunProperties(formatting.RunFormatting, pPr.Element(W + "rPr"));
    }

    public static void ApplyRunProperties(WordRunFormatting formatting, XElement? rPr)
    {
        if (rPr is null) return;
        foreach (XElement property in rPr.Elements())
        {
            string name = property.Name.LocalName;
            switch (name)
            {
                case "b": case "bCs": formatting.IsBold = IsOn(property.Attribute(W + "val")); break;
                case "i": case "iCs": formatting.IsItalic = IsOn(property.Attribute(W + "val")); break;
                case "u": formatting.IsUnderlined = ((string?)property.Attribute(W + "val") ?? "single") != "none"; break;
                case "strike": case "dstrike": formatting.IsStrikethrough = IsOn(property.Attribute(W + "val")); break;
                case "vanish": formatting.IsHidden = IsOn(property.Attribute(W + "val")); break;
                case "rtl": formatting.IsRightToLeft = IsOn(property.Attribute(W + "val")); break;
                case "vertAlign":
                    string? vertical = (string?)property.Attribute(W + "val");
                    formatting.IsSuperscript = vertical == "superscript";
                    formatting.IsSubscript = vertical == "subscript";
                    break;
                case "sz": formatting.FontSizeHalfPoints = TwipsAttribute(property, "val") ?? formatting.FontSizeHalfPoints; break;
                case "szCs": formatting.ComplexScriptFontSizeHalfPoints = TwipsAttribute(property, "val") ?? formatting.ComplexScriptFontSizeHalfPoints; break;
                case "color":
                    string? color = (string?)property.Attribute(W + "val");
                    formatting.ColorHex = color is null || color.Equals("auto", StringComparison.OrdinalIgnoreCase) || color.Length != 6 ? null : color;
                    break;
                case "highlight": formatting.HighlightCssColor = WordHtmlText.HighlightCss((string?)property.Attribute(W + "val")); break;
                case "shd":
                    string? fill = (string?)property.Attribute(W + "fill");
                    if (fill is { Length: 6 } && !fill.Equals("auto", StringComparison.OrdinalIgnoreCase)) formatting.HighlightCssColor = "#" + fill;
                    break;
                case "rFonts":
                    formatting.FontName = (string?)property.Attribute(W + "ascii") ?? (string?)property.Attribute(W + "hAnsi") ?? formatting.FontName;
                    formatting.ComplexScriptFontName = (string?)property.Attribute(W + "cs") ?? formatting.ComplexScriptFontName;
                    break;
            }
        }
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
