using System.Globalization;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

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
    private readonly Dictionary<string, List<XElement>> _styleChains = new(StringComparer.Ordinal);
    private readonly Dictionary<(string tableStyleId, DocxTableConditions conditions), List<XElement>> _tableCellProperties = new();
    private readonly string? _defaultParagraphStyleId;
    private readonly string? _defaultCharacterStyleId;
    private readonly DocxTheme? _theme;

    public WordParagraphFormatting DocumentDefaults { get; } = new();

    /// <summary>settings.xml compatibility options that change the cascade (table style vs default paragraph style).</summary>
    public DocxCompatibilitySettings Compatibility { get; init; } = DocxCompatibilitySettings.Default;

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
    /// Inside a table, Word layers the table style's pPr/rPr between the document defaults and the paragraph style
    /// ([MS-OI29500] / ECMA-376 §17.7.2): the whole-table part first, then the conditional parts that apply to the cell
    /// (<see cref="DocxTableConditionOrder.Layers"/>). Then LibreOffice's Word rule (DomainMapperTableHandler.cxx, "a very
    /// strange compatibility rule"): unless settings.xml has overrideTableStyleFontSizeAndJustification, a paragraph in the
    /// default paragraph style whose size is 11 or 12 pt takes the table style's size, and one whose alignment is left takes
    /// the table style's. Cached per (table style, conditions, paragraph style).
    /// </summary>
    public WordParagraphFormatting ResolveParagraphStyleInTable(string? tableStyleId, string? styleId, DocxTableConditions conditions = DocxTableConditions.None)
    {
        if (tableStyleId is null) return ResolveParagraphStyle(styleId);
        string paragraphStyle = styleId ?? _defaultParagraphStyleId ?? "";
        string key = tableStyleId + "\u0001" + ((int)conditions).ToString(CultureInfo.InvariantCulture) + "\u0001" + paragraphStyle;
        if (_resolvedParagraphStyles.TryGetValue(key, out WordParagraphFormatting? cached)) return cached;
        WordParagraphFormatting formatting = DocumentDefaults.Clone();
        foreach ((XElement? pPr, XElement? rPr) in TableStyleTextLayers(tableStyleId, conditions))
        {
            ApplyParagraphProperties(formatting, pPr, _theme);
            ApplyRunProperties(formatting.RunFormatting, rPr, _theme, WordStyleLayer.TableStyle);
        }
        WordParagraphFormatting tableOnly = formatting.Clone();
        foreach (XElement style in StyleChain(styleId ?? _defaultParagraphStyleId))
        {
            ApplyParagraphProperties(formatting, style.Element(W + "pPr"), _theme);
            ApplyRunProperties(formatting.RunFormatting, style.Element(W + "rPr"), _theme, WordStyleLayer.ParagraphStyle);
        }
        if (!Compatibility.OverridesTableStyleFontSizeAndJustification && paragraphStyle == (_defaultParagraphStyleId ?? ""))
        {
            WordParagraphFormatting paragraphOnly = ResolveParagraphStyle(styleId);
            if (paragraphOnly.RunFormatting.FontSizeHalfPoints is 22 or 24 && tableOnly.RunFormatting.FontSizeHalfPoints != DocumentDefaults.RunFormatting.FontSizeHalfPoints)
                formatting.RunFormatting.FontSizeHalfPoints = tableOnly.RunFormatting.FontSizeHalfPoints;
            if (paragraphOnly.Alignment is WordParagraphAlignment.Start && tableOnly.Alignment is { } tableAlignment)
                formatting.Alignment = tableAlignment;
        }
        ResolveToggles(formatting.RunFormatting);
        _resolvedParagraphStyles[key] = formatting;
        return formatting;
    }

    /// <summary>The table style's pPr / rPr pairs in layering order: each chain style's whole-table part, then the applying conditional parts.</summary>
    private IEnumerable<(XElement? pPr, XElement? rPr)> TableStyleTextLayers(string tableStyleId, DocxTableConditions conditions)
    {
        List<XElement> chain = StyleChain(tableStyleId);
        foreach (XElement style in chain) yield return (style.Element(W + "pPr"), style.Element(W + "rPr"));
        foreach (XElement part in TableConditionalParts(chain, conditions)) yield return (part.Element(W + "pPr"), part.Element(W + "rPr"));
    }

    private static IEnumerable<XElement> TableConditionalParts(List<XElement> chain, DocxTableConditions conditions)
    {
        if (conditions == DocxTableConditions.None) yield break;
        foreach ((DocxTableConditions condition, string type) in DocxTableConditionOrder.Layers)
        {
            if ((conditions & condition) == 0) continue;
            foreach (XElement style in chain)
                foreach (XElement part in style.Elements(W + "tblStylePr"))
                    if ((string?)part.Attribute(W + "type") == type) yield return part;
        }
    }

    /// <summary>Whether the table style (through its chain) has a conditional part of this type.</summary>
    public bool HasTableConditionalPart(string? tableStyleId, DocxTableConditions condition)
        => tableStyleId is not null && TableConditionalParts(StyleChain(tableStyleId), condition).Any();

    /// <summary>The cell properties (w:tcPr) a table style gives a cell: whole-table first, then the applying conditional parts.</summary>
    public List<XElement> TableStyleCellProperties(string? tableStyleId, DocxTableConditions conditions)
    {
        if (tableStyleId is null) return [];
        if (_tableCellProperties.TryGetValue((tableStyleId, conditions), out List<XElement>? cached)) return cached;
        var layers = new List<XElement>();
        _tableCellProperties[(tableStyleId, conditions)] = layers;
        List<XElement> chain = StyleChain(tableStyleId);
        foreach (XElement style in chain) if (style.Element(W + "tcPr") is { } tcPr) layers.Add(tcPr);
        foreach (XElement part in TableConditionalParts(chain, conditions)) if (part.Element(W + "tcPr") is { } tcPr) layers.Add(tcPr);
        return layers;
    }

    /// <summary>A table style's own w:tblPr / w:tcPr children through its basedOn chain, most specific last (borders, cell margins, shading).</summary>
    public IEnumerable<XElement> TableStyleProperties(string? tableStyleId, string elementName)
    {
        foreach (XElement style in StyleChain(tableStyleId))
            if (style.Element(W + elementName) is { } properties) yield return properties;
    }

    /// <summary>Full paragraph formatting: style chain + direct pPr (+ its paragraph-mark rPr as the base run).</summary>
    public WordParagraphFormatting ResolveParagraph(string? styleId, XElement? directParagraphProperties, string? tableStyleId = null,
                                                    DocxTableConditions tableConditions = DocxTableConditions.None,
                                                    XElement? numberingParagraphProperties = null)
    {
        WordParagraphFormatting formatting = ResolveParagraphStyleInTable(tableStyleId, styleId, tableConditions).Clone();
        if (numberingParagraphProperties is not null) ApplyNumberingLevel(formatting, styleId, numberingParagraphProperties, directParagraphProperties);
        ApplyParagraphProperties(formatting, directParagraphProperties, _theme);
        ResolveToggles(formatting.RunFormatting);
        formatting.StyleKey = styleId ?? _defaultParagraphStyleId ?? "";
        return formatting;
    }

    /// <summary>
    /// A list paragraph's numbering level pPr (17.9.23: indents, tab stops) between the style and the direct pPr. Numbering
    /// that comes from the paragraph style leaves the style's own indents in charge (LibreOffice DomainMapper_Impl
    /// finishParagraph: "when numbering is defined by the paragraph style, then the para-style indents have priority");
    /// direct numbering (numPr in the paragraph) overrides them, and direct w:ind overrides both.
    /// </summary>
    private void ApplyNumberingLevel(WordParagraphFormatting formatting, string? styleId, XElement levelProperties, XElement? directProperties)
    {
        bool isFromStyle = directProperties?.Element(W + "numPr")?.Element(W + "numId") is null;
        int left = formatting.LeftIndentTwips, firstLine = formatting.FirstLineIndentTwips;
        ApplyParagraphProperties(formatting, levelProperties, _theme);
        if (!isFromStyle) return;
        foreach (XElement style in StyleChain(styleId ?? _defaultParagraphStyleId))
        {
            if (style.Element(W + "pPr")?.Element(W + "ind") is not { } ind) continue;
            if (ind.Attribute(W + "start") is not null || ind.Attribute(W + "left") is not null) formatting.LeftIndentTwips = left;
            if (ind.Attribute(W + "hanging") is not null || ind.Attribute(W + "firstLine") is not null) formatting.FirstLineIndentTwips = firstLine;
        }
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

    /// <summary>
    /// A numbering level's w:rPr as CSS for its label (only what the level sets). A bullet whose symbol-font character was
    /// mapped to a Unicode bullet drops the symbol font, which would draw another glyph for it.
    /// </summary>
    public string LevelRunCss(XElement? levelRunProperties, bool isBulletMapped)
    {
        if (levelRunProperties is null) return "";
        var empty = new WordRunFormatting();
        WordRunFormatting formatting = empty.Clone();
        ApplyRunProperties(formatting, levelRunProperties, _theme, WordStyleLayer.Direct);
        ResolveToggles(formatting);
        if (isBulletMapped) { formatting.FontName = null; formatting.ComplexScriptFontName = null; formatting.EastAsianFontName = null; }
        return formatting.ToCss(empty);
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

    /// <summary>A style and its basedOn ancestors, base first — cached per style (asked for every run and paragraph).</summary>
    private List<XElement> StyleChain(string? styleId)
    {
        if (styleId is null) return [];
        if (_styleChains.TryGetValue(styleId, out List<XElement>? cached)) return cached;
        List<XElement> built = BuildStyleChain(styleId);
        _styleChains[styleId] = built;
        return built;
    }

    private List<XElement> BuildStyleChain(string? styleId)
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
            // beforeAutospacing / afterAutospacing win over before / after in the same element; an explicit before / after in a
            // later layer overrides an inherited auto value (LibreOffice DomainMapper / DomainMapper_Impl: "apply INHERITED
            // autospacing only if top margin is not set").
            if (TwipsAttribute(spacing, "before") is { } before) { formatting.SpaceBeforeTwips = before; formatting.IsSpaceBeforeAuto = false; }
            if (TwipsAttribute(spacing, "after") is { } after) { formatting.SpaceAfterTwips = after; formatting.IsSpaceAfterAuto = false; }
            if (spacing.Attribute(W + "beforeAutospacing") is { } beforeAuto) formatting.IsSpaceBeforeAuto = IsOn(beforeAuto);
            if (spacing.Attribute(W + "afterAutospacing") is { } afterAuto) formatting.IsSpaceAfterAuto = IsOn(afterAuto);
            if (TwipsAttribute(spacing, "line") is { } line)
            {
                formatting.LineSpacingTwipsOrAuto = line;
                string lineRule = (string?)spacing.Attribute(W + "lineRule") ?? "auto";
                formatting.IsLineSpacingAuto = lineRule == "auto";
                formatting.IsLineSpacingExact = lineRule == "exact";   // ST_LineSpacingRule: auto, exact, atLeast (17.18.48)
            }
        }
        if (TwipsAttribute(pPr.Element(W + "outlineLvl"), "val") is { } outline) formatting.OutlineLevel = outline;
        if (pPr.Element(W + "pBdr") is { } borders)
        {
            formatting.BorderTopCss = DocxTableFormatting.BorderCss(borders.Element(W + "top"), theme) ?? formatting.BorderTopCss;
            formatting.BorderBottomCss = DocxTableFormatting.BorderCss(borders.Element(W + "bottom"), theme) ?? formatting.BorderBottomCss;
            formatting.BorderStartCss = DocxTableFormatting.BorderCss(borders.Element(W + "start") ?? borders.Element(W + "left"), theme) ?? formatting.BorderStartCss;
            formatting.BorderEndCss = DocxTableFormatting.BorderCss(borders.Element(W + "end") ?? borders.Element(W + "right"), theme) ?? formatting.BorderEndCss;
            formatting.BorderBetweenCss = DocxTableFormatting.BorderCss(borders.Element(W + "between"), theme) ?? formatting.BorderBetweenCss;
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
            // Stops merge across layers: a later layer adds or replaces the stop at a position, "clear" removes it (§17.3.1.37).
            var stops = new List<WordTabStop>(formatting.TabStops);
            foreach (XElement tab in tabs.Elements(W + "tab"))
            {
                if (TwipsAttribute(tab, "pos") is not { } position) continue;
                string alignment = (string?)tab.Attribute(W + "val") ?? "left";
                stops.RemoveAll(existing => existing.PositionTwips == position);
                if (alignment != "clear") stops.Add(new WordTabStop(position, alignment, (string?)tab.Attribute(W + "leader") ?? "none"));
            }
            stops.Sort((a, b) => a.PositionTwips.CompareTo(b.PositionTwips));
            formatting.TabStops = stops.ToArray();
        }
        if (pPr.Element(W + "numPr") is { } numPr)
        {
            formatting.NumberingId = TwipsAttribute(numPr.Element(W + "numId"), "val") ?? -1;
            formatting.NumberingLevel = TwipsAttribute(numPr.Element(W + "ilvl"), "val") ?? 0;
        }
        // pPr/rPr formats the paragraph mark only (§17.3.1.29), not the paragraph's runs: the renderer uses it for the list
        // label and the height of an empty paragraph (DocxBlockHtmlRenderer.RenderParagraph).
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
                    formatting.ShadingHex = FillHex(property, theme);
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

    /// <summary>
    /// The color a w:shd paints (§17.3.5 / ST_Shd §17.18.78), as LibreOffice's CellColorHandler computes it: the pattern's
    /// share of the pattern color (w:color / themeColor; auto = black) mixed over the fill (w:fill / themeFill; auto = white):
    /// clear = fill only, solid = color only, pctN = N %, pct12 / pct37 / pct62 / pct87 = 12.5 / 37.5 / 62.5 / 87.5 %, stripes
    /// and crosses ≈ 33.3 %; per channel (color·p + fill·(1000 − p)) / 1000. nil = no shading. A result of white is no background.
    /// </summary>
    public static string? FillHex(XElement shading, DocxTheme? theme)
    {
        string pattern = (string?)shading.Attribute(W + "val") ?? "clear";
        if (pattern == "nil") return null;
        string? fill = DocxTheme.ResolveColorHex(shading, "fill", theme, "themeFill", "themeFillTint", "themeFillShade");
        int share = PatternShare(pattern);
        string? result;
        if (share == 0) result = fill;
        else
        {
            int foreground = HexToRgb(DocxTheme.ResolveColorHex(shading, "color", theme) ?? "000000");
            int background = HexToRgb(fill ?? "FFFFFF");
            int Mix(int shift) => (((foreground >> shift) & 0xFF) * share + ((background >> shift) & 0xFF) * (1000 - share)) / 1000;
            result = (Mix(16) << 16 | Mix(8) << 8 | Mix(0)).ToString("X6", CultureInfo.InvariantCulture);
        }
        return result is not null && !result.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase) ? result : null;
    }

    /// <summary>ST_Shd → the pattern color's share in thousandths (LibreOffice CellColorHandler::getProperties nWW8BrushStyle).</summary>
    private static int PatternShare(string pattern) => pattern switch
    {
        "clear" => 0,
        "solid" => 1000,
        "pct12" => 125, "pct37" => 375, "pct62" => 625, "pct87" => 875,
        "horzStripe" or "vertStripe" or "reverseDiagStripe" or "diagStripe" or "horzCross" or "diagCross"
            or "thinHorzStripe" or "thinVertStripe" or "thinReverseDiagStripe" or "thinDiagStripe" or "thinHorzCross" or "thinDiagCross" => 333,
        _ when pattern.StartsWith("pct", StringComparison.Ordinal) && int.TryParse(pattern.Substring(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out int percent) => percent * 10,
        _ => 0,
    };

    private static int HexToRgb(string hex) => int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb) ? rgb : 0;

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
