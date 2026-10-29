using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// The OpenDocument style resolver: style:default-style per family → the style:parent-style-name
/// chain → the style itself, over the named styles of styles.xml (or a .fodt's office:styles) and
/// the automatic styles of content.xml (captured during the index pass, never read again from the
/// body). Property names follow LibreOffice's import maps (xmloff/source/text/txtprmap.cxx for
/// text and paragraph properties, sw/source/filter/xml/xmlitemm.cxx for table, column and cell
/// properties). Also holds the list styles (text:list-style) and text:notes-configuration.
/// </summary>
public sealed class OdtStyleCascade
{
    public static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    public static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    public static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    public static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    public static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    public static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    public static readonly XNamespace Xml = XNamespace.Xml;

    /// <summary>Attributes of style:text-properties this cascade maps (the audit counts the rest).</summary>
    public static readonly HashSet<string> HandledTextProperties =
    [
        "fo:font-weight", "style:font-weight-complex", "fo:font-style", "style:font-style-complex",
        "style:text-underline-style", "style:text-underline-type", "style:text-underline-width", "style:text-underline-color", "style:text-underline-mode",
        "style:text-line-through-style", "style:text-line-through-type", "fo:font-size", "style:font-size-complex",
        "style:font-name", "style:font-name-complex", "fo:font-family", "style:font-family-complex", "fo:color", "fo:background-color",
        "style:text-position", "fo:text-transform", "fo:font-variant", "fo:letter-spacing", "text:display", "style:text-outline",
        "fo:text-shadow", "style:font-relief", "fo:border",
    ];

    /// <summary>Attributes of style:paragraph-properties this cascade maps.</summary>
    public static readonly HashSet<string> HandledParagraphProperties =
    [
        "fo:text-align", "style:writing-mode", "fo:margin", "fo:margin-left", "fo:margin-right", "fo:margin-top", "fo:margin-bottom",
        "fo:text-indent", "fo:line-height", "style:line-height-at-least", "fo:background-color", "fo:border", "fo:border-top",
        "fo:border-bottom", "fo:border-left", "fo:border-right", "fo:padding", "fo:break-before",
    ];

    /// <summary>Attributes of table, column and cell properties the table renderer maps.</summary>
    public static readonly HashSet<string> HandledTableProperties =
    [
        "style:width", "style:rel-width", "table:align", "style:writing-mode", "fo:margin-left", "fo:margin-right", "fo:background-color",
        "style:column-width", "style:rel-column-width", "fo:border", "fo:border-top", "fo:border-bottom", "fo:border-left", "fo:border-right",
        "fo:padding", "fo:padding-top", "fo:padding-bottom", "fo:padding-left", "fo:padding-right", "style:vertical-align",
    ];

    private readonly Dictionary<(string family, string name), XElement> _styles = new();
    private readonly Dictionary<string, XElement> _defaultStyles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, OdtListStyle> _listStyles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _fontFamilies = new(StringComparer.Ordinal);
    private readonly Dictionary<string, OdtParagraphStyle> _resolvedParagraphs = new(StringComparer.Ordinal);
    private readonly Dictionary<(string? paragraph, string text), WordRunFormatting> _resolvedSpans = new();
    private readonly OdtNotesConfiguration[] _notesConfigurations = [new(), new()];
    private readonly List<XElement> _sections = new();

    /// <summary>A resolved paragraph style: formatting, page break before, the list style it names.</summary>
    public sealed record OdtParagraphStyle(WordParagraphFormatting Formatting, bool IsPageBreakBefore, string? ListStyleName);

    /// <summary>text:notes-configuration of one note class (footnote / endnote).</summary>
    public sealed class OdtNotesConfiguration
    {
        public string NumberFormat = "1";
        public string Prefix = "";
        public string Suffix = "";
        public int StartOffset;
    }

    public OdtNotesConfiguration NotesConfiguration(WordDocumentNoteKind kind) => _notesConfigurations[(int)kind];

    /// <summary>Adds an office:font-face-decls, office:styles or office:automatic-styles section (later sections win on name clashes).</summary>
    public void AddSection(XElement section)
    {
        _sections.Add(section);
        _resolvedParagraphs.Clear();
        _resolvedSpans.Clear();
        foreach (XElement element in section.Elements())
        {
            string local = element.Name.LocalName;
            if (element.Name == Style + "font-face")
            {
                string? name = (string?)element.Attribute(Style + "name");
                string? family = (string?)element.Attribute(Svg + "font-family");
                if (name is not null && family is not null) _fontFamilies[name] = family.Trim('\'', '"', ' ');
            }
            else if (element.Name == Style + "style")
            {
                string? family = (string?)element.Attribute(Style + "family");
                string? name = (string?)element.Attribute(Style + "name");
                if (family is not null && name is not null) _styles[(family, name)] = element;
            }
            else if (element.Name == Style + "default-style")
            {
                if ((string?)element.Attribute(Style + "family") is { } family) _defaultStyles[family] = element;
            }
            else if (element.Name == Text + "list-style")
            {
                if ((string?)element.Attribute(Style + "name") is { } name) _listStyles[name] = new OdtListStyle(element);
            }
            else if (element.Name == Text + "notes-configuration")
            {
                var configuration = new OdtNotesConfiguration
                {
                    NumberFormat = (string?)element.Attribute(Style + "num-format") ?? "1",
                    Prefix = (string?)element.Attribute(Style + "num-prefix") ?? "",
                    Suffix = (string?)element.Attribute(Style + "num-suffix") ?? "",
                };
                if (int.TryParse((string?)element.Attribute(Text + "start-value"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int offset)) configuration.StartOffset = offset;
                _notesConfigurations[(string?)element.Attribute(Text + "note-class") == "endnote" ? 1 : 0] = configuration;
            }
            else if (local == "outline-style") { /* heading numbering: not rendered (see NOTES.md) */ }
        }
    }

    /// <summary>
    /// The sections as stored in the index: only what rendering reads — styles with their name,
    /// family, parent and the mapped property attributes, font faces with their family. List styles,
    /// notes configuration (labels and numbers are in the index), tab stops, language tags and every
    /// unmapped attribute are dropped: a LibreOffice conversion of a Word file carries 14 MB of list
    /// styles and MBs of unmapped attributes that would otherwise be parsed on every reopen.
    /// </summary>
    public string[] SectionsForRendering()
        => _sections.Select(section =>
        {
            var copy = new XElement(section.Name);
            foreach (XElement element in section.Elements())
            {
                if (element.Name == Style + "font-face")
                    copy.Add(new XElement(element.Name, element.Attribute(Style + "name"), element.Attribute(Svg + "font-family")));
                else if (element.Name == Style + "style" || element.Name == Style + "default-style")
                {
                    var style = new XElement(element.Name, element.Attribute(Style + "name"), element.Attribute(Style + "family"),
                                             element.Attribute(Style + "parent-style-name"), element.Attribute(Style + "list-style-name"));
                    foreach (XElement properties in element.Elements())
                    {
                        HashSet<string>? handled = properties.Name.LocalName switch
                        {
                            "text-properties" => HandledTextProperties,
                            "paragraph-properties" => HandledParagraphProperties,
                            "table-properties" or "table-cell-properties" or "table-column-properties" => HandledTableProperties,
                            _ => null,
                        };
                        if (handled is null) continue;
                        var kept = new XElement(properties.Name, properties.Attributes().Where(attribute => handled.Contains(QualifiedName(attribute.Name))));
                        if (kept.HasAttributes) style.Add(kept);
                    }
                    copy.Add(style);
                }
            }
            return copy.ToString(SaveOptions.DisableFormatting | SaveOptions.OmitDuplicateNamespaces);
        }).ToArray();

    private static string QualifiedName(XName name)
    {
        string prefix = name.Namespace == Fo ? "fo" : name.Namespace == Style ? "style" : name.Namespace == Text ? "text"
                      : name.Namespace == Table ? "table" : name.Namespace == Svg ? "svg" : "?";
        return prefix + ":" + name.LocalName;
    }

    public OdtListStyle? ListStyle(string? name) => name is not null && _listStyles.TryGetValue(name, out OdtListStyle? style) ? style : null;

    public XElement? StyleElement(string family, string? name)
        => name is not null && _styles.TryGetValue((family, name), out XElement? style) ? style : null;

    /// <summary>The default paragraph formatting (style:default-style family paragraph).</summary>
    public OdtParagraphStyle DefaultParagraph => ResolveParagraph(null);

    public bool IsDocumentRightToLeft => DefaultParagraph.Formatting.IsRightToLeft || ResolveParagraph("Standard").Formatting.IsRightToLeft;

    /// <summary>Resolves a paragraph style (named or automatic) through its parent chain; cached.</summary>
    public OdtParagraphStyle ResolveParagraph(string? styleName)
    {
        string key = styleName ?? "";
        if (_resolvedParagraphs.TryGetValue(key, out OdtParagraphStyle? cached)) return cached;

        List<XElement> chain = Chain("paragraph", styleName);
        var raw = new RawParagraph();
        string? listStyleName = null;
        foreach (XElement style in chain)
        {
            ReadParagraphProperties(raw, style.Element(Style + "paragraph-properties"));
            if ((string?)style.Attribute(Style + "list-style-name") is { } list) listStyleName = list.Length == 0 ? null : list;
        }
        WordParagraphFormatting formatting = raw.ToFormatting();
        foreach (XElement style in chain) ApplyTextProperties(formatting.RunFormatting, style.Element(Style + "text-properties"), formatting.IsRightToLeft);
        formatting.RunFormatting.IsRightToLeft = formatting.IsRightToLeft;
        formatting.StyleKey = styleName;
        var resolved = new OdtParagraphStyle(formatting, raw.IsPageBreakBefore, listStyleName);
        _resolvedParagraphs[key] = resolved;
        return resolved;
    }

    /// <summary>A text:span's formatting over its paragraph's base (family "text" chain); cached per pair.</summary>
    public WordRunFormatting ResolveSpan(string? paragraphStyleName, WordRunFormatting paragraphBase, string textStyleName)
    {
        if (_resolvedSpans.TryGetValue((paragraphStyleName, textStyleName), out WordRunFormatting? cached)) return cached;
        WordRunFormatting run = ApplySpan(paragraphBase, textStyleName);
        _resolvedSpans[(paragraphStyleName, textStyleName)] = run;
        return run;
    }

    /// <summary>A text style's chain applied over <paramref name="outer"/> (a copy is returned).</summary>
    public WordRunFormatting ApplySpan(WordRunFormatting outer, string textStyleName)
    {
        WordRunFormatting run = outer.Clone();
        foreach (XElement style in Chain("text", textStyleName, includeDefault: false))
            ApplyTextProperties(run, style.Element(Style + "text-properties"), outer.IsRightToLeft);
        return run;
    }

    /// <summary>Page widths per master page (index time only; a reopen reads the sections from the index).</summary>
    public OdtPageSetup PageSetup { get; set; } = OdtPageSetup.Empty;

    /// <summary>The master page a paragraph / table style starts (style:master-page-name, through its parents), or null.</summary>
    public string? MasterPageNameOf(string family, string? styleName)
    {
        if (styleName is null) return null;
        if (_masterPageNames.TryGetValue((family, styleName), out string? cached)) return cached;
        string? found = null;
        List<XElement> chain = Chain(family, styleName, includeDefault: false);
        for (int i = chain.Count - 1; i >= 0 && found is null; i--)
            if ((string?)chain[i].Attribute(Style + "master-page-name") is { Length: > 0 } master) found = master;
        _masterPageNames[(family, styleName)] = found;
        return found;
    }
    private readonly Dictionary<(string family, string name), string?> _masterPageNames = new();

    /// <summary>default-style → root parent … → the style itself.</summary>
    private List<XElement> Chain(string family, string? name, bool includeDefault = true)
    {
        var chain = new List<XElement>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (string? current = name; current is not null && seen.Add(current);)
        {
            if (!_styles.TryGetValue((family, current), out XElement? style)) break;
            chain.Add(style);
            current = (string?)style.Attribute(Style + "parent-style-name");
        }
        if (includeDefault && _defaultStyles.TryGetValue(family, out XElement? defaultStyle)) chain.Add(defaultStyle);
        chain.Reverse();
        return chain;
    }

    // ── paragraph properties ─────────────────────────────────────────────────────────

    /// <summary>Physical sides as written (left/right), turned into start/end once the direction is known.</summary>
    private sealed class RawParagraph
    {
        public string? Align;
        public bool IsRightToLeft;
        public int? MarginLeft, MarginRight, MarginTop, MarginBottom, TextIndent;
        public int LineHeight; public bool IsLineHeightAuto = true;
        public string? Background;
        public string? BorderTop, BorderBottom, BorderLeft, BorderRight;
        public bool IsPageBreakBefore;

        public WordParagraphFormatting ToFormatting()
        {
            var formatting = new WordParagraphFormatting { IsRightToLeft = IsRightToLeft };
            formatting.Alignment = Align switch
            {
                "center" => WordParagraphAlignment.Center,
                "justify" => WordParagraphAlignment.Justify,
                "start" => WordParagraphAlignment.Start,
                "end" => WordParagraphAlignment.End,
                "left" => IsRightToLeft ? WordParagraphAlignment.End : WordParagraphAlignment.Start,
                "right" => IsRightToLeft ? WordParagraphAlignment.Start : WordParagraphAlignment.End,
                _ => null,
            };
            int left = MarginLeft ?? 0, right = MarginRight ?? 0;
            formatting.LeftIndentTwips = IsRightToLeft ? right : left;      // WordParagraphFormatting "left" = start side
            formatting.RightIndentTwips = IsRightToLeft ? left : right;
            formatting.FirstLineIndentTwips = TextIndent ?? 0;
            formatting.SpaceBeforeTwips = MarginTop ?? 0;
            formatting.SpaceAfterTwips = MarginBottom ?? 0;
            formatting.LineSpacingTwipsOrAuto = LineHeight;
            formatting.IsLineSpacingAuto = IsLineHeightAuto;
            formatting.ShadingHex = Background;
            formatting.BorderTopCss = BorderTop;
            formatting.BorderBottomCss = BorderBottom;
            formatting.BorderStartCss = IsRightToLeft ? BorderRight : BorderLeft;
            formatting.BorderEndCss = IsRightToLeft ? BorderLeft : BorderRight;
            return formatting;
        }
    }

    private static void ReadParagraphProperties(RawParagraph raw, XElement? properties)
    {
        if (properties is null) return;
        if ((string?)properties.Attribute(Fo + "text-align") is { } align) raw.Align = align;
        switch ((string?)properties.Attribute(Style + "writing-mode"))
        {
            case "rl-tb" or "rl" or "tb-rl": raw.IsRightToLeft = true; break;
            case "lr-tb" or "lr" or "tb-lr": raw.IsRightToLeft = false; break;
        }
        if (Twips(properties.Attribute(Fo + "margin")) is { } all) raw.MarginLeft = raw.MarginRight = raw.MarginTop = raw.MarginBottom = all;
        raw.MarginLeft = Twips(properties.Attribute(Fo + "margin-left")) ?? raw.MarginLeft;
        raw.MarginRight = Twips(properties.Attribute(Fo + "margin-right")) ?? raw.MarginRight;
        raw.MarginTop = Twips(properties.Attribute(Fo + "margin-top")) ?? raw.MarginTop;
        raw.MarginBottom = Twips(properties.Attribute(Fo + "margin-bottom")) ?? raw.MarginBottom;
        raw.TextIndent = Twips(properties.Attribute(Fo + "text-indent")) ?? raw.TextIndent;
        if ((string?)properties.Attribute(Fo + "line-height") is { } lineHeight)
        {
            if (lineHeight == "normal") { raw.LineHeight = 0; raw.IsLineHeightAuto = true; }
            else if (Percent(lineHeight) is { } percent) { raw.LineHeight = (int)Math.Round(240 * percent / 100); raw.IsLineHeightAuto = true; }
            else if (Twips(lineHeight) is { } fixedHeight) { raw.LineHeight = fixedHeight; raw.IsLineHeightAuto = false; }
        }
        if (Twips(properties.Attribute(Style + "line-height-at-least")) is { } atLeast && atLeast > 0) { raw.LineHeight = atLeast; raw.IsLineHeightAuto = false; }
        if ((string?)properties.Attribute(Fo + "background-color") is { } background) raw.Background = ColorHex(background);
        if (BorderCss(properties.Attribute(Fo + "border")) is { } border) raw.BorderTop = raw.BorderBottom = raw.BorderLeft = raw.BorderRight = border;
        raw.BorderTop = BorderCss(properties.Attribute(Fo + "border-top")) ?? raw.BorderTop;
        raw.BorderBottom = BorderCss(properties.Attribute(Fo + "border-bottom")) ?? raw.BorderBottom;
        raw.BorderLeft = BorderCss(properties.Attribute(Fo + "border-left")) ?? raw.BorderLeft;
        raw.BorderRight = BorderCss(properties.Attribute(Fo + "border-right")) ?? raw.BorderRight;
        if ((string?)properties.Attribute(Fo + "break-before") is { } breakBefore) raw.IsPageBreakBefore = breakBefore == "page";
    }

    // ── text properties ───────────────────────────────────────────────────────────────

    /// <summary>
    /// style:text-properties onto a run. Western and complex-script variants: the complex one
    /// (style:*-complex) wins in a right-to-left paragraph, the western one otherwise — the
    /// renderer has no per-character script detection.
    /// </summary>
    public void ApplyTextProperties(WordRunFormatting run, XElement? properties, bool isRightToLeft)
    {
        if (properties is null) return;
        string? Pick(XName western, XName complex)
            => isRightToLeft ? (string?)properties.Attribute(complex) ?? (string?)properties.Attribute(western)
                             : (string?)properties.Attribute(western);

        if (Pick(Fo + "font-weight", Style + "font-weight-complex") is { } weight)
            run.IsBold = weight == "bold" || (int.TryParse(weight, out int numeric) && numeric >= 600);
        if (Pick(Fo + "font-style", Style + "font-style-complex") is { } posture) run.IsItalic = posture is "italic" or "oblique";

        if ((string?)properties.Attribute(Style + "text-underline-style") is { } underline)
        {
            run.IsUnderlined = underline != "none";
            run.UnderlineStyleCss = underline switch { "dotted" => "dotted", "dash" or "long-dash" or "dot-dash" or "dot-dot-dash" => "dashed", "wave" => "wavy", _ => null };
        }
        if ((string?)properties.Attribute(Style + "text-underline-type") == "double" && run.IsUnderlined) run.UnderlineStyleCss = "double";
        if ((string?)properties.Attribute(Style + "text-underline-width") is { } underlineWidth) run.IsUnderlineThick = underlineWidth is "bold" or "thick";
        if ((string?)properties.Attribute(Style + "text-underline-color") is { } underlineColor) run.UnderlineColorHex = underlineColor == "font-color" ? null : ColorHex(underlineColor);
        if ((string?)properties.Attribute(Style + "text-line-through-style") is { } lineThrough) run.IsStrikethrough = lineThrough != "none";

        if ((string?)properties.Attribute(Fo + "font-size") is { } size) run.FontSizeHalfPoints = HalfPoints(size, run.FontSizeHalfPoints);
        if ((string?)properties.Attribute(Style + "font-size-complex") is { } complexSize)
            run.ComplexScriptFontSizeHalfPoints = HalfPoints(complexSize, run.ComplexScriptFontSizeHalfPoints > 0 ? run.ComplexScriptFontSizeHalfPoints : run.FontSizeHalfPoints);
        if (FontName(properties, Style + "font-name", Fo + "font-family") is { } font) run.FontName = font;
        if (FontName(properties, Style + "font-name-complex", Style + "font-family-complex") is { } complexFont) run.ComplexScriptFontName = complexFont;

        if ((string?)properties.Attribute(Fo + "color") is { } color) run.ColorHex = ColorHex(color);
        if ((string?)properties.Attribute(Fo + "background-color") is { } background)
            run.HighlightCssColor = ColorHex(background) is { } hex ? "#" + hex : null;
        if ((string?)properties.Attribute(Style + "text-position") is { } position) ApplyTextPosition(run, position);
        if ((string?)properties.Attribute(Fo + "text-transform") is { } transform) run.IsAllCaps = transform == "uppercase";
        if ((string?)properties.Attribute(Fo + "font-variant") is { } variant) run.IsSmallCaps = variant == "small-caps";
        if ((string?)properties.Attribute(Fo + "letter-spacing") is { } spacing) run.CharacterSpacingTwips = spacing == "normal" ? 0 : Twips(spacing) ?? 0;
        if ((string?)properties.Attribute(Text + "display") is { } display) run.IsHidden = display == "none";
        if ((string?)properties.Attribute(Style + "text-outline") is { } outline) run.IsOutline = outline == "true";
        if ((string?)properties.Attribute(Fo + "text-shadow") is { } shadow) run.IsShadow = shadow != "none";
        if ((string?)properties.Attribute(Style + "font-relief") is { } relief) { run.IsEmboss = relief == "embossed"; run.IsImprint = relief == "engraved"; }
        if (BorderCss(properties.Attribute(Fo + "border")) is { } border) run.BorderCss = border;
    }

    /// <summary>"super 58%", "sub 58%", "33% 58%", "-33% 100%", "0% 100%" (ODF 1.3 part 3 §20.374).</summary>
    private static void ApplyTextPosition(WordRunFormatting run, string value)
    {
        string first = value.Split(' ')[0];
        run.IsSuperscript = first == "super" || (Percent(first) is { } up && up > 0);
        run.IsSubscript = first == "sub" || (Percent(first) is { } down && down < 0);
    }

    private string? FontName(XElement properties, XName nameAttribute, XName familyAttribute)
    {
        if ((string?)properties.Attribute(nameAttribute) is { } name)
            return _fontFamilies.TryGetValue(name, out string? family) ? family : name;
        return ((string?)properties.Attribute(familyAttribute))?.Trim('\'', '"', ' ');
    }

    // ── tables ────────────────────────────────────────────────────────────────────────

    /// <summary>CSS for a table:table-cell's style: borders, padding, background, vertical alignment.</summary>
    public string CellCss(string? cellStyleName)
    {
        XElement? properties = StyleElement("table-cell", cellStyleName)?.Element(Style + "table-cell-properties");
        if (properties is null) return "";
        var css = new StringBuilder();
        if (BorderCss(properties.Attribute(Fo + "border")) is { } border) css.Append("border:").Append(border).Append(';');
        foreach ((string side, string cssSide) in new[] { ("top", "top"), ("bottom", "bottom"), ("left", "left"), ("right", "right") })
            if (BorderCss(properties.Attribute(Fo + "border-" + side)) is { } sideBorder) css.Append("border-").Append(cssSide).Append(':').Append(sideBorder).Append(';');
        if (Twips(properties.Attribute(Fo + "padding")) is { } padding) css.Append("padding:").Append(WordParagraphFormatting.TwipsToPt(padding)).Append("pt;");
        foreach (string side in new[] { "top", "bottom", "left", "right" })
            if (Twips(properties.Attribute(Fo + "padding-" + side)) is { } sidePadding) css.Append("padding-").Append(side).Append(':').Append(WordParagraphFormatting.TwipsToPt(sidePadding)).Append("pt;");
        if ((string?)properties.Attribute(Fo + "background-color") is { } background && ColorHex(background) is { } hex) css.Append("background-color:#").Append(hex).Append(';');
        if ((string?)properties.Attribute(Style + "vertical-align") is "middle" or "bottom" or "top")
            css.Append("vertical-align:").Append((string?)properties.Attribute(Style + "vertical-align")).Append(';');
        return css.ToString();
    }

    /// <summary>A column's width in points (style:column-width), or null.</summary>
    public string? ColumnWidthCss(string? columnStyleName)
    {
        XElement? properties = StyleElement("table-column", columnStyleName)?.Element(Style + "table-column-properties");
        return Twips(properties?.Attribute(Style + "column-width")) is { } width ? WordParagraphFormatting.TwipsToPt(width) + "pt" : null;
    }

    /// <summary>A table's CSS (width, alignment, background) and direction (style:writing-mode).</summary>
    public (string css, bool? isRightToLeft) TableCss(string? tableStyleName)
    {
        XElement? properties = StyleElement("table", tableStyleName)?.Element(Style + "table-properties");
        if (properties is null) return ("", null);
        var css = new StringBuilder();
        if ((string?)properties.Attribute(Style + "rel-width") is { } relative && Percent(relative) is { } percent) css.Append("width:").Append(percent.ToString("0.#", CultureInfo.InvariantCulture)).Append("%;");
        else if (Twips(properties.Attribute(Style + "width")) is { } width) css.Append("width:").Append(WordParagraphFormatting.TwipsToPt(width)).Append("pt;");
        if ((string?)properties.Attribute(Table + "align") == "center") css.Append("margin-left:auto;margin-right:auto;");
        if ((string?)properties.Attribute(Fo + "background-color") is { } background && ColorHex(background) is { } hex) css.Append("background-color:#").Append(hex).Append(';');
        bool? rtl = (string?)properties.Attribute(Style + "writing-mode") switch { "rl-tb" or "rl" => true, "lr-tb" or "lr" => false, _ => null };
        return (css.ToString(), rtl);
    }

    // ── value parsing (ODF 1.3 part 3 §18: length = number + unit cm|mm|in|pt|pc|px) ──

    public static int? Twips(XAttribute? attribute) => attribute is null ? null : Twips(attribute.Value);

    public static int? Twips(string value)
    {
        value = value.Trim();
        int unitStart = value.Length;
        while (unitStart > 0 && char.IsLetter(value[unitStart - 1])) unitStart--;
        if (!double.TryParse(value.Substring(0, unitStart), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return null;
        double points = value.Substring(unitStart) switch
        {
            "pt" => number, "in" => number * 72, "cm" => number * 72 / 2.54, "mm" => number * 72 / 25.4,
            "pc" => number * 12, "px" => number * 0.75, _ => double.NaN,
        };
        return double.IsNaN(points) ? null : (int)Math.Round(points * 20);
    }

    private static double? Percent(string value)
        => value.EndsWith("%", StringComparison.Ordinal) && double.TryParse(value.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) ? percent : null;

    /// <summary>fo:font-size: a length or a percentage of the inherited size.</summary>
    private static int HalfPoints(string value, int inheritedHalfPoints)
    {
        if (Percent(value) is { } percent) return inheritedHalfPoints > 0 ? (int)Math.Round(inheritedHalfPoints * percent / 100) : inheritedHalfPoints;
        return Twips(value) is { } twips ? (int)Math.Round(twips / 10.0) : inheritedHalfPoints;
    }

    /// <summary>"#rrggbb" → "RRGGBB"; "transparent" → null.</summary>
    public static string? ColorHex(string value)
        => value.Length == 7 && value[0] == '#' ? value.Substring(1).ToUpperInvariant() : null;

    /// <summary>fo:border "0.06pt solid #000000" is already CSS; "none" → "none".</summary>
    private static string? BorderCss(XAttribute? attribute)
    {
        if (attribute is null) return null;
        string value = attribute.Value.Trim();
        if (value == "none") return "none";
        // CSS has no "0.06pt" hairline problem, but very thin lines vanish; keep at least 0.5pt.
        string[] parts = value.Split(' ');
        if (parts.Length > 0 && Twips(parts[0]) is { } width && width < 10) parts[0] = "0.5pt";
        return string.Join(" ", parts);
    }
}
