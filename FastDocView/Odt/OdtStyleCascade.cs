using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.DocxHtml;
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
    public static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    public static readonly XNamespace LoExt = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
    public static readonly XNamespace Xml = XNamespace.Xml;
    public static readonly XNamespace DataStyle = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

    /// <summary>Attributes of style:text-properties this cascade maps (the audit counts the rest).</summary>
    public static readonly HashSet<string> HandledTextProperties =
    [
        "fo:font-weight", "style:font-weight-complex", "fo:font-style", "style:font-style-complex",
        "style:text-underline-style", "style:text-underline-type", "style:text-underline-width", "style:text-underline-color", "style:text-underline-mode",
        "style:text-line-through-style", "style:text-line-through-type", "fo:font-size", "style:font-size-complex",
        "style:font-name", "style:font-name-complex", "fo:font-family", "style:font-family-complex", "fo:color", "fo:background-color",
        "style:text-position", "fo:text-transform", "fo:font-variant", "fo:letter-spacing", "text:display", "style:text-outline",
        "fo:text-shadow", "style:font-relief", "fo:border", "style:text-emphasize", "style:text-scale", "fo:hyphenate", "fo:language", "fo:country",
        // the three languages a list number follows (LanguagesOf)
        "fo:script", "style:rfc-language-tag", "style:language-asian", "style:script-asian", "style:country-asian", "style:rfc-language-tag-asian",
        "style:language-complex", "style:script-complex", "style:country-complex", "style:rfc-language-tag-complex",
    ];

    /// <summary>Attributes of style:paragraph-properties this cascade maps.</summary>
    public static readonly HashSet<string> HandledParagraphProperties =
    [
        "fo:text-align", "style:writing-mode", "fo:margin", "fo:margin-left", "fo:margin-right", "fo:margin-top", "fo:margin-bottom",
        "fo:text-indent", "fo:line-height", "style:line-height-at-least", "fo:background-color", "fo:border", "fo:border-top",
        "fo:border-bottom", "fo:border-left", "fo:border-right", "fo:padding", "fo:break-before", "style:tab-stop-distance",
    ];

    /// <summary>Attributes of table, column and cell properties the table renderer maps.</summary>
    public static readonly HashSet<string> HandledTableProperties =
    [
        "style:width", "style:rel-width", "table:align", "style:writing-mode", "fo:margin-left", "fo:margin-right", "fo:background-color",
        "style:column-width", "style:rel-column-width", "fo:border", "fo:border-top", "fo:border-bottom", "fo:border-left", "fo:border-right",
        "fo:padding", "fo:padding-top", "fo:padding-bottom", "fo:padding-left", "fo:padding-right", "style:vertical-align", "style:row-height", "style:min-row-height",
        "loext:writing-mode",
    ];

    /// <summary>Attributes of style:graphic-properties (frames: pictures, text boxes) the frame renderer maps.</summary>
    public static readonly HashSet<string> HandledGraphicProperties =
    [
        "style:wrap", "style:horizontal-pos", "style:horizontal-rel", "fo:margin", "fo:margin-left", "fo:margin-right", "fo:margin-top",
        "fo:margin-bottom", "text:anchor-type", "style:vertical-pos", "style:vertical-rel", "fo:clip",
        // text boxes and shapes (ODF 1.3 part 3 20.118 draw:fill, 20.119 draw:fill-color, 20.120 draw:fill-gradient-name,
        // 20.167 draw:stroke, 20.408 svg:stroke-color, 20.410 svg:stroke-width; fo:padding / fo:border / fo:background-color)
        "draw:fill", "draw:fill-color", "draw:fill-gradient-name", "draw:stroke", "svg:stroke-color", "svg:stroke-width",
        "fo:padding", "fo:padding-left", "fo:padding-right", "fo:padding-top", "fo:padding-bottom", "fo:border", "fo:border-left",
        "fo:border-right", "fo:border-top", "fo:border-bottom", "fo:background-color", "fo:min-height", "fo:min-width",
    ];

    private readonly Dictionary<(string family, string name), XElement> _styles = new();
    /// <summary>style:style elements cut from the bytes as the part streamed by (<see cref="OdtStylesListStyleFilterStream"/>), parsed
    /// when first looked up: a LibreOffice conversion of a Word file declares thousands (the sample: 5,716 named, 2,702 automatic),
    /// a window uses a few. Keys never overlap <see cref="_styles"/>: adding to one removes the key from the other (later wins).</summary>
    private readonly Dictionary<(string family, string name), UnparsedStyle> _unparsedStyles = new();
    /// <summary>The keys of the styles that came cut (parsed since or not): they are in no section, so the index's copy lists them itself.</summary>
    private readonly Dictionary<(string family, string name), long> _cutStyleKeys = new();   // key -> the order it was cut in (the copy is deterministic)
    private long _cutStyleCount;
    /// <summary>A first open's rendering cascade reads its styles from the indexing cascade, each filtered as it is first used
    /// (<see cref="ForRendering"/>); null for every other cascade.</summary>
    private OdtStyleCascade? _renderingSource;

    /// <summary>One style:style element's bytes, with the root start tag (namespace declarations) it is parsed inside.</summary>
    internal sealed record UnparsedStyle(byte[] Element, byte[] RootOpenTag, string RootElementName);
    private readonly Dictionary<string, XElement> _defaultStyles = new(StringComparer.Ordinal);
    /// <summary>The named drawing styles of office:styles (ODF 1.3 part 3 16.40-16.44 / 16.46: draw:gradient, svg:linearGradient,
    /// svg:radialGradient, draw:hatch, draw:fill-image, draw:opacity, draw:marker, draw:stroke-dash) by (element local name, draw:name).</summary>
    private readonly Dictionary<(string kind, string name), XElement> _drawingStyles = new();
    private readonly Dictionary<string, OdtListStyle> _listStyles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _fontFamilies = new(StringComparer.Ordinal);
    private readonly Dictionary<string, OdtParagraphStyle> _resolvedParagraphs = new(StringComparer.Ordinal);
    private readonly Dictionary<(string? paragraph, string text), WordRunFormatting> _resolvedSpans = new();
    private readonly OdtNotesConfiguration[] _notesConfigurations = [new(), new()];
    private readonly List<XElement> _sections = new();
    /// <summary>number:date-style / number:time-style by style:name (ODF 1.3 part 3 16.29.11 / 16.29.19): what a date or time field is shown in.</summary>
    private readonly Dictionary<string, XElement> _dataStyles = new(StringComparer.Ordinal);

    public XElement? DataStyleNamed(string? name) => name is not null && _dataStyles.TryGetValue(name, out XElement? style) ? style : null;

    private static bool IsDateTimeDataStyle(XName name) => name == DataStyle + "date-style" || name == DataStyle + "time-style";

    /// <summary>A resolved paragraph style: formatting, page break before, the list style it names.</summary>
    public sealed record OdtParagraphStyle(WordParagraphFormatting Formatting, bool IsPageBreakBefore, string? ListStyleName);

    /// <summary>text:notes-configuration of one note class (footnote / endnote).</summary>
    public sealed class OdtNotesConfiguration
    {
        public string NumberFormat = "1";
        public string Prefix = "";
        public string Suffix = "";
        public int StartOffset;
        /// <summary>Footnotes only: text:start-numbering-at="chapter" restarts the count at each outline-level-1 heading
        /// (LibreOffice xmloff XMLFootnoteConfigurationImportContext.cxx:199-201, 312-315: FootnoteCounting, footnotes only).
        /// "page" cannot be followed in a flow view; a note's citation text (what LibreOffice printed) is shown when there is one.</summary>
        public bool RestartsAtChapter;
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
                if (family is not null && name is not null)
                {
                    _styles[(family, name)] = element;
                    _unparsedStyles.Remove((family, name));
                    _cutStyleKeys.Remove((family, name));
                }
            }
            else if (element.Name == Style + "default-style")
            {
                if ((string?)element.Attribute(Style + "family") is { } family) _defaultStyles[family] = element;
            }
            else if (IsDrawingStyle(element.Name))
            {
                if ((string?)element.Attribute(Draw + "name") is { } name) _drawingStyles[(element.Name.LocalName, name)] = element;
            }
            else if (IsDateTimeDataStyle(element.Name))
            {
                if ((string?)element.Attribute(Style + "name") is { } name) _dataStyles[name] = element;
            }
            else if (element.Name == Text + "list-style")
            {
                if ((string?)element.Attribute(Style + "name") is { } name) _listStyles[name] = new OdtListStyle(element);
            }
            else if (element.Name == Text + "notes-configuration")
            {
                var configuration = new OdtNotesConfiguration
                {
                    NumberFormat = OdtListStyle.NumberFormatOf((string?)element.Attribute(Style + "num-format"), (string?)element.Attribute(Style + "num-letter-sync")),
                    Prefix = (string?)element.Attribute(Style + "num-prefix") ?? "",
                    Suffix = (string?)element.Attribute(Style + "num-suffix") ?? "",
                };
                if (int.TryParse((string?)element.Attribute(Text + "start-value"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int offset)) configuration.StartOffset = offset;
                bool isEndnote = (string?)element.Attribute(Text + "note-class") == "endnote";
                configuration.RestartsAtChapter = !isEndnote && (string?)element.Attribute(Text + "start-numbering-at") == "chapter";
                _notesConfigurations[isEndnote ? 1 : 0] = configuration;
            }
            else if (local == "outline-style") _listStyles[OdtListStyle.OutlineStyleName] = new OdtListStyle(element);   // heading numbering
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
        => SectionElementsForRendering().Select(section => section.ToString(SaveOptions.DisableFormatting | SaveOptions.OmitDuplicateNamespaces)).ToArray();

    /// <summary>
    /// <see cref="SectionsForRendering"/> as trees: a first open builds its rendering cascade from these directly
    /// (the strings are only for the saved index - the same content, without a serialise-and-parse round trip).
    /// </summary>
    public XElement[] SectionElementsForRendering()
    {
        var sections = _sections.Select(section => SectionForRendering(section, includeStyles: true)).ToList();
        if (_cutStyleKeys.Count > 0)
        {
            // the styles that came cut, in a section of their own after the others (their keys are in no other section); the
            // ones still unparsed are parsed for this copy only: the whole trees are not kept (a window parses the ones it uses)
            var rest = new XElement(Office + "styles");
            foreach ((string family, string name) key in _cutStyleKeys.OrderBy(pair => pair.Value).Select(pair => pair.Key))
            {
                XElement? style = _styles.TryGetValue(key, out XElement? parsed) ? parsed
                                : _unparsedStyles.TryGetValue(key, out UnparsedStyle? unparsed) ? Parse(unparsed) : null;
                if (style is not null) rest.Add(StyleForRendering(style));
            }
            sections.Add(rest);
        }
        return sections.ToArray();
    }

    private static XElement SectionForRendering(XElement section, bool includeStyles)
    {
        var copy = new XElement(section.Name);
        foreach (XElement element in section.Elements())
        {
            if (element.Name == Style + "font-face")
                copy.Add(new XElement(element.Name, element.Attribute(Style + "name"), element.Attribute(Svg + "font-family")));
            else if (element.Name == Style + "default-style" || (includeStyles && element.Name == Style + "style"))
                copy.Add(StyleForRendering(element));
            else if (IsDrawingStyle(element.Name) || IsDateTimeDataStyle(element.Name))
                copy.Add(new XElement(element));             // whole: gradient stops, marker paths, fill images, date / time parts
        }
        return copy;
    }

    /// <summary>A style:style / style:default-style with only what rendering reads: name, family, parent, list style and the mapped properties.</summary>
    private static XElement StyleForRendering(XElement element)
    {
        var style = new XElement(element.Name, element.Attribute(Style + "name"), element.Attribute(Style + "family"),
                                 element.Attribute(Style + "parent-style-name"), element.Attribute(Style + "list-style-name"));
        foreach (XElement properties in element.Elements())
        {
            HashSet<string>? handled = properties.Name.LocalName switch
            {
                "text-properties" => HandledTextProperties,
                "paragraph-properties" => HandledParagraphProperties,
                "table-properties" or "table-cell-properties" or "table-column-properties" or "table-row-properties" => HandledTableProperties,
                "graphic-properties" => HandledGraphicProperties,
                _ => null,
            };
            if (handled is null) continue;
            // graphic properties are kept whole: every drawing attribute (fill, stroke, markers, text area, wrap) is read
            var kept = new XElement(properties.Name, properties.Attributes().Where(attribute => properties.Name.LocalName == "graphic-properties" || handled.Contains(QualifiedName(attribute.Name))));
            if (properties.Element(Style + "drop-cap") is { } dropCap) kept.Add(new XElement(dropCap.Name, dropCap.Attributes()));   // style:drop-cap is a child element
            if (properties.Element(Style + "tab-stops") is { } tabStops) kept.Add(new XElement(tabStops));        // style:tab-stops too
            if (kept.HasAttributes || kept.HasElements) style.Add(kept);
        }
        return style;
    }

    private static string QualifiedName(XName name)
    {
        string prefix = name.Namespace == Fo ? "fo" : name.Namespace == Style ? "style" : name.Namespace == Text ? "text"
                      : name.Namespace == Table ? "table" : name.Namespace == Svg ? "svg" : name.Namespace == Draw ? "draw" : name.Namespace == LoExt ? "loext" : "?";
        return prefix + ":" + name.LocalName;
    }

    /// <summary>A list style read without a DOM (styles.xml pass, see OdtPackage.ReadStyles); later ones win on name clashes.</summary>
    internal void AddListStyle(string name, OdtListStyle style) => _listStyles[name] = style;

    /// <summary>A style:style cut from the part's bytes (its family and name read from the start tag); parsed when first looked up.</summary>
    internal void AddUnparsedStyle(string family, string name, UnparsedStyle style)
    {
        _unparsedStyles[(family, name)] = style;
        _cutStyleKeys[(family, name)] = _cutStyleCount++;
        _styles.Remove((family, name));
        _resolvedParagraphs.Clear();
        _resolvedSpans.Clear();
    }

    /// <summary>The style:style of that family and name: parsed now when it was cut unparsed, filtered now in a rendering view.</summary>
    private bool TryGetStyle(string family, string name, out XElement style)
    {
        if (_styles.TryGetValue((family, name), out style!)) return true;
        if (_unparsedStyles.TryGetValue((family, name), out UnparsedStyle? unparsed))
        {
            _unparsedStyles.Remove((family, name));
            if (Parse(unparsed) is not { } parsed) { style = null!; return false; }
            _styles[(family, name)] = style = parsed;
            return true;
        }
        if (_renderingSource is { } source && source.TryGetStyle(family, name, out XElement whole))
        {
            _styles[(family, name)] = style = StyleForRendering(whole);
            return true;
        }
        return false;
    }

    /// <summary>The style:style of the bytes; null when they do not parse to one (broken bytes: as if the style were absent).</summary>
    private static XElement? Parse(UnparsedStyle unparsed)
    {
        XElement parsed = OdtXmlFragment.Parse(unparsed.RootOpenTag, unparsed.RootElementName, unparsed.Element, 0, unparsed.Element.Length);
        return parsed.Name == Style + "style" ? parsed : null;
    }

    /// <summary>
    /// A first open's rendering cascade over the indexing cascade: what <see cref="SectionElementsForRendering"/> keeps, with the
    /// style:style elements filtered one by one as they are first used instead of all at once (thousands in a converted Word file).
    /// The indexing cascade is complete by then (every style section precedes office:body) and both are used under the document's lock.
    /// </summary>
    internal static OdtStyleCascade ForRendering(OdtStyleCascade source)
    {
        var view = new OdtStyleCascade { _renderingSource = source };
        foreach (XElement section in source._sections) view.AddSection(SectionForRendering(section, includeStyles: false));
        return view;
    }

    /// <summary>A list style, its bullet style:font-name values resolved through the font-face-decls read so far (all of them by now).</summary>
    public OdtListStyle? ListStyle(string? name)
    {
        if (name is null || !_listStyles.TryGetValue(name, out OdtListStyle? style)) return null;
        style.ResolveFontNames(_fontFamilies);
        return style;
    }

    public XElement? StyleElement(string family, string? name)
        => name is not null && TryGetStyle(family, name, out XElement? style) ? style : null;

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

    /// <summary>
    /// The character formatting a drawing shape's graphic style gives its text (style:text-properties of the graphic style and its
    /// parents over the document default): LibreOffice's shape text takes the shape's item set as its defaults.
    /// </summary>
    public WordRunFormatting GraphicTextFormatting(string? graphicStyleName)
    {
        WordRunFormatting run = DefaultParagraph.Formatting.RunFormatting.Clone();
        if (_defaultStyles.TryGetValue("graphic", out XElement? defaults)) ApplyTextProperties(run, defaults.Element(Style + "text-properties"), run.IsRightToLeft);
        foreach (XElement style in Chain("graphic", graphicStyleName, includeDefault: false))
            ApplyTextProperties(run, style.Element(Style + "text-properties"), run.IsRightToLeft);
        return run;
    }

    /// <summary>The Western, Asian and complex-script language tags of a paragraph style, with a text style over it (null = none set).</summary>
    public readonly record struct OdtLanguages(string? Latin, string? Asian, string? Complex);

    private readonly Dictionary<(string?, string?), OdtLanguages> _languages = new();

    /// <summary>
    /// The three languages (fo:language / fo:script / fo:country, style:*-asian, style:*-complex, or style:rfc-language-tag*) the
    /// default paragraph style, the paragraph style chain and the text style chain set, the last one winning per script
    /// (xmloff: the CharLocale / CharLocaleAsian / CharLocaleComplex properties).
    /// </summary>
    public OdtLanguages LanguagesOf(string? paragraphStyleName, string? textStyleName)
    {
        if (_languages.TryGetValue((paragraphStyleName, textStyleName), out OdtLanguages cached)) return cached;
        string? latin = null, asian = null, complex = null;
        var chain = Chain("paragraph", paragraphStyleName);
        if (textStyleName is not null) chain.AddRange(Chain("text", textStyleName, includeDefault: false));
        foreach (XElement style in chain)
        {
            if (style.Element(Style + "text-properties") is not { } properties) continue;
            latin = LanguageTag(properties, Style + "rfc-language-tag", Fo + "language", Fo + "script", Fo + "country") ?? latin;
            asian = LanguageTag(properties, Style + "rfc-language-tag-asian", Style + "language-asian", Style + "script-asian", Style + "country-asian") ?? asian;
            complex = LanguageTag(properties, Style + "rfc-language-tag-complex", Style + "language-complex", Style + "script-complex", Style + "country-complex") ?? complex;
        }
        var languages = new OdtLanguages(latin, asian, complex);
        _languages[(paragraphStyleName, textStyleName)] = languages;
        return languages;
    }

    private static string? LanguageTag(XElement properties, XName rfc, XName language, XName script, XName country)
    {
        if ((string?)properties.Attribute(rfc) is { Length: > 0 } tag) return tag;
        if ((string?)properties.Attribute(language) is not { Length: > 0 } code) return null;
        string result = code;
        if ((string?)properties.Attribute(script) is { Length: > 0 } scriptCode) result += "-" + scriptCode;
        if ((string?)properties.Attribute(country) is { Length: > 0 } countryCode && countryCode != "none") result += "-" + countryCode;
        return result;
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
    public WordRunFormatting ApplySpan(WordRunFormatting outer, string textStyleName, bool isComplexStrict = false)
    {
        WordRunFormatting run = outer.Clone();
        foreach (XElement style in Chain("text", textStyleName, includeDefault: false))
            ApplyTextProperties(run, style.Element(Style + "text-properties"), outer.IsRightToLeft, isComplexStrict);
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
            if (!TryGetStyle(family, current, out XElement? style)) break;
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
        public int LineHeight; public bool IsLineHeightAuto = true; public bool IsLineHeightExact;
        public string? Background;
        public string? BorderTop, BorderBottom, BorderLeft, BorderRight;
        public bool IsPageBreakBefore;
        public int DropCapLines, DropCapLength, DropCapDistance;
        public WordTabStop[]? TabStops;       // style:tab-stops: a style's own list replaces the inherited one (one SvxTabStopItem)
        public int? TabStopDistance;          // style:tab-stop-distance (usually on the default style)

        public WordParagraphFormatting ToFormatting()
        {
            var formatting = new WordParagraphFormatting
            {
                IsRightToLeft = IsRightToLeft, DropCapLines = DropCapLines, DropCapCharacterCount = DropCapLength, DropCapDistanceTwips = DropCapDistance,
            };
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
            formatting.IsLineSpacingExact = IsLineHeightExact;
            formatting.ShadingHex = Background;
            formatting.BorderTopCss = BorderTop;
            formatting.BorderBottomCss = BorderBottom;
            formatting.BorderStartCss = IsRightToLeft ? BorderRight : BorderLeft;
            formatting.BorderEndCss = IsRightToLeft ? BorderLeft : BorderRight;
            formatting.TabStops = TabStops ?? [];
            formatting.DefaultTabStopTwips = TabStopDistance ?? 0;
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
            else if (Twips(lineHeight) is { } fixedHeight) { raw.LineHeight = fixedHeight; raw.IsLineHeightAuto = false; raw.IsLineHeightExact = true; }
        }
        if (Twips(properties.Attribute(Style + "line-height-at-least")) is { } atLeast && atLeast > 0) { raw.LineHeight = atLeast; raw.IsLineHeightAuto = false; raw.IsLineHeightExact = false; }
        if ((string?)properties.Attribute(Fo + "background-color") is { } background) raw.Background = ColorHex(background);
        if (BorderCss(properties.Attribute(Fo + "border")) is { } border) raw.BorderTop = raw.BorderBottom = raw.BorderLeft = raw.BorderRight = border;
        raw.BorderTop = BorderCss(properties.Attribute(Fo + "border-top")) ?? raw.BorderTop;
        raw.BorderBottom = BorderCss(properties.Attribute(Fo + "border-bottom")) ?? raw.BorderBottom;
        raw.BorderLeft = BorderCss(properties.Attribute(Fo + "border-left")) ?? raw.BorderLeft;
        raw.BorderRight = BorderCss(properties.Attribute(Fo + "border-right")) ?? raw.BorderRight;
        if ((string?)properties.Attribute(Fo + "break-before") is { } breakBefore) raw.IsPageBreakBefore = breakBefore == "page";
        if (Twips(properties.Attribute(Style + "tab-stop-distance")) is > 0 and var distance) raw.TabStopDistance = distance;
        if (properties.Element(Style + "tab-stops") is { } tabStops) raw.TabStops = ReadTabStops(tabStops);
        // style:drop-cap (ODF 1.3 17.9), read as LibreOffice xmloff txtdropi.cxx ProcessAttrs: style:lines below 2 = no drop cap,
        // style:length a count or "word" (-1 here), with lines and no length 1 character; style:distance the gap to the text.
        if (properties.Element(Style + "drop-cap") is { } dropCap)
        {
            int lines = int.TryParse((string?)dropCap.Attribute(Style + "lines"), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int parsedLines) ? parsedLines : 1;
            raw.DropCapLines = lines is >= 2 and <= 255 ? lines : 0;
            string? length = (string?)dropCap.Attribute(Style + "length");
            raw.DropCapLength = length == "word" ? -1
                : int.TryParse(length, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int count) && count is >= 1 and <= 255 ? count : 1;
            raw.DropCapDistance = Math.Max(0, Twips(dropCap.Attribute(Style + "distance")) ?? 0);
        }
    }

    // ── text properties ───────────────────────────────────────────────────────────────

    /// <summary>
    /// style:text-properties onto a run. Western and complex-script variants are kept apart (size, font, weight, posture):
    /// the renderer draws each piece of text in the set LibreOffice picks for it (WordTextScript). While no complex weight /
    /// posture is set anywhere in the chain the complex-script text uses the western one, as before. <paramref name="isComplexStrict"/>:
    /// a list label whose script is known (WordListLabelScript) takes only the complex-script variants, as LibreOffice draws a CTL portion.
    /// </summary>
    public void ApplyTextProperties(WordRunFormatting run, XElement? properties, bool isRightToLeft, bool isComplexStrict = false)
    {
        if (properties is null) return;
        static bool IsBoldWeight(string weight) => weight == "bold" || (int.TryParse(weight, out int numeric) && numeric >= 600);
        // a western value set after the complex one (or in a strict label) must not reach the complex-script text
        void KeepComplexApart()
        {
            if (run.HasComplexScriptBoldItalic) return;
            run.HasComplexScriptBoldItalic = true;
            run.IsBoldComplexScript = run.IsBold;
            run.IsItalicComplexScript = run.IsItalic;
        }
        if (isComplexStrict && (properties.Attribute(Fo + "font-weight") is not null || properties.Attribute(Fo + "font-style") is not null)) KeepComplexApart();
        if ((string?)properties.Attribute(Fo + "font-weight") is { } weight) run.IsBold = IsBoldWeight(weight);
        if ((string?)properties.Attribute(Fo + "font-style") is { } posture) run.IsItalic = posture is "italic" or "oblique";
        if ((string?)properties.Attribute(Style + "font-weight-complex") is { } complexWeight) { KeepComplexApart(); run.IsBoldComplexScript = IsBoldWeight(complexWeight); }
        if ((string?)properties.Attribute(Style + "font-style-complex") is { } complexPosture) { KeepComplexApart(); run.IsItalicComplexScript = complexPosture is "italic" or "oblique"; }

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
        if (FontName(properties, Style + "font-name", Fo + "font-family") is { } font) run.FontName = FontWithoutFeatures(font, out run.FontFeaturesCss);
        if (FontName(properties, Style + "font-name-complex", Style + "font-family-complex") is { } complexFont)
            run.ComplexScriptFontName = FontWithoutFeatures(complexFont, out run.ComplexScriptFontFeaturesCss);

        if ((string?)properties.Attribute(Fo + "color") is { } color) run.ColorHex = ColorHex(color);
        if ((string?)properties.Attribute(Fo + "background-color") is { } background)
            run.HighlightCssColor = ColorHex(background) is { } hex ? "#" + hex : null;
        if ((string?)properties.Attribute(Style + "text-position") is { } position) ApplyTextPosition(run, position);
        if ((string?)properties.Attribute(Fo + "text-transform") is { } transform) run.IsAllCaps = transform == "uppercase";
        if ((string?)properties.Attribute(Fo + "font-variant") is { } variant) run.IsSmallCaps = variant == "small-caps";
        if ((string?)properties.Attribute(Fo + "letter-spacing") is { } spacing) run.CharacterSpacingTwips = spacing == "normal" ? 0 : Twips(spacing) ?? 0;
        // style:text-scale (ODF 1.3 20.393, a percent): LibreOffice maps it to CharScaleWidth (xmloff txtprmap.cxx:254)
        if ((string?)properties.Attribute(Style + "text-scale") is { } textScale)
            run.CharacterScalePercent = int.TryParse(textScale.TrimEnd('%'), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int scalePercent) ? scalePercent : 100;
        if ((string?)properties.Attribute(Text + "display") is { } display) run.IsHidden = display == "none";
        if ((string?)properties.Attribute(Fo + "hyphenate") is { } hyphenate) run.Hyphenate = hyphenate == "true";
        // fo:language + fo:country (ODF 20.218 / 20.183): the Latin-text language, for the hyphenated paragraph lang attribute
        if ((string?)properties.Attribute(Fo + "language") is { } language)
            run.Language = language is "none" or "zxx" ? null
                : (string?)properties.Attribute(Fo + "country") is { Length: > 0 } country && country != "none" ? language + "-" + country : language;
        if ((string?)properties.Attribute(Style + "text-outline") is { } outline) run.IsOutline = outline == "true";
        if ((string?)properties.Attribute(Fo + "text-shadow") is { } shadow) run.IsShadow = shadow != "none";
        if ((string?)properties.Attribute(Style + "font-relief") is { } relief) { run.IsEmboss = relief == "embossed"; run.IsImprint = relief == "engraved"; }
        if (BorderCss(properties.Attribute(Fo + "border")) is { } border) run.BorderCss = border;
        if ((string?)properties.Attribute(Style + "text-emphasize") is { } emphasis) ApplyTextEmphasis(run, emphasis);
    }

    /// <summary>
    /// style:text-emphasize (ODF 1.3 part 3 20.381): "none" or a mark (dot / circle / disc / accent) plus "above" / "below",
    /// read as LibreOffice's XMLTextEmphasizePropHdl_Impl::importXML (xmloff/source/text/txtprhdl.cxx:907-950). The CSS
    /// marks follow the .docx w:em mapping: dot / disc = filled dot, circle = open circle, accent = filled sesame (the comma mark).
    /// </summary>
    private static void ApplyTextEmphasis(WordRunFormatting run, string value)
    {
        string[] tokens = value.Split([' '], StringSplitOptions.RemoveEmptyEntries);
        run.EmphasisMarkCss = null;
        run.IsEmphasisMarkBelow = false;
        foreach (string token in tokens)
        {
            if (token == "below") run.IsEmphasisMarkBelow = true;
            else if (token is "dot" or "disc") run.EmphasisMarkCss = "filled dot";
            else if (token == "circle") run.EmphasisMarkCss = "open circle";
            else if (token == "accent") run.EmphasisMarkCss = "filled sesame";
        }
        if (run.EmphasisMarkCss is null) run.IsEmphasisMarkBelow = false;
    }

    /// <summary>"super 58%", "sub 58%", "33% 58%", "-33% 100%", "0% 100%" (ODF 1.3 part 3 §20.374).</summary>
    private static void ApplyTextPosition(WordRunFormatting run, string value)
    {
        string first = value.Split(' ')[0];
        run.IsSuperscript = first == "super" || (Percent(first) is { } up && up > 0);
        run.IsSubscript = first == "sub" || (Percent(first) is { } down && down < 0);
    }

    /// <summary>
    /// A LibreOffice font name with its OpenType features ("Gabriola:ss05&amp;-liga") split into the family and a
    /// font-feature-settings list ("'ss05' 1,'liga' 0"; null = none). LibreOffice vcl/source/font/FeatureParser.cxx:24-56:
    /// the features follow the first ':' (include/vcl/font/FeatureParser.hxx:23-24, FeaturePrefix ':' and FeatureSeparator '&amp;'),
    /// each parsed by HarfBuzz hb_feature_from_string (vcl/source/font/Feature.cxx:44-58: "liga", "+liga", "-liga", "liga=0",
    /// a [start:end] range); "lang=xx" sets the language, not a feature. Without features LibreOffice shapes with HarfBuzz's
    /// defaults (liga, clig, calt on), which is the browser's default too, so nothing is written then.
    /// </summary>
    internal static string FontWithoutFeatures(string font, out string? featuresCss)
    {
        featuresCss = null;
        int colon = font.IndexOf(':');
        if (colon < 0) return font;
        var css = new StringBuilder();
        foreach (string token in font.Substring(colon + 1).Split('&'))
        {
            string feature = token.Trim();
            int range = feature.IndexOf('[');
            int equals = feature.IndexOf('=');
            string tag = feature.Substring(0, range >= 0 ? range : equals >= 0 ? equals : feature.Length).Trim();
            if (tag == "lang") continue;
            int value = 1;
            if (tag.StartsWith("-", StringComparison.Ordinal)) { value = 0; tag = tag.Substring(1); }
            else if (tag.StartsWith("+", StringComparison.Ordinal)) tag = tag.Substring(1);
            if (equals >= 0)
            {
                string number = feature.Substring(equals + 1).Trim();
                if (number is "on" or "true") value = 1;
                else if (number is "off" or "false") value = 0;
                else if (!int.TryParse(number, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out value)) continue;
            }
            if (tag.Length is 0 or > 4) continue;
            bool isTag = true;
            foreach (char c in tag) if (c is < '!' or > '~' or '\'' or '"' or '\\') isTag = false;
            if (!isTag) continue;
            if (css.Length > 0) css.Append(',');
            css.Append('\'').Append(tag.PadRight(4)).Append("' ").Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        if (css.Length > 0) featuresCss = css.ToString();
        return font.Substring(0, colon).Trim();
    }

    private string? FontName(XElement properties, XName nameAttribute, XName familyAttribute)
    {
        if ((string?)properties.Attribute(nameAttribute) is { } name)
            return _fontFamilies.TryGetValue(name, out string? family) ? family : name;
        return ((string?)properties.Attribute(familyAttribute))?.Trim('\'', '"', ' ');
    }

    // ── tables ────────────────────────────────────────────────────────────────────────

    /// <summary>CSS for a table:table-cell's style: borders, padding, background, vertical alignment.</summary>
    /// <summary>
    /// A table row's height (style:table-row-properties): style:row-height is exact, style:min-row-height at least (LibreOffice
    /// sw/source/filter/xml/xmlitemm.cxx:131-132: MID_FRMSIZE_FIX_HEIGHT / MID_FRMSIZE_MIN_HEIGHT) - twips, negative = exact,
    /// 0 = none.
    /// </summary>
    public int RowHeightTwips(string? rowStyleName)
    {
        XElement? properties = StyleElement("table-row", rowStyleName)?.Element(Style + "table-row-properties");
        if (properties is null) return 0;
        if (Twips(properties.Attribute(Style + "row-height")) is > 0 and var exact) return -exact;
        return Twips(properties.Attribute(Style + "min-row-height")) is > 0 and var least ? least : 0;
    }

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
        if ((string?)properties.Attribute(Fo + "background-color") is { } background && ColorHex(background) is { } hex) css.Append("background-color:").Append(WordHtmlColor.Css(hex)).Append(';');
        if ((string?)properties.Attribute(Style + "vertical-align") is "middle" or "bottom" or "top")
            css.Append("vertical-align:").Append((string?)properties.Attribute(Style + "vertical-align")).Append(';');
        if (CellTextDirectionName((string?)properties.Attribute(LoExt + "writing-mode") ?? (string?)properties.Attribute(Style + "writing-mode")) is { } name
            && DocxTableFormatting.TextDirectionCss(name) is { } direction)
            css.Append(direction);
        return css.ToString();
    }

    /// <summary>
    /// A cell's writing mode as the .docx ST_TextDirection it draws like (the .docx cell mapping is reused). As LibreOffice reads a
    /// cell (sw/source/filter/xml/xmlitemm.cxx :280-282 style: / loext:writing-mode -> RES_FRAMEDIR; xmlimpit.cxx :1008-1032: loext
    /// bt-lr -> Vertical_LR_BT, tb-rl90 -> Vertical_RL_TB90, else the writing-mode handler: tb-rl / tb -> Vertical_RL_TB, tb-lr ->
    /// Vertical_LR_TB). style:direction (ttb, stacked letters) is not read by Writer for cells, so it is not drawn either.
    /// </summary>
    private static string? CellTextDirectionName(string? writingMode)
    {
        return writingMode switch
        {
            "tb-rl" or "tb" => "tbRl",
            "tb-rl90" => "tbRlV",
            "bt-lr" => "btLr",
            "tb-lr" => "tbLrV",
            _ => null,
        };
    }

    /// <summary>A column's width in points (style:column-width), or null.</summary>
    public string? ColumnWidthCss(string? columnStyleName)
    {
        XElement? properties = StyleElement("table-column", columnStyleName)?.Element(Style + "table-column-properties");
        return Twips(properties?.Attribute(Style + "column-width")) is { } width ? WordParagraphFormatting.TwipsToPt(width) + "pt" : null;
    }

    private static bool IsDrawingStyle(XName name) => name.Namespace == Draw && name.LocalName is "gradient" or "hatch" or "fill-image" or "opacity" or "marker" or "stroke-dash";

    /// <summary>A draw:gradient by its draw:name (office:styles), null when unknown.</summary>
    public XElement? Gradient(string? name) => DrawingStyle("gradient", name);

    /// <summary>A named drawing style (draw:gradient / hatch / fill-image / opacity / marker / stroke-dash) by its draw:name; null when unknown.</summary>
    public XElement? DrawingStyle(string kind, string? name) => name is not null && _drawingStyles.TryGetValue((kind, name), out XElement? style) ? style : null;

    /// <summary>A graphic style's property through its parent chain (style:parent-style-name), then the graphic default style.</summary>
    public string? GraphicProperty(string? styleName, XName attribute)
    {
        for (int depth = 0; styleName is not null && depth < 64; depth++)
        {
            XElement? style = StyleElement("graphic", styleName);
            if (style is null) break;
            if ((string?)style.Element(Style + "graphic-properties")?.Attribute(attribute) is { } value) return value;
            styleName = (string?)style.Attribute(Style + "parent-style-name");
        }
        return _defaultStyles.TryGetValue("graphic", out XElement? defaults) ? (string?)defaults.Element(Style + "graphic-properties")?.Attribute(attribute) : null;
    }

    /// <summary>A table's CSS (width, alignment, background) and direction (style:writing-mode).</summary>
    public (string css, bool? isRightToLeft) TableCss(string? tableStyleName)
    {
        XElement? properties = StyleElement("table", tableStyleName)?.Element(Style + "table-properties");
        if (properties is null) return ("", null);
        var css = new StringBuilder();
        if ((string?)properties.Attribute(Style + "rel-width") is { } relative && Percent(relative) is { } percent) css.Append("width:").Append(percent.ToString("0.#", CultureInfo.InvariantCulture)).Append("%;");
        else if (Twips(properties.Attribute(Style + "width")) is { } width) css.Append("width:").Append(WordParagraphFormatting.TwipsToPt(width)).Append("pt;");
        // table:align (ODF 1.3 part 3 20.414): left / right / center against the margins, "margins" fills the space between
        // them (fo:margin-left / fo:margin-right, 20.237 / 20.239); physical sides. Absent = left (LibreOffice writes it always).
        string? Margin(string side) => Twips(properties.Attribute(Fo + "margin-" + side)) is { } twips and not 0 ? WordParagraphFormatting.TwipsToPt(twips) + "pt" : null;
        switch ((string?)properties.Attribute(Table + "align"))
        {
            case "center": css.Append("margin-left:auto;margin-right:auto;"); break;
            case "right": css.Append("margin-left:auto;margin-right:").Append(Margin("right") ?? "0").Append(';'); break;
            case "margins": css.Append("margin-left:").Append(Margin("left") ?? "0").Append(";margin-right:").Append(Margin("right") ?? "0").Append(';'); break;
            default: css.Append("margin-left:").Append(Margin("left") ?? "0").Append(";margin-right:auto;"); break;
        }
        if ((string?)properties.Attribute(Fo + "background-color") is { } background && ColorHex(background) is { } hex) css.Append("background-color:").Append(WordHtmlColor.Css(hex)).Append(';');
        bool? rtl = (string?)properties.Attribute(Style + "writing-mode") switch { "rl-tb" or "rl" => true, "lr-tb" or "lr" => false, _ => null };
        return (css.ToString(), rtl);
    }

    // ── value parsing (ODF 1.3 part 3 §18: length = number + unit cm|mm|in|pt|pc|px) ──

    /// <summary>
    /// style:tab-stops (ODF 1.3 part 3 17.8, 19.497-19.506): style:position (from the indent, LibreOffice's default
    /// TABS_RELATIVE_TO_INDENT), style:type left / center / right / char (a decimal stop on style:char), the leader as its text
    /// (style:leader-text, LibreOffice xmloff/source/style/xmltabi.cxx: the fill character; a leader style without text draws ".")
    /// - the fill characters Word has map to its leaders, any other to dots.
    /// </summary>
    private static WordTabStop[] ReadTabStops(XElement tabStops)
    {
        var stops = new List<WordTabStop>();
        foreach (XElement stop in tabStops.Elements(Style + "tab-stop"))
        {
            if (Twips(stop.Attribute(Style + "position")) is not { } position) continue;
            string alignment = (string?)stop.Attribute(Style + "type") switch { "center" => "center", "right" => "right", "char" => "decimal", _ => "left" };
            string? text = (string?)stop.Attribute(Style + "leader-text");
            string? style = (string?)stop.Attribute(Style + "leader-style");
            if (string.IsNullOrEmpty(text) && style is not null and not "none") text = ".";
            string leader = text switch
            {
                null or "" or " " => "none", "." => "dot", "-" => "hyphen", "_" => "underscore", "\u00B7" => "middleDot", _ => "dot",
            };
            stops.Add(new WordTabStop(position, alignment, leader));
        }
        stops.Sort((a, b) => a.PositionTwips.CompareTo(b.PositionTwips));
        return stops.ToArray();
    }

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
