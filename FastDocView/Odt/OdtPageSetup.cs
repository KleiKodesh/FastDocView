using System.Xml;

namespace FastDocView.Odt;

/// <summary>
/// The page text width of each master page of an OpenDocument text: style:master-page
/// (style:page-layout-name) -> style:page-layout -> style:page-layout-properties fo:page-width minus
/// fo:margin-left, fo:margin-right (LibreOffice writes its loext:margin-gutter inside them). Read from styles.xml (or a
/// flat .fodt's head, stopping at office:body) with XmlReader: page layouts and master pages only.
/// A paragraph or table whose style names a master page (style:master-page-name) starts a section with
/// that page's width; the text before the first such block uses the default master page "Standard"
/// (LibreOffice's default page style), or the first master page.
/// </summary>
public sealed class OdtPageSetup
{
    private const string StyleNamespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    private const string FoNamespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    private const string OfficeNamespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private const string LibreOfficeNamespace = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

    private readonly Dictionary<string, int> _masterPageWidths = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _masterPageColumnLefts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, (int Top, int Bottom)> _masterPageVerticalMargins = new(StringComparer.Ordinal);

    /// <summary>
    /// A master page's top and bottom margins (fo:margin-top / fo:margin-bottom, or fo:margin; a gutter at the top is
    /// already inside fo:margin-top - xmloff PageMasterImportPropMapper.cxx); null = not known.
    /// </summary>
    public (int Top, int Bottom)? VerticalMarginsOfMasterPage(string? masterPageName)
        => masterPageName is not null && _masterPageVerticalMargins.TryGetValue(masterPageName, out var margins) ? margins : null;

    /// <summary>
    /// What the file does not say takes Word's default (the rule when Word and LibreOffice differ): Letter (8.5 in) with
    /// 1 in margins on every side (<see cref="WordDocumentPageSetup"/>). LibreOffice is cited only for how the attributes
    /// are read; its own default page style (A4 / 2 cm in a metric locale, docdesc.cxx:68-113) is not followed.
    /// </summary>
    public int DefaultTopMarginTwips { get; private set; } = WordDocumentPageSetup.DefaultMarginTwips;
    public int DefaultBottomMarginTwips { get; private set; } = WordDocumentPageSetup.DefaultMarginTwips;

    /// <summary>Where the text column of a master page's layout starts on the page (fo:margin-left, which holds a left gutter); null = not known.</summary>
    public int? ColumnLeftOfMasterPage(string? masterPageName)
        => masterPageName is not null && _masterPageColumnLefts.TryGetValue(masterPageName, out int left) ? left : null;

    /// <summary>The width of the text before any block names a master page.</summary>
    public int DefaultTextWidthTwips { get; private set; } = WordDocumentPageSetup.DefaultTextWidthTwips;
    /// <summary>The default master page's column start on the page (twips); -1 = not known.</summary>
    public int DefaultColumnLeftTwips { get; private set; } = -1;

    public static OdtPageSetup Empty { get; } = new();

    /// <summary>The text width of a master page, or null when it is not known.</summary>
    public int? TextWidthOfMasterPage(string? masterPageName)
        => masterPageName is not null && _masterPageWidths.TryGetValue(masterPageName, out int width) ? width : null;

    public static OdtPageSetup Read(Stream xml)
    {
        var collector = new Collector();
        try
        {
            using XmlReader reader = WordDocumentXml.CreateReader(xml, tolerateUndeclaredPrefixes: true);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "body" && reader.NamespaceURI == OfficeNamespace) break;      // .fodt: the text follows
                collector.Visit(reader);
            }
        }
        catch (XmlException) { }          // malformed: keep what was read
        return collector.Finish();
    }

    /// <summary>
    /// Collects page layouts and master pages from the elements of a styles pass it is shown (each start element, in
    /// document order) - lets the styles.xml pass that reads the named styles take the page setup in the same pass.
    /// </summary>
    internal sealed class Collector
    {
        private readonly Dictionary<string, int> _layoutWidths = new(StringComparer.Ordinal);
        private readonly Dictionary<string, int> _layoutColumnLefts = new(StringComparer.Ordinal);
        private readonly Dictionary<string, (int Top, int Bottom)> _layoutVerticalMargins = new(StringComparer.Ordinal);
        private readonly List<(string master, string layout)> _masterLayouts = new();
        private string? _currentLayout;
        private int _currentLayoutDepth = -1;

        /// <summary>The reader is on a start element; its attributes are read, the reader is not moved.</summary>
        public void Visit(XmlReader reader)
        {
            if (reader.NamespaceURI != StyleNamespace) return;
            if (_currentLayout is not null && reader.Depth <= _currentLayoutDepth) _currentLayout = null;
            switch (reader.LocalName)
            {
                case "page-layout":
                    _currentLayout = reader.GetAttribute("name", StyleNamespace);
                    _currentLayoutDepth = reader.Depth;
                    break;
                case "page-layout-properties" when _currentLayout is not null && reader.Depth == _currentLayoutDepth + 1:
                {
                    int? page = WordDocumentPageSetup.ParseTwips(reader.GetAttribute("page-width", FoNamespace));
                    int? margin = WordDocumentPageSetup.ParseTwips(reader.GetAttribute("margin", FoNamespace));
                    int? left = WordDocumentPageSetup.ParseTwips(reader.GetAttribute("margin-left", FoNamespace)) ?? margin;
                    int? right = WordDocumentPageSetup.ParseTwips(reader.GetAttribute("margin-right", FoNamespace)) ?? margin;
                    // loext:margin-gutter is already inside fo:margin-left (or fo:margin-right with a right-to-left gutter,
                    // fo:margin-top with GutterAtTop): LibreOffice writes the margin with the gutter and subtracts it on import
                    // (xmloff PageMasterImportPropMapper.cxx:412-452), so the text width and the column start are the
                    // margins as written, wherever the gutter is
                    // missing values take Word's defaults (the rule when Word and LibreOffice differ): Letter, 1 in margins
                    const int WordMargin = WordDocumentPageSetup.DefaultMarginTwips;
                    _layoutWidths[_currentLayout] = WordDocumentPageSetup.TextWidth(page ?? WordDocumentPageSetup.LetterPageWidthTwips, left ?? WordMargin, right ?? WordMargin, 0);
                    _layoutColumnLefts[_currentLayout] = WordDocumentPageSetup.ColumnLeft(left ?? WordMargin, 0);
                    _layoutVerticalMargins[_currentLayout] = (
                        WordDocumentPageSetup.VerticalMargin(WordDocumentPageSetup.ParseTwips(reader.GetAttribute("margin-top", FoNamespace)) ?? margin ?? WordMargin),
                        WordDocumentPageSetup.VerticalMargin(WordDocumentPageSetup.ParseTwips(reader.GetAttribute("margin-bottom", FoNamespace)) ?? margin ?? WordMargin));
                    break;
                }
                case "master-page":
                    if (reader.GetAttribute("name", StyleNamespace) is { } master && reader.GetAttribute("page-layout-name", StyleNamespace) is { } layout)
                        _masterLayouts.Add((master, layout));
                    break;
            }
        }

        public OdtPageSetup Finish()
        {
            var setup = new OdtPageSetup();
            foreach ((string master, string layout) in _masterLayouts)
            {
                if (_layoutWidths.TryGetValue(layout, out int width) && !setup._masterPageWidths.ContainsKey(master)) setup._masterPageWidths[master] = width;
                if (_layoutColumnLefts.TryGetValue(layout, out int left) && !setup._masterPageColumnLefts.ContainsKey(master)) setup._masterPageColumnLefts[master] = left;
                if (_layoutVerticalMargins.TryGetValue(layout, out var margins) && !setup._masterPageVerticalMargins.ContainsKey(master)) setup._masterPageVerticalMargins[master] = margins;
            }
            string? defaultMaster = setup.TextWidthOfMasterPage("Standard") is not null ? "Standard" : _masterLayouts.Count > 0 ? _masterLayouts[0].master : null;
            if (setup.TextWidthOfMasterPage(defaultMaster) is { } width0) setup.DefaultTextWidthTwips = width0;
            if (setup.ColumnLeftOfMasterPage(defaultMaster) is { } left0) setup.DefaultColumnLeftTwips = left0;
            if (setup.VerticalMarginsOfMasterPage(defaultMaster) is { } margins0) (setup.DefaultTopMarginTwips, setup.DefaultBottomMarginTwips) = margins0;
            return setup;
        }
    }
}
