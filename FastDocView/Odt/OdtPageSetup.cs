using System.Xml;

namespace FastDocView.Odt;

/// <summary>
/// The page text width of each master page of an OpenDocument text: style:master-page
/// (style:page-layout-name) -> style:page-layout -> style:page-layout-properties fo:page-width minus
/// fo:margin-left, fo:margin-right (and LibreOffice's loext:margin-gutter). Read from styles.xml (or a
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

    /// <summary>The width of the text before any block names a master page.</summary>
    public int DefaultTextWidthTwips { get; private set; } = WordDocumentPageSetup.DefaultTextWidthTwips;

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
                    int? gutter = WordDocumentPageSetup.ParseTwips(reader.GetAttribute("margin-gutter", LibreOfficeNamespace));
                    // ODF has no default page size; LibreOffice's own default page style is A4 with 2 cm margins,
                    // but a layout without fo:page-width is rare - fall back to the Word default like other formats.
                    _layoutWidths[_currentLayout] = page is { } width
                        ? WordDocumentPageSetup.TextWidth(width, left ?? 0, right ?? 0, gutter ?? 0)
                        : WordDocumentPageSetup.DefaultTextWidthTwips;
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
                if (_layoutWidths.TryGetValue(layout, out int width) && !setup._masterPageWidths.ContainsKey(master)) setup._masterPageWidths[master] = width;
            if (setup.TextWidthOfMasterPage("Standard") is { } standard) setup.DefaultTextWidthTwips = standard;
            else if (_masterLayouts.Count > 0 && setup.TextWidthOfMasterPage(_masterLayouts[0].master) is { } first) setup.DefaultTextWidthTwips = first;
            return setup;
        }
    }
}
