using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// The drawing parts a w:drawing names, as rendering and search both read them (on demand, for the block being shown or
/// searched): a SmartArt's drawing (the dgm:relIds r:dm data part's dsp:dataModelExt relId, resolved against the part the
/// drawing is in, as LibreOffice does: oox datamodelcontext.cxx:403-404, diagram.cxx:615-631; a relId that resolves to no
/// part, or a drawing without shapes, = no drawing) and a chart (c:chart r:id -> the chart part).
/// </summary>
internal static class DocxDrawingParts
{
    public delegate (string Type, string PartName, XDocument Xml)? RelatedPartReader(string relationshipId);

    /// <summary>
    /// The namespace-free dsp:spTree of a SmartArt's drawing and the drawing part's name. Without a drawing part the diagram is
    /// laid out from its data model and the layout definition (r:lo), with the colors (r:cs) and quickStyle (r:qs) parts, in a
    /// frame of the given EMU size (DocxDiagramLayout; the part name is then the data part's). Null when neither works.
    /// </summary>
    public static (XElement Tree, string PartName)? DiagramDrawing(RelatedPartReader related, XElement diagram, double widthEmu = DefaultDiagramWidthEmu, double heightEmu = DefaultDiagramHeightEmu)
    {
        if (related((string?)diagram.Attribute("dm") ?? "") is not { } data) return null;
        foreach (XElement extension in data.Xml.Descendants().Where(e => e.Name.LocalName == "dataModelExt"))
        {
            if ((string?)extension.Attribute("relId") is not { Length: > 0 } relId || related(relId) is not { } drawing) continue;
            XElement? tree = drawing.Xml.Root?.Elements().FirstOrDefault(e => e.Name.LocalName == "spTree");
            if (tree is null || !tree.Elements().Any(e => e.Name.LocalName is "sp" or "grpSp")) continue;
            return (DocxShapeGeometry.WithoutNamespaces(tree), drawing.PartName);
        }
        if ((string?)diagram.Attribute("lo") is not { Length: > 0 } layoutId || related(layoutId) is not { } layout || layout.Xml.Root is null || data.Xml.Root is null) return null;
        XElement? Part(string attribute) => (string?)diagram.Attribute(attribute) is { Length: > 0 } id && related(id)?.Xml.Root is { } root ? DocxShapeGeometry.WithoutNamespaces(root) : null;
        XElement? laidOut = DocxDiagramLayout.Lay(DocxShapeGeometry.WithoutNamespaces(data.Xml.Root), DocxShapeGeometry.WithoutNamespaces(layout.Xml.Root),
                                                  Part("cs"), Part("qs"), widthEmu, heightEmu);
        return laidOut is null ? null : (laidOut, data.PartName);
    }

    /// <summary>Word's default SmartArt size (5486400 x 3200400 EMU), for search, where the frame size doesn't matter.</summary>
    public const double DefaultDiagramWidthEmu = 5486400, DefaultDiagramHeightEmu = 3200400;

    public static DocxChart? Chart(RelatedPartReader related, string chartRelationshipId, DocxTheme? theme)
        => related(chartRelationshipId) is { } chart && chart.Type.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) && chart.Xml.Root is { } root
           ? DocxChart.Read(DocxShapeGeometry.WithoutNamespaces(root), theme) : null;

    public static DocxChartEx? ChartEx(RelatedPartReader related, string chartRelationshipId, DocxTheme? theme)
        => related(chartRelationshipId) is { } chart && chart.Type.EndsWith("/chartEx", StringComparison.OrdinalIgnoreCase) && chart.Xml.Root is { } root
           ? DocxChartEx.Read(DocxShapeGeometry.WithoutNamespaces(root), theme) : null;

    /// <summary>The text a SmartArt or chart drawing shows, for search: shape texts in drawing order / the chart's texts, one per line.</summary>
    public static string? SearchText(RelatedPartReader related, string relationshipId, DocxTheme? theme)
    {
        // a SmartArt comes as "dm lo" (DocxSearchTextScanner / WordXmlSearchText): its data part and, if named, its layout part
        string layoutId = "";
        int space = relationshipId.IndexOf(' ');
        if (space >= 0) (relationshipId, layoutId) = (relationshipId.Substring(0, space), relationshipId.Substring(space + 1));
        if (related(relationshipId) is not { } part) return null;
        if (part.Type.EndsWith("/chartEx", StringComparison.OrdinalIgnoreCase))
            return ChartEx(related, relationshipId, theme) is { CanDraw: true } chartEx ? string.Join("\n", chartEx.Texts()) : null;
        if (part.Type.EndsWith("/chart", StringComparison.OrdinalIgnoreCase))
            return Chart(related, relationshipId, theme) is { } chart ? string.Join("\n", chart.Texts()) : null;
        if (DiagramDrawing(related, new XElement("diagram", new XAttribute("dm", relationshipId))) is { } drawing)
            return string.Concat(drawing.Tree.Descendants("txBody").Select(DocxDrawingTextHtml.Text));
        // no drawing part: drawn from the layout part when there is one (DiagramDrawing), showing the data points' own text;
        // search reads it in data-model order
        if (layoutId.Length == 0 || related(layoutId)?.Xml.Root is null || part.Xml.Root is not { } data) return null;
        XElement model = DocxShapeGeometry.WithoutNamespaces(data);
        string text = string.Concat(model.Element("ptLst")?.Elements("pt")
                                         .Where(p => (string?)p.Attribute("type") is null or "node" or "asst")
                                         .Select(p => p.Element("t")).OfType<XElement>()
                                         .Select(t => DocxDrawingTextHtml.Text(new XElement("txBody", t.Elements("p")))) ?? []);
        return text.Length > 0 ? text : null;
    }
}
