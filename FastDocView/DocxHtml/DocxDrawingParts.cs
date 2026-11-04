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

    /// <summary>The namespace-free dsp:spTree of a SmartArt's drawing and the drawing part's name; null when it has none.</summary>
    public static (XElement Tree, string PartName)? DiagramDrawing(RelatedPartReader related, string dataRelationshipId)
    {
        if (related(dataRelationshipId) is not { } data) return null;
        foreach (XElement extension in data.Xml.Descendants().Where(e => e.Name.LocalName == "dataModelExt"))
        {
            if ((string?)extension.Attribute("relId") is not { Length: > 0 } relId || related(relId) is not { } drawing) continue;
            XElement? tree = drawing.Xml.Root?.Elements().FirstOrDefault(e => e.Name.LocalName == "spTree");
            if (tree is null || !tree.Elements().Any(e => e.Name.LocalName is "sp" or "grpSp")) continue;
            return (DocxShapeGeometry.WithoutNamespaces(tree), drawing.PartName);
        }
        return null;
    }

    public static DocxChart? Chart(RelatedPartReader related, string chartRelationshipId, DocxTheme? theme)
        => related(chartRelationshipId) is { } chart && chart.Type.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) && chart.Xml.Root is { } root
           ? DocxChart.Read(DocxShapeGeometry.WithoutNamespaces(root), theme) : null;

    /// <summary>The text a SmartArt or chart drawing shows, for search: shape texts in drawing order / the chart's texts, one per line.</summary>
    public static string? SearchText(RelatedPartReader related, string relationshipId, DocxTheme? theme)
    {
        if (related(relationshipId) is not { } part) return null;
        if (part.Type.EndsWith("/chart", StringComparison.OrdinalIgnoreCase))
            return Chart(related, relationshipId, theme) is { } chart ? string.Join("\n", chart.Texts()) : null;
        if (DiagramDrawing(related, relationshipId) is not { } drawing) return null;
        return string.Concat(drawing.Tree.Descendants("txBody").Select(DocxDrawingTextHtml.Text));
    }
}
