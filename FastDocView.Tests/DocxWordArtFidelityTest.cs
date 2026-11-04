using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// .docx WordArt (--fidelity, seconds): a DrawingML shape with a:prstTxWarp textArchUp (its text box text laid on the arch,
/// in the run's w14:textFill colour and font, not boxed) and a VML v:shape of type 136 (plain text) with v:textpath, in its fill
/// colour; both through the shared WordOfficeArtGeometry / WordOfficeArtFontwork. Search: the DrawingML text box text and the
/// VML textpath string are found once each (byte scanner and reference reader); a v:textpath in mc:Fallback is not.
/// </summary>
public static class DocxWordArtFidelityTest
{
    private const string Namespaces = "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
        + "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"";

    public static void Run()
    {
        TestReport.Step("WordArt: DrawingML prstTxWarp and VML textpath on the shared text-effect drawing");
        string drawingMl = $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"2743200\" cy=\"914400\"/><wp:docPr id=\"1\" name=\"w\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
            + "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2743200\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>"
            + "<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:rFonts w:ascii=\"Impact\" w:hAnsi=\"Impact\"/><w14:textFill xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w14:solidFill><w14:srgbClr w14:val=\"FF0000\"/></w14:solidFill></w14:textFill></w:rPr><w:t>Arched</w:t></w:r></w:p></w:txbxContent></wps:txbx>"
            + "<wps:bodyPr><a:prstTxWarp prst=\"textArchUp\"><a:avLst><a:gd name=\"adj\" fmla=\"val 10800000\"/></a:avLst></a:prstTxWarp></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        string vml = $"<w:p><w:r><w:pict {Namespaces}><v:shape type=\"#_x0000_t136\" style=\"width:216pt;height:72pt\" fillcolor=\"blue\" stroked=\"f\">"
            + "<v:textpath style=\"font-family:&quot;Arial Black&quot;\" string=\"VmlArt\"/></v:shape></w:pict></w:r></w:p>"
            + "<w:p><w:r><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"wps\"><w:t>choice</w:t></mc:Choice><mc:Fallback><w:pict " + Namespaces + "><v:shape type=\"#_x0000_t136\" style=\"width:72pt;height:36pt\"><v:textpath string=\"FallbackArt\"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>";
        // a footnote holding VML WordArt: its textpath string makes the referencing block a hit
        string footnotes = "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id=\"1\"><w:p><w:r><w:pict " + Namespaces + "><v:shape type=\"#_x0000_t136\" style=\"width:72pt;height:36pt\"><v:textpath string=\"NoteArt\"/></v:shape></w:pict></w:r></w:p></w:footnote>";
        vml += "<w:p><w:r><w:t>noted</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>";
        List<string> html = DocxFidelityTest.RenderAll("wordart", drawingMl + vml, footnotes: footnotes,
            inspect: document =>
            {
                foreach ((string text, int expected) in new[] { ("Arched", 1), ("VmlArt", 1), ("FallbackArt", 0), ("NoteArt", 1) })
                {
                    int hits = 0;
                    for (int from = 0; from < document.BlockCount;)
                    {
                        WordDocumentSearchBatch batch = document.Search(text, WordDocumentSearchOptions.Default, from, 4);
                        hits += batch.HitCounts.Sum();
                        if (batch.ReachedEnd) break;
                        from = batch.NextBlock;
                    }
                    TestReport.Check(hits == expected, "search: \"" + text + "\" found " + expected + " time(s) (" + hits + ")");
                    int referenceHits = 0;
                    if (document is FastDocView.Docx.DocxWindowedDocument docx)
                        for (int from = 0; from < document.BlockCount;)
                        {
                            WordDocumentSearchBatch batch = docx.Search(text, WordDocumentSearchOptions.Default, from, 4, FastDocView.Docx.DocxWindowedDocument.DocxSearchTextMode.ReferenceXmlReader);
                            referenceHits += batch.HitCounts.Sum();
                            if (batch.ReachedEnd) break;
                            from = batch.NextBlock;
                        }
                    TestReport.Check(referenceHits == expected, "reference reader: \"" + text + "\" found " + expected + " time(s) (" + referenceHits + ")");
                }
            });
        // textArchUp -> fontwork-arch-up-curve -> mso_sptTextArchUpCurve 144 (presetgeometrynames.cxx, EnhancedCustomShapeTypeNames.cxx, msdffdef.hxx)
        TestReport.Check(html[0].Contains("<textPath") && html[0].Contains(">Arched</textPath>") && html[0].Contains("fill=\"#FF0000\"") && html[0].Contains("font-family=\"Impact\"")
                         && Regex.Matches(html[0], "Arched").Count == 1 && !html[0].Contains("class=\"text-box word-shape-text\""), "DrawingML WordArt: the text on the arch in its run colour and font, no text box (" + html[0] + ")");
        TestReport.Check(html[1].Contains(">VmlArt</textPath>") && html[1].Contains("fill=\"blue\"") && html[1].Contains("font-family=\"Arial Black\""), "VML WordArt: plain text shape 136 in its fill colour (" + html[1] + ")");
    }
}
