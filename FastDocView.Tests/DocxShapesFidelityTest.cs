using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// .docx drawn shapes (--fidelity, seconds): wps presets / custGeom / fills / theme colours / lines and arrow ends, a wpg group
/// with child-space scaling, nested group, picture and text box, a shape with text, VML shapes with text boxes, and
/// mc:AlternateContent (Choice drawn, Fallback skipped): every shape text shows and is found exactly once.
/// </summary>
public static class DocxShapesFidelityTest
{
    private const string Namespaces = "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" "
        + "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
        + "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\"";

    private const string Theme = """
        <?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>
        <a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
        <a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
        <a:fmtScheme name="T"><a:fillStyleLst/><a:lnStyleLst><a:ln w="6350"/><a:ln w="12700"/><a:ln w="19050"/></a:lnStyleLst></a:fmtScheme>
        </a:themeElements></a:theme>
        """;

    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEElEQVR4nGP4z8DAwMDAAAAJ/gH/0wkYjQAAAABJRU5ErkJggg==");

    public static void Run()
    {
        TestReport.Step("Shapes: .docx wps presets, fills, lines, groups, VML, alternate content");
        string body = Inline("rightArrow", "<a:solidFill><a:schemeClr val=\"accent1\"><a:lumMod val=\"75000\"/></a:schemeClr></a:solidFill><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln>")
            + Inline("ellipse", "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"accent6\"><a:tint val=\"50000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\"/></a:gradFill>")
            + Inline("straightConnector1", "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill><a:prstDash val=\"dash\"/><a:tailEnd type=\"triangle\"/></a:ln>")
            + Inline("star5", "<a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill>")
            + Inline("rect", "<a:noFill/>", rotation: " rot=\"2700000\" flipH=\"1\"", style: "<wps:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent2\"/></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:effectRef><a:fontRef idx=\"minor\"/></wps:style>")
            + InlineGeometry("<a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"100\"/></a:moveTo><a:lnTo><a:pt x=\"50\" y=\"0\"/></a:lnTo><a:lnTo><a:pt x=\"100\" y=\"100\"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val=\"123456\"/></a:solidFill>")
            + Inline("ellipse", "<a:solidFill><a:srgbClr val=\"FFFF00\"/></a:solidFill>", text: "ovaltext", bodyPr: "<wps:bodyPr lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\"><a:spAutoFit/></wps:bodyPr>")
            + Group()
            + Vml()
            + Alternate();
        List<string> html = DocxFidelityTest.RenderAll("shapes", body,
            extraDocumentRelationships: "<Relationship Id=\"rIdTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
                                      + "<Relationship Id=\"rIdP\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/p.png\"/>",
            binaryParts: new Dictionary<string, byte[]> { ["word/theme/theme1.xml"] = new UTF8Encoding(false).GetBytes(Theme), ["word/media/p.png"] = Png },
            inspect: document =>
            {
                foreach (string text in new[] { "ovaltext", "grouptext", "nestedtext", "vmltext", "choicetext" })
                {
                    int hits = 0;
                    for (int from = 0; from < document.BlockCount;)
                    {
                        WordDocumentSearchBatch batch = document.Search(text, WordDocumentSearchOptions.Default, from, 4);
                        hits += batch.HitCounts.Sum();
                        if (batch.ReachedEnd) break;
                        from = batch.NextBlock;
                    }
                    TestReport.Check(hits == 1, "search: \"" + text + "\" found exactly once (" + hits + ")");
                }
            });
        // lumMod 75000 of 4472C4 in HSL luminance as LibreOffice color.cxx: 2F5597 (Word's "accent 1, darker 25%")
        TestReport.Check(html[0].Contains("<svg class=\"word-shape-geometry\"") && html[0].Contains("width:72pt;height:36pt;") && html[0].Contains("fill=\"#2F5597\" stroke=\"#FF0000\" stroke-width=\"1\"")
                         && html[0].Contains("<path d=\"M0 9L54 9L54 0L72 18L54 36L54 27L0 27Z\""),
                         "rightArrow: SVG at wp:extent, schemeClr accent1 lumMod 75% = #2F5597, red 1pt line, preset outline (" + html[0] + ")");
        TestReport.Check(html[1].Contains("<linearGradient") && html[1].Contains("x1=\"0.5\" y1=\"0\" x2=\"0.5\" y2=\"1\"") && html[1].Contains("stop-color=\"#FF0000\"") && html[1].Contains("<ellipse"),
                         "ellipse: gradFill as a top-to-bottom linear gradient (ang 90 deg) (" + html[1] + ")");
        TestReport.Check(html[2].Contains("stroke=\"#0000FF\" stroke-width=\"2\" stroke-dasharray=\"6 2\"") && html[2].Contains("<path d=\"M72 36L") && html[2].Contains("fill=\"#0000FF\" stroke=\"none\""),
                         "connector: dashed 2pt blue line with a filled arrow head at its tail end (" + html[2] + ")");
        TestReport.Check(html[3].Contains("<rect width=\"72\" height=\"36\" fill=\"#00FF00\""), "unknown preset (star5): a box of the right size (" + html[3] + ")");
        TestReport.Check(html[4].Contains("rotate(45) scale(-1 1)") && html[4].Contains("fill=\"none\" stroke=\"#ED7D31\" stroke-width=\"1\""),
                         "rect: rot 45 deg + flipH; spPr noFill over fillRef; lnRef idx 2 = accent2 at the theme's 1pt (" + html[4] + ")");
        TestReport.Check(html[5].Contains("<path d=\"M0 36L36 0L72 36Z\" fill=\"#123456\""), "custGeom: the path scaled from its w / h to the shape (" + html[5] + ")");
        TestReport.Check(html[6].Contains("min-height:36pt;") && html[6].Contains("vector-effect=\"non-scaling-stroke\"") && html[6].Contains("word-shape-text") && Count(html[6], "ovaltext") == 1,
                         "ellipse with growing text: the frame grows, the outline stretches, the text once (" + html[6] + ")");
        string group = html[7];
        TestReport.Check(group.Contains("<rect width=\"36\" height=\"18\"") && group.Contains("translate(54 9)") && group.Contains("translate(18 27)"), "group: a child at chOff scaled by ext / chExt (half size) (" + group + ")");
        TestReport.Check(group.Contains("position:absolute;left:36pt;top:0pt;width:36pt;height:18pt;") && Count(group, "grouptext") == 1, "group: the child's text box over its outline");
        TestReport.Check(group.Contains("left:0pt;top:18pt;") && Count(group, "nestedtext") == 1, "nested group: its child scaled through both groups");
        TestReport.Check(group.Contains("data-word-picture=\"m:rIdP\"") && group.Contains("left:36pt;top:18pt;"), "group picture: by id at its box");
        string vml = html[8];
        TestReport.Check(vml.Contains("float:right;") && vml.Contains("width:144pt;height:72pt;") && vml.Contains("rx=\"") && vml.Contains("fill=\"yellow\" stroke=\"#FF0000\" stroke-width=\"2\"")
                         && vml.Contains("padding:1pt 2pt 3pt 4pt;") && Count(vml, "vmltext") == 1,
                         "VML roundrect: style size, square wrap right, fillcolor / strokecolor / strokeweight, v:textbox inset (left,top,right,bottom) (" + vml + ")");
        TestReport.Check(html[9].Contains("<ellipse") && Count(html[9], "choicetext") == 1 && !html[9].Contains("fallbacktext"), "AlternateContent: the Choice drawn, the VML Fallback not (" + html[9] + ")");
    }

    private static int Count(string html, string text) => Regex.Matches(html, text).Count;

    private static string Inline(string preset, string fill, string rotation = "", string style = "", string? text = null, string bodyPr = "<wps:bodyPr/>")
        => InlineGeometry($"<a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{fill}", rotation, style, text, bodyPr);

    private static string InlineGeometry(string geometryAndFill, string rotation = "", string style = "", string? text = null, string bodyPr = "<wps:bodyPr/>")
        => $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"s\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
         + $"<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm{rotation}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm>{geometryAndFill}<a:effectLst/><a:extLst><a:ext uri=\"x\"><big/></a:ext></a:extLst></wps:spPr>{style}"
         + (text is null ? "" : $"<wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>") + bodyPr
         + "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    /// <summary>A 72 x 36 pt group whose child space is twice its size: a text box at the right top, a nested group bottom left, a picture bottom right.</summary>
    private static string Group()
    {
        static string Xfrm(string inner) => $"<a:xfrm>{inner}</a:xfrm>";
        string textShape = "<wps:wsp><wps:cNvSpPr/><wps:spPr>" + Xfrm("<a:off x=\"1914400\" y=\"1000000\"/><a:ext cx=\"914400\" cy=\"457200\"/>")
            + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></wps:spPr>"
            + "<wps:txbx><w:txbxContent><w:p><w:r><w:t>grouptext</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>";
        string nested = "<wpg:grpSp><wpg:cNvGrpSpPr/><wpg:grpSpPr>" + Xfrm("<a:off x=\"1000000\" y=\"1457200\"/><a:ext cx=\"914400\" cy=\"457200\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"914400\" cy=\"457200\"/>") + "</wpg:grpSpPr>"
            + "<wps:wsp><wps:cNvSpPr/><wps:spPr>" + Xfrm("<a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/>") + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>"
            + "<wps:txbx><w:txbxContent><w:p><w:r><w:t>nestedtext</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></wpg:grpSp>";
        string picture = "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"3\" name=\"p\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rIdP\"/></pic:blipFill><pic:spPr>"
            + Xfrm("<a:off x=\"1914400\" y=\"1457200\"/><a:ext cx=\"914400\" cy=\"457200\"/>") + "</pic:spPr></pic:pic>";
        return $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"2\" name=\"g\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\">"
             + "<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>" + Xfrm("<a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/><a:chOff x=\"1000000\" y=\"1000000\"/><a:chExt cx=\"1828800\" cy=\"914400\"/>") + "</wpg:grpSpPr>"
             + textShape + nested + picture + "</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
    }

    private static string Vml()
        => $"<w:p><w:r><w:pict {Namespaces}><v:roundrect style=\"position:absolute;margin-left:0;margin-top:0;width:2in;height:1in;mso-position-horizontal:right;z-index:1\" arcsize=\"10923f\" fillcolor=\"yellow\" strokecolor=\"#f00\" strokeweight=\"2pt\">"
         + "<v:textbox inset=\"4pt,1pt,2pt,3pt\"><w:txbxContent><w:p><w:r><w:t>vmltext</w:t></w:r></w:p></w:txbxContent></v:textbox><w10:wrap type=\"square\"/></v:roundrect></w:pict></w:r><w:r><w:t>after</w:t></w:r></w:p>";

    private static string Alternate()
        => $"<w:p><w:r><mc:AlternateContent {Namespaces}><mc:Choice Requires=\"wps\"><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"4\" name=\"c\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
         + "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=\"00FFFF\"/></a:solidFill></wps:spPr>"
         + "<wps:txbx><w:txbxContent><w:p><w:r><w:t>choicetext</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>"
         + "<mc:Fallback><w:pict><v:oval style=\"width:1in;height:.5in\"><v:textbox><w:txbxContent><w:p><w:r><w:t>fallbacktext</w:t></w:r></w:p></w:txbxContent></v:textbox></v:oval></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>";
}
