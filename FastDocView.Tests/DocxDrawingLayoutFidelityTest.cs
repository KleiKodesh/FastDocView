using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// .docx drawing layout (--fidelity, seconds), expected values worked out by hand from the LibreOffice source each check
/// cites: the preset text rectangle, rectangular gradients, pattern (hatch) and picture fills, group rotation and flip with
/// MSO's width / height swap, drawing canvases (wpc), and VML groups with their coordinate systems. Hand-written XML in the
/// shape Word writes.
/// </summary>
public static class DocxDrawingLayoutFidelityTest
{
    private const string Namespaces = "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" "
        + "xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"";

    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEElEQVR4nGP4z8DAwMDAAAAJ/gH/0wkYjQAAAABJRU5ErkJggg==");

    public static void Run()
    {
        TestReport.Step("Drawing layout: text rectangles, rect gradients, hatches, picture fills, group turns, canvases, VML groups");
        string Text(string text) => $"<wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>";
        var body = new List<string>
        {
            /* 0 */ Shape("ellipse", "<a:solidFill><a:srgbClr val=\"FFFF00\"/></a:solidFill>", text: Text("ellipsetext"), bodyPr: "<wps:bodyPr lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\"/>"),
            /* 1 */ Shape("rect", "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst>"
                             + "<a:path path=\"rect\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill>"),
            /* 2 */ Shape("rect", "<a:pattFill prst=\"horz\"><a:fgClr><a:srgbClr val=\"000000\"/></a:fgClr><a:bgClr><a:srgbClr val=\"FFFFFF\"/></a:bgClr></a:pattFill>"),
            /* 3 */ Shape("ellipse", "<a:blipFill><a:blip r:embed=\"rIdP\"><a:alphaModFix amt=\"50000\"/></a:blip><a:stretch><a:fillRect l=\"10000\"/></a:stretch></a:blipFill><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>"),
            /* 4 */ Group("<a:xfrm rot=\"5400000\" flipH=\"1\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"914400\" cy=\"457200\"/></a:xfrm>",
                           "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"457200\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
                           + "<a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></wps:spPr>" + Text("turnedtext") + "<wps:bodyPr/></wps:wsp>"),
            /* 5 */ Group("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"1828800\" cy=\"457200\"/></a:xfrm>",
                           "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm rot=\"5400000\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
                           + "<a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>"),
            /* 6 */ Canvas(),
            /* 7 */ VmlGroup(),
        };
        List<string> html = DocxFidelityTest.RenderAll("drawing-layout", string.Concat(body),
            extraDocumentRelationships: "<Relationship Id=\"rIdP\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/p.png\"/>",
            binaryParts: new Dictionary<string, byte[]> { ["word/media/p.png"] = Png },
            inspect: document =>
            {
                foreach (string text in new[] { "ellipsetext", "turnedtext", "canvastext", "vmlgrouptext" })
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

        // ellipse a:rect il / it / ir / ib = hc -+ wd2 cos 45, vc -+ hd2 sin 45 (presetShapeDefinitions.xml ellipse; EnhancedCustomShape2d::GetTextRect)
        TestReport.Check(html[0].Contains("left:10.54pt;top:5.27pt;width:50.91pt;height:25.46pt;") && Count(html[0], "ellipsetext") == 1,
                         "ellipse text in the preset's text rectangle (" + html[0] + ")");
        // rect gradient centred: the top region from its inner edge (pos 0) to the top edge (pos 1); m = 36 / 2
        TestReport.Check(html[1].Contains("<path d=\"M0 0L72 0L54 18L18 18Z\"") && html[1].Contains("x1=\"0\" y1=\"18\" x2=\"0\" y2=\"0\"") && html[1].Contains("fill=\"#0000FF\""),
                         "a:path rect: LibreOffice's rectangular gradient, the last stop outside the inner regions (" + html[1] + ")");
        // horz = SINGLE, 1 mm, 0 degrees (hatchmap.hxx): 15 lines each side of the centre over the half diagonal 40.25
        TestReport.Check(html[2].Contains("fill=\"#FFFFFF\"") && html[2].Contains("M-4.25 -24.52L76.25 -24.52") && html[2].Contains("stroke=\"#000000\" stroke-width=\"0.75\""),
                         "pattFill horz: black 1 mm hatch over the white background (" + html[2] + ")");
        TestReport.Check(html[3].Contains("class=\"word-shape-fill-picture\"") && html[3].Contains("data-word-picture=\"m:rIdP\"") && html[3].Contains("left:7.2pt;top:0pt;width:64.8pt;height:36pt;")
                         && html[3].Contains("opacity:0.5;") && html[3].Contains("clipPathUnits=\"objectBoundingBox\"") && Regex.Matches(html[3], "<svg class=\"word-shape-geometry\"").Count == 2,
                         "blipFill: the picture stretched into the box less fillRect l 10 %, half transparent, clipped, between two SVG layers (" + html[3] + ")");
        // group rot 90 + flipH (shape.cxx:1394-1493): child centre (18, 18) mirrored to (54, 18), turned about (36, 18) to (36, 36); rotation -0 + 90; flipH carried
        TestReport.Check(html[4].Contains("translate(36 36) rotate(90) scale(-1 1)") && html[4].Contains("left:18pt;top:18pt;width:36pt;height:36pt;") && html[4].Contains("transform:rotate(90deg);"),
                         "group turned and flipped: the child and its text box carried along (" + html[4] + ")");
        // MSO width / height swap for a child turned 90 in a 1:2 scaled group (shape.cxx:1456-1473): 72 x 18, centre (18, 18)
        TestReport.Check(html[5].Contains("translate(18 18) rotate(90)") && html[5].Contains("<path d=\"M0 0L72 0L72 18L0 18Z\""),
                         "child turned 90 degrees in a scaled group: MSO's swapped scaling (" + html[5] + ")");
        TestReport.Check(html[6].Contains("<path d=\"M0 0L72 0L72 36L0 36Z\" fill-rule=\"evenodd\" fill=\"#FFFF00\"") && html[6].Contains("translate(54 18)") && Count(html[6], "canvastext") == 1,
                         "wpc canvas: background from wpc:bg, a child at its EMU offset (" + html[6] + ")");
        // VML group coordsize 200,100 on 144 x 72 pt: rect 100,0 100x50 -> (72, 0) 72 x 36; nested group 0,0 100x50 coordsize 10,10: rect 5,5 5x5 -> (36, 18) 36 x 18
        string vml = html[7];
        TestReport.Check(vml.Contains("translate(108 18)") && vml.Contains("fill=\"red\"") && vml.Contains("translate(54 27)") && vml.Contains("translate(72 36)")
                         && vml.Contains("left:0pt;top:36pt;width:72pt;height:36pt;") && Count(vml, "vmlgrouptext") == 1,
                         "v:group: children in its coordinate system, a nested group, a line, the oval's text box (" + vml + ")");
    }

    private static int Count(string html, string text) => Regex.Matches(html, text).Count;

    private static string Shape(string preset, string fill, string text = "", string bodyPr = "<wps:bodyPr/>")
        => $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"s\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
         + $"<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{fill}</wps:spPr>{text}{bodyPr}"
         + "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    private static string Group(string transform, string children)
        => $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"2\" name=\"g\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\">"
         + $"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>{transform}</wpg:grpSpPr>{children}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    private static string Canvas()
        => $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"3\" name=\"c\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\">"
         + "<wpc:wpc><wpc:bg><a:solidFill><a:srgbClr val=\"FFFF00\"/></a:solidFill></wpc:bg><wpc:whole><a:ln><a:noFill/></a:ln></wpc:whole>"
         + "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"457200\" y=\"0\"/><a:ext cx=\"457200\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
         + "<a:solidFill><a:srgbClr val=\"FF00FF\"/></a:solidFill></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>canvastext</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>"
         + "</wpc:wpc></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    private static string VmlGroup()
        => $"<w:p><w:r><w:pict {Namespaces}><v:group style=\"width:144pt;height:72pt\" coordorigin=\"0,0\" coordsize=\"200,100\">"
         + "<v:rect style=\"position:absolute;left:100;top:0;width:100;height:50\" fillcolor=\"red\"/>"
         + "<v:oval style=\"position:absolute;left:0;top:50;width:100;height:50\"><v:textbox inset=\"0,0,0,0\"><w:txbxContent><w:p><w:r><w:t>vmlgrouptext</w:t></w:r></w:p></w:txbxContent></v:textbox></v:oval>"
         + "<v:group style=\"position:absolute;left:0;top:0;width:100;height:50\" coordsize=\"10,10\"><v:rect style=\"position:absolute;left:5;top:5;width:5;height:5\" fillcolor=\"blue\"/></v:group>"
         + "<v:line from=\"0,0\" to=\"200,100\" strokecolor=\"black\"/>"
         + "</v:group></w:pict></w:r></w:p>";
}
