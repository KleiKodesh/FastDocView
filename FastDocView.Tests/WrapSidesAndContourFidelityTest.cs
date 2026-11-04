using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Two-sided and contour wrap in every format (WordPictureHtml.AppendPlacement, WordPictureContour): text on both sides of an
/// object floats it flush to its side (never text above and below), a one-sided wrap floats it opposite its text, a tight /
/// through wrap becomes shape-outside:polygon from the wrap polygon or, without one, the shape's outline. A hand-written .rtf
/// (synthetic) and the LibreOffice test files picture-wrap-polygon.docx / .rtf and tdf128501.doc (a tight / through wrap).
/// </summary>
public static class WrapSidesAndContourFidelityTest
{
    private static string Shape(string words, string properties, string text = "")
        => @"{\shp{\*\shpinst" + words + properties + (text.Length > 0 ? @"{\shptxt \pard\plain " + text + "}" : "") + @"}{\shprslt {\pard RESULTCOPY\par}}}";

    private static string Property(string name, string value) => @"{\sp{\sn " + name + @"}{\sv " + value + "}}";

    private static readonly string[] RtfBody =
    [
        // 0 centred, text on both sides: floats left, no centred block
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440\shpwr2\shpwrk0", Property("shapeType", "1") + Property("posh", "2")) + @"both\par",
        // 1 placed right, text on both sides: floats right
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440\shpwr2\shpwrk0", Property("shapeType", "1") + Property("posh", "3")) + @"right\par",
        // 2 text only on the left: floats right
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440\shpwr2\shpwrk1", Property("shapeType", "1")) + @"leftonly\par",
        // 3 tight with a wrap polygon (a triangle, apex top-left): shape-outside from it
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440\shpwr4\shpwrk0", Property("shapeType", "1") + Property("pWrapPolygonVertices", "8;3;(0,0);(21600,21600);(0,21600)")) + @"tight\par",
        // 4 tight without a polygon, an ellipse: its outline
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440\shpwr4\shpwrk0", Property("shapeType", "3")) + @"ellipse\par",
        // 5 through with a wrap polygon: Word's through is tight plus holes - a contour too
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440\shpwr5\shpwrk0", Property("shapeType", "1") + Property("pWrapPolygonVertices", "8;3;(0,0);(21600,21600);(0,21600)")) + @"through\par",
        // 6 placed by offset alone, centre right of the column's (default 12240 - 2 x 1800 = 8640 twips): floats right
        @"\pard " + Shape(@"\shpleft6000\shptop0\shpright7440\shpbottom1440\shpwr2\shpwrk0", Property("shapeType", "1")) + @"nearright\par",
        // 7 placed by offset alone near the left: floats left
        @"\pard " + Shape(@"\shpleft500\shptop0\shpright1940\shpbottom1440\shpwr2\shpwrk0", Property("shapeType", "1")) + @"nearleft\par",
        // 8 the same offset as 6 but from the page edge (\shpbxpage): it includes the margin - no nearer-edge choice, left
        @"\pard " + Shape(@"\shpbxpage\shpleft6000\shptop0\shpright7440\shpbottom1440\shpwr2\shpwrk0", Property("shapeType", "1")) + @"pageleft\par",
    ];

    public static void Run()
    {
        TestReport.Step("Two-sided and contour wrap: .rtf (synthetic + LibreOffice), .docx, .doc");
        string rtf = Path.Combine(TestReport.OutputDirectory, "fidelity-v3-rtf-wrap-sides-contour.rtf");
        string document = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}" + string.Join("", RtfBody) + "}";
        if (!File.Exists(rtf) || File.ReadAllText(rtf) != document) File.WriteAllText(rtf, document, new UTF8Encoding(false));
        List<string> html = Render(rtf);
        TestReport.Check(html.Count >= 9, $"rtf: one block per paragraph ({html.Count})");
        if (html.Count >= 9)
        {
            TestReport.Check(html[0].Contains("float:left;") && !html[0].Contains("display:block;margin"), "rtf centred both sides: floats flush left (" + html[0] + ")");
            TestReport.Check(html[1].Contains("float:right;"), "rtf right both sides: floats right (" + html[1] + ")");
            TestReport.Check(html[2].Contains("float:right;"), "rtf \\shpwrk1 (text on the left): floats right (" + html[2] + ")");
            // text on the right of a left float: the polygon runs from the left edge along the triangle's right side x = y
            TestReport.Check(html[3].Contains("shape-outside:polygon(0pt 0pt,0pt 0pt,1pt 1pt,") && html[3].Contains("72pt 72pt,0pt 72pt) border-box;"),
                             "rtf \\shpwr4 + pWrapPolygonVertices: shape-outside from the polygon (" + html[3] + ")");
            TestReport.Check(html[5].Contains("shape-outside:polygon(0pt 0pt,0pt 0pt,1pt 1pt,"), "rtf \\shpwr5 (through): a contour like tight (" + html[5] + ")");
            TestReport.Check(html[6].Contains("float:right;"), "rtf offset alone, centre right of the column's: floats right (" + html[6] + ")");
            TestReport.Check(html[8].Contains("float:left;"), "rtf page-relative offset: floats left (" + html[8] + ")");
            TestReport.Check(html[7].Contains("float:left;"), "rtf offset alone near the left: floats left (" + html[7] + ")");
            TestReport.Check(html[4].Contains("shape-outside:polygon(") && html[4].Contains(",72pt 36pt,"), "rtf \\shpwr4 ellipse, no polygon: its outline (widest at mid-height) (" + html[4] + ")");
        }
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "libreoffice", "sw", "qa", "extras");
        foreach (string file in new[] { Path.Combine(folder, "ooxmlexport", "data", "picture-wrap-polygon.docx"), Path.Combine(folder, "rtfexport", "data", "picture-wrap-polygon.rtf") })
        {
            if (!File.Exists(file)) { TestReport.Check(false, $"{Path.GetFileName(file)} present"); continue; }
            TestReport.Check(Render(file).Exists(b => b.Contains("shape-outside:polygon(", StringComparison.Ordinal)), $"{Path.GetFileName(file)}: its wrap polygon as shape-outside");
        }
        // .doc: FSPA wr 4 / 5 along pWrapPolygonVertices or the outline (LibreOffice's tdf128501.doc has one)
        string doc = Path.Combine(folder, "ww8export", "data", "tdf128501.doc");
        TestReport.Check(File.Exists(doc) && Render(doc).Exists(b => b.Contains("shape-outside:polygon(", StringComparison.Ordinal)), "tdf128501.doc: a tight / through wrap drawn as shape-outside");
    }

    private static List<string> Render(string path)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        var blocks = new List<string>();
        for (int first = 0; first < document.BlockCount; first += 100)
            foreach (WordDocumentBlockHtml block in document.RenderBlocks(first, 100)) blocks.Add(block.Html);
        return blocks;
    }
}
