using System.Text;
using FastDocView.Html;

namespace FastDocView.Tests;

/// <summary>
/// .odt / .fodt text boxes and drawing shapes: a hand-written .fodt with each object type, and the same document saved by
/// LibreOffice (headless conversion into TestOutput, never committed) - checked for placement, size, look and search.
/// </summary>
public static class OdtShapesFidelityTest
{
    private const string Namespaces =
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
      + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" "
      + "xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" "
      + "xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"";

    private const string Styles = """
        <office:styles>
          <draw:gradient draw:name="G1" draw:style="linear" draw:start-color="#ff0000" draw:end-color="#0000ff" draw:angle="0deg" draw:border="0%"/>
        </office:styles>
        <office:automatic-styles>
          <style:style style:name="Rtl" style:family="paragraph"><style:paragraph-properties style:writing-mode="rl-tb"/></style:style>
          <style:style style:name="BoxLeft" style:family="graphic"><style:graphic-properties style:wrap="right" style:horizontal-pos="left" style:horizontal-rel="paragraph" style:vertical-pos="top" style:vertical-rel="paragraph" fo:margin-left="0in" fo:margin-right="0.25in" fo:margin-top="0in" fo:margin-bottom="0.1in" fo:padding="0.1in" fo:border="0.03in solid #008000" draw:fill="solid" draw:fill-color="#ffff00"/></style:style>
          <style:style style:name="BoxCenter" style:family="graphic"><style:graphic-properties style:wrap="parallel" style:horizontal-pos="center" style:horizontal-rel="paragraph" fo:padding="0in" fo:border="none" draw:fill="none"/></style:style>
          <style:style style:name="BoxPlain" style:family="graphic"><style:graphic-properties style:wrap="dynamic" fo:padding="0in" fo:border="none" draw:fill="none"/></style:style>
          <style:style style:name="ShapeRed" style:family="graphic"><style:graphic-properties style:wrap="none" draw:fill="solid" draw:fill-color="#ff0000" draw:stroke="solid" svg:stroke-color="#000080" svg:stroke-width="0.03in" style:horizontal-pos="center" style:horizontal-rel="paragraph"/></style:style>
          <style:style style:name="ShapeGradient" style:family="graphic"><style:graphic-properties style:wrap="run-through" draw:fill="gradient" draw:fill-gradient-name="G1" draw:stroke="dash" svg:stroke-color="#000000" style:horizontal-pos="left" style:horizontal-rel="paragraph"/></style:style>
          <style:style style:name="ShapeNone" style:family="graphic"><style:graphic-properties style:wrap="none" draw:fill="none" draw:stroke="solid" svg:stroke-color="#00ff00" style:horizontal-pos="left" style:horizontal-rel="paragraph"/></style:style>
        </office:automatic-styles>
        """;

    private static string Shape(string type, string text = "")
        => "<draw:custom-shape draw:style-name=\"ShapeRed\" text:anchor-type=\"paragraph\" svg:width=\"1in\" svg:height=\"0.5in\" svg:x=\"0in\" svg:y=\"0in\">"
         + text + "<draw:enhanced-geometry svg:viewBox=\"0 0 21600 21600\" draw:type=\"" + type + "\"/></draw:custom-shape>";

    /// <summary>Blocks in order; the test checks each by its index.</summary>
    private static readonly string[] Body =
    [
        // 0 text box anchored to the paragraph, text right of it, min-height = grows
        "<text:p><draw:frame draw:style-name=\"BoxLeft\" text:anchor-type=\"paragraph\" svg:width=\"1.5in\" svg:height=\"0.5in\"><draw:text-box fo:min-height=\"0.5in\">"
            + "<text:p>boxtext one</text:p><text:p>second <text:hidden-text text:string-value=\"x\" text:is-hidden=\"true\">hiddenword</text:hidden-text>line</text:p></draw:text-box></draw:frame>outside text</text:p>",
        // 1 text box in the line
        "<text:p>before <draw:frame draw:style-name=\"BoxPlain\" text:anchor-type=\"as-char\" svg:width=\"1in\" svg:height=\"0.3in\"><draw:text-box><text:p>inlinebox</text:p></draw:text-box></draw:frame> after</text:p>",
        // 2 a centred box with text on both sides: above and below (CSS cannot wrap on both sides)
        "<text:p><draw:frame draw:style-name=\"BoxCenter\" text:anchor-type=\"char\" svg:width=\"1in\" svg:height=\"0.3in\"><draw:text-box><text:p>centrebox</text:p></draw:text-box></draw:frame>around</text:p>",
        // 3 right-to-left paragraph, box without a horizontal position: at the paragraph's start (right)
        "<text:p text:style-name=\"Rtl\"><draw:frame draw:style-name=\"BoxPlain\" text:anchor-type=\"paragraph\" svg:width=\"1in\" svg:height=\"0.3in\"><draw:text-box><text:p>rtlbox</text:p></draw:text-box></draw:frame>rtl text</text:p>",
        // 4 custom shapes
        "<text:p>" + Shape("rectangle", "<text:p>shapetext</text:p>") + "</text:p>",
        "<text:p>" + Shape("ellipse") + "</text:p>",
        "<text:p>" + Shape("round-rectangle") + "</text:p>",
        "<text:p>" + Shape("right-arrow") + "</text:p>",
        "<text:p>" + Shape("mso-spt999") + "</text:p>",
        // 9 line, gradient rect, group
        "<text:p><draw:line draw:style-name=\"ShapeNone\" text:anchor-type=\"paragraph\" svg:x1=\"0in\" svg:y1=\"0in\" svg:x2=\"2in\" svg:y2=\"0.5in\"/></text:p>",
        "<text:p><draw:rect draw:style-name=\"ShapeGradient\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"1in\"><text:p>gradtext</text:p></draw:rect></text:p>",
        "<text:p><draw:g text:anchor-type=\"paragraph\">"
            + "<draw:rect draw:style-name=\"ShapeNone\" svg:x=\"0.5in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"0.5in\"><text:p>groupone</text:p></draw:rect>"
            + "<draw:ellipse draw:style-name=\"ShapeRed\" svg:x=\"1in\" svg:y=\"0.25in\" svg:width=\"1in\" svg:height=\"0.5in\"/></draw:g></text:p>",
    ];

    public static void Run()
    {
        TestReport.Step("ODT shapes: hand-written .fodt");
        string fodt = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-shapes.fodt");
        string document = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document " + Namespaces + " office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
                        + Styles + "<office:body><office:text>" + string.Concat(Body) + "</office:text></office:body></office:document>";
        if (!File.Exists(fodt) || File.ReadAllText(fodt) != document) File.WriteAllText(fodt, document, new UTF8Encoding(false));   // unchanged: the conversion is reused
        CheckHandWritten(fodt);

        TestReport.Step("ODT shapes: the same document saved by LibreOffice");
        if (LibreOfficeConversion.Convert(fodt, ".odt", "odt") is { } odt) CheckLibreOfficeSaved(odt);
    }

    private static List<string> Render(string path)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }

    private static int SearchCount(string path, string query)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        WordDocumentSearchBatch batch = document.Search(query, new WordDocumentSearchOptions(), 0, document.BlockCount);
        return batch.HitCounts.Sum();
    }

    private static void CheckHandWritten(string path)
    {
        List<string> html = Render(path);
        TestReport.Check(html.Count == Body.Length, $"one block per paragraph ({html.Count})");
        string box = html[0];
        TestReport.Check(box.Contains("<span class=\"text-box\" style=\"box-sizing:border-box;overflow-wrap:break-word;width:108pt;min-height:36pt;float:left;margin:0pt 18pt 7.2pt 0pt;padding:7.2pt 7.2pt 7.2pt 7.2pt;border:2.15pt solid #008000;background:#ffff00;\">"),
                         "text box: float left (wrap right), its size, min-height grows, gaps, padding, border, fill (" + box + ")");
        TestReport.Check(box.Contains("<span class=\"text-box-paragraph\" dir=\"ltr\" style=\"display:block;\">boxtext one</span>") && box.IndexOf("outside text", StringComparison.Ordinal) > box.IndexOf("</span>", box.IndexOf("line", StringComparison.Ordinal), StringComparison.Ordinal),
                         "text box: its paragraphs as block spans, the anchoring paragraph's text after the box");
        TestReport.Check(!box.Contains("hiddenword") && !box.Contains("<p>boxtext"), "text box: hidden text not shown, no <p> inside the <p>");
        TestReport.Check(html[1].Contains("display:inline-block;vertical-align:bottom;width:72pt;height:21.6pt;overflow:hidden;") && html[1].IndexOf("before", StringComparison.Ordinal) < html[1].IndexOf("inlinebox", StringComparison.Ordinal),
                         "as-char text box: in the line at its fixed size (" + html[1] + ")");
        TestReport.Check(html[2].Contains("float:left;margin:0pt 0pt 0pt 0pt;") && !html[2].Contains("margin:0pt auto"), "centred parallel wrap: floats flush left, the text along its other side (" + html[2] + ")");
        TestReport.Check(html[3].Contains("float:right;") && html[3].StartsWith("<p dir=\"rtl\"", StringComparison.Ordinal), "right-to-left paragraph: the box floats to its start side, right (" + html[3] + ")");
        string rectangle = html[4];
        TestReport.Check(rectangle.Contains("<span class=\"word-shape\" data-shape-page=\"72 36\" style=\"position:relative;box-sizing:border-box;width:72pt;height:36pt;display:block;margin:0pt auto 0pt auto;\">"),
                         "custom shape: a frame of its size, centred, wrap none = above and below, marked for the page (it grows with its text: auto-grow default on) (" + rectangle + ")");
        // "rectangle" is drawn as a plain rectangle (EnhancedCustomShape2d.cxx:3042-3055); 0.03 in = 2.16 pt
        TestReport.Check(rectangle.Contains("<path d=\"M0 0L72 0L72 36L0 36Z\" fill-rule=\"evenodd\" fill=\"#ff0000\" stroke=\"none\"/>")
                         && rectangle.Contains("stroke=\"#000080\" stroke-width=\"2.16\" stroke-linejoin=\"round\""), "rectangle preset: fill and stroke from the graphic style, LibreOffice's round join");
        TestReport.Check(rectangle.Contains("class=\"text-box word-shape-text\"") && rectangle.Contains(">shapetext</span>"), "custom shape text in its text area");
        TestReport.Check(html[5].Contains("<path d=\"M72 18L71.83 19.76"), "ellipse preset: the full angle ellipse from 3 o'clock (msoEllipse)");
        // msoRoundRectangle stretches (nXRef / nYRef 10800): the 3600 / 21600 corner of the shorter side, 6 pt both ways
        TestReport.Check(html[6].Contains("<path d=\"M6 0L6 0L5.41 0.03"), "round-rectangle preset: circular corners of 6 pt (" + html[6] + ")");
        TestReport.Check(html[7].Contains("<path d=\"M0 9L54 9L54 0L72 18L54 36L54 27L0 27Z\""), "right-arrow preset at its default adjustments (" + html[7] + ")");
        // a type without a preset and without a path draws nothing (CreatePathObj :2890-2893); its frame keeps its size
        TestReport.Check(html[8].Contains("width:72pt;height:36pt;") && !html[8].Contains("<path"), "unknown geometry: nothing drawn, the frame keeps its size (" + html[8] + ")");
        TestReport.Check(html[9].Contains("<path d=\"M0 0L144 36\" fill=\"none\" stroke=\"#00ff00\" stroke-width=\"0.75\"") && html[9].Contains("width:144pt;height:36pt;"),
                         "line: its bounding box, a hairline (" + html[9] + ")");
        // draw:stroke dash without a draw:stroke-dash: the pool's XDash() - 0.2 mm parts, at least 26.95 hundredths of a mm = 0.76 pt
        TestReport.Check(html[10].Contains("<linearGradient") && html[10].Contains("stop-color=\"#ff0000\"") && html[10].Contains("stop-color=\"#0000ff\"")
                         && html[10].Contains("stroke-dasharray=\"0.76 0.76 0.76 0.76\"") && html[10].Contains("display:block;margin:0pt auto 0pt 0pt;"),
                         "linear gradient, the default dash, run-through = a block without gaps (" + html[10] + ")");
        TestReport.Check(html[11].Contains("width:108pt;height:54pt;") && html[11].Contains("<g transform=\"matrix(1 0 0 1 0 0)\">") && html[11].Contains("<g transform=\"matrix(1 0 0 1 36 18)\">")
                         && html[11].Contains(">groupone</span>"), "group: the children's bounding box, each child at its offset (" + html[11] + ")");
        CheckSearch(path);
    }

    private static void CheckSearch(string path, int hiddenExpected = 0)
    {
        foreach ((string query, int expected) in new[] { ("boxtext", 1), ("inlinebox", 1), ("shapetext", 1), ("gradtext", 1), ("groupone", 1), ("hiddenword", hiddenExpected), ("outside text", 1) })
        {
            int count = SearchCount(path, query);
            TestReport.Check(count == expected, $"{Path.GetExtension(path)} search \"{query}\": {expected} hit(s) ({count})");
        }
    }

    /// <summary>LibreOffice's own writing of the same objects: the boxes and shapes keep their placement, text and search.</summary>
    private static void CheckLibreOfficeSaved(string path)
    {
        List<string> html = Render(path);
        string all = string.Concat(html);
        TestReport.Check(html.Any(block => block.Contains("class=\"text-box\"") && block.Contains("float:left;") && block.Contains("boxtext one")), "LibreOffice .odt: the text box floats left with its text");
        TestReport.Check(all.Contains("inlinebox") && all.Contains("display:inline-block;vertical-align:bottom;"), "LibreOffice .odt: the as-char box in the line");
        // LibreOffice writes an explicit horizontal-pos from-left svg:x 0 for the right-to-left box: physical left (ODF 20.298)
        TestReport.Check(html.Any(block => block.Contains("rtlbox") && block.Contains("float:left;")), "LibreOffice .odt: the right-to-left box at its explicit from-left position");
        TestReport.Check(all.Contains("<path d=\"M72") && all.Contains("<path d=\"M0 9") && all.Contains("groupone") && html.Count(block => block.Contains("class=\"word-shape\"")) >= 8,
                         "LibreOffice .odt: ellipse, arrow, line, group drawn from the saved enhanced paths");
        // LibreOffice 25.x saves the text:hidden-text of the source as plain text (shown): the search must find it exactly once
        TestReport.Check(all.Contains("hiddenword"), "LibreOffice .odt: the former hidden text is plain text there");
        TestReport.Check(all.Contains("min-height:36pt;") || all.Contains("min-height:41.47pt;"), "LibreOffice .odt: fo:min-height from the graphic style grows the box");
        CheckSearch(path, hiddenExpected: 1);
    }
}
