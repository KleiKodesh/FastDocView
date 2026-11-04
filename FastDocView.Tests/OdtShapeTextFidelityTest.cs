using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// OpenDocument shape text the page finishes with its own text metrics (2026-11-04): custom shapes that grow with their text
/// (draw:auto-grow-height, LibreOffice svdoashp.cxx:2179-2194 / 2260-2464) carry their text area, shape size, distances,
/// minimum and alignment for the page; dimension lines carry their layout inputs. The page's part is checked live by
/// vue-demo/playwright/shapeTextCheck.mjs on the fixture written here (fidelity-v2-odt-shape-text.fodt).
/// </summary>
public static class OdtShapeTextFidelityTest
{
    private const string Namespaces =
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
      + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" "
      + "xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\"";

    private const string Styles = """
        <office:styles><draw:marker draw:name="Arrow" svg:viewBox="0 0 20 30" svg:d="M10 0l-10 30h20z"/></office:styles>
        <office:automatic-styles>
          <style:style style:name="Grow" style:family="graphic"><style:graphic-properties style:wrap="none" style:horizontal-pos="left" style:horizontal-rel="paragraph" draw:fill="solid" draw:fill-color="#ffcc00" draw:stroke="solid" svg:stroke-color="#000000" svg:stroke-width="0.01in" draw:auto-grow-height="true" draw:textarea-vertical-align="top" fo:padding-left="0.1in" fo:padding-right="0.1in" fo:padding-top="0.05in" fo:padding-bottom="0.05in"/></style:style>
          <style:style style:name="GrowUp" style:family="graphic" style:parent-style-name="Grow"><style:graphic-properties draw:textarea-vertical-align="bottom"/></style:style>
          <style:style style:name="GrowWide" style:family="graphic" style:parent-style-name="Grow"><style:graphic-properties fo:wrap-option="no-wrap" draw:textarea-horizontal-align="left"/></style:style>
          <style:style style:name="Shrink" style:family="graphic" style:parent-style-name="Grow"><style:graphic-properties draw:textarea-vertical-align="middle" fo:min-height="0.5in"/></style:style>
          <style:style style:name="Fixed" style:family="graphic" style:parent-style-name="Grow"><style:graphic-properties draw:auto-grow-height="false"/></style:style>
          <style:style style:name="NoAttribute" style:family="graphic"><style:graphic-properties style:wrap="none" style:horizontal-pos="left" style:horizontal-rel="paragraph" draw:fill="solid" draw:fill-color="#99ccff" draw:stroke="solid" svg:stroke-color="#000000"/></style:style>
          <style:style style:name="Contour" style:family="graphic" style:parent-style-name="Grow"><style:graphic-properties style:wrap="right" style:wrap-contour="true" style:wrap-contour-mode="outside"/></style:style>
          <style:style style:name="Arrows" style:family="graphic" style:parent-style-name="Measure"><style:graphic-properties draw:marker-start="Arrow" draw:marker-start-width="0.15in" draw:marker-end="Arrow" draw:marker-end-width="0.15in"/></style:style>
          <style:style style:name="Measure" style:family="graphic"><style:graphic-properties style:wrap="none" style:horizontal-pos="left" draw:stroke="solid" svg:stroke-color="#000000" svg:stroke-width="0in"/></style:style>
        </office:automatic-styles>
        """;

    private const string Long = "growtext one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen";

    private const string At = "svg:x=\"0.5in\" svg:y=\"0.5in\"";

    private static string Shape(string style, string text, string size = "svg:width=\"1.5in\" svg:height=\"0.5in\"", string place = At, string geometry = "")
        => "<text:p>" + Custom(style, text, size, place, geometry) + "</text:p>";

    private static string Custom(string style, string text, string size = "svg:width=\"1.5in\" svg:height=\"0.5in\"", string place = At, string geometry = "")
        => "<draw:custom-shape draw:style-name=\"" + style + "\" text:anchor-type=\"paragraph\" " + size + " " + place + "><text:p>" + text
         + "</text:p><draw:enhanced-geometry svg:viewBox=\"0 0 21600 21600\" draw:type=\"rectangle\" " + geometry + "draw:enhanced-path=\"M 0 0 L 21600 0 21600 21600 0 21600 Z N\"/></draw:custom-shape>";

    private static string Measure(string x2, string text, string style = "Measure")
        => "<text:p><draw:measure draw:style-name=\"" + style + "\" text:anchor-type=\"paragraph\" svg:x1=\"0.5in\" svg:y1=\"1in\" svg:x2=\"" + x2 + "\" svg:y2=\"1in\"><text:p>" + text + "</text:p></draw:measure></text:p>";

    /// <summary>Blocks in order; the checks (here and in the live check) name each by its index.</summary>
    internal static readonly string[] Body =
    [
        // 0 grows down: much text in a 1.5 x 0.5 in shape, top-aligned
        Shape("Grow", Long),
        // 1 grows up: bottom-aligned
        Shape("GrowUp", Long),
        // 2 no wrap: grows to the right (left-aligned) and in height
        Shape("GrowWide", "widetext on one long line that does not wrap"),
        // 3 shrinks to the minimum (0.5 in) plus distances, both ways (middle)
        Shape("Shrink", "shrinktext", "svg:width=\"1.5in\" svg:height=\"2in\""),
        // 4 auto-grow off: keeps its size
        Shape("Fixed", Long),
        // 5 a dimension line with a long label: wider than the line, so outside on the left
        Measure("1in", "a long measure label wider than its line"),
        // 6 a short label: centred above the line
        Measure("3in", "short"),
        // 7 turned 90 degrees (counter-clockwise), grows down in its own box: to the right on the page
        Shape("Grow", Long, place: "draw:transform=\"rotate (1.5707963267949) translate (0.5in 2in)\""),
        // 8 mirrored horizontally
        Shape("Grow", Long, geometry: "draw:mirror-horizontal=\"true\" "),
        // 9 a group: a fixed shape and a growing one under it
        "<text:p><draw:g>" + Custom("Fixed", "groupfixed", place: "svg:x=\"0.5in\" svg:y=\"0in\"") + Custom("Grow", Long) + "</draw:g></text:p>",
        // 10 text turned by draw:text-rotate-angle 90: the text runs along the tall side
        Shape("Grow", Long, "svg:width=\"0.5in\" svg:height=\"1.5in\"", geometry: "draw:text-rotate-angle=\"90\" "),
        // 11 draw:auto-grow-height absent: LibreOffice's default (on)
        Shape("NoAttribute", Long),
        // 12 a contour-wrapped growing shape with text beside it
        "<text:p>" + Custom("Contour", Long, place: "svg:x=\"0in\" svg:y=\"0in\"") + "beside the contour shape</text:p>",
        // 13 a dimension line with arrows at both ends
        Measure("3in", "arrows", "Arrows"),
        "<text:p>last</text:p>",
    ];

    public static void Run()
    {
        TestReport.Step("ODT shape text: growing shapes and dimension-line labels (fixture for the page)");
        string fodt = Path.Combine(TestReport.OutputDirectory, "fidelity-v2-odt-shape-text.fodt");
        string document = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document " + Namespaces + " office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
                        + Styles + "<office:body><office:text>" + string.Concat(Body) + "</office:text></office:body></office:document>";
        if (!File.Exists(fodt) || File.ReadAllText(fodt) != document) File.WriteAllText(fodt, document, new UTF8Encoding(false));
        List<string> html = OdtDrawingFidelityTest.Render(fodt);
        TestReport.Check(html.Count == Body.Length, $"blocks: {html.Count} of {Body.Length}");
        if (html.Count < Body.Length) return;
        // 0.1 in = 7.2 pt, 0.05 in = 3.6 pt; the rectangle's text area is its box (108 x 36 pt)
        TestReport.Check(html[0].Contains("data-autogrow=\"h\" data-autogrow-text=\"0 0 108 36\" data-autogrow-shape=\"108 36\" data-autogrow-pad=\"7.2 3.6 7.2 3.6\" data-autogrow-min=\"0 0\" data-autogrow-align=\"top justify\""),
                         "grows down: height only, text area, shape, distances, alignment (" + Clip(html[0]) + ")");
        TestReport.Check(html[1].Contains("data-autogrow=\"h\"") && html[1].Contains("data-autogrow-align=\"bottom justify\""), "grows up: bottom-aligned");
        TestReport.Check(html[2].Contains("data-autogrow=\"wh\"") && html[2].Contains("data-autogrow-align=\"top left\""), "no wrap: width and height grow (svdoashp.cxx:2187-2194)");
        TestReport.Check(html[3].Contains("data-autogrow-min=\"0 36\"") && html[3].Contains("data-autogrow-align=\"middle justify\""), "minimum height 0.5 in, centred");
        TestReport.Check(!html[4].Contains("data-autogrow"), "auto-grow off: no growth");
        TestReport.Check(html[5].Contains("class=\"fast-doc-view-measure\" data-measure=\"36 72 72 72 22.6772 0 0 0 0 0 0 0 0 0 0 0 0 0\"")
                         && html[5].Contains("data-measure-matrix=\"") && html[5].Contains(">a long measure label wider than its line</span></span>"),
                         "measure: layout inputs (line 36,72 -> 72,72 pt, default distance 8 mm) and the label inside its wrapper (" + Clip(html[5]) + ")");
        TestReport.Check(System.Text.RegularExpressions.Regex.Matches(html[6], "<svg class=\"word-shape-geometry\"").Count == 2, "measure: help lines and dimension line drawn apart");
        TestReport.Check(html[0].Contains("data-shape-page=\"108 36\"") && html[0].Contains("data-shape-box=\"108 36\" data-autogrow-id=\"0\"")
                         && html[0].Contains("data-autogrow-matrix=\"1 0 0 1 0 0\" data-autogrow-mirror=\"0 0\""), "the frame and the shape's box marked for the page");
        TestReport.Check(html[7].Contains("data-autogrow-matrix=\"0 -1 1 0 ") && html[7].Contains("data-autogrow-turn=\"-1.570796"), "turned shape: its own matrix and the text's turn (" + Clip(html[7]) + ")");
        TestReport.Check(html[8].Contains("data-autogrow-mirror=\"1 0\""), "mirrored shape: mirroring passed apart from the matrix");
        TestReport.Check(html[9].Contains("data-shape-box=\"108 36\">") && html[9].Contains("data-autogrow-id=\"1\"") && !html[9].Contains("data-autogrow-id=\"0\""), "group: every shape's box, only the growing one marked");
        TestReport.Check(html[10].Contains("data-autogrow-text-rotate=\"90\""), "draw:text-rotate-angle passed");
        TestReport.Check(html[11].Contains("data-autogrow=\"h\""), "draw:auto-grow-height absent: grows (pool default on, svdattr.cxx:247)");
        TestReport.Check(html[12].Contains("shape-outside:polygon(") && html[12].Contains("data-autogrow=\"h\""), "contour-wrapped growing shape");
        // 0.15 in = 10.8 pt wide, the 20 x 30 view box: 16.2 pt long (sdrmeasureprimitive2d.cxx:150-170)
        TestReport.Check(html[13].Contains(" 10.8 16.2 10.8 16.2 ") && html[13].Contains("data-measure-markers=\"10.8 ") && html[13].Contains("|10.8 "),
                         "measure arrows: the markers' width and length reach the layout and the page (" + Clip(html[13]) + ")");
    }

    private static string Clip(string html) => html.Length > 1500 ? html.Substring(0, 1500) + "..." : html;
}
