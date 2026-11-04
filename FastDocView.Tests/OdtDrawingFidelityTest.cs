using System.Text;
using FastDocView.Html;
using FastDocView.Odt;

namespace FastDocView.Tests;

/// <summary>
/// OpenDocument drawing (2026-11-04): custom-shape geometry (enhanced paths, formulas, presets by draw:type), Fontwork, strokes
/// (dashes, joins, markers), fills (every gradient style, steps, borders, transparency gradients, hatches, bitmaps), every
/// drawing object kind with draw:transform, text areas, text boxes with lists and tables, contour wrap, page-anchored objects,
/// and search over all of their text. Hand-worked values from the ODF 1.3 spec and LibreOffice's source (cited in the code under
/// test); a hand-written .fodt, and its LibreOffice conversion (presets saved with their full enhanced path) must draw the same.
/// </summary>
public static class OdtDrawingFidelityTest
{
    private const string Namespaces =
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
      + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" "
      + "xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" "
      + "xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" "
      + "xmlns:loext=\"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0\"";

    private const string Png = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAAEUlEQVR4nGP4z8AARwgWXg4ArpMP8aaUSCMAAAAASUVORK5CYII=";

    private const string Styles = """
        <office:styles>
          <draw:gradient draw:name="Axial" draw:style="axial" draw:start-color="#000000" draw:end-color="#ffffff" draw:angle="0deg" draw:border="0%"/>
          <draw:gradient draw:name="Radial" draw:style="radial" draw:cx="50%" draw:cy="50%" draw:start-color="#000000" draw:end-color="#ffffff" draw:border="0%"/>
          <draw:gradient draw:name="Ellipsoid" draw:style="ellipsoid" draw:cx="50%" draw:cy="50%" draw:start-color="#000000" draw:end-color="#ffffff" draw:angle="0" draw:border="0%"/>
          <draw:gradient draw:name="Square" draw:style="square" draw:cx="50%" draw:cy="50%" draw:start-color="#000000" draw:end-color="#ffffff" draw:angle="0" draw:border="0%"/>
          <draw:gradient draw:name="Rectangular" draw:style="rectangular" draw:cx="50%" draw:cy="50%" draw:start-color="#000000" draw:end-color="#ffffff" draw:angle="0" draw:border="0%"/>
          <draw:gradient draw:name="Border" draw:style="linear" draw:start-color="#ff0000" draw:end-color="#0000ff" draw:angle="90deg" draw:border="20%" draw:start-intensity="50%"/>
          <draw:gradient draw:name="Multi" draw:style="linear" draw:start-color="#ff0000" draw:end-color="#0000ff" draw:angle="0">
            <loext:gradient-stop svg:offset="0" loext:color-type="rgb" loext:color-value="#ff0000"/>
            <loext:gradient-stop svg:offset="0.5" loext:color-type="rgb" loext:color-value="#00ff00"/>
            <loext:gradient-stop svg:offset="1" loext:color-type="rgb" loext:color-value="#0000ff"/>
          </draw:gradient>
          <draw:opacity draw:name="Fade" draw:style="linear" draw:start="100%" draw:end="0%" draw:angle="0"/>
          <draw:hatch draw:name="Cross" draw:style="double" draw:color="#0000ff" draw:distance="0.1in" draw:rotation="450"/>
          <draw:stroke-dash draw:name="DotDash" draw:style="rect" draw:dots1="1" draw:dots1-length="0.05in" draw:dots2="1" draw:dots2-length="0.15in" draw:distance="0.05in"/>
          <draw:stroke-dash draw:name="Relative" draw:style="round" draw:dots1="2" draw:dots1-length="100%" draw:distance="200%"/>
          <draw:marker draw:name="Arrow" svg:viewBox="0 0 20 30" svg:d="M10 0l-10 30h20z"/>
          <draw:marker draw:name="Notched" svg:viewBox="0 0 20 30" svg:d="M10 0l10 30-10-10-10 10z"/>
          <draw:fill-image draw:name="Tile"><office:binary-data>
        """ + Png + """
        </office:binary-data></draw:fill-image>
        </office:styles>
        <office:automatic-styles>
          <style:style style:name="Plain" style:family="graphic"><style:graphic-properties style:wrap="none" style:horizontal-pos="left" style:horizontal-rel="paragraph" draw:fill="solid" draw:fill-color="#ff0000" draw:stroke="solid" svg:stroke-color="#000000" svg:stroke-width="0.01in"/></style:style>
          <style:style style:name="Defaults" style:family="graphic"><style:graphic-properties style:wrap="none" style:horizontal-pos="left" style:horizontal-rel="paragraph"/></style:style>
          <style:style style:name="Grey" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill-color="#808080"/></style:style>
          <style:style style:name="Lines" style:family="graphic"><style:graphic-properties style:wrap="none" style:horizontal-pos="left" draw:fill="none" draw:stroke="dash" draw:stroke-dash="DotDash" svg:stroke-color="#000080" svg:stroke-width="0.02in" draw:stroke-linejoin="bevel" svg:stroke-linecap="round" draw:marker-start="Arrow" draw:marker-start-width="0.2in" draw:marker-end="Notched" draw:marker-end-width="0.2in" draw:marker-end-center="true"/></style:style>
          <style:style style:name="Relative" style:family="graphic"><style:graphic-properties style:wrap="none" draw:fill="none" draw:stroke="dash" draw:stroke-dash="Relative" svg:stroke-color="#000000" svg:stroke-width="0.1in" draw:stroke-linejoin="none"/></style:style>
          <style:style style:name="FillAxial" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Axial"/></style:style>
          <style:style style:name="FillRadial" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Radial"/></style:style>
          <style:style style:name="FillEllipsoid" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Ellipsoid"/></style:style>
          <style:style style:name="FillSquare" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Square"/></style:style>
          <style:style style:name="FillRectangular" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Rectangular"/></style:style>
          <style:style style:name="FillBorder" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Border" draw:gradient-step-count="4"/></style:style>
          <style:style style:name="FillMulti" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Multi"/></style:style>
          <style:style style:name="FillFade" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="solid" draw:fill-color="#00ff00" draw:opacity-name="Fade"/></style:style>
          <style:style style:name="FillHalf" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="solid" draw:fill-color="#00ff00" draw:opacity="40%"/></style:style>
          <style:style style:name="FillHatch" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="hatch" draw:fill-hatch-name="Cross" draw:fill-hatch-solid="true" draw:fill-color="#ffff00"/></style:style>
          <style:style style:name="FillTile" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:fill="bitmap" draw:fill-image-name="Tile" style:repeat="repeat" draw:fill-image-width="0.25in" draw:fill-image-height="0.25in" draw:fill-image-ref-point="top-left" draw:tile-repeat-offset="50% horizontal"/></style:style>
          <style:style style:name="TextArea" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:textarea-vertical-align="bottom" draw:textarea-horizontal-align="center" fo:padding-left="0.1in" fo:padding-right="0.1in" fo:padding-top="0.05in" fo:padding-bottom="0.05in"/></style:style>
          <style:style style:name="Fontwork" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties draw:stroke="none" draw:fill-color="#0000ff" draw:textarea-horizontal-align="center"/><style:text-properties style:font-name="Liberation Serif" fo:font-family="'Liberation Serif'"/></style:style>
          <style:style style:name="Middle" style:family="graphic"><style:graphic-properties style:wrap="none" fo:padding="0in" fo:border="none" draw:fill="none" draw:textarea-vertical-align="middle"/></style:style>
          <style:style style:name="Contour" style:family="graphic" style:parent-style-name="Plain"><style:graphic-properties style:wrap="right" style:wrap-contour="true" style:wrap-contour-mode="outside" style:horizontal-pos="left" fo:margin-right="0.1in"/></style:style>
          <text:list-style style:name="Numbers"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>
        </office:automatic-styles>
        """;

    private static string Custom(string style, string geometry, string text = "", string size = "svg:width=\"2in\" svg:height=\"1in\"", string extra = "")
        => "<draw:custom-shape draw:style-name=\"" + style + "\" text:anchor-type=\"paragraph\" " + size + " svg:x=\"0in\" svg:y=\"0in\" " + extra + ">" + text
         + "<draw:enhanced-geometry " + geometry + "</draw:enhanced-geometry></draw:custom-shape>";

    private const string OwnGeometry = "svg:viewBox=\"0 0 21600 21600\" draw:modifiers=\"5400\" draw:text-areas=\"0 0 ?f0 ?f1\" draw:enhanced-path=\"M 0 0 L ?f0 0 ?f0 ?f1 0 21600 Z N\"";
    private const string OwnEquations = "<draw:equation draw:name=\"f0\" draw:formula=\"$0 *2\"/><draw:equation draw:name=\"f1\" draw:formula=\"height-$0 \"/>";

    private static string Rect(string style, string text = "") => "<draw:rect draw:style-name=\"" + style + "\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"1in\">" + text + "</draw:rect>";

    /// <summary>Blocks in order; the checks name each by its index.</summary>
    internal static readonly string[] Body =
    [
        // 0 an enhanced path with formulas, modifiers and a text area (144 x 72 pt)
        "<text:p>" + Custom("Plain", OwnGeometry + ">" + OwnEquations, "<text:p>owntext</text:p>") + "</text:p>",
        // 1 the same mirrored horizontally
        "<text:p>" + Custom("Plain", "draw:mirror-horizontal=\"true\" " + OwnGeometry + ">" + OwnEquations) + "</text:p>",
        // 2 a preset by type only, its adjustment given
        "<text:p>" + Custom("Plain", "svg:viewBox=\"0 0 21600 21600\" draw:type=\"isosceles-triangle\" draw:modifiers=\"5400\">") + "</text:p>",
        // 3 the cube preset (colour data: lighter / darker faces)
        "<text:p>" + Custom("Grey", "svg:viewBox=\"0 0 21600 21600\" draw:type=\"cube\">", size: "svg:width=\"1in\" svg:height=\"1in\"") + "</text:p>",
        // 4 Fontwork: plain text on its outlines
        "<text:p>" + Custom("Fontwork", "svg:viewBox=\"0 0 21600 21600\" draw:type=\"fontwork-plain-text\" draw:text-path=\"true\" draw:text-path-mode=\"shape\" draw:text-path-scale=\"path\">", "<text:p>fontworkword</text:p>") + "</text:p>",
        // 5 a poly-line with a dash, a bevel join, round caps and two markers
        "<text:p><draw:polyline draw:style-name=\"Lines\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"2in\" svg:height=\"0.5in\" svg:viewBox=\"0 0 2000 500\" draw:points=\"0,0 1000,500 2000,0\"/></text:p>",
        // 6 a closed path with an arc (filled) and an open one (not filled), relative dashes, no join
        "<text:p><draw:path draw:style-name=\"Plain\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1.5in\" svg:height=\"1in\" svg:viewBox=\"0 0 1500 1000\" svg:d=\"M0 0L1000 0A500 500 0 0 1 1000 1000L0 1000Z\"/>"
            + "<draw:path draw:style-name=\"Relative\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"1in\" svg:viewBox=\"0 0 1000 1000\" svg:d=\"M0 1000L500 0 1000 1000\"/></text:p>",
        // 7 circle section, cut and arc
        "<text:p><draw:circle draw:style-name=\"Plain\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"1in\" draw:kind=\"section\" draw:start-angle=\"0\" draw:end-angle=\"90\"/>"
            + "<draw:ellipse draw:style-name=\"Plain\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"1in\" draw:kind=\"arc\" draw:start-angle=\"90\" draw:end-angle=\"180\"/></text:p>",
        // 8 a five-pointed star (regular polygon, 50 % sharpness)
        "<text:p><draw:regular-polygon draw:style-name=\"Plain\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"1in\" draw:corners=\"5\" draw:concave=\"true\" draw:sharpness=\"50%\"/></text:p>",
        // 9 a connector with its routed path
        "<text:p><draw:connector draw:style-name=\"Lines\" text:anchor-type=\"paragraph\" draw:type=\"standard\" svg:x1=\"0in\" svg:y1=\"0in\" svg:x2=\"1in\" svg:y2=\"0.5in\" svg:d=\"M0 0L2540 0 2540 1270\"/></text:p>",
        // 10 a rectangle turned 30 degrees and moved; a sheared one
        "<text:p><draw:rect draw:style-name=\"Plain\" text:anchor-type=\"paragraph\" svg:width=\"1in\" svg:height=\"0.5in\" draw:transform=\"rotate (0.523598775598299) translate (1in 0.5in)\"><text:p>turnedtext</text:p></draw:rect></text:p>",
        "<text:p><draw:rect draw:style-name=\"Plain\" text:anchor-type=\"paragraph\" svg:width=\"1in\" svg:height=\"0.5in\" draw:transform=\"skewX (-0.785398163397448) translate (0in 0in)\"/></text:p>",
        // 12.. fills
        "<text:p>" + Rect("FillAxial") + "</text:p>",
        "<text:p>" + Rect("FillRadial") + "</text:p>",
        "<text:p>" + Rect("FillEllipsoid") + "</text:p>",
        "<text:p>" + Rect("FillSquare") + "</text:p>",
        "<text:p>" + Rect("FillRectangular") + "</text:p>",
        "<text:p>" + Rect("FillBorder") + "</text:p>",
        "<text:p>" + Rect("FillMulti") + "</text:p>",
        "<text:p>" + Rect("FillFade") + "</text:p>",
        "<text:p>" + Rect("FillHalf") + "</text:p>",
        "<text:p>" + Rect("FillHatch") + "</text:p>",
        "<text:p>" + Rect("FillTile") + "</text:p>",
        // 23 a text box: middle, a numbered list and a table
        "<text:p><draw:frame draw:style-name=\"Middle\" text:anchor-type=\"paragraph\" svg:width=\"2in\" svg:height=\"1.5in\"><draw:text-box>"
            + "<text:list text:style-name=\"Numbers\"><text:list-item><text:p>listone</text:p></text:list-item><text:list-item><text:p>listtwo</text:p></text:list-item></text:list>"
            + "<table:table><table:table-column/><table:table-row><table:table-cell><text:p>cellone</text:p></table:table-cell></table:table-row></table:table>"
            + "</draw:text-box></draw:frame>after box</text:p>",
        // 24 a shape's text area: centred at the bottom with padding
        "<text:p>" + Rect("TextArea", "<text:p>areatext</text:p>") + "</text:p>",
        // 25 contour wrap around an ellipse
        "<text:p><draw:ellipse draw:style-name=\"Contour\" text:anchor-type=\"paragraph\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"1in\"/>wrapped text</text:p>",
        // 26 a page-anchored rectangle between paragraphs (its own block), then a paragraph
        "<draw:rect draw:style-name=\"Defaults\" text:anchor-type=\"page\" text:anchor-page-number=\"1\" svg:x=\"0in\" svg:y=\"0in\" svg:width=\"1in\" svg:height=\"0.5in\"><text:p>pagetext</text:p></draw:rect>",
        "<text:p>last</text:p>",
    ];

    public static void Run()
    {
        UnpaddedBinaryPicture();
        TestReport.Step("ODT drawing: formulas and paths");
        CheckFormulas();
        TestReport.Step("ODT drawing: hand-written .fodt");
        string fodt = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-drawing.fodt");
        string document = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document " + Namespaces + " office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
                        + Styles + "<office:body><office:text>" + string.Concat(Body) + "</office:text></office:body></office:document>";
        if (!File.Exists(fodt) || File.ReadAllText(fodt) != document) File.WriteAllText(fodt, document, new UTF8Encoding(false));
        CheckHandWritten(fodt);
        TestReport.Step("ODT drawing: gradient angle units by generator, automatic steps");
        CheckGradientAnglesAndSteps();
        TestReport.Step("ODT drawing: the same document saved by LibreOffice");
        if (LibreOfficeConversion.Convert(fodt, ".odt", "odt") is { } odt) CheckLibreOfficeSaved(fodt, odt);
    }

    private const string MetaNamespace = "xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\"";

    /// <summary>A gradient with a unitless draw:angle of 900 and one with automatic steps, in an ODF 1.2 file by <paramref name="generator"/>.</summary>
    private static (string Meta, string Styles, string Body) AngleFixture(string? generator)
        => (generator is null ? "" : "<office:meta><meta:generator>" + generator + "</meta:generator></office:meta>",
            "<office:styles>"
            + "<draw:gradient draw:name=\"Turned\" draw:style=\"linear\" draw:start-color=\"#ff0000\" draw:end-color=\"#0000ff\" draw:angle=\"900\" draw:border=\"0%\"/>"
            + "<draw:gradient draw:name=\"Auto\" draw:style=\"linear\" draw:start-color=\"#ff0000\" draw:end-color=\"#0000ff\" draw:angle=\"0deg\" draw:border=\"0%\"/>"
            + "<style:style style:name=\"FillTurned\" style:family=\"graphic\"><style:graphic-properties style:wrap=\"none\" style:horizontal-pos=\"left\" draw:fill=\"gradient\" draw:fill-gradient-name=\"Turned\" draw:gradient-step-count=\"2\"/></style:style>"
            + "<style:style style:name=\"FillAuto\" style:family=\"graphic\"><style:graphic-properties style:wrap=\"none\" style:horizontal-pos=\"left\" draw:fill=\"gradient\" draw:fill-gradient-name=\"Auto\" draw:opacity=\"50%\"/></style:style>"
            + "<draw:gradient draw:name=\"Ring\" draw:style=\"radial\" draw:cx=\"50%\" draw:cy=\"50%\" draw:start-color=\"#ff0000\" draw:end-color=\"#0000ff\" draw:border=\"20%\"/>"
            + "<style:style style:name=\"FillRing\" style:family=\"graphic\"><style:graphic-properties style:wrap=\"none\" style:horizontal-pos=\"left\" draw:fill=\"gradient\" draw:fill-gradient-name=\"Ring\"/></style:style>"
            + "<style:style style:name=\"FillSmooth\" style:family=\"graphic\"><style:graphic-properties style:wrap=\"none\" style:horizontal-pos=\"left\" draw:fill=\"gradient\" draw:fill-gradient-name=\"Auto\"/></style:style>"
            + "</office:styles>",
            "<text:p>" + Rect("FillTurned") + "</text:p><text:p>" + Rect("FillAuto") + "</text:p><text:p>" + Rect("FillSmooth") + "</text:p><text:p>" + Rect("FillRing") + "</text:p>");

    private static void CheckGradientAnglesAndSteps()
    {
        // generator classification (xmlmetai.cxx:232-298, xmlimp.cxx:171-338)
        (string? generator, ushort expected)[] generators =
        [
            (null, OdtGeneratorVersion.Unknown),
            ("OpenOffice.org/3.2$Win32 OpenOffice.org_project/320m12$Build-9483", OdtGeneratorVersion.OOo32),
            ("OpenOffice.org/2.4$Win32 OpenOffice.org_project/680m17$Build-9310", OdtGeneratorVersion.OOo2),
            ("StarOffice 7", OdtGeneratorVersion.OOo1),
            ("OpenOffice/4.1.6$Win32 OpenOffice.org_project/416m1$Build-9790", OdtGeneratorVersion.Aoo4),
            ("LibreOffice/6.4.7.2$Linux_X86_64 LibreOffice_project/40$Build-2", OdtGeneratorVersion.LibreOffice63),
            ("LibreOffice/7.3.7.2$Windows_X86_64 LibreOffice_project/e114eadc50a9ff8d8c8a0567d6da8f454beeb84f", OdtGeneratorVersion.LibreOffice7),
            ("LibreOffice/24.8.4.2$Windows_X86_64 LibreOffice_project/bb3cfa12c7b1bf994ecc5649a80400d06cd71002", OdtGeneratorVersion.LibreOffice248),
            ("LibreOffice/27.2.0.1$Windows_X86_64 LibreOffice_project/0", OdtGeneratorVersion.LibreOfficeNew),
            ("Microsoft Word 16.0", OdtGeneratorVersion.Unknown),
        ];
        foreach ((string? generator, ushort expected) in generators)
            TestReport.Check(OdtGeneratorVersion.Of(generator) == expected, $"generator \"{generator}\": version code {OdtGeneratorVersion.Of(generator)} (expected {expected})");
        // tdf#89475 (GradientStyle.cxx:115-128): tenths before ODF 1.2, and in 1.2 from OOo / AOO / LibreOffice before 7; degrees otherwise
        TestReport.Check(OdtGeneratorVersion.GradientAngleIsTenths(1.1, OdtGeneratorVersion.LibreOffice7)
                         && OdtGeneratorVersion.GradientAngleIsTenths(1.2, OdtGeneratorVersion.OOo32)
                         && OdtGeneratorVersion.GradientAngleIsTenths(1.2, OdtGeneratorVersion.Aoo4)
                         && OdtGeneratorVersion.GradientAngleIsTenths(1.2, OdtGeneratorVersion.LibreOffice63)
                         && !OdtGeneratorVersion.GradientAngleIsTenths(1.2, OdtGeneratorVersion.LibreOffice7)
                         && !OdtGeneratorVersion.GradientAngleIsTenths(1.2, OdtGeneratorVersion.Unknown)
                         && !OdtGeneratorVersion.GradientAngleIsTenths(1.3, OdtGeneratorVersion.OOo32),
                         "unitless gradient angles: tenths of a degree only for old ODF or old writers");

        const string oldWriter = "OpenOffice.org/3.2$Win32 OpenOffice.org_project/320m12$Build-9483", newWriter = "LibreOffice/7.3.7.2$Windows_X86_64 LibreOffice_project/e114ead";
        foreach ((string? generator, string label, string rotation) in new[] { ((string?)oldWriter, "OOo 3.2", "rotate(-90)"), (newWriter, "LibreOffice 7.3", "rotate(-180)"), (null, "no generator", "rotate(-180)") })
        {
            (string meta, string styles, string body) = AngleFixture(generator);
            string flat = Path.Combine(TestReport.OutputDirectory, $"fidelity-v1-odt-angle-{(generator is null ? "none" : generator == oldWriter ? "ooo" : "lo")}.fodt");
            string document = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document " + Namespaces + " " + MetaNamespace + " office:version=\"1.2\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
                            + meta + styles + "<office:body><office:text>" + body + "</office:text></office:body></office:document>";
            File.WriteAllText(flat, document, new UTF8Encoding(false));
            List<string> html = Render(flat);
            // 900 unitless: 90 degrees as tenths (start on the left: rotate(-90)), else 900 degrees = 180
            TestReport.Check(html.Count >= 1 && html[0].Contains(rotation), $".fodt by {label}: draw:angle=\"900\" draws as {rotation} ({(html.Count > 0 ? Clip(html[0]) : "")})");
            if (generator != oldWriter) continue;

            // automatic steps (no draw:gradient-step-count = 0) under a uniform transparency (a gradient with alpha: decomposed): red to blue is a colour distance of 1 = 128 bands (gradienttools.cxx:479-497),
            // band k coloured interpolate(start, end, k / 127) (texture.cxx:169-222): band 1 = #fd0002, two stops per band
            string auto = html.Count > 1 ? html[1] : "";
            TestReport.Check(Count(auto, "<stop ") == 256 && auto.Contains("stop-color=\"#fd0002\"") && auto.Contains("stop-color=\"#0000ff\""),
                             $"automatic steps with transparency: 128 bands as LibreOffice's decomposition draws them ({Count(auto, "<stop ")} stops)");
            // the same without transparency: Skia draws it natively, smooth (OdtFillPaint.IsDrawnSmoothBySkia: gdiimpl.cxx drawGradient)
            string smooth = html.Count > 2 ? html[2] : "";
            TestReport.Check(Count(smooth, "<stop ") == 2 && smooth.Contains("<stop offset=\"0\" stop-color=\"#ff0000\"/><stop offset=\"1\" stop-color=\"#0000ff\"/>"),
                             $"automatic steps, plain two-colour linear: smooth as Skia draws it ({Count(smooth, "<stop ")} stops)");
            // radial with a 20 % border, as Skia draws it (gdiimpl.cxx:2006-2014): radius half the diagonal (1 in square: 50.91 pt),
            // the end colour solid out to 20 % of it, then smooth to the start colour
            string ring = html.Count > 3 ? html[3] : "";
            TestReport.Check(ring.Contains("<stop offset=\"0\" stop-color=\"#0000ff\"/><stop offset=\"0.2\" stop-color=\"#0000ff\"/><stop offset=\"1\" stop-color=\"#ff0000\"/>")
                             && ring.Contains("scale(50.911688 50.911688)"), $"radial with border: Skia's layout ({Clip(ring)})");

            string package = Path.ChangeExtension(flat, ".odt");
            if (File.Exists(package)) File.Delete(package);
            using (var zip = System.IO.Compression.ZipFile.Open(package, System.IO.Compression.ZipArchiveMode.Create))
            {
                void Entry(string name, string text, bool stored = false)
                {
                    var entry = zip.CreateEntry(name, stored ? System.IO.Compression.CompressionLevel.NoCompression : System.IO.Compression.CompressionLevel.Optimal);
                    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                    writer.Write(text);
                }
                Entry("mimetype", "application/vnd.oasis.opendocument.text", stored: true);
                Entry("meta.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-meta " + Namespaces + " " + MetaNamespace + " office:version=\"1.2\">" + meta + "</office:document-meta>");
                Entry("styles.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles " + Namespaces + " office:version=\"1.2\">" + styles + "</office:document-styles>");
                Entry("content.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content " + Namespaces + " office:version=\"1.2\"><office:body><office:text>" + body + "</office:text></office:body></office:document-content>");
            }
            foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
            {
                var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-angle"), ForcedMode = mode };
                for (int open = 0; open < 2; open++)     // the second open reuses the sidecar: the generator is read again, not stored
                {
                    using IWordWindowedDocument opened = WordWindowedDocumentOpener.Open(package, policy);
                    while (!opened.IsIndexComplete) opened.ScanAhead(100_000);
                    string first = opened.RenderBlocks(0, 1).FirstOrDefault()?.Html ?? "";
                    TestReport.Check(first.Contains("rotate(-90)"), $".odt package by {label}, {mode}, open {open + 1}: meta.xml's generator makes 900 tenths ({Clip(first)})");
                }
            }
        }
    }

    internal static List<string> Render(string path)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }

    private static void CheckFormulas()
    {
        var engine = new WordCustomShapeEngine { LogicWidth = 2540, LogicHeight = 1270, Adjust = [5400, 2.5] };
        engine.SetPathSize(0);
        double Eval(string formula) => WordCustomShapeFormula.Parse(formula)?.Evaluate(engine) ?? double.NaN;
        TestReport.Check(Eval("$0 *2") == 10800 && Eval("width-$0") == 16200 && Eval("-$1*2") == -5, "formula: adjustments, keywords, unary minus");
        TestReport.Check(Eval("if(-1, 1, 2)") == 2 && Eval("if(0.5,1,2)") == 1, "formula: if() takes the second argument only for a positive condition (EnhancedCustomShapeFunctionParser.cxx:702-705)");
        TestReport.Check(Math.Abs(Eval("atan2(1,1)") - Math.PI / 4) < 1e-12 && Math.Abs(Eval("sin(pi/2)") - 1) < 1e-12 && Eval("max(3,min(7,5))") == 5 && Eval("abs(-4)+sqrt(16)") == 8,
                         "formula: radians, atan2(y, x), min / max / abs / sqrt");
        TestReport.Check(Eval("1.5E2") == 150 && double.IsNaN(Eval("1e2")) && double.IsNaN(Eval("2 3")), "formula: exponent only with E; a formula not read whole is not a value");
        TestReport.Check(Eval("logwidth/2") == 1270 && Eval("right") == 21600 && Eval("2*(3+4)/7") == 2, "formula: logwidth in 1/100 mm, right = the view box width, precedence");
        // OfficeArt calculation records as LibreOffice turns them into formulas (GetEquation)
        TestReport.Check(WordCustomShapeFormula.OfficeArtEquation(0x2001, 327, 1, 2) == "$0 /2" && WordCustomShapeFormula.OfficeArtEquation(0x8000, 21600, 0, 0x400) == "21600-?0 "
                         && WordCustomShapeFormula.OfficeArtEquation(0x0008, 1, 2, 0) == "atan2(2,1)/(pi/180)", "OfficeArt records become formulas (EnhancedCustomShape2d.cxx:88-330)");
        // enhanced-path syntax: a moveto with more pairs continues as lineto; repeated commands merge
        (List<WordShapeParameter> coordinates, List<WordPathSegment> segments) = OdtEnhancedGeometry.Path("M 0 0 10 10 20 0 Z N", null, new Dictionary<string, int>());
        TestReport.Check(coordinates.Count == 6 && segments.Count == 4 && segments[0] == new WordPathSegment(WordPathCommand.MoveTo, 1) && segments[1] == new WordPathSegment(WordPathCommand.LineTo, 2)
                         && segments[2].Command == WordPathCommand.CloseSubpath && segments[3].Command == WordPathCommand.EndSubpath, "enhanced path: M pairs after the first are L (ODF 19.145)");
        // dashes: absolute lengths, then relative to the line width with a zero length = a dot of the line width
        var dash = System.Xml.Linq.XElement.Parse("<d xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" draw:dots1=\"1\" draw:dots1-length=\"0.05in\" draw:dots2=\"1\" draw:dots2-length=\"0.15in\" draw:distance=\"0.05in\"/>");
        TestReport.Check(string.Join(",", OdtShapeStyle.DashArray(dash, 1.44).Select(v => Math.Round(v, 6))) == "3.6,3.6,10.8,3.6", "dash: dot, distance, dash, distance in points (XDash::CreateDotDashArray)");
        var relative = System.Xml.Linq.XElement.Parse("<d xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" draw:dots1=\"2\" draw:dots1-length=\"100%\" draw:distance=\"200%\"/>");
        TestReport.Check(string.Join(",", OdtShapeStyle.DashArray(relative, 7.2).Select(v => Math.Round(v, 6))) == "7.2,14.4,7.2,14.4", "dash: percentages of the line width (7.2 pt line)");
        // transforms: rotate(a) turns counter-clockwise on screen, then translate (xexptran.cxx:480-540)
        WordAffine t = OdtShapeBuilder.Transform("rotate (1.5707963267949) translate (1in 0in)");
        (double x, double y) = t.Apply(10, 0);
        TestReport.Check(Math.Abs(x - 72) < 1e-6 && Math.Abs(y + 10) < 1e-6, "draw:transform: rotate(pi/2) turns (10, 0) to (0, -10), then the 1 in move");
    }

    private static void CheckHandWritten(string path)
    {
        List<string> html = Render(path);
        TestReport.Check(html.Count == Body.Length, $"blocks: one per paragraph and one for the page-anchored rectangle ({html.Count} of {Body.Length})");
        if (html.Count < Body.Length) return;
        // 0: f0 = 5400 * 2 = 10800 -> 72 pt of 144; f1 = 21600 - 5400 = 16200 -> 54 pt of 72
        TestReport.Check(html[0].Contains("<path d=\"M0 0L72 0L72 54L0 72Z\" fill-rule=\"evenodd\" fill=\"#ff0000\""), "enhanced path: ?f0 / ?f1 from the formulas, scaled to 144 x 72 pt (" + Clip(html[0]) + ")");
        TestReport.Check(html[0].Contains("width:72pt;height:54pt;") && html[0].Contains(">owntext</span>"), "enhanced path: the text area 0 0 ?f0 ?f1");
        TestReport.Check(html[1].Contains("<g transform=\"matrix(-1 0 0 1 144 0)\">"), "draw:mirror-horizontal: mirrored about the box centre (" + Clip(html[1]) + ")");
        // 2: msoIsocelesTriangle apex at $0 = 5400 of 21600 -> 36 pt of 144
        TestReport.Check(html[2].Contains("<path d=\"M36 0L144 72L0 72Z\""), "isosceles-triangle preset with draw:modifiers 5400: apex at 36 pt (" + Clip(html[2]) + ")");
        // 3: cube, colour data 0x302e0000: faces 0, +20 %, -20 % luminance of #808080 (HSV: value +20 % toward 1 = 0.6, s 0; value 0.8 x 0.502)
        TestReport.Check(html[3].Contains("fill=\"#808080\"") && html[3].Contains("fill=\"#999999\"") && html[3].Contains("fill=\"#666666\""), "cube: the three faces by its colour data (" + Clip(html[3]) + ")");
        TestReport.Check(html[4].Contains("data-fontwork=\"pair\"") && html[4].Contains(">fontworkword</textPath>") && html[4].Contains("font-family=\"Liberation Serif\"")
                         && html[4].Contains("fill=\"#0000ff\"") && html[4].Contains("data-fontwork-align=\"1\"") && !html[4].Contains("class=\"text-box word-shape-text\""),
                         "Fontwork: the text on the preset's outlines, the shape's fill, centred, no text box (" + Clip(html[4]) + ")");
        // 5: 0.02 in = 1.44 pt line; dash 3.6 / 3.6 / 10.8 / 3.6 pt = 2.5, 2.5, 7.5, 2.5 line widths; bevel join; round cap; two markers
        TestReport.Check(html[5].Contains("stroke-dasharray=\"3.6 3.6 10.8 3.6\"") && html[5].Contains("stroke-linejoin=\"bevel\"") && html[5].Contains("stroke-linecap=\"round\"")
                         && Count(html[5], "fill=\"#000080\" stroke=\"none\"") == 2, "poly-line: dash, bevel join, round cap, start and end markers (" + Clip(html[5]) + ")");
        TestReport.Check(html[6].Contains("fill=\"#ff0000\"") && html[6].Contains("stroke-dasharray=\"7.2 14.4 7.2 14.4\"") && Count(html[6], "<path d=\"M") >= 4,
                         "paths: the closed one filled, the open one stroked only, relative dashes, segments stroked one by one without a join (" + Clip(html[6]) + ")");
        // svdocirc.cxx:265-286: the segment runs from 360 - end (top) to 360 - start (right) in y-down angles, the centre first
        TestReport.Check(html[7].Contains("<path d=\"M36 36L36 0L39.53 0.17") && html[7].Contains("L72 36Z"), "circle section 0..90: the centre, then from 12 to 3 o'clock (" + Clip(html[7]) + ")");
        TestReport.Check(html[8].Contains("<path d=\"M36 0L") && html[8].Contains("fill=\"#ff0000\""), "regular polygon: a five-pointed star from the top (" + Clip(html[8]) + ")");
        // the route's ends cut back by the markers: the start arrow 0.2 in = 14.4 pt wide, 21.6 pt long, the line meets it 6 / 15 of the width in (b2dlinegeometry / AppendWithLineEnds)
        TestReport.Check(html[9].Contains("L72 0L72 ") && html[9].Contains("<path d=\"M0 0L21.6 7.2L21.6 -7.2Z\""), "connector: its routed svg:d with its markers (" + Clip(html[9]) + ")");
        TestReport.Check(html[10].Contains("matrix(0.866025 -0.5 0.5 0.866025 ") && html[10].Contains(">turnedtext</span>"), "rotate 30 degrees: counter-clockwise; the text turns with it (" + Clip(html[10]) + ")");
        TestReport.Check(html[11].Contains("matrix(1 0 1 1 "), "skewX(-45 degrees): x' = x + y (" + Clip(html[11]) + ")");
        TestReport.Check(html[12].Contains("<stop offset=\"0\" stop-color=\"#ffffff\"/><stop offset=\"0.5\" stop-color=\"#000000\"/><stop offset=\"1\" stop-color=\"#ffffff\"/>"),
                         "axial: the start colour in the middle, the end colour at both edges, smooth (Skia) (" + Clip(html[12]) + ")");
        TestReport.Check(html[13].Contains("gradientUnits=\"userSpaceOnUse\" cx=\"0\" cy=\"0\" r=\"1\"><stop offset=\"0\" stop-color=\"#ffffff\"/><stop offset=\"1\" stop-color=\"#000000\"/>")
                         && html[13].Contains("scale(50.911688 50.911688)"), "radial: the end colour in the centre, radius half the diagonal (72 sqrt 2 / 2 = 50.91) (" + Clip(html[13]) + ")");
        TestReport.Check(html[14].Contains("scale(50.911688 50.911688)") && Count(html[14], "<stop ") == 256,
                         "ellipsoid: the box times sqrt 2, half of it; automatic steps banded (not drawn by Skia natively) (" + Clip(html[14]) + ")");
        TestReport.Check(Count(html[15], "<linearGradient") == 4 && Count(html[16], "<linearGradient") == 4, "square / rectangular: four edge regions");
        // 17: border 20 %, angle 90 (start on the left), 4 steps, start intensity 50 %: start #800000 (half red)
        TestReport.Check(html[17].Contains("rotate(-90)") && html[17].Contains("<stop offset=\"0.2\" stop-color=\"#800000\"/><stop offset=\"0.4\" stop-color=\"#800000\"/>"),
                         "linear: 90 degrees, 20 % border, four bands, start intensity (" + Clip(html[17]) + ")");
        TestReport.Check(html[18].Contains("stop-color=\"#00ff00\""), "LibreOffice multi-colour gradient stops (" + Clip(html[18]) + ")");
        TestReport.Check(html[19].Contains("<mask") && html[19].Contains("stop-color=\"#ffffff\"") && html[19].Contains("stop-color=\"#000000\"") && Count(html[19], "<stop ") == 4, "transparency gradient: a smooth luminance mask (Skia draws the alpha gradient natively) from opaque to clear (" + Clip(html[19]) + ")");
        TestReport.Check(html[20].Contains("fill-opacity=\"0.4\""), "draw:opacity 40 %");
        TestReport.Check(html[21].Contains("fill=\"#ffff00\"") && html[21].Contains("stroke=\"#0000ff\"") && html[21].Contains("rotate(-45)") && html[21].Contains("rotate(45)"),
                         "hatch: double (45 and -45 degrees), its colour, the solid background (" + Clip(html[21]) + ")");
        TestReport.Check(html[22].Contains("data-word-picture=\"f:Tile\"") && html[22].Contains("width:18pt;height:18pt"), "bitmap fill: tiles of 0.25 in from the fill image in the styles (" + Clip(html[22]) + ")");
        TestReport.Check(html[23].StartsWith("<div", StringComparison.Ordinal) && html[23].Contains("justify-content:center") && html[23].Contains(">1.</span>") && html[23].Contains("listtwo")
                         && html[23].Contains("<table") && html[23].Contains("cellone"), "text box: centred vertically, a numbered list, a table (no table inside a <p>) (" + Clip(html[23]) + ")");
        TestReport.Check(html[24].Contains("justify-content:flex-end;align-items:center;") && html[24].Contains("padding:3.6pt 7.2pt 3.6pt 7.2pt;"), "text area: bottom, centred, padding (" + Clip(html[24]) + ")");
        TestReport.Check(html[25].Contains("shape-outside:polygon(") && html[25].Contains("shape-margin:7.2pt") && html[25].Contains("float:left"), "contour wrap: the outline as shape-outside (" + Clip(html[25]) + ")");
        TestReport.Check(html[26].Contains("pagetext") && html[26].Contains("fill=\"#729fcf\"") && html[26].Contains("stroke=\"#3465a4\""),
                         "page-anchored: drawn where it stands; LibreOffice's default fill and line (" + Clip(html[26]) + ")");
        foreach ((string query, int expected) in new[] { ("owntext", 1), ("fontworkword", 1), ("turnedtext", 1), ("listone", 1), ("listtwo", 1), ("cellone", 1), ("areatext", 1), ("pagetext", 1), ("after box", 1), ("last", 1) })
        {
            int count = SearchCount(path, query);
            TestReport.Check(count == expected, $"search \"{query}\": {expected} ({count})");
        }
    }

    private static int SearchCount(string path, string query)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.Search(query, new WordDocumentSearchOptions(), 0, document.BlockCount).HitCounts.Sum();
    }

    /// <summary>LibreOffice writes presets with their full path and equations: drawn the same as from the tables.</summary>
    private static void CheckLibreOfficeSaved(string fodt, string odt)
    {
        List<string> own = Render(fodt), saved = Render(odt);
        string all = string.Concat(saved);
        // LibreOffice stores 1/100 mm more than written (ximpshap.cxx:880-897): 2.0004 in = 144.03 pt
        TestReport.Check(all.Contains("<path d=\"M36.01 0L144.03 72.03L0 72.03Z\""), "LibreOffice .odt: the triangle saved with its enhanced path draws the same");
        TestReport.Check(all.Contains("fill=\"#999999\"") && all.Contains("fill=\"#666666\""), "LibreOffice .odt: the cube's faces");
        TestReport.Check(all.Contains(">fontworkword</textPath>"), "LibreOffice .odt: Fontwork");
        // LibreOffice 25.x drops the list numbering and the table of the hand-written text box, and the page-anchored shape, on this conversion
        foreach (string query in new[] { "owntext", "turnedtext", "listone", "areatext" })
            TestReport.Check(SearchCount(odt, query) == 1, $"LibreOffice .odt search \"{query}\": 1");
        _ = own;
    }

    private static int Count(string text, string part)
    {
        int count = 0;
        for (int at = text.IndexOf(part, StringComparison.Ordinal); at >= 0; at = text.IndexOf(part, at + part.Length, StringComparison.Ordinal)) count++;
        return count;
    }

    private static string Clip(string html) => html.Length > 1500 ? html.Substring(0, 1500) + "..." : html;

    /// <summary>
    /// office:binary-data without its '=' padding (LibreOffice's own tdf145542 .fodt files): LibreOffice decodes the complete
    /// groups (comphelper Base64::decodeSomeChars), so the picture shows; the id the block carries must be served, never a 404.
    /// </summary>
    private static void UnpaddedBinaryPicture()
    {
        TestReport.Step("ODT drawing: a binary picture without base64 padding");
        string unpadded = Png.TrimEnd('=') + "\n  ";
        string path = Path.Combine(TestReport.OutputDirectory, "odt-unpadded-picture-v1.fodt");
        File.WriteAllText(path, "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document " + Namespaces + " office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\"><office:body><office:text>"
            + "<text:p>a<draw:frame text:anchor-type=\"as-char\" svg:width=\"1cm\" svg:height=\"1cm\"><draw:image><office:binary-data>\n  " + unpadded + "</office:binary-data></draw:image></draw:frame>b</text:p>"
            + "</office:text></office:body></office:document>");
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") });
        string html = document.RenderBlocks(0, 1)[0].Html;
        string? id = System.Text.RegularExpressions.Regex.Match(html, "data-word-picture=\"([^\"]*)\"") is { Success: true } m ? m.Groups[1].Value : null;
        WordDocumentPicture? picture = id is null ? null : document.ReadPicture(id);
        TestReport.Check(id is not null && picture is { ContentType: "image/png" } && picture.Bytes.Length > 8, $"the picture's id ({id}) is served as a PNG ({picture?.Bytes.Length ?? 0} bytes)");
    }
}
