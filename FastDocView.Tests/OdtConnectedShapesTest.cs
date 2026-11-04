using System.Text;
using System.Text.RegularExpressions;
using FastDocView.Odt;

namespace FastDocView.Tests;

/// <summary>
/// A draw:connector glued to shapes in OTHER blocks (draw:start-shape / draw:end-shape): the shapes' geometry is recorded at index time
/// (OdtConnectedShapeTable, index format 10) and the connector without a usable svg:d is routed around them as LibreOffice does
/// (SdrEdgeObj::ImpCalcEdgeTrack, svx/source/svdraw/svdoedge.cxx) - never by reading the other blocks at render time.
/// </summary>
public static class OdtConnectedShapesTest
{
    private const string Fixture = "cross-block-connectors-v1.fodt";

    public static void Run()
    {
        TestReport.Step("ODT connectors glued to shapes in other blocks");
        string path = Path.Combine(TestReport.OutputDirectory, Fixture);
        File.WriteAllText(path, Document(), new UTF8Encoding(false));
        foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
        {
            string cache = Path.Combine(TestReport.OutputDirectory, "index-cache-connectors-" + mode);
            if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
            var policy = new WordIndexPolicy { CacheDirectory = cache, ForcedMode = mode };
            List<string> first, second;
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
                first = Enumerable.Range(0, document.BlockCount).Select(i => document.RenderBlocks(i, 1)[0].Html).ToList();
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))       // sidecar: the stored table
                second = Enumerable.Range(0, document.BlockCount).Select(i => document.RenderBlocks(i, 1)[0].Html).ToList();
            TestReport.Check(first.SequenceEqual(second), $"{mode}: a reopen draws the connectors the same");
            // page-anchored, block 2: from A's right glue point (vertex 1) out to the right first (escape right), into B's left side
            string page = ConnectorPath(first[2]);
            TestReport.Check(FirstSegmentIsHorizontal(page), $"{mode}: page-anchored connector leaves shape A sideways through its right glue point ({page})");
            // paragraph-anchored, its shapes in blocks 3 and 5, best glue points (-1): routed around the shapes too
            string paragraph = ConnectorPath(first[4]);
            TestReport.Check(FirstSegmentIsHorizontal(paragraph), $"{mode}: paragraph-anchored connector routed around shapes of other paragraphs ({paragraph})");
            // without the shapes the same ends would be routed as free points: a vertical Z (OdtConnectorRouteTest)
        }
    }

    private static string ConnectorPath(string html)
    {
        MatchCollection paths = Regex.Matches(html, "<path[^>]* d=\"([^\"]+)\"");
        return paths.Count == 0 ? "" : paths[paths.Count - 1].Groups[1].Value;
    }

    private static bool FirstSegmentIsHorizontal(string d)
    {
        MatchCollection numbers = Regex.Matches(d, @"-?\d+(\.\d+)?");
        if (numbers.Count < 4) return false;
        double y0 = double.Parse(numbers[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        double x0 = double.Parse(numbers[0].Value, System.Globalization.CultureInfo.InvariantCulture);
        double x1 = double.Parse(numbers[2].Value, System.Globalization.CultureInfo.InvariantCulture);
        double y1 = double.Parse(numbers[3].Value, System.Globalization.CultureInfo.InvariantCulture);
        return Math.Abs(y1 - y0) < 0.01 && Math.Abs(x1 - x0) > 0.01;
    }

    private static string Document() => """
        <?xml version="1.0" encoding="UTF-8"?>
        <office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
          xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
          xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
          office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
        <office:automatic-styles>
          <style:style style:name="G" style:family="graphic"><style:graphic-properties style:wrap="none" draw:fill="solid" draw:fill-color="#ffcc00" draw:stroke="solid" svg:stroke-color="#000000"/></style:style>
          <style:style style:name="L" style:family="graphic"><style:graphic-properties style:wrap="none" draw:fill="none" draw:stroke="solid" svg:stroke-color="#0000ff"/></style:style>
        </office:automatic-styles>
        <office:body><office:text>
        <draw:rect draw:style-name="G" draw:id="A" text:anchor-type="page" text:anchor-page-number="1" svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="1cm"/>
        <draw:rect draw:style-name="G" draw:id="B" text:anchor-type="page" text:anchor-page-number="1" svg:x="6cm" svg:y="4cm" svg:width="2cm" svg:height="1cm"/>
        <draw:connector draw:style-name="L" draw:type="standard" text:anchor-type="page" text:anchor-page-number="1" draw:start-shape="A" draw:start-glue-point="1" draw:end-shape="B" draw:end-glue-point="3" svg:x1="3cm" svg:y1="1.5cm" svg:x2="6cm" svg:y2="4.5cm"/>
        <text:p><draw:rect draw:style-name="G" draw:id="C" text:anchor-type="paragraph" svg:x="1cm" svg:y="0cm" svg:width="2cm" svg:height="1cm"/>first shape</text:p>
        <text:p><draw:connector draw:style-name="L" draw:type="standard" text:anchor-type="paragraph" draw:start-shape="C" draw:end-shape="D" svg:x1="3cm" svg:y1="0.5cm" svg:x2="6cm" svg:y2="3.5cm"/>connector</text:p>
        <text:p><draw:rect draw:style-name="G" draw:id="D" text:anchor-type="paragraph" svg:x="6cm" svg:y="3cm" svg:width="2cm" svg:height="1cm"/>second shape</text:p>
        </office:text></office:body></office:document>
        """;
}
