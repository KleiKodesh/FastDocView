using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// .docx SmartArt and charts (--fidelity, seconds): a SmartArt drawn from its diagramDrawing part (shapes, text in the
/// txXfrm box, centred), one without a drawing part falling back to its mc:Fallback picture, a chartex chart falling back
/// likewise, and charts from their cached values: clustered columns (bar heights in the values' ratio, automatic colours =
/// accents 1 and 2 of chart style 2, LibreOffice objectformatter.cxx), a varyColors pie (one accent per point, slice sweeps in
/// the values' ratio), a line chart, a scatter chart and a stacked area chart; every SmartArt and chart text found by search
/// exactly once (title, legend, categories, data labels, shape texts).
/// </summary>
public static class DocxDrawingPartsFidelityTest
{
    private const string Namespaces = "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"";
    private const string A = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";
    private const string C = "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";

    internal const string Theme = """
        <?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>
        <a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
        <a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
        <a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst></a:fmtScheme>
        </a:themeElements></a:theme>
        """;

    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEElEQVR4nGP4z8DAwMDAAAAJ/gH/0wkYjQAAAABJRU5ErkJggg==");

    public static void Run()
    {
        TestReport.Step("SmartArt and charts: drawing parts, fallbacks, cached-value charts, search");
        string body = Graphic("diagram", "<dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:dm=\"rIdDm\" r:lo=\"rIdLo\" r:qs=\"rIdQs\" r:cs=\"rIdCs\"/>")
            + Alternate("<dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:dm=\"rIdDm2\" r:lo=\"rIdLo\" r:qs=\"rIdQs\" r:cs=\"rIdCs\"/>", "diagram")
            + Chart("rIdC1") + Chart("rIdC2") + Chart("rIdC3") + Chart("rIdC4") + Chart("rIdC5")
            + Alternate("<cx:chart xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdCx\"/>", "chartex");
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        string relationships = "<Relationship Id=\"rIdTheme\" Type=\"" + R + "theme\" Target=\"theme/theme1.xml\"/>"
            + "<Relationship Id=\"rIdP\" Type=\"" + R + "image\" Target=\"media/p.png\"/>"
            + "<Relationship Id=\"rIdDm\" Type=\"" + R + "diagramData\" Target=\"diagrams/data1.xml\"/>"
            + "<Relationship Id=\"rIdDm2\" Type=\"" + R + "diagramData\" Target=\"diagrams/data2.xml\"/>"
            + "<Relationship Id=\"rIdDr\" Type=\"http://schemas.microsoft.com/office/2007/relationships/diagramDrawing\" Target=\"diagrams/drawing1.xml\"/>"
            + string.Concat(Enumerable.Range(1, 5).Select(i => "<Relationship Id=\"rIdC" + i + "\" Type=\"" + R + "chart\" Target=\"charts/chart" + i + ".xml\"/>"));
        var parts = new Dictionary<string, byte[]>
        {
            ["word/theme/theme1.xml"] = Utf8(Theme), ["word/media/p.png"] = Png,
            ["word/diagrams/data1.xml"] = Utf8(DataModel("rIdDr")), ["word/diagrams/data2.xml"] = Utf8(DataModel(null)), ["word/diagrams/drawing1.xml"] = Utf8(Drawing()),
            ["word/charts/chart1.xml"] = Utf8(ChartSpace("<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>"
                + Series(0, "Alpha", ["Q1", "Q2", "Q3"], [10, 20, 40], "<c:dLbls><c:showVal val=\"1\"/></c:dLbls>") + Series(1, "Beta", ["Q1", "Q2", "Q3"], [5, 15, 30])
                + "<c:gapWidth val=\"150\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" + Axes(), "Sales")),
            ["word/charts/chart2.xml"] = Utf8(ChartSpace("<c:pieChart><c:varyColors val=\"1\"/>" + Series(0, "Share", ["North", "South", "East"], [50, 25, 25]) + "<c:firstSliceAng val=\"0\"/></c:pieChart>", null)),
            ["word/charts/chart3.xml"] = Utf8(ChartSpace("<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + Series(0, "Trend", ["Jan", "Feb", "Mar"], [1, 3, 2]) + "<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" + Axes(), null)),
            ["word/charts/chart4.xml"] = Utf8(ChartSpace("<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx><c:v>Points</c:v></c:tx>"
                + "<c:xVal><c:numLit><c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt><c:pt idx=\"1\"><c:v>2</c:v></c:pt><c:pt idx=\"2\"><c:v>4</c:v></c:pt></c:numLit></c:xVal>"
                + "<c:yVal><c:numLit><c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>2</c:v></c:pt><c:pt idx=\"1\"><c:v>4</c:v></c:pt><c:pt idx=\"2\"><c:v>8</c:v></c:pt></c:numLit></c:yVal></c:ser>"
                + "<c:axId val=\"3\"/><c:axId val=\"2\"/></c:scatterChart><c:valAx><c:axId val=\"3\"/><c:scaling/><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:valAx>"
                + "<c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:crossAx val=\"3\"/></c:valAx>", null)),
            ["word/charts/chart5.xml"] = Utf8(ChartSpace("<c:areaChart><c:grouping val=\"stacked\"/><c:varyColors val=\"0\"/>" + Series(0, "Lower", ["A", "B"], [2, 2]) + Series(1, "Upper", ["A", "B"], [2, 6])
                + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:areaChart>" + Axes(), null)),
        };
        List<string> html = DocxFidelityTest.RenderAll("drawing-parts", body, extraDocumentRelationships: relationships, binaryParts: parts,
            inspect: document =>
            {
                foreach ((string text, int expected) in new[] { ("smartone", 1), ("smarttwo", 1), ("Sales", 1), ("Alpha", 1), ("Beta", 1), ("Q2", 1), ("North", 1), ("Trend", 1), ("Jan", 1), ("Points", 1), ("Upper", 1), ("missingsmart", 0) })
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
                }
            });

        // SmartArt: two dsp:sp, the first a roundRect in accent1 with its text in the txXfrm box (36pt in, 9pt down: 457200 / 114300 EMU), centred
        TestReport.Check(html[0].Contains("fill=\"#4472C4\"") && html[0].Contains("left:36pt;top:9pt;width:72pt;height:18pt;") && html[0].Contains("justify-content:center")
                         && html[0].Contains(">smartone<") && html[0].Contains(">smarttwo<") && html[0].Contains("font-size:14pt;"), "SmartArt drawn from its drawing part (" + html[0] + ")");
        TestReport.Check(html[1].Contains("data-word-picture=\"m:rIdP\"") && !html[1].Contains("word-shape-geometry"), "SmartArt without a drawing part: the mc:Fallback picture (" + html[1] + ")");
        string bars = html[2];
        var rects = Regex.Matches(bars, "<rect x=\"([0-9.]+)\" y=\"([0-9.]+)\" width=\"([0-9.]+)\" height=\"([0-9.]+)\" fill=\"(#[0-9A-F]{6})\"").Cast<Match>()
            .Where(m => m.Groups[3].Value != "7")   // not the legend keys
            .Select(m => (H: double.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture), Fill: m.Groups[5].Value)).ToList();
        var alpha = rects.Where(r => r.Fill == "#4472C4").ToList();
        var beta = rects.Where(r => r.Fill == "#ED7D31").ToList();
        TestReport.Check(alpha.Count == 3 && beta.Count == 3 && Math.Abs(alpha[1].H / alpha[0].H - 2) < 0.02 && Math.Abs(alpha[2].H / beta[2].H - 40.0 / 30) < 0.02,
                         "columns: 3 + 3 bars in accent 1 / accent 2, heights in the values' ratio (" + bars + ")");
        TestReport.Check(bars.Contains(">Sales<") && bars.Contains(">Alpha<") && bars.Contains(">Q3<") && bars.Contains(">40<"), "column chart: title, legend, categories, data labels (showVal)");
        var slices = Regex.Matches(html[3], "<path d=\"[^\"]*\" fill=\"(#[0-9A-F]{6})\"").Cast<Match>().Select(m => m.Groups[1].Value).ToList();
        TestReport.Check(slices.SequenceEqual(new[] { "#4472C4", "#ED7D31", "#A5A5A5" }) && html[3].Contains(" 0 1 1 ") == false && html[3].Contains(">North<"),
                         "pie with varyColors: one accent per point, a half and two quarters, the legend lists the points (" + html[3] + ")");
        TestReport.Check(Regex.IsMatch(html[4], "<path d=\"M[^\"]*L[^\"]*L[^\"]*\" fill=\"none\" stroke=\"#4472C4\"") && Regex.Matches(html[4], "<circle").Count == 3 && html[4].Contains(">Feb<"),
                         "line chart: one line in accent 1 through 3 markers, categories (" + html[4] + ")");
        TestReport.Check(Regex.Matches(html[5], "<circle").Count == 3 && html[5].Contains("stroke=\"#4472C4\""), "scatter chart: 3 points joined (" + html[5] + ")");
        TestReport.Check(Regex.Matches(html[6], "<path d=\"M[^\"]*Z\" fill=\"#(4472C4|ED7D31)\"").Count == 2, "stacked area: two filled bands (" + html[6] + ")");
        TestReport.Check(html[7].Contains("data-word-picture=\"m:rIdP\"") && !html[7].Contains("word-chart"), "chartex: the mc:Fallback picture (" + html[7] + ")");
    }

    private static byte[] Utf8(string text) => new UTF8Encoding(false).GetBytes(text);

    private static string Graphic(string kind, string inner)
        => $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"2743200\" cy=\"1371600\"/><wp:docPr id=\"1\" name=\"d\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/{kind}\">{inner}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    private static string Chart(string id)
        => Graphic("chart", "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"" + id + "\"/>");

    private static string Alternate(string inner, string kind)
        => $"<w:p><w:r><mc:AlternateContent {Namespaces}><mc:Choice Requires=\"cx1\"><w:drawing><wp:inline><wp:extent cx=\"2743200\" cy=\"1371600\"/><wp:docPr id=\"2\" name=\"x\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/{kind}\">{inner}</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>"
         + "<mc:Fallback><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"3\" name=\"p\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"3\" name=\"p\"/><pic:cNvPicPr/></pic:nvPicPr>"
         + "<pic:blipFill><a:blip r:embed=\"rIdP\"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Fallback></mc:AlternateContent></w:r></w:p>";

    private static string DataModel(string? drawing)
        => "<?xml version=\"1.0\"?><dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" " + A + "><dgm:ptLst><dgm:pt modelId=\"1\"><dgm:t><a:bodyPr/><a:p><a:r><a:t>missingsmart</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst>"
         + (drawing is null ? "" : "<dgm:extLst><a:ext uri=\"http://schemas.microsoft.com/office/drawing/2008/diagram\"><dsp:dataModelExt xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" relId=\"" + drawing + "\" minVer=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/></a:ext></dgm:extLst>")
         + "</dgm:dataModel>";

    private static string Drawing()
        => "<?xml version=\"1.0\"?><dsp:drawing xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" " + A + "><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>"
         + "<dsp:sp modelId=\"{1}\"><dsp:nvSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1371600\" cy=\"685800\"/></a:xfrm><a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom>"
         + "<a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></dsp:spPr><dsp:style><a:lnRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:lnRef><a:fillRef idx=\"1\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:fillRef><a:effectRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></dsp:style>"
         + "<dsp:txBody><a:bodyPr spcFirstLastPara=\"0\" vert=\"horz\" wrap=\"square\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"ctr\" anchorCtr=\"0\"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr lang=\"en-US\" sz=\"1400\"/><a:t>smartone</a:t></a:r></a:p></dsp:txBody>"
         + "<dsp:txXfrm><a:off x=\"457200\" y=\"114300\"/><a:ext cx=\"914400\" cy=\"228600\"/></dsp:txXfrm></dsp:sp>"
         + "<dsp:sp modelId=\"{2}\"><dsp:nvSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x=\"1371600\" y=\"685800\"/><a:ext cx=\"1371600\" cy=\"685800\"/></a:xfrm><a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>"
         + "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill></dsp:spPr><dsp:txBody><a:bodyPr anchor=\"ctr\"/><a:lstStyle/><a:p><a:r><a:t>smarttwo</a:t></a:r></a:p></dsp:txBody></dsp:sp>"
         + "</dsp:spTree></dsp:drawing>";

    private static string ChartSpace(string plot, string? title)
        => "<?xml version=\"1.0\"?><c:chartSpace " + C + "><c:chart>"
         + (title is null ? "<c:autoTitleDeleted val=\"1\"/>" : "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>" + title + "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title><c:autoTitleDeleted val=\"0\"/>")
         + "<c:plotArea><c:layout/>" + plot + "</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:overlay val=\"0\"/></c:legend><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>";

    private static string Axes()
        => "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>"
         + "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:crossAx val=\"1\"/></c:valAx>";

    private static string Series(int index, string name, string[] categories, double[] values, string labels = "")
        => $"<c:ser><c:idx val=\"{index}\"/><c:order val=\"{index}\"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>{labels}"
         + "<c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val=\"" + categories.Length + "\"/>" + string.Concat(categories.Select((c, i) => $"<c:pt idx=\"{i}\"><c:v>{c}</c:v></c:pt>")) + "</c:strCache></c:strRef></c:cat>"
         + "<c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"" + values.Length + "\"/>"
         + string.Concat(values.Select((v, i) => $"<c:pt idx=\"{i}\"><c:v>{v.ToString(CultureInfo.InvariantCulture)}</c:v></c:pt>")) + "</c:numCache></c:numRef></c:val></c:ser>";
}
