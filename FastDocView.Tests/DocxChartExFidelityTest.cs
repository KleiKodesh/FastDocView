using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// .docx chartex charts ([MS-ODRAWXML] cx:chartSpace) drawn from their cached data (--fidelity, seconds): histogram with
/// binSize, pareto over aggregated categories, waterfall with a subtotal, funnel, box and whisker (exclusive quartiles, an
/// outlier), treemap and sunburst over two category levels, a region map from a "clear" geography cache (Bing point
/// compression rings, encoded here by the algorithm's inverse), and a region map cached only as "binary" falling back to its
/// mc:Fallback picture; every chart's document text found by search exactly once. Expected values worked out by hand.
/// </summary>
public static class DocxChartExFidelityTest
{
    private const string Namespaces = "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"";
    private const string Cx = "xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";

    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEElEQVR4nGP4z8DAwMDAAAAJ/gH/0wkYjQAAAABJRU5ErkJggg==");

    public static void Run()
    {
        TestReport.Step("chartex: histogram, pareto, waterfall, funnel, box and whisker, treemap, sunburst, region map");
        var charts = new[]
        {
            Space("Hist2", Data(Num([1, 5, 12, 18, 25, 33, 38])), Series("clusteredColumn", "Ages", "<cx:layoutPr><cx:binning intervalClosed=\"r\"><cx:binSize>10</cx:binSize></cx:binning></cx:layoutPr>")),
            Space(null, Data(Str(["A", "B", "A", "C"]) + Num([1, 4, 2, 1])), Series("clusteredColumn", "Defects", "<cx:layoutPr><cx:aggregation/></cx:layoutPr>") + "<cx:series layoutId=\"paretoLine\" ownerIdx=\"0\"><cx:axisId val=\"2\"/></cx:series>"),
            Space(null, Data(Str(["Start", "Loss", "End"]) + Num([10, -4, 6])), Series("waterfall", "Flow", "<cx:layoutPr><cx:subtotals><cx:idx val=\"2\"/></cx:subtotals></cx:layoutPr>")),
            Space(null, Data(Str(["Lead", "Deal", "Won"]) + Num([100, 50, 25])), Series("funnel", "Pipeline", "")),
            Space(null, Data(Num([1, 2, 3, 4, 5, 6, 7, 8, 100])), Series("boxWhisker", "Scores", "<cx:layoutPr><cx:statistics quartileMethod=\"exclusive\"/></cx:layoutPr>")),
            Space(null, Data(Levels(["Xone", "Xtwo", "Yone"], ["Xtop", "", "Ytop"]) + Num([4, 2, 2])), Series("treemap", "Tree", "<cx:layoutPr><cx:parentLabelLayout val=\"banner\"/></cx:layoutPr>")),
            Space(null, Data(Levels(["Sone", "Stwo", "Sthree"], ["Sleft", "", "Sright"]) + Num([1, 1, 2])), Series("sunburst", "Sun", "")),
            Space(null, Data(Str(["Northland", "Southland"]) + Num([1, 3])), Series("regionMap", "Map", Geography(), "<cx:valueColors><cx:minColor><a:srgbClr val=\"FF0000\"/></cx:minColor><cx:maxColor><a:srgbClr val=\"0000FF\"/></cx:maxColor></cx:valueColors>")),
            Space(null, Data(Str(["Northland"]) + Num([1])), Series("regionMap", "Opaque", "<cx:layoutPr><cx:geography cultureLanguage=\"en-US\" cultureRegion=\"US\" attribution=\"x\"><cx:geoCache provider=\"{x}\"><cx:binary>AAAA</cx:binary></cx:geoCache></cx:geography></cx:layoutPr>")),
        };
        string body = string.Concat(Enumerable.Range(0, charts.Length).Select(i => Alternate("rIdX" + i)));
        string relationships = "<Relationship Id=\"rIdTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rIdP\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/p.png\"/>"
            + string.Concat(Enumerable.Range(0, charts.Length).Select(i => "<Relationship Id=\"rIdX" + i + "\" Type=\"http://schemas.microsoft.com/office/2014/relationships/chartEx\" Target=\"charts/chartEx" + i + ".xml\"/>"));
        var parts = new Dictionary<string, byte[]> { ["word/media/p.png"] = Png, ["word/theme/theme1.xml"] = new UTF8Encoding(false).GetBytes(DocxDrawingPartsFidelityTest.Theme) };
        for (int i = 0; i < charts.Length; i++) parts["word/charts/chartEx" + i + ".xml"] = new UTF8Encoding(false).GetBytes(charts[i]);
        List<string> html = DocxFidelityTest.RenderAll("chartex", body, extraDocumentRelationships: relationships, binaryParts: parts,
            inspect: document =>
            {
                foreach ((string text, int expected) in new[] { ("Hist2", 1), ("Lead", 1), ("Won", 1), ("Xone", 1), ("Sthree", 1), ("Northland", 1), ("Southland", 1), ("Defects", 1), ("[1, 11]", 0) })
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

        static List<double> Heights(string svg, string fill) => Regex.Matches(svg, "<rect x=\"[0-9.]+\" y=\"[0-9.]+\" width=\"([0-9.]+)\" height=\"([0-9.]+)\" fill=\"" + fill + "\"").Cast<Match>()
            .Where(m => m.Groups[1].Value != "7").Select(m => double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture)).ToList();
        static bool Ratio(double a, double b, double ratio) => b > 0 && Math.Abs(a / b - ratio) < 0.03;

        // binSize 10 from the minimum 1, right-closed: [1, 11] (11, 21] (21, 31] (31, 41] hold 2, 2, 1, 2
        var hist = Heights(html[0], "#4472C4");
        TestReport.Check(html[0].Contains(">[1, 11]<") && html[0].Contains(">(31, 41]<") && hist.Count == 4 && Ratio(hist[0], hist[2], 2) && Ratio(hist[3], hist[2], 2), "histogram: bins of 10 from the minimum, counts 2 2 1 2 (" + html[0] + ")");
        // aggregation: A 3, B 4, C 1, sorted for the pareto: B A C
        var pareto = Heights(html[1], "#4472C4");
        TestReport.Check(pareto.Count == 3 && Ratio(pareto[0], pareto[2], 4) && Ratio(pareto[1], pareto[2], 3) && html[1].Contains("stroke=\"#ED7D31\" stroke-width=\"2.25\""), "pareto: aggregated columns largest first and the cumulative line (" + html[1] + ")");
        TestReport.Check(Heights(html[2], "#4472C4").Count == 1 && Heights(html[2], "#ED7D31").Count == 1 && Heights(html[2], "#A5A5A5").Count == 1 && Ratio(Heights(html[2], "#4472C4")[0], Heights(html[2], "#ED7D31")[0], 2.5),
                         "waterfall: increase 10, decrease 4 (floating), the subtotal 6 from zero (" + html[2] + ")");
        var funnel = Regex.Matches(html[3], "<rect x=\"[0-9.]+\" y=\"[0-9.]+\" width=\"([0-9.]+)\" height=\"[0-9.]+\" fill=\"#4472C4\"").Cast<Match>().Select(m => double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture)).Where(w => w != 7).ToList();
        TestReport.Check(funnel.Count == 3 && Ratio(funnel[0], funnel[1], 2) && Ratio(funnel[1], funnel[2], 2), "funnel: bar widths 100 : 50 : 25 (" + html[3] + ")");
        TestReport.Check(Regex.Matches(html[4], "<circle").Count == 1, "box and whisker: Q1 2.5, Q3 7.5, the whisker to 8, 100 an outlier (" + html[4] + ")");
        TestReport.Check(html[5].Contains(">Xone<") && html[5].Contains(">Xtop<") && Regex.Matches(html[5], "stroke=\"#FFFFFF\"/>").Count == 5, "treemap: banners for Xtop / Ytop, three leaves (" + html[5] + ")");
        TestReport.Check(Regex.Matches(html[6], "<path d=\"M[^\"]*Z\" fill=\"#[0-9A-F]{6}\" stroke=\"#FFFFFF\"/>").Count == 5 && html[6].Contains(">Sthree<"), "sunburst: two parents and three leaves (" + html[6] + ")");
        TestReport.Check(html[7].Contains("fill=\"#FF0000\" fill-rule=\"evenodd\"") && html[7].Contains("fill=\"#0000FF\" fill-rule=\"evenodd\""), "region map: Northland at the minimum colour, Southland at the maximum (" + html[7] + ")");
        TestReport.Check(html[8].Contains("data-word-picture=\"m:rIdP\"") && !html[8].Contains("word-chart"), "region map with a binary geography cache: the mc:Fallback picture (" + html[8] + ")");
    }

    private static string Space(string? title, string data, string series)
        => "<?xml version=\"1.0\"?><cx:chartSpace " + Cx + "><cx:chartData>" + data + "</cx:chartData><cx:chart>"
         + (title is null ? "" : "<cx:title pos=\"t\" align=\"ctr\" overlay=\"0\"><cx:tx><cx:txData><cx:v>" + title + "</cx:v></cx:txData></cx:tx></cx:title>")
         + "<cx:plotArea><cx:plotAreaRegion>" + series + "</cx:plotAreaRegion><cx:axis id=\"0\"><cx:catScaling gapWidth=\"0\"/></cx:axis><cx:axis id=\"1\"><cx:valScaling/></cx:axis></cx:plotArea>"
         + "<cx:legend pos=\"r\" align=\"ctr\" overlay=\"0\"/></cx:chart></cx:chartSpace>";

    private static string Data(string dims) => "<cx:data id=\"0\">" + dims + "</cx:data>";

    private static string Num(double[] values)
        => "<cx:numDim type=\"val\"><cx:f>Sheet1!$B$2</cx:f><cx:lvl ptCount=\"" + values.Length + "\" formatCode=\"General\">"
         + string.Concat(values.Select((v, i) => "<cx:pt idx=\"" + i + "\">" + v.ToString(CultureInfo.InvariantCulture) + "</cx:pt>")) + "</cx:lvl></cx:numDim>";

    private static string Str(string[] values) => "<cx:strDim type=\"cat\"><cx:f>Sheet1!$A$2</cx:f>" + Level(values) + "</cx:strDim>";

    private static string Levels(string[] leaves, string[] parents) => "<cx:strDim type=\"cat\"><cx:f>Sheet1!$A$2</cx:f>" + Level(leaves) + Level(parents) + "</cx:strDim>";

    private static string Level(string[] values)
        => "<cx:lvl ptCount=\"" + values.Length + "\">" + string.Concat(values.Select((v, i) => v.Length == 0 ? "" : "<cx:pt idx=\"" + i + "\">" + v + "</cx:pt>")) + "</cx:lvl>";

    private static string Series(string layout, string name, string layoutPr, string valueColors = "")
        => "<cx:series layoutId=\"" + layout + "\" uniqueId=\"{1}\"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>" + name + "</cx:v></cx:txData></cx:tx>" + valueColors
         + "<cx:dataId val=\"0\"/>" + layoutPr + "</cx:series>";

    /// <summary>Two square entities (0..1 and -2..-1 degrees) with their locations and point-to-entity mapping, rings encoded by the Bing point compression algorithm.</summary>
    private static string Geography()
    {
        string Square(double lat, double lon) => "1," + Encode([(lat, lon), (lat, lon + 1), (lat + 1, lon + 1), (lat + 1, lon), (lat, lon)]);
        string Entity(string id, string name, double lat, double lon)
            => "<cx:geoDataEntityQueryResult><cx:geoDataEntityQuery entityType=\"CountryRegion\" entityId=\"" + id + "\"/><cx:geoData entityName=\"" + name + "\" entityId=\"" + id + "\" east=\"0\" west=\"0\" north=\"0\" south=\"0\">"
             + "<cx:geoPolygons><cx:geoPolygon polygonId=\"p" + id + "\" numPoints=\"5\" pcaRings=\"" + Square(lat, lon) + "\"/></cx:geoPolygons></cx:geoData></cx:geoDataEntityQueryResult>";
        string Location(string name, double lat, double lon)
            => "<cx:geoLocationQueryResult><cx:geoLocationQuery countryRegion=\"" + name + "\" entityType=\"CountryRegion\"/><cx:geoLocations><cx:geoLocation latitude=\"" + lat.ToString(CultureInfo.InvariantCulture)
             + "\" longitude=\"" + lon.ToString(CultureInfo.InvariantCulture) + "\" entityName=\"" + name + "\" entityType=\"CountryRegion\"/></cx:geoLocations></cx:geoLocationQueryResult>";
        string ToEntity(string id, double lat, double lon)
            => "<cx:geoDataPointToEntityQueryResult><cx:geoDataPointQuery entityType=\"CountryRegion\" latitude=\"" + lat.ToString(CultureInfo.InvariantCulture) + "\" longitude=\"" + lon.ToString(CultureInfo.InvariantCulture)
             + "\"/><cx:geoDataPointToEntityQuery entityType=\"CountryRegion\" entityId=\"" + id + "\"/></cx:geoDataPointToEntityQueryResult>";
        return "<cx:layoutPr><cx:geography projectionType=\"mercator\" cultureLanguage=\"en-US\" cultureRegion=\"US\" attribution=\"x\"><cx:geoCache provider=\"{x}\"><cx:clear>"
             + "<cx:geoLocationQueryResults>" + Location("Northland", 0.5, 0.5) + Location("Southland", -1.5, -1.5) + "</cx:geoLocationQueryResults>"
             + "<cx:geoDataEntityQueryResults>" + Entity("1", "Northland", 0, 0) + Entity("2", "Southland", -2, -2) + "</cx:geoDataEntityQueryResults>"
             + "<cx:geoDataPointToEntityQueryResults>" + ToEntity("1", 0.5, 0.5) + ToEntity("2", -1.5, -1.5) + "</cx:geoDataPointToEntityQueryResults>"
             + "</cx:clear></cx:geoCache></cx:geography></cx:layoutPr>";
    }

    /// <summary>The Bing Maps point compression algorithm (the inverse of its documented decoder): delta, zig-zag, diagonal pairing, 5-bit groups.</summary>
    private static string Encode((double Lat, double Lon)[] points)
    {
        const string safe = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-";
        var text = new StringBuilder();
        long latitude = 0, longitude = 0;
        foreach ((double lat, double lon) in points)
        {
            long newLat = (long)Math.Round(lat * 100000), newLon = (long)Math.Round(lon * 100000);
            long dy = newLat - latitude, dx = newLon - longitude;
            latitude = newLat; longitude = newLon;
            dy = (dy << 1) ^ (dy >> 63); dx = (dx << 1) ^ (dx >> 63);
            long index = (dy + dx) * (dy + dx + 1) / 2 + dy;
            do
            {
                long rem = index & 31;
                index = (index - rem) / 32;
                if (index > 0) rem += 32;
                text.Append(safe[(int)rem]);
            } while (index > 0);
        }
        return text.ToString();
    }

    private static string Alternate(string id)
        => $"<w:p><w:r><mc:AlternateContent {Namespaces}><mc:Choice Requires=\"cx1\"><w:drawing><wp:inline><wp:extent cx=\"2743200\" cy=\"1828800\"/><wp:docPr id=\"2\" name=\"x\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2014/chartex\">"
         + $"<cx:chart xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{id}\"/></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>"
         + "<mc:Fallback><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"3\" name=\"p\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"3\" name=\"p\"/><pic:cNvPicPr/></pic:nvPicPr>"
         + "<pic:blipFill><a:blip r:embed=\"rIdP\"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Fallback></mc:AlternateContent></w:r></w:p>";
}
