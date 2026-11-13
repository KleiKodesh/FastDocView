using System.Xml.Linq;
using System.Text.RegularExpressions;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Charts in .doc and RTF (--fidelity / --charts, seconds): test-data/word-charts-binary/word-charts-binary-v1, made by Word 16
/// through COM (scripts/make-word-binary-charts.ps1: column, bar, line, pie, area, scatter and 3-D column charts with ASCII
/// data, saved as .docx, .doc and .rtf). Word writes each chart into the .doc / RTF as an embedded "Excel.Chart.8" object
/// (BIFF8 chart records + their cached values) - except the line chart, which Word 16 saves as a plain picture. Each object
/// chart is drawn as SVG by the .docx chart renderer, not as its preview picture, with the data, titles and Word's colours,
/// and its texts are found by search.
/// </summary>
public static class BinaryChartsFidelityTest
{
    private static readonly string[] Titles = ["Column Chart", "Bar Chart", "Pie Chart", "Area Chart", "Scatter Chart", "Depth Column Chart"];

    public static void Run()
    {
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-charts-binary");
        // the same charts as Word's .docx, for comparing layouts (scripts/compare-chart-text-positions.py)
        string docx = Path.Combine(folder, "word-charts-binary-v3.docx");
        if (File.Exists(docx)) File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "charts-binary-docx.html"), string.Concat(Render(docx)));
        foreach (string extension in new[] { ".doc", ".rtf" })
        {
            TestReport.Step("Charts in " + extension + ": Word's embedded Excel chart objects drawn from their chart records");
            string path = Path.Combine(folder, "word-charts-binary-v3" + extension);
            TestReport.Check(File.Exists(path), extension + ": chart fixture exists");
            if (!File.Exists(path)) continue;
            List<string> html = Render(path);
            File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "charts-binary" + extension.Replace('.', '-') + ".html"), string.Concat(html));
            string all = string.Concat(html);
            int drawn = Regex.Matches(all, "<svg class=\"word-chart\"").Count;
            TestReport.Check(drawn == Titles.Length, extension + ": every chart object drawn as SVG (" + drawn + " of " + Titles.Length + ")");
            foreach (string title in Titles) TestReport.Check(all.Contains(">" + title + "<"), extension + ": chart title drawn: " + title);
            string Block(string marker) => html.FirstOrDefault(h => h.Contains(">" + marker + "<")) ?? "";
            string column = Block("Column Chart");
            TestReport.Check(Block("Area Chart").Contains(">Q1<") && Block("Area Chart").Contains(">Q4<") && column.Contains(">North<") && column.Contains(">West<"), extension + ": categories and series names (legend)");
            TestReport.Check(column.Contains("#4472C4") && column.Contains("#ED7D31") && column.Contains("#A5A5A5"), extension + ": Word's series colours (from the OfficeArt format records)");
            TestReport.Check(column.Contains(">5.0<") && column.Contains(">0.0<"), extension + ": value axis scaled from the cached values (0 to 5)");
            string pie = Block("Pie Chart");
            TestReport.Check(Regex.Matches(pie, "<path[^>]*fill=\"#[0-9A-F]{6}\"").Count >= 4, extension + ": pie slices");
            TestReport.Check(Block("Scatter Chart").Contains("<circle") || Block("Scatter Chart").Contains("word-chart-marker"), extension + ": scatter points");
            TestReport.Check(Block("Depth Column Chart").Contains("word-chart-box"), extension + ": 3-D columns as projected boxes");
            TestReport.Check(Block("Bar Chart").Contains(">Q1<"), extension + ": horizontal bars with category names");
            TestReport.Check(SearchCount(path, "Scatter Chart") == 1 && SearchCount(path, "Depth Column Chart") == 1, extension + ": chart titles found by search, once each");
            TestReport.Check(SearchCount(path, "PieChart paragraph") == 1, extension + ": the paragraph's own text still found");
            // v2's column chart (scripts/make-word-binary-charts-v2.py): one point's label, a linear trend line, y error bars, the value axis as 0.0, no category labels, a red title
            TestReport.Check(column.Contains(">2.4<") || column.Contains(">2.5<"), extension + ": the one point's data label");
            TestReport.Check(column.Contains("word-chart-trendline"), extension + ": trend line from the trend line series");
            TestReport.Check(System.Text.RegularExpressions.Regex.Matches(column, "word-chart-error-bar").Count == 4, extension + ": an error bar per point of the first series");
            TestReport.Check(column.Contains(">1.0<") && column.Contains(">5.0<"), extension + ": value axis in its 0.0 number format");
            TestReport.Check(!column.Contains(">Q1<"), extension + ": no category labels (axis labels turned off)");
            TestReport.Check(Regex.IsMatch(column, "fill=\"#FF0000\"[^>]*>Column Chart<"), extension + ": the title's own colour");
            TestReport.Check(Regex.Matches(all, "<svg class=\"word-chart\"").Count == Titles.Length, extension + ": the trend line and error bar series are not drawn as series");
            // v3: a blue value axis, a green legend; the scatter chart's first series with custom plus-only x and fixed y error bars
            TestReport.Check(Regex.IsMatch(column, "fill=\"#0000FF\"[^>]*>5[.]0<") && Regex.IsMatch(column, "fill=\"#339966\"[^>]*>North<"), extension + ": axis label and legend text colours (the legend's 00B050 is saved as its nearest palette colour, 339966, in both the RGB and the index)");
            string scatter = Block("Scatter Chart");
            TestReport.Check(Regex.Matches(scatter, "word-chart-error-bar").Count == 8, extension + ": x (custom) and y (fixed) error bars, one per point each (" + Regex.Matches(scatter, "word-chart-error-bar").Count + ")");
        }
    }

    /// <summary>Word 16's own default line chart (AddChart2, xlLineMarkers) saved as .doc / .rtf: an Excel chart object like the others.</summary>
    public static void RunLineChart()
    {
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-charts-binary");
        foreach (string extension in new[] { ".doc", ".rtf" })
        {
            string path = Path.Combine(folder, "word-line-chart-v1" + extension);
            if (!File.Exists(path)) { TestReport.Check(false, extension + ": line chart fixture exists"); continue; }
            string all = string.Concat(Render(path));
            TestReport.Check(Regex.Matches(all, "<svg class=\"word-chart\"").Count == 1 && all.Contains(">Line Chart<") && all.Contains("fill=\"none\" stroke=") && all.Contains("stroke-linejoin=\"round\""),
                             extension + ": Word's line chart object drawn as a line chart");
        }
    }

    /// <summary>RTF objects that are not charts (Apache Tika's testRTFEmbeddedFiles.rtf: Word-made embedded files) show their \result preview picture, not a placeholder.</summary>
    public static void RunObjectPreviews()
    {
        string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "apache-tika", "testRTFEmbeddedFiles.rtf");
        if (!File.Exists(path)) { TestReport.Check(false, "rtf: embedded-files fixture exists"); return; }
        string all = string.Concat(Render(path));
        // the previews are EMF (shown as the picture placeholder, as every EMF picture is): no object placeholder is left
        int pictures = Regex.Matches(all, "class=\"word-picture\"").Count + Regex.Matches(all, ">[[]image[]]<").Count, placeholders = Regex.Matches(all, ">[[]object[]]<").Count;
        TestReport.Check(pictures >= 14 && placeholders == 0, "rtf: the 14 non-chart objects drawn as their preview pictures (" + pictures + " pictures, " + placeholders + " object placeholders)");
    }

    /// <summary>--chart-html &lt;file&gt; &lt;out.html&gt;: every block of a document as HTML, for scripts/compare-chart-text-positions.py.</summary>
    public static void WriteHtml(string path, string output) => File.WriteAllText(output, string.Concat(Render(path)));

    /// <summary>Office's automatic trend line names in a Hebrew interface, as Word 16 printed them (its PDF of a chart with all six types, 2026-11-12; code points).</summary>
    public static void RunTrendNames()
    {
        TestReport.Step("Charts: trend line legend names as a Hebrew Office writes them");
        var saved = System.Globalization.CultureInfo.CurrentUICulture;
        try
        {
            System.Globalization.CultureInfo.CurrentUICulture = new System.Globalization.CultureInfo("he-IL");
            var poly = new System.Xml.Linq.XElement("trendline");
            var moving = new System.Xml.Linq.XElement("trendline", new System.Xml.Linq.XElement("period", new System.Xml.Linq.XAttribute("val", "2")));
            (string Type, System.Xml.Linq.XElement Trend, string Name)[] cases =
            [
                ("exp", poly, "\u05DE\u05E2\u05E8\u05D9\u05DB\u05D9"), ("log", poly, "\u05DC\u05D5\u05D2\u05E8\u05D9\u05EA\u05DE\u05D9"), ("poly", poly, "\u05E4\u05D5\u05DC\u05D9\u05E0\u05D5\u05DE\u05D9\u05D0\u05DC\u05D9"), ("power", poly, "\u05D7\u05D6\u05E7\u05D4"), ("movingAvg", moving, "2 \u05EA\u05E7\u05D5\u05E4\u05EA \u05DE\u05DE\u05D5\u05E6\u05E2 \u05E0\u05E2"), ("linear", poly, "\u05DC\u05D9\u05E0\u05D9\u05D0\u05E8\u05D9"),
            ];
            foreach ((string type, System.Xml.Linq.XElement trend, string name) in cases)
                TestReport.Check(FastDocView.DocxHtml.DocxChart.TrendLineName(type, trend, "S1") == name + " (S1)", "Hebrew trend name: " + type);
        }
        finally { System.Globalization.CultureInfo.CurrentUICulture = saved; }
    }

    /// <summary>c:dispEq / c:dispRSqr shown unless val is false (Word 16's PDFs of probes, 2026-11-12; oox seriesmodel.cxx:90-91).</summary>
    public static void RunTrendLabelDefaults()
    {
        TestReport.Step("Charts: trend equation / R squared shown unless turned off");
        static System.Xml.Linq.XElement Flag(string? val) => val is null ? new System.Xml.Linq.XElement("dispRSqr") : new System.Xml.Linq.XElement("dispRSqr", new System.Xml.Linq.XAttribute("val", val));
        TestReport.Check(FastDocView.DocxHtml.DocxChart.TrendLabelShown(null), "absent element: shown");
        TestReport.Check(FastDocView.DocxHtml.DocxChart.TrendLabelShown(Flag(null)), "no val: shown");
        TestReport.Check(FastDocView.DocxHtml.DocxChart.TrendLabelShown(Flag("1")), "val 1: shown");
        TestReport.Check(!FastDocView.DocxHtml.DocxChart.TrendLabelShown(Flag("0")), "val 0: hidden");
        TestReport.Check(!FastDocView.DocxHtml.DocxChart.TrendLabelShown(Flag("false")), "val false: hidden");
    }

    /// <summary>A chart's texts relative to its title: (text, x, y, anchor).</summary>
    private static List<(string Text, double X, double Y, string Anchor)> ChartTexts(List<string> html, string title)
    {
        string svg = Regex.Match(string.Concat(html), "<svg class=\"word-chart\"(?:(?!</svg>).)*>" + Regex.Escape(title) + "<.*?</svg>", RegexOptions.Singleline).Value;
        var texts = Regex.Matches(svg, "<text x=\"([-0-9.]+)\" y=\"([-0-9.]+)\"[^>]*text-anchor=\"([a-z]+)\"[^>]*>([^<]*)<")
            .Select(m => (Text: m.Groups[4].Value, X: double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture), Y: double.Parse(m.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture), Anchor: m.Groups[3].Value)).ToList();
        var head = texts.FirstOrDefault(t => t.Text == title);
        return texts.Select(t => (t.Text, t.X - head.X, t.Y - head.Y, t.Anchor)).ToList();
    }

    /// <summary>
    /// 3-D line / area scenes and axis crossings against Word 16's PDFs (test-data/word-charts-binary/word-charts-3d-rows-v1 and
    /// word-charts-crosses-v1, scripts/make-word-charts-3d-crosses-v1.py): label positions relative to the title, in points.
    /// </summary>
    public static void RunCrossingsAndRows()
    {
        TestReport.Step("Charts: 3-D line / area scenes and axis crossings (c:crosses, c:crossesAt; BIFF8; ODF chart:axis-position) as Word draws them");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-charts-binary");
        bool Near(List<(string Text, double X, double Y, string Anchor)> texts, string text, double? x, double? y, double tolerance)
            => texts.Any(t => t.Text == text && (x is null || Math.Abs(t.X - x.Value) <= tolerance) && (y is null || Math.Abs(t.Y - y.Value) <= tolerance));
        string rows = Path.Combine(folder, "word-charts-3d-rows-v1.docx"), crosses = Path.Combine(folder, "word-charts-crosses-v1.docx");
        TestReport.Check(File.Exists(rows) && File.Exists(crosses), "fixtures exist");
        if (!File.Exists(rows) || !File.Exists(crosses)) return;
        List<string> html = Render(rows);
        var right = ChartTexts(html, "Area3D Right");
        TestReport.Check(Near(right, "Q1", -186.2, 186.4, 2) && Near(right, "Q4", 112.0, 186.4, 2) && Near(right, "5", -194.6, 81.4, 2), "right-angled 3-D area: Word's scene (categories within 2 pt)");
        var flat = ChartTexts(html, "Area3D Flat");
        TestReport.Check(Near(flat, "Q1", -186.0, 178.7, 2) && Near(flat, "Q4", 161.7, 178.7, 2), "right-angled 3-D area, depth 60 %: Word's scene");
        var perspective = ChartTexts(html, "Area3D Persp");
        TestReport.Check(Near(perspective, "Q1", -153.7, 150.2, 5) && Near(perspective, "Q4", 136.4, 187.6, 5) && !perspective.Any(t => t.Text == "0.5"), "3-D area with perspective: Word's scene within 5 pt, ticks by 1");
        html = Render(crosses);
        var columnMax = ChartTexts(html, "Col CrossMax");
        TestReport.Check(Near(columnMax, "Q1", null, 19.3, 1.5), "category axis at the maximum: its labels above the plot");
        var lineMax = ChartTexts(html, "Line ValMax");
        TestReport.Check(lineMax.Any(t => t.Text == "4" && t.Anchor == "start" && Math.Abs(t.X - 195.7) <= 2), "value axis crossing at the last category: its labels at the right");
        var bar = ChartTexts(html, "Bar CrossAt");
        TestReport.Check(bar.Any(t => t.Text == "Q1" && t.Anchor == "end" && Math.Abs(t.X - 57.0) <= 3) && Near(bar, "2", null, 18.7, 1.5), "horizontal bars: category axis at a value, value axis at the top");
        var area = ChartTexts(html, "Area CrossAt");
        TestReport.Check(Near(area, "Q1", null, 116.2, 1.5) && area.Any(t => t.Text == "3" && t.Anchor == "end" && Math.Abs(t.X + 71.6) <= 5), "area: both axes crossing at values inside the plot");
        string binary = Path.Combine(folder, "word-charts-crosses-v1.doc");
        if (File.Exists(binary))
        {
            List<string> doc = Render(binary);
            // (Word saves the line charts as pictures; its .doc keeps the workbook's own values, all positive)
            TestReport.Check(ChartTexts(doc, "Col CrossAt").Any(t => t.Text == "3" && t.Anchor == "end" && Math.Abs(t.X + 8.3) <= 4) && Near(ChartTexts(doc, "Bar CrossAt"), "2", null, 18.7, 1.5),
                             ".doc (BIFF8 CHVALUERANGE / CHLABELRANGE crossings): value axis at category 3, horizontal bars' value labels at the top");
        }
        // ODF: chart:axis-position on the axis style becomes c:crosses / c:crossesAt
        XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0", style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        var styles = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach ((string name, string position) in new[] { ("x", "end"), ("y", "2.5") })
            styles[name] = new XElement(style + "style", new XAttribute(style + "name", name), new XElement(style + "chart-properties", new XAttribute(chart + "axis-position", position)));
        var odf = new XElement(chart + "chart", new XAttribute(chart + "class", "chart:bar"),
            new XElement(chart + "plot-area", new XElement(chart + "axis", new XAttribute(chart + "dimension", "x"), new XAttribute(chart + "style-name", "x")),
                                              new XElement(chart + "axis", new XAttribute(chart + "dimension", "y"), new XAttribute(chart + "style-name", "y"))));
        string converted = FastDocView.Odt.OdtChart.ChartSpace(odf, styles).ToString();
        TestReport.Check(Regex.IsMatch(converted, "<catAx[^>]*>.*<crosses val=\"max\" />.*</catAx>", RegexOptions.Singleline) && Regex.IsMatch(converted, "<valAx[^>]*>.*<crossesAt val=\"2.5\" />.*</valAx>", RegexOptions.Singleline),
                         "ODT: chart:axis-position end / a value become c:crosses max / c:crossesAt");
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
}
