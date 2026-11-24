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
            TestReport.Check(Regex.IsMatch(column, "fill=\"#0000FF\"[^>]*>5[.]0<") && Regex.IsMatch(column, "fill=\"#00B050\"[^>]*>North<"), extension + ": axis label and legend text colours (the legend's 00B050, saved as its nearest palette colour 339966 in both the RGB and the index, read exactly from its TextPropsStream)");
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

    /// <summary>Office's automatic trend line names in an English interface, as Word 16 printed them (its PDF of a chart with all six types, Office's UI language switched to en-us, 2026-11-16).</summary>
    public static void RunTrendNamesEnglish()
    {
        TestReport.Step("Charts: trend line legend names as an English Office writes them");
        var saved = System.Globalization.CultureInfo.CurrentUICulture;
        try
        {
            System.Globalization.CultureInfo.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            var plain = new System.Xml.Linq.XElement("trendline");
            var moving = new System.Xml.Linq.XElement("trendline", new System.Xml.Linq.XElement("period", new System.Xml.Linq.XAttribute("val", "3")));
            (string Type, System.Xml.Linq.XElement Trend, string Name)[] cases =
            [
                ("linear", plain, "Linear"), ("exp", plain, "Expon."), ("log", plain, "Log."), ("poly", plain, "Poly."), ("power", plain, "Power"), ("movingAvg", moving, "3 per. Mov. Avg."), ("movingAvg", plain, "2 per. Mov. Avg."),
            ];
            foreach ((string type, System.Xml.Linq.XElement trend, string name) in cases)
                TestReport.Check(FastDocView.DocxHtml.DocxChart.TrendLineName(type, trend, "North") == name + " (North)", "English trend name: " + name);
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
        // Word's camera with LibreOffice's placement (NOTES.md 2026-11-24, the mix): locked at our positions; Word's PDF has Q1 at
        // (-153.7, 150.2) and Q4 at (136.4, 187.6), 0.3 / 0.5 and 7.5 / 2.5 pt away
        TestReport.Check(Near(perspective, "Q1", -153.5, 149.8, 1) && Near(perspective, "Q4", 128.9, 190.1, 1) && !perspective.Any(t => t.Text == "0.5"), "3-D area with perspective: Word's camera, LibreOffice's placement, ticks by 1");
        var seriesAxis = ChartTexts(html, "Line3D SerAx");
        TestReport.Check(Near(seriesAxis, "0", -170.8, 140.6, 1.5) && Near(seriesAxis, "5", -177.1, 48.1, 1.5), "3-D line with a shown series axis: Word's room for the series names on the right (value labels within 1.5 pt of Word)");
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
        // scatter / bubble: the y axis at an x value (x 0 for autoZero), at the right for max; bubbles widen both axes (Word: -4..7, -8..8)
        string scatterPath = Path.Combine(folder, "word-charts-scatter-crosses-v1.docx");
        TestReport.Check(File.Exists(scatterPath), "scatter crossings fixture exists");
        if (File.Exists(scatterPath))
        {
            List<string> scatter = Render(scatterPath);
            TestReport.Check(ChartTexts(scatter, "Sc Zero").Any(t => t.Text == "2" && t.Anchor == "end" && Math.Abs(t.X + 72.0) <= 2), "scatter: y axis at x 0 (autoZero), its labels beside it");
            TestReport.Check(ChartTexts(scatter, "Sc MaxBoth").Any(t => t.Text == "2" && t.Anchor == "start" && Math.Abs(t.X - 193.2) <= 2), "scatter: y axis crossing at the x maximum, labels at the right");
            TestReport.Check(ChartTexts(scatter, "Sc AtBoth").Any(t => t.Text == "2" && t.Anchor == "end" && Math.Abs(t.X + 7.6) <= 2), "scatter: c:crossesAt on both axes");
            var bubble = ChartTexts(scatter, "Bub AtBoth");
            TestReport.Check(bubble.Any(t => t.Text == "8") && bubble.Any(t => t.Text == "7") && bubble.Any(t => t.Text == "-4"), "bubble: the axes widened for the bubbles as Word does (-4..7, -8..8)");
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

    /// <summary>
    /// Series axes (c:serAx; BIFF8 the third AXIS block of a deep 3-D chart) and exact text colours (BIFF8 TextPropsStream / RichTextStream)
    /// against Word 16's PDF of test-data/word-charts-binary/word-charts-seraxis-v1 (scripts/make-word-charts-seraxis-v1.py), in the .docx
    /// and in Word's own .doc and .rtf of it: series names left-aligned beside their rows (positions relative to the title, in points).
    /// </summary>
    public static void RunSeriesAxes()
    {
        TestReport.Step("Charts: 3-D series axis labels and exact .doc / RTF text colours as Word draws them");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-charts-binary");
        foreach (string extension in new[] { "docx", "doc", "rtf" })
        {
            string path = Path.Combine(folder, "word-charts-seraxis-v1." + extension);
            TestReport.Check(File.Exists(path), "fixture exists: " + extension);
            if (!File.Exists(path)) continue;
            List<string> html = Render(path);
            bool Start(string title, string text, double x, double y) => ChartTexts(html, title).Any(t => t.Text == text && t.Anchor == "start" && Math.Abs(t.X - x) <= 6 && Math.Abs(t.Y - y) <= 4);
            TestReport.Check(Start("Col3D SerAx R", "North", 121.4, 164.7) && Start("Col3D SerAx R", "West", 170.1, 127.8), extension + ": right-angled 3-D columns, one row per series: the series names beside the rows (within 6 pt)");
            // perspective: Word's camera with LibreOffice's placement (NOTES.md 2026-11-24), locked at our positions within 1.5 pt; Word's PDF
            // has North at (136.1, 161.8) / West at (153.5, 119.4), 7-11 / 2.5 and 4-8 / 0.5 pt away
            bool Ours(string title, string text, double x, double y) => ChartTexts(html, title).Any(t => t.Text == text && t.Anchor == "start" && Math.Abs(t.X - x) <= 1.5 && Math.Abs(t.Y - y) <= 1.5);
            TestReport.Check(Ours("Col3D SerAx P", "North", 129.6, 164.9) && Ours("Col3D SerAx P", "West", 149.5, 119.4), extension + ": 3-D columns with perspective: the series names, Word's camera, LibreOffice's placement");
            TestReport.Check(Ours("Area3D SerAx P", "North", 125.7, 165.0) && Ours("Area3D SerAx P", "West", 145.6, 119.9), extension + ": 3-D area with perspective: the series names, Word's camera, LibreOffice's placement");
            string colours = Regex.Match(string.Concat(html), "<svg class=\"word-chart\"(?:(?!</svg>).)*>Col Colors<.*?</svg>", RegexOptions.Singleline).Value;
            string colours3D = Regex.Match(string.Concat(html), "<svg class=\"word-chart\"(?:(?!</svg>).)*>Col3D Colors<.*?</svg>", RegexOptions.Singleline).Value;
            bool TextIn(string svg, string text, string fill) => Regex.IsMatch(svg, "<text[^>]*fill=\"#" + fill + "\"[^>]*>" + Regex.Escape(text) + "<", RegexOptions.IgnoreCase);
            TestReport.Check(TextIn(colours, "Col Colors", "7A2E9B") && TextIn(colours, "Q1", "8A4F1C") && TextIn(colours, "3", "8A4F1C") && TextIn(colours, "North", "2F6EA8"),
                             extension + ": title, axis and legend text in Word's exact colours (not the nearest palette colours 333399, 993300, 666699)");
            TestReport.Check(TextIn(colours3D, "Col3D Colors", "C0392B") && Regex.Matches(colours3D, "<text[^>]*fill=\"#1E7A6B\"[^>]*>(North|South|West)<").Count == 3 && TextIn(colours3D, "Q1", "1E7A6B"),
                             extension + ": 3-D title, category and series axis labels in Word's exact colours");
            TestReport.Check(Regex.IsMatch(colours, "fill=\"#12AB9C\"") && Regex.IsMatch(colours, "fill=\"#5B2C83\""), extension + ": series fills in Word's exact colours");
            // word-charts-right-angled-v1 (scripts/make-word-charts-right-theme-v1.py): right-angled scene with the names fitted beside it
            string rightAngled = Path.Combine(folder, "word-charts-right-angled-v1." + extension);
            if (File.Exists(rightAngled))
            {
                List<string> right = Render(rightAngled);
                bool Name(string text, double x, double y) => ChartTexts(right, "R 20 30 S").Any(t => t.Text == text && t.Anchor == "start" && Math.Abs(t.X - x) <= 4 && Math.Abs(t.Y - y) <= 3);
                TestReport.Check(Name("North", 101.2, 164.5) && Name("West", 166.2, 120.1), extension + ": right-angled 3-D area (rotX 20, rotY 30): the scene fitted with the series names as Word does (within 4 pt)");
            }
            // word-charts-theme-text-v1: chart text in theme colours under a document theme with non-default accents
            string themed = Path.Combine(folder, "word-charts-theme-text-v1." + extension);
            TestReport.Check(File.Exists(themed), "fixture exists: theme text " + extension);
            if (File.Exists(themed))
            {
                string all = string.Concat(Render(themed));
                string flat = Regex.Match(all, "<svg class=\"word-chart\"(?:(?!</svg>).)*>Theme Text<.*?</svg>", RegexOptions.Singleline).Value;
                string deep = Regex.Match(all, "<svg class=\"word-chart\"(?:(?!</svg>).)*>Theme Text 3D<.*?</svg>", RegexOptions.Singleline).Value;
                TestReport.Check(TextIn(flat, "Theme Text", "E65100") && TextIn(flat, "Q1", "751D84") && TextIn(flat, "3", "751D84") && TextIn(flat, "North", "58D9BA"),
                                 extension + ": title, axis and legend text in theme colours (accent6, accent2 75 %, accent1 60 / 40 %) of the chart's own theme");
                TestReport.Check(TextIn(deep, "Theme Text 3D", "5D4037") && TextIn(deep, "Q1", "1D2556") && Regex.Matches(deep, "<text[^>]*fill=\"#1D2556\"[^>]*>(North|South|West)<").Count == 3 && Regex.Matches(deep, "<text[^>]*fill=\"#C2185B\"[^>]*>(North|South|West)<").Count == 3,
                                 extension + ": 3-D title (tx2), axis and series axis labels (accent5 50 %) and legend (accent4) in theme colours");
            }
        }
    }

    /// <summary>
    /// The 3-D scene placed and sized as LibreOffice does it (VDiagram.cxx:409-423 adjustPosAndSize_3d: the projected scene box scaled
    /// keeping its aspect and centred in the plot; the camera distance of ThreeDHelper.cxx:635-645; the eye on the axis through the
    /// scene's centre), one chart per type from test-data/word-charts-scene-types (Word's box sizes, perspective 30 / 120, parallel,
    /// right-angled) and the 3-D pie of odt-charts-v5 (LibreOffice's 1 x 1 x 0.1 pie volume, PieChart.cxx:252-257): each wall's corners
    /// locked within 0.05 pt. Word's residuals per type are in NOTES.md (scripts/chart3d_walls_vs_word.py).
    /// </summary>
    public static void RunScenePlacement()
    {
        TestReport.Step("Charts: 3-D scenes with Word's camera placed as LibreOffice places them (VDiagram adjustPosAndSize_3d), every type");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-charts-scene-types");
        (string Name, int Index, string Back, string Side, string Floor)[] expected =
        [
            ("col-clustered_x10_y20_p30_r0", 0, "81.18,175.86 370.6,203.75 377.22,20.87 76.9,14.1", "43.78,194.25 81.18,175.86 76.9,14.1 37.68,18.56", "43.78,194.25 356.66,227.6 370.6,203.75 81.18,175.86"),
            ("col-clustered_x10_y20_p120_r0", 1, "95.74,159.89 349.74,192.39 363.11,20 89.04,25.24", "54.58,180.43 95.74,159.89 89.04,25.24 43.34,21.95", "54.58,180.43 354.21,227.6 349.74,192.39 95.74,159.89"),
            ("col-clustered_x15_y20_p0_r0", 18, "75.38,181.42 365.96,208.8 365.96,41.47 75.38,14.1", "48.94,200.23 75.38,181.42 75.38,14.1 48.94,32.9", "48.94,200.23 339.52,227.6 365.96,208.8 75.38,181.42"),
            ("col-stacked_x10_y20_p30_r0", 20, "81.18,175.86 370.6,203.75 377.22,20.87 76.9,14.1", "43.78,194.25 81.18,175.86 76.9,14.1 37.68,18.56", "43.78,194.25 356.66,227.6 370.6,203.75 81.18,175.86"),
            ("col-stacked_x10_y20_p120_r0", 21, "95.74,159.89 349.74,192.39 363.11,20 89.04,25.24", "54.58,180.43 95.74,159.89 89.04,25.24 43.34,21.95", "54.58,180.43 354.21,227.6 349.74,192.39 95.74,159.89"),
            ("col-stacked_x15_y20_p0_r0", 38, "75.38,181.42 365.96,208.8 365.96,41.47 75.38,14.1", "48.94,200.23 75.38,181.42 75.38,14.1 48.94,32.9", "48.94,200.23 339.52,227.6 365.96,208.8 75.38,181.42"),
            ("col-standard_x10_y20_p30_r0", 40, "102.14,145.64 393.13,170.18 398.5,33.47 99.26,24.78", "22.35,181.11 102.14,145.64 99.26,24.78 16.4,37.36", "22.35,181.11 365.5,216.92 393.13,170.18 102.14,145.64"),
            ("col-standard_x10_y20_p120_r0", 41, "118.53,123.96 371.53,149.55 381.08,24.29 114.39,24.18", "28.14,162.82 118.53,123.96 114.39,24.18 16.4,24.35", "28.14,162.82 382.31,217.52 371.53,149.55 118.53,123.96"),
            ("col-standard_x15_y20_p0_r0", 58, "75.23,154.51 398.5,184.97 398.5,45.35 75.23,14.9", "16.4,196.35 75.23,154.51 75.23,14.9 16.4,56.73", "16.4,196.35 339.67,226.8 398.5,184.97 75.23,154.51"),
            ("area-standard_x10_y20_p30_r0", 60, "102.14,145.64 393.13,170.18 398.5,33.47 99.26,24.78", "22.35,181.11 102.14,145.64 99.26,24.78 16.4,37.36", "22.35,181.11 365.5,216.92 393.13,170.18 102.14,145.64"),
            ("area-standard_x10_y20_p120_r0", 61, "118.53,123.96 371.53,149.55 381.08,24.29 114.39,24.18", "28.14,162.82 118.53,123.96 114.39,24.18 16.4,24.35", "28.14,162.82 382.31,217.52 371.53,149.55 118.53,123.96"),
            ("area-standard_x15_y20_p0_r0", 78, "75.23,154.51 398.5,184.97 398.5,45.35 75.23,14.9", "16.4,196.35 75.23,154.51 75.23,14.9 16.4,56.73", "16.4,196.35 339.67,226.8 398.5,184.97 75.23,154.51"),
            ("area-stacked_x10_y20_p30_r0", 80, "81.18,175.86 370.6,203.75 377.22,20.87 76.9,14.1", "43.78,194.25 81.18,175.86 76.9,14.1 37.68,18.56", "43.78,194.25 356.66,227.6 370.6,203.75 81.18,175.86"),
            ("area-stacked_x10_y20_p120_r0", 81, "95.74,159.89 349.74,192.39 363.11,20 89.04,25.24", "54.58,180.43 95.74,159.89 89.04,25.24 43.34,21.95", "54.58,180.43 354.21,227.6 349.74,192.39 95.74,159.89"),
            ("area-stacked_x15_y20_p0_r0", 98, "75.38,181.42 365.96,208.8 365.96,41.47 75.38,14.1", "48.94,200.23 75.38,181.42 75.38,14.1 48.94,32.9", "48.94,200.23 339.52,227.6 365.96,208.8 75.38,181.42"),
            ("surface_x10_y20_p30_r0", 100, "102.14,145.64 393.13,170.18 398.5,33.47 99.26,24.78", "22.35,181.11 102.14,145.64 99.26,24.78 16.4,37.36", "22.35,181.11 365.5,216.92 393.13,170.18 102.14,145.64"),
            ("surface_x10_y20_p120_r0", 101, "118.53,123.96 371.53,149.55 381.08,24.29 114.39,24.18", "28.14,162.82 118.53,123.96 114.39,24.18 16.4,24.35", "28.14,162.82 382.31,217.52 371.53,149.55 118.53,123.96"),
            ("surface_x15_y20_p0_r0", 118, "75.23,154.51 398.5,184.97 398.5,45.35 75.23,14.9", "16.4,196.35 75.23,154.51 75.23,14.9 16.4,56.73", "16.4,196.35 339.67,226.8 398.5,184.97 75.23,154.51"),
            ("line-standard_x10_y20_p30_r0", 120, "102.14,145.64 393.13,170.18 398.5,33.47 99.26,24.78", "22.35,181.11 102.14,145.64 99.26,24.78 16.4,37.36", "22.35,181.11 365.5,216.92 393.13,170.18 102.14,145.64"),
            ("line-standard_x10_y20_p120_r0", 121, "118.53,123.96 371.53,149.55 381.08,24.29 114.39,24.18", "28.14,162.82 118.53,123.96 114.39,24.18 16.4,24.35", "28.14,162.82 382.31,217.52 371.53,149.55 118.53,123.96"),
            ("line-standard_x15_y20_p0_r0", 138, "75.23,154.51 398.5,184.97 398.5,45.35 75.23,14.9", "16.4,196.35 75.23,154.51 75.23,14.9 16.4,56.73", "16.4,196.35 339.67,226.8 398.5,184.97 75.23,154.51"),
            ("bar-clustered_x10_y20_p30_r0", 140, "66.14,183.51 369.39,213.65 376.45,21.01 61.25,14.1", "44.43,194.19 361.6,227.6 369.39,213.65 66.14,183.51", "44.43,194.19 66.14,183.51 61.25,14.1 38.45,16.54"),
            ("bar-clustered_x10_y20_p120_r0", 141, "76.63,170.09 353.21,207 368.03,17.55 68.74,23.53", "52.47,182.15 355.82,227.6 353.21,207 76.63,170.09", "52.47,182.15 76.63,170.09 68.74,23.53 41.93,21.6"),
            ("line-ra_x20_y30_p30_r1", 170, "54.66,154.83 403.9,154.83 403.9,57.01 54.66,57.01", "11,184.69 54.66,154.83 54.66,57.01 11,86.87", "11,184.69 360.24,184.69 403.9,154.83 54.66,154.83"),
            ("col-ra_x20_y40_p30_r1", 208, "62.75,201.75 403.62,201.75 403.62,10.8 62.75,10.8", "11.28,230.9 62.75,201.75 62.75,10.8 11.28,39.95", "11.28,230.9 352.15,230.9 403.62,201.75 62.75,201.75")
        ];
        var documents = new Dictionary<int, List<string>>();
        foreach (var probe in expected)
        {
            string path = Path.Combine(folder, "types-" + (probe.Index / 24).ToString("000", System.Globalization.CultureInfo.InvariantCulture) + ".docx");
            if (!File.Exists(path)) { TestReport.Check(false, probe.Name + ": fixture exists"); continue; }
            if (!documents.TryGetValue(probe.Index / 24, out var svgs))
                documents[probe.Index / 24] = svgs = Regex.Matches(string.Concat(Render(path)), "<svg class=\"word-chart\".*?</svg>", RegexOptions.Singleline).Select(m => m.Value).ToList();
            string svg = probe.Index % 24 < svgs.Count ? svgs[probe.Index % 24] : "";
            // the probes paint the floor red, the back wall green and the side wall blue
            string Face(string fill) => string.Join(" ", Regex.Matches(Regex.Match(svg, "<path class=\"word-chart-(?:wall|floor)\" d=\"([^\"]+)\" fill=\"" + fill + "\"", RegexOptions.IgnoreCase).Groups[1].Value, "[ML]([-0-9.]+) ([-0-9.]+)")
                .Select(m => m.Groups[1].Value + "," + m.Groups[2].Value));
            bool Same(string actual, string wanted)
            {
                double[] Numbers(string text) => text.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries).Select(v => double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                double[] a = Numbers(actual), w = Numbers(wanted);
                return a.Length == w.Length && a.Zip(w, (x, y) => Math.Abs(x - y)).All(d => d <= 0.05);
            }
            TestReport.Check(Same(Face("#00FF00"), probe.Back) && Same(Face("#0000FF"), probe.Side) && Same(Face("#FF0000"), probe.Floor), probe.Name + ": walls where LibreOffice's placement puts them (0.05 pt)");
        }
        // the 3-D pie: LibreOffice's volume fitted and centred - the slice tops' box and the rim's lowest point
        foreach ((string extension, double minX, double minY, double maxX, double maxY, double rim) in new[] { ("odt", 30.89, 46.60, 261.62, 160.31, 181.85), ("docx", 14.68, 35.02, 273.93, 163.44, 186.78) })
        {
            string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "odt-charts", "odt-charts-v5." + extension);
            if (!File.Exists(path)) { TestReport.Check(false, "3-D pie fixture exists: " + extension); continue; }
            string pie = Regex.Matches(string.Concat(Render(path)), "<svg class=\"word-chart\".*?</svg>", RegexOptions.Singleline).Select(m => m.Value).FirstOrDefault(v => v.Contains(">Pie Depth<")) ?? "";
            List<(double X, double Y)> Points(string pattern) => Regex.Matches(pie, pattern).SelectMany(m => Regex.Matches(m.Groups[1].Value, "[ML]([-0-9.]+) ([-0-9.]+)").Select(p => (double.Parse(p.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture), double.Parse(p.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture)))).ToList();
            var tops = Points("<path class=\"word-chart-pie-top\" d=\"([^\"]+)\"");
            var sides = Points("class=\"word-chart-pie-side[^\"]*\" d=\"([^\"]+)\"");
            bool near(double a, double b) => Math.Abs(a - b) <= 0.05;
            TestReport.Check(tops.Count > 0 && sides.Count > 0 && near(tops.Min(p => p.X), minX) && near(tops.Min(p => p.Y), minY) && near(tops.Max(p => p.X), maxX) && near(tops.Max(p => p.Y), maxY) && near(sides.Max(p => p.Y), rim),
                             extension + ": 3-D pie as LibreOffice's 1 x 1 x 0.1 volume, fitted and centred (0.05 pt)");
        }
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
