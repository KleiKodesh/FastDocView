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
        foreach (string extension in new[] { ".doc", ".rtf" })
        {
            TestReport.Step("Charts in " + extension + ": Word's embedded Excel chart objects drawn from their chart records");
            string path = Path.Combine(folder, "word-charts-binary-v2" + extension);
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
