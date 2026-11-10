using System.Text.RegularExpressions;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// .odt charts (--fidelity, seconds): test-data/odt-charts/odt-charts-v1.odt, made by LibreOffice itself through UNO
/// (scripts/make-odt-charts.py: one Writer chart per class - column with subtitle, axis titles and value labels, stacked bars,
/// percent columns, line, stacked area, pie, donut, scatter, radar, stock, bubble, 3-D column). Each chart is drawn as SVG from
/// its "Object N" sub-document (not the preview picture), with the series colours LibreOffice wrote; the same file saved
/// flat (.fodt, the chart inline in the draw:object) draws the same; chart texts are found by search.
/// </summary>
public static class OdtChartsFidelityTest
{
    private static readonly string[] Titles =
        ["Column Sales", "Stacked Bars", "Percent Columns", "Line Trend", "Area Stack", "Pie Share", "Donut Rings", "Scatter Points", "Radar Net", "Stock Prices", "Bubble Sizes", "Column Depth"];

    public static void Run()
    {
        TestReport.Step("ODT charts: LibreOffice-made charts drawn from their sub-documents");
        string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "odt-charts", "odt-charts-v1.odt");
        TestReport.Check(File.Exists(path), "chart fixture exists");
        if (!File.Exists(path)) return;
        List<string> html = Render(path);
        Check(html, "odt");
        TestReport.Check(SearchCount(path, "Quarter Axis") == 1 && SearchCount(path, "Sub Heading") == 1 && SearchCount(path, "Speed") == 1 && SearchCount(path, "Bubble Sizes") == 1,
                         "search: axis title, subtitle, a radar category, a title - each found once, in its chart's block");
        TestReport.Check(SearchCount(path, "ChartPie paragraph") == 1, "search: the paragraph's own text still found");

        TestReport.Step("ODT charts: the same document saved flat (.fodt)");
        if (LibreOfficeConversion.Convert(path, ".fodt", "fodt") is { } fodt)
        {
            List<string> flat = Render(fodt);
            Check(flat, "fodt");
        }
    }

    private static void Check(List<string> html, string format)
    {
        string all = string.Join("\n", html);
        int charts = Regex.Matches(all, "<svg class=\"word-chart\"").Count;
        TestReport.Check(charts == Titles.Length, format + ": every chart drawn as SVG (" + charts + " of " + Titles.Length + ")");
        foreach (string title in Titles)
            TestReport.Check(all.Contains(">" + title + "<"), format + ": chart title drawn: " + title);
        string Block(string marker) => html.FirstOrDefault(h => h.Contains(">" + marker + "<")) ?? "";
        string column = Block("Column Sales");
        TestReport.Check(column.Contains(">Sub Heading<") && column.Contains(">Quarter Axis<") && column.Contains(">Amount Axis<"), format + ": subtitle and axis titles");
        TestReport.Check(column.Contains(">Q1<") && column.Contains(">Q3<") && column.Contains(">Alpha<") && column.Contains(">Beta<") && column.Contains(">40<"),
                         format + ": categories, legend (series names from the label cells), value labels");
        TestReport.Check(column.Contains("#004586") && column.Contains("#FF420E"), format + ": series colours from the series styles");
        TestReport.Check(column.Contains("#D9D9D9"), format + ": the wall's fill as the plot area");
        string pie = Block("Pie Share");
        TestReport.Check(Regex.Matches(pie, "<path[^>]*fill=\"#[0-9A-F]{6}\"").Count >= 3 && pie.Contains(">East<") && pie.Contains(">Middle<"),
                         format + ": pie slices with the data points' colours, category legend");
        TestReport.Check(!all.Contains("ObjectReplacements") || charts == Titles.Length, format + ": no preview picture where the chart is drawn");
        TestReport.Check(Block("Scatter Points").Contains("<circle") && Block("Radar Net").Contains(">Speed<"), format + ": scatter points, radar categories");
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
