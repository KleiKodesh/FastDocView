using System.Text.RegularExpressions;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// .odt charts (--fidelity, seconds): test-data/odt-charts/odt-charts-v5.odt, made by LibreOffice itself through UNO
/// (scripts/make-odt-charts.py: one Writer chart per class - column with subtitle, axis titles and value labels, stacked bars,
/// percent columns, line, stacked area, pie, donut, scatter, radar, stock, bubble, 3-D column). Each chart is drawn as SVG from
/// its "Object N" sub-document (not the preview picture), with the series colours LibreOffice wrote; the same file saved
/// flat (.fodt, the chart inline in the draw:object) draws the same; chart texts are found by search.
/// </summary>
public static class OdtChartsFidelityTest
{
    private static readonly string[] Titles =
        ["Column Sales", "Stacked Bars", "Percent Columns", "Line Trend", "Area Stack", "Pie Share", "Donut Rings", "Scatter Points", "Radar Net", "Stock Prices", "Bubble Sizes", "Column Depth",
         "Log Growth", "Two Axes", "Error Bars", "Trend Fit", "Pie Depth", "Range Errors", "Line Depth", "Area Depth", "Bar Depth", "Crowded Axis", "Wrapped Labels", "Long Legend", "Legend Below", "Wide Values", "Small Chart", "Many Series", "Candle Stock"];

    public static void Run()
    {
        TestReport.Step("ODT charts: LibreOffice-made charts drawn from their sub-documents");
        string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "odt-charts", "odt-charts-v5.odt");
        TestReport.Check(File.Exists(path), "chart fixture exists");
        if (!File.Exists(path)) return;
        List<string> html = Render(path);
        File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "charts-odt.html"), string.Concat(html));
        Check(html, "odt");
        TestReport.Check(SearchCount(path, "Quarter Axis") == 1 && SearchCount(path, "Sub Heading") == 1 && SearchCount(path, "Speed") == 1 && SearchCount(path, "Bubble Sizes") == 1,
                         "search: axis title, subtitle, a radar category, a title - each found once, in its chart's block");
        TestReport.Check(SearchCount(path, "ChartPie paragraph") == 1, "search: the paragraph's own text still found");

        CheckParts(html, "odt");

        TestReport.Step("ODT charts: the same charts saved as .docx by LibreOffice (DrawingML), drawn by the same renderer");
        string docx = Path.ChangeExtension(path, ".docx");
        if (File.Exists(docx))
        {
            List<string> word = Render(docx);
            File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "charts-docx.html"), string.Concat(word));
            int drawn = Regex.Matches(string.Concat(word), "<svg class=\"word-chart\"").Count;
            TestReport.Check(drawn == Titles.Length, "docx: every chart drawn (" + drawn + ")");
            CheckParts(word, "docx", layout: false);
        }

        TestReport.Step("Word-made charts in Calibri, Cambria and Times New Roman (text widths from the fonts' own advance widths)");
        string wordFonts = Path.Combine(Path.GetDirectoryName(path)!, "word-chart-fonts-v1.docx");
        if (File.Exists(wordFonts))
        {
            List<string> fonts = Render(wordFonts);
            File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "charts-word-fonts.html"), string.Concat(fonts));
            string allFonts = string.Concat(fonts);
            TestReport.Check(Regex.Matches(allFonts, "<svg class=\"word-chart\"").Count == 4 && allFonts.Contains(">Calibri Chart<") && allFonts.Contains(">Cambria Chart<") && allFonts.Contains(">Times Chart<"),
                             "Word-made charts: all four drawn with their titles");
            // Word's default data in a Hebrew Word: right-to-left category and series names, drawn right to left as Word draws them
            TestReport.Check(Regex.Matches(allFonts, "direction=\"rtl\"").Count >= 20, "Word-made charts: right-to-left names drawn right to left (" + Regex.Matches(allFonts, "direction=\"rtl\"").Count + ")");
        }

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

    /// <summary>The chart parts beyond series: log scale ticks, secondary axis, error bars, trend line with equation and R squared, stock boxes, 3-D floor, the stored plot position.</summary>
    private static void CheckParts(List<string> html, string format, bool layout = true)
    {
        string Block(string marker) => html.FirstOrDefault(h => h.Contains(">" + marker + "<")) ?? "";
        string log = Block("Log Growth");
        TestReport.Check(log.Contains(">1<") && log.Contains(">10<") && log.Contains(">100<") && log.Contains(">1000<") && !log.Contains(">500<"), format + ": logarithmic axis, a tick per power of ten");
        string two = Block("Two Axes");
        TestReport.Check(Regex.IsMatch(two, "text-anchor=\"start\">0[.][89]<"), format + ": secondary axis labelled on the right");
        string errors = Block("Error Bars");
        TestReport.Check(Regex.Matches(errors, "word-chart-error-bar").Count == 3, format + ": an error bar per point (" + Regex.Matches(errors, "word-chart-error-bar").Count + ")");
        string trend = Block("Trend Fit");
        var trendXs = Regex.Matches(trend, "class=\"word-chart-trendline\" d=\"([^\"]*)\"").Cast<Match>().SelectMany(m => Regex.Matches(m.Groups[1].Value, "[ML]([0-9.]+) ").Cast<Match>()).Select(m => double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)).ToList();
        var pointXs = Regex.Matches(trend, "<circle cx=\"([0-9.]+)\"").Cast<Match>().Select(m => double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)).ToList();
        TestReport.Check(trendXs.Count > 0 && pointXs.Count > 0 && trendXs.Max() > pointXs.Max() + 1 && trendXs.Min() < pointXs.Min() - 1, format + ": trend line extended forward and backward past the points");
        TestReport.Check(trend.Contains("word-chart-trendline") && trend.Contains("R&#178; = 1<") && Regex.IsMatch(trend, ">(f[(]x[)]|y) = "), format + ": trend line with its equation and R squared (" + string.Join(" | ", Regex.Matches(trend, ">([^<]*= [^<]*)<").Cast<Match>().Select(m => m.Groups[1].Value)) + ")");
        string candle = Block("Candle Stock");
        TestReport.Check(Regex.Matches(candle, "word-chart-hilow").Count == 3 && Regex.Matches(candle, "word-chart-updown").Count == 3, format + ": stock high-low lines and open-close boxes");
        TestReport.Check(candle.Contains("word-chart-updown\" x=") && candle.Contains("#FFFFFF") , format + ": a rising day is a white box");
        string depth = Block("Column Depth");
        var boxFills = Regex.Matches(depth, "class=\"word-chart-box\" d=\"[^\"]*\" fill=\"(#[0-9A-F]{6})\"").Cast<Match>().Select(m => m.Groups[1].Value).ToList();
        TestReport.Check(boxFills.Count >= 4 && boxFills.Distinct().Count() >= 2, format + ": 3-D columns as projected boxes, faces shaded differently (" + boxFills.Count + " faces, " + boxFills.Distinct().Count() + " shades)");
        TestReport.Check(Block("Line Depth").Contains("word-chart-ribbon") && Block("Area Depth").Contains("word-chart-area3d") && Block("Bar Depth").Contains("word-chart-box"),
                         format + ": 3-D line ribbons, area slabs and horizontal bar boxes");
        string crowded = Block("Crowded Axis");
        TestReport.Check(Regex.Matches(crowded, "rotate[(]-45 ").Count >= 6, format + ": crowded category labels turned 45 degrees (" + Regex.Matches(crowded, "rotate[(]-45 ").Count + ")");
        string wrappedLabels = Block("Wrapped Labels");
        bool wrapsInWord = format == "docx";
        TestReport.Check(wrapsInWord ? wrappedLabels.Contains(">region<") && !wrappedLabels.Contains("rotate(-45") : Regex.Matches(wrappedLabels, "rotate[(]-45 ").Count == 6,
                         format + (wrapsInWord ? ": long category names wrapped at spaces (Word)" : ": long names with a word wider than the slot turned 45 degrees (LibreOffice gives up the line break)"));
        string pie3D = Block("Pie Depth");
        TestReport.Check(Regex.Matches(pie3D, "word-chart-pie-top").Count == 4 && pie3D.Contains("word-chart-pie-side"), format + ": 3-D pie: four tilted slices over a darker rim");
        TestReport.Check(format == "docx" ? !pie3D.Contains("word-chart-smooth") : pie3D.Contains("word-chart-smooth") && pie3D.Contains("<linearGradient"),
                         format + (format == "docx" ? ": 3-D pie rim flat shaded (Office's flat shade mode)" : ": 3-D pie rim Gouraud shaded (dr3d:shade-mode gouraud) with a gradient"));
        TestReport.Check(Regex.Matches(Block("Range Errors"), "word-chart-error-bar").Count == 3, format + ": error amounts from a cell range, one bar per point");
        string columnSales = Block("Column Sales");
        TestReport.Check(columnSales.Contains(format == "odt" || format == "fodt" ? ">50<" : ">45<"), format + ": automatic value scale as " + (format == "docx" ? "Word (5 % headroom: 40 -> 45)" : "LibreOffice (one more step near the top: 40 -> 50)"));
        TestReport.Check(depth.Contains("word-chart-floor"), format + ": 3-D column chart drawn with wall and floor");
        if (layout)
        {
            TestReport.Check(depth.Contains("fill=\"#CCE5FF\""), format + ": the floor's own colour");
            // chart:coordinate-region of the first chart: x 1.94 cm of 11.999 cm; the frame is 4.7244 in = 340.16 pt wide -> 55.0 pt
            string column = Block("Column Sales");
            var lefts = Regex.Matches(column, "<path d=\"M([0-9.]+) [0-9.]+H").Cast<Match>().Select(m => double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)).ToList();
            double left = lefts.Count > 0 ? lefts.Min() : -1;
            TestReport.Check(Math.Abs(left - 1.94 / 11.999 * 340.16) < 0.6, format + ": plot area where LibreOffice put it (left " + left + " pt, LibreOffice " + Math.Round(1.94 / 11.999 * 340.16, 1) + ")");
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
