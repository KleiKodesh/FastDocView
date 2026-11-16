using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Vertical cell text, RTF border spacing / shadow / border groups and ODT drop caps across spans:
/// test-data/borders-cell-direction/borders-cell-direction-v1 .docx / .doc / .rtf saved by Word 16
/// (make-borders-cell-direction-fixtures.ps1, Word's PDF beside it) and .odt converted from the .docx by LibreOffice; cells
/// run horizontal, downward, upward and East-Asian vertical; two paragraphs share one shadowed box (4 pt / 10 pt from the
/// text), a third has a box of its own. drop-cap-spans-v1.fodt (hand-written, checked against LibreOffice's PDF): drop caps
/// whose characters run over several spans, a whole word split by a span, an initial cut short by a tab, and cells with
/// style:writing-mode tb-rl and loext:writing-mode bt-lr.
/// </summary>
public static class BordersAndCellDirectionFidelityTest
{
    private static string Folder => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "borders-cell-direction"));

    public static void Run()
    {
        TestReport.Step("Vertical cell text, RTF border spacing / shadow / groups, ODT drop caps over spans");
        foreach (string extension in new[] { ".docx", ".doc", ".rtf", ".odt" })
        {
            string path = Path.Combine(Folder, "borders-cell-direction-v1" + extension);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/borders-cell-direction/borders-cell-direction-v1{extension} present"); continue; }
            CheckCells(path, extension);
            if (extension == ".rtf") CheckRtfBorders(path);
        }
        string groups = Path.Combine(Folder, "borders-cell-groups-v1.rtf");
        if (File.Exists(groups)) CheckRtfGroupsInCellsAndTextBoxes(groups);
        else TestReport.Check(false, "test-data/borders-cell-direction/borders-cell-groups-v1.rtf present");
        string spans = Path.Combine(Folder, "drop-cap-spans-v1.fodt");
        if (File.Exists(spans)) CheckDropCapSpans(spans);
        else TestReport.Check(false, "test-data/borders-cell-direction/drop-cap-spans-v1.fodt present");
    }

    private static List<string> Render(string path, string name)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-borders-cells-" + name) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }

    /// <summary>The style of the td holding <paramref name="text"/>.</summary>
    private static string CellStyle(string html, string text)
    {
        foreach (Match cell in Regex.Matches(html, "<td([^>]*)>(.*?)</td>", RegexOptions.Singleline))
            if (WordFieldsFidelityTest.Strip(cell.Groups[2].Value).Trim() == text)
                return Regex.Match(cell.Groups[1].Value, "style=\"([^\"]*)\"").Groups[1].Value;
        return "(no cell)";
    }

    private static void CheckCells(string path, string format)
    {
        string html = string.Join("\n", Render(path, format.TrimStart('.')));
        // Word's PDF: Down and Asian run top to bottom (direction 0,1), Up bottom to top (0,-1), Flat stays horizontal
        TestReport.Check(!CellStyle(html, "Flat").Contains("writing-mode"), $"{format}: the horizontal cell stays horizontal");
        TestReport.Check(CellStyle(html, "Down").Contains("writing-mode:vertical-rl;"), $"{format}: the downward cell runs top to bottom (got [{CellStyle(html, "Down")}])");
        TestReport.Check(CellStyle(html, "Up").Contains("writing-mode:sideways-lr;"), $"{format}: the upward cell runs bottom to top (got [{CellStyle(html, "Up")}])");
        TestReport.Check(CellStyle(html, "Asian").Contains("writing-mode:vertical-rl;"), $"{format}: the East-Asian vertical cell runs top to bottom (got [{CellStyle(html, "Asian")}])");
    }

    private static void CheckRtfBorders(string path)
    {
        List<string> blocks = Render(path, "rtf-borders");
        string Block(string start) => blocks.FirstOrDefault(html => WordFieldsFidelityTest.Strip(html).Trim().StartsWith(start, StringComparison.Ordinal)) ?? "";
        string Style(string html) => Regex.Match(html, "<p[^>]*style=\"([^\"]*)\"").Groups[1].Value;
        string one = Style(Block("Boxed one")), two = Style(Block("Boxed two")), alone = Style(Block("Alone"));
        // Word's PDF: one box around both (top line above "Boxed one", bottom line + shadow below "Boxed two", no line between)
        TestReport.Check(one.Contains("border-top:") && !one.Contains("border-bottom:"), $".rtf: the first paragraph of the border group is open at the bottom (got [{one}])");
        TestReport.Check(!two.Contains("border-top:") && two.Contains("border-bottom:"), $".rtf: the second paragraph of the border group is open at the top (got [{two}])");
        TestReport.Check(one.Contains("padding-top:4pt;") && one.Contains("padding-inline-start:10pt;") && one.Contains("padding-inline-end:10pt;") || one.Contains("padding:4pt 10pt "),
                         $".rtf: \\brsp keeps the text 4 pt below the top line and 10 pt from the sides (got [{one}])");
        TestReport.Check(two.Contains("padding:") && Regex.IsMatch(two, "padding:[^;]* 10pt 4pt;"), $".rtf: 4 pt between the last text and the bottom line (got [{two}])");
        TestReport.Check(one.Contains("box-shadow:0.5pt 0 0 0 #000;") && two.Contains("box-shadow:0.5pt 0.5pt 0 0 #000;"),
                         $".rtf: \\brdrsh shadows the box on the right, and below its last paragraph");
        TestReport.Check(alone.Contains("border-top:1.5pt solid") && alone.Contains("border-bottom:1.5pt solid") && !alone.Contains("box-shadow"),
                         $".rtf: a different border is a box of its own (got [{alone}])");
    }

    private static void CheckDropCapSpans(string path)
    {
        List<string> blocks = Render(path, "fodt-drop-caps");
        string Block(string contains) => blocks.FirstOrDefault(html => WordFieldsFidelityTest.Strip(html).Contains(contains, StringComparison.Ordinal)) ?? "";
        string DropCap(string html) => Regex.Match(html, "<span class=\"drop-cap\"[^>]*>((?:<span[^>]*>[^<]*</span>|[^<])*)</span>").Groups[1].Value;
        string Text(string html) => WordFieldsFidelityTest.Strip(html);
        // LibreOffice's PDF: "Abc" (21.2 / 35.3 pt) drops, "def" runs on in the text; "Word" drops whole; "X" alone drops (the tab ends it)
        string three = DropCap(Block("def three"));
        TestReport.Check(Text(three) == "Abc" && three.Contains("color:") && three.Contains("font-weight:bold"),
                         $".fodt: three drop-cap characters over two spans, each in its span's look (got [{three}])");
        TestReport.Check(!three.Contains("font-size"), ".fodt: the spans inside the initial keep the drop cap's size");
        TestReport.Check(Text(DropCap(Block("rest: the whole"))) == "Word", $".fodt: a whole-word drop cap split by a span (got [{DropCap(Block("rest: the whole"))}])");
        TestReport.Check(Text(DropCap(Block("Tab ends"))) == "X", $".fodt: a tab ends the drop cap (got [{DropCap(Block("Tab ends"))}])");
        string table = string.Join("\n", blocks);
        TestReport.Check(CellStyle(table, "Down").Contains("writing-mode:vertical-rl;") && CellStyle(table, "Up").Contains("writing-mode:sideways-lr;")
                         && !CellStyle(table, "Flat").Contains("writing-mode"), $".fodt: style:writing-mode tb-rl and loext:writing-mode bt-lr cells");
    }

    /// <summary>Word's PDF of borders-cell-groups-v1: one box around "Cell one" + "Cell two" (no line between), one around "Box one" + "Box two".</summary>
    private static void CheckRtfGroupsInCellsAndTextBoxes(string path)
    {
        string html = string.Join(" ", Render(path, "rtf-cell-groups"));
        string Open(string text)
        {
            int at = html.IndexOf(text, StringComparison.Ordinal);
            if (at < 0) return "(missing)";
            int open = Math.Max(html.LastIndexOf("<p ", at, StringComparison.Ordinal), html.LastIndexOf("<span class=\"text-box-paragraph\"", at, StringComparison.Ordinal));
            return open < 0 ? "(no tag)" : html.Substring(open, html.IndexOf('>', open) - open);
        }
        foreach ((string first, string second) in new[] { ("Cell one", "Cell two"), ("Box one", "Box two") })
        {
            string one = Open(first), two = Open(second);
            TestReport.Check(one.Contains("border-top:") && !one.Contains("border-bottom:") && !two.Contains("border-top:") && two.Contains("border-bottom:"),
                             $".rtf: {first} and {second} share one box (got [{one}] [{two}])");
        }
    }
}
