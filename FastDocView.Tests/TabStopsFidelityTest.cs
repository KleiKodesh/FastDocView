namespace FastDocView.Tests;

/// <summary>
/// Tab stops in every format: test-data/tab-stops/tab-stops-v2 .docx (scripts/make-tab-stops-docx.py), .doc / .rtf saved from it by
/// Word 16 and .odt by LibreOffice (scripts/save-tab-stops.ps1). Each paragraph carries the tab layout data the page's tab pass reads
/// (data-tabs: default interval, start indent, first-line indent, relative-to-indent flag, then the stops with alignment and
/// leader), and the text keeps its tab characters. Where the text after each tab lands is checked against Word's and LibreOffice's
/// PDFs live (vue-demo/playwright/tabStopsCheck.mjs).
/// </summary>
public static class TabStopsFidelityTest
{
    public static void Run()
    {
        TestReport.Step("Tab stops: Word-saved .docx / .doc / .rtf, LibreOffice .odt");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "tab-stops");
        RunRtfGroupScope();
        foreach (string extension in new[] { ".docx", ".doc", ".rtf", ".odt" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "tab-stops-v2" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/tab-stops/tab-stops-v2{extension} present"); continue; }
            Check(path, extension);
        }
    }

    /// <summary>RTF tab stops are paragraph properties of the group state: a group's stops end with its '}' (as Word reads them).</summary>
    public static void RunRtfGroupScope()
    {
        TestReport.Step("Tab stops: an RTF group scopes its stops");
        string path = Path.Combine(TestReport.OutputDirectory, "tab-stops-group-v1.rtf");
        File.WriteAllText(path, @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\pard\tx1440 Outer\tab one\par {\pard\tqr\tldot\tx5000 Inner\tab two\par}Back\tab three\par}");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-tab-stops-group") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string Block(string start) => blocks.FirstOrDefault(html => WordFieldsFidelityTest.Strip(html).Trim().StartsWith(start, StringComparison.Ordinal)) ?? "";
        TestReport.Check(Block("Inner").Contains(";250 r d\""), $"rtf: the group's own right stop (got [{Shorten(Block("Inner"))}])");
        TestReport.Check(Block("Back").Contains(";72 l \"") && !Block("Back").Contains("250"), $"rtf: after the group's '}}' the outer stop is back (got [{Shorten(Block("Back"))}])");
    }

    private static void Check(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-tab-stops-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string Text(string html) => WordFieldsFidelityTest.Strip(html).Trim();
        string Block(string start) => blocks.FirstOrDefault(html => Text(html).StartsWith(start, StringComparison.Ordinal)) ?? "";
        bool odt = format == ".odt";
        // the right stop at 9350 twips: 467.5 pt (LibreOffice writes 6.4929in, which rounds to the same)
        const string right = "467.5";

        TestReport.Check(Block("Chapter One").Contains($";{right} r d\""), $"{format}: the TOC line's right stop with a dot leader (from its style)");
        TestReport.Check(Block("A much longer").Contains($";{right} r d\""), $"{format}: the next TOC line keeps the style's stop (.rtf: carried over a \\par without \\pard)");
        TestReport.Check(Block("Hyphen leader").Contains(" r h\"") && Block("Underscore leader").Contains(" r u\""), $"{format}: hyphen and underscore leaders");
        // .doc keeps heavy (tlc 4) and middle dot (tlc 5); RTF and LibreOffice have no heavy leader (\tlth and ODF write an underscore)
        if (format is ".docx" or ".doc")
            TestReport.Check(Block("Heavy leader").Contains(" r H\"") && Block("Middle dot leader").Contains(" r m\""), $"{format}: heavy and middle dot leaders");
        TestReport.Check(Block("Left").Contains(" c ;") && Block("Left").Contains($";{right} r \""), $"{format}: centre and right stops");
        TestReport.Check(Block("Price").Contains(";200 d \"") || Block("Price").Contains(";199.99 d \""), $"{format}: the decimal stop (got [{Shorten(Block("Price"))}])");
        if (!odt) TestReport.Check(Block("Bar").Contains(";100 b ;150 l \""), $"{format}: a bar stop and a left stop (LibreOffice drops bar stops)");
        TestReport.Check(Block("in").Contains("data-tabs=\"36 50 0 "), $"{format}: an indented paragraph says its start indent (default stops count from the margin)");
        TestReport.Check(Block("1.").Contains("data-tabs=\"36 72 -72 "), $"{format}: the hanging indent (the start indent is a stop)");
        TestReport.Check(Block("Right to left").Contains(" dir=\"rtl\"") && Block("Right to left").Contains($";{right} r d\""), $"{format}: the right-to-left paragraph's end stop");
        TestReport.Check(!Block("Cleared").Contains($";{right} r"), $"{format}: a cleared style stop is gone");
        TestReport.Check(Block("Chapter One").Contains("Chapter One\t") || Text(Block("Chapter One")).Contains('\t'), $"{format}: the tab stays a tab character");
        if (format == ".docx") TestReport.Check(Block("Absolute").Contains("data-ptab=\"r d m\""), ".docx: an absolute position tab, right-aligned at the margin with dots");
        TestReport.Check(!Block("a").Contains("data-tabs"), $"{format}: a paragraph with only the default 36 pt stops needs no tab data (tab-size places them)");
    }

    private static string Shorten(string html) => html.Length > 300 ? html.Substring(0, 300) + "..." : html;
}
