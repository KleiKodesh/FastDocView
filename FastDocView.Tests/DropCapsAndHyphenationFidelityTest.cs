using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Drop caps and automatic hyphenation in every format Word saves: test-data/drop-caps-hyphenation/drop-caps-hyphenation-v1
/// .docx / .doc / .rtf / .odt, saved by Word 16 from one document (make-drop-caps-hyphenation-fixtures.ps1): automatic
/// hyphenation on, a three-line drop cap "L" (6 pt from the text), a paragraph with hyphenation off, a two-line drop cap "M"
/// in the margin. Word keeps each initial in a framed paragraph of its own (.docx / .doc / .rtf); it must be drawn inside
/// the paragraph after it, and its own block left empty. ODT carries the drop cap on the text paragraph itself.
/// </summary>
public static class DropCapsAndHyphenationFidelityTest
{
    public static void Run()
    {
        TestReport.Step("Drop caps and hyphenation: Word-saved .docx / .doc / .rtf / .odt");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "drop-caps-hyphenation");
        foreach (string extension in new[] { ".docx", ".doc", ".rtf", ".odt" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "drop-caps-hyphenation-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/drop-caps-hyphenation/drop-caps-hyphenation-v1{extension} present"); continue; }
            Check(path, extension);
        }
    }

    private static void Check(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-drop-caps-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string Text(string html) => WordFieldsFidelityTest.Strip(html).Trim();
        string Block(string start) => blocks.FirstOrDefault(html => Text(html).StartsWith(start, StringComparison.Ordinal)) ?? "";

        string first = Block("Lorem");
        TestReport.Check(first.Contains("class=\"drop-cap\"") && first.Contains("--drop-cap-lines:3;") && Regex.IsMatch(first, "class=\"drop-cap\"[^>]*>(<[^>]+>)*L(<[^>]+>)*</span>"),
                         $"{format}: the three-line drop cap L is drawn at the start of its paragraph (got [{Shorten(first)}])");
        TestReport.Check(first.Contains("margin-inline-end:6pt;"), $"{format}: the drop cap keeps Word's 6 pt distance from the text");
        TestReport.Check(!first.Contains("font-size:") || !Regex.IsMatch(first, "class=\"drop-cap\"[^>]*>(<span[^>]*font-size:[^>]*>)"),
                         $"{format}: the initial loses the frame-filling font size Word gave it");
        if (format != ".odt")
            TestReport.Check(!blocks.Any(html => Text(html) == "L") && !blocks.Any(html => Text(html) == "M"),
                             $"{format}: the drop-cap frame paragraphs draw nothing of their own");
        string margin = Block("Margin");
        // ODF's style:drop-cap has no margin position (Word's .odt export writes the margin drop cap as an ordinary one)
        TestReport.Check((format == ".odt" || margin.Contains("drop-cap-in-margin")) && margin.Contains("--drop-cap-lines:2;"), $"{format}: the margin drop cap M over two lines (got [{Shorten(margin)}])");

        TestReport.Check(first.Contains("hyphens:auto;") && first.Contains("lang=\"en-US\""), $"{format}: automatic hyphenation with the paragraph's language");
        string unhyphenated = Block("Unhyphenated:");
        TestReport.Check(unhyphenated.Length > 0 && !unhyphenated.Contains("hyphens:auto") && !unhyphenated.Contains("lang="), $"{format}: \"Don't hyphenate\" paragraph not hyphenated");
    }

    private static string Shorten(string html) => html.Length > 300 ? html.Substring(0, 300) + "..." : html;
}
