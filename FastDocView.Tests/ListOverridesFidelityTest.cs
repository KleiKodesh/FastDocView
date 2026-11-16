namespace FastDocView.Tests;

/// <summary>
/// RTF list cases against Word 16's print (test-data/word-lists/make-list-overrides.py): a level whose restart limit Word saves as
/// \levelnorestart alone (never restarts: 2.1.4), a list linked to a numbering style (\liststyleid; the emptied linked list draws the
/// style's levels and still counts on its own), an \lfolevel holding a whole \listlevel ((e), (f), (g)), and counted labels in the level's
/// own colour and size. Plus the .odt side of label colour / size: LibreOffice Writer prints a bullet level's fo:color, fo:font-size
/// percentage and text:bullet-relative-size as the plain label (soffice PDF: black 12 pt; sw/source/core/text/txtfld.cxx builds the
/// bullet font from the paragraph and the character style only), so FastDocView ignores them too.
/// </summary>
public static class ListOverridesFidelityTest
{
    public static void Run()
    {
        TestReport.Step("RTF lists: restart limit, list style links, levels replaced in an override, label colour / size (Word 16 print); ODT label look (LibreOffice print)");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-lists"));
        string expectedPath = Path.Combine(folder, "list-overrides-v1.expected.txt");
        if (!File.Exists(expectedPath)) { TestReport.Check(false, "test-data/word-lists/list-overrides-v1.expected.txt present"); return; }
        // Word's printed labels for every item line except the bullets (pdftotext gives their glyphs as private-use codes)
        List<string> expected = File.ReadAllLines(expectedPath).Where(line => line.Contains("item ") && !line.StartsWith("item", StringComparison.Ordinal)
                                                                          && line[0] < 0x80).Select(line => line.Trim()).ToList();
        foreach (string name in new[] { "list-overrides-nolisttext-v1.rtf", "list-overrides-stylelink-v1.rtf" })
        {
            string path = Path.Combine(folder, name);
            if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-lists/" + name + " present"); continue; }
            List<string> blocks = Render(path, name);
            List<string> got = blocks.Select(html => WordFieldsFidelityTest.Strip(html).Trim()).Where(text => text.Contains("item ") && text[0] < 0x80 && !text.StartsWith("item", StringComparison.Ordinal))
                                     .Select(text => System.Text.RegularExpressions.Regex.Replace(text, @"^(\S+?)\s*item", "$1 item")).ToList();
            TestReport.Check(expected.Count == 22 && got.SequenceEqual(expected), $"{name}: every counted label as Word 16 prints it (got [{string.Join(" | ", got)}], Word [{string.Join(" | ", expected)}])");
            if (name.Contains("nolisttext"))
            {
                string bullet = blocks.FirstOrDefault(html => WordFieldsFidelityTest.Strip(html).EndsWith("item a", StringComparison.Ordinal) && html.Contains("●")) ?? "";
                string number = blocks.FirstOrDefault(html => WordFieldsFidelityTest.Strip(html).EndsWith("2.item c", StringComparison.Ordinal)) ?? "";
                TestReport.Check(bullet.Contains("font-size:20pt") && bullet.Contains("#FF0000") && number.Contains("font-size:8pt") && number.Contains("#0000FF"),
                                 $"{name}: counted bullet 20 pt red, counted number 8 pt blue (the level's \\fs \\cf)");
            }
        }

        string odt = Path.Combine(folder, "list-bullet-look-v1.fodt");
        if (!File.Exists(odt)) { TestReport.Check(false, "test-data/word-lists/list-bullet-look-v1.fodt present"); return; }
        List<string> odtBlocks = Render(odt, "fodt");
        foreach (string item in new[] { "item a", "item b" })
        {
            string html = odtBlocks.FirstOrDefault(block => WordFieldsFidelityTest.Strip(block).EndsWith(item, StringComparison.Ordinal)) ?? "";
            int marker = html.IndexOf("list-marker", StringComparison.Ordinal);
            int text = html.IndexOf(item, StringComparison.Ordinal);
            string label = marker >= 0 && text > marker ? html.Substring(marker, text - marker) : "";
            TestReport.Check(label.Length > 0 && !label.Contains("color:") && !label.Contains("font-size:"),
                             $".fodt {item}: label drawn plain as LibreOffice prints it (no level colour or relative size) (got [{label}])");
        }
    }

    private static List<string> Render(string path, string tag)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-list-overrides-" + tag.Replace('.', '-')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }
}
