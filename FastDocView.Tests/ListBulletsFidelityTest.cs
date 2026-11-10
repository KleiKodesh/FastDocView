using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Picture bullets, symbol-font bullets and a label with its own text style: test-data/word-lists/list-bullets-v1.docx
/// (Word 16, make-list-bullets-fixture.ps1), list-bullets-v1.odt (LibreOffice's conversion of it) and Word 16's .doc / .rtf saves
/// of it. The picture bullet is drawn as an &lt;img&gt; whose id the document resolves to the 40 x 20 picture, at Word's 2:1 ratio
/// and font-relative height (font size less 2 pt);
/// Symbol U+F0B7 becomes U+2022 in the text font, Wingdings U+F0A7 keeps its code and font; the numbered label is bold red.
/// </summary>
public static class ListBulletsFidelityTest
{
    public static void Run()
    {
        TestReport.Step("List bullets: picture bullets, symbol-font bullets, label text style (.docx / .odt / .doc / .rtf)");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-lists");
        foreach (string extension in new[] { ".docx", ".odt", ".doc", ".rtf" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "list-bullets-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-lists/list-bullets-v1{extension} present"); continue; }
            Check(path, extension);
        }
        // the same .rtf without its flat label groups (strip-list-labels.py): every label counted and drawn from the level's formatting
        string bare = Path.GetFullPath(Path.Combine(folder, "list-bullets-nolisttext-v1.rtf"));
        if (File.Exists(bare)) Check(bare, ".rtf", hasFlatLabels: false);
        else TestReport.Check(false, "test-data/word-lists/list-bullets-nolisttext-v1.rtf present");
        LegacyListsFidelityTest.Run();
    }

    private static void Check(string path, string format, bool hasFlatLabels = true)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-list-bullets-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string Block(string text) => blocks.FirstOrDefault(html => WordFieldsFidelityTest.Strip(html).Contains(text)) ?? "";

        foreach (string size in new[] { "11", "24" })
        {
            string picture = Block("picture " + size);
            Match image = Regex.Match(picture, "<img class=\"word-picture word-picture-bullet\"[^>]*data-word-picture=\"([^\"]+)\"[^>]*>");
            TestReport.Check(image.Success && image.Value.Contains("height:max(calc(1em - 2pt),1px);") && Regex.IsMatch(image.Value, @"aspect-ratio:(\d+(\.\d+)?)/(\d+(\.\d+)?)")
                             && Math.Abs(Ratio(image.Value) - 2) < 0.05,
                             $"{format}: picture bullet ({size} pt) drawn as a 2:1 picture 2 pt less high than its font (got [{Shorten(picture)}])");
            Match pictureMarker = Regex.Match(picture, "<span class=\"list-marker\"[^>]*>");
            TestReport.Check(pictureMarker.Success && pictureMarker.Value.Contains("font-size:" + size + "pt"),
                             $"{format}: picture bullet label sized by the label's own font ({size} pt), not the paragraph's (got [{Shorten(pictureMarker.Value)}])");
            WordDocumentPicture? bytes = image.Success ? document.ReadPicture(System.Net.WebUtility.HtmlDecode(image.Groups[1].Value)) : null;
            TestReport.Check(bytes is not null && bytes.ContentType == "image/png", $"{format}: picture bullet ({size} pt) resolves to its PNG");
        }
        string symbol = Block("symbol Symbol");
        TestReport.Check(symbol.Contains("\u2022") && !symbol.Contains("\uF0B7") && !symbol.Contains("font-family:'Symbol'"),
                         $"{format}: Symbol bullet recoded to U+2022 (got [{Shorten(symbol)}])");
        string wingdings = Block("symbol Wingdings");
        TestReport.Check(wingdings.Contains("\uF0A7") && wingdings.Contains("Wingdings"), $"{format}: Wingdings bullet keeps its code and font (got [{Shorten(wingdings)}])");
        // symbol bullets draw in the Latin set: the mark's 11 pt, not the complex-script 12 pt (Word 16 PDF: SymbolMT / Wingdings 11)
        TestReport.Check(Label(symbol).Contains("font-size:11pt") && Label(wingdings).Contains("font-size:11pt"),
                         $"{format}: symbol bullets sized by the mark's Latin size (got [{Shorten(Label(symbol))}] [{Shorten(Label(wingdings))}])");
        // "1." in a right-to-left paragraph draws in the complex-script set (WordListLabelScript): Word 16 PDF shows it in
        // Arial 12, red, NOT bold - the level set only the Latin bold (b / sprmCFBold / ), not bCs / sprmCFBoldBi / b
        string styled = Block("styled label");
        string label = Label(styled);
        TestReport.Check(label.Contains("#FF0000") && !label.Contains("font-weight:bold") && !label.Contains("font-size:11pt"),
                         $"{format}: the numbered label is red, complex-script size and not bold, as Word draws it (got [{Shorten(label)}])");
        // Word 16's re-save of the label-less .rtf writes the label's complex-script font as the mark's (\af0), not Arial: no font check
        if (format is ".doc" || (format == ".rtf" && hasFlatLabels))
            TestReport.Check(label.Contains("font-family:'Arial'") || (format == ".doc" && Regex.IsMatch(styled, "<p [^>]*font-family:'Arial'")),
                             $"{format}: the numbered label draws in the complex-script font Arial (got [{Shorten(styled)}])");
    }

    /// <summary>The list-marker span and what it holds, up to the paragraph text.</summary>
    private static string Label(string html)
    {
        int start = html.IndexOf("<span class=\"list-marker\"", StringComparison.Ordinal);
        if (start < 0) return "";
        int end = html.IndexOf("</span></span>", start, StringComparison.Ordinal);
        return end < 0 ? html.Substring(start) : html.Substring(start, end - start);
    }

    private static double Ratio(string image)
    {
        Match m = Regex.Match(image, @"aspect-ratio:(\d+(?:\.\d+)?)/(\d+(?:\.\d+)?)");
        return double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) / double.Parse(m.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string Shorten(string html) => html.Length > 400 ? html.Substring(0, 400) + "..." : html;
}
