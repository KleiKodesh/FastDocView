using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Picture bullets, symbol-font bullets and a label with its own text style: test-data/word-lists/list-bullets-v1.docx
/// (Word 16, make-list-bullets-fixture.ps1) and list-bullets-v1.odt (LibreOffice's conversion of it). The picture bullet is
/// drawn as an &lt;img&gt; whose id the document resolves to the 40 x 20 picture, at Word's 2:1 ratio and font-relative height;
/// Symbol U+F0B7 becomes U+2022 in the text font, Wingdings U+F0A7 keeps its code and font; the numbered label is bold red.
/// </summary>
public static class ListBulletsFidelityTest
{
    public static void Run()
    {
        TestReport.Step("List bullets: picture bullets, symbol-font bullets, label text style (.docx / .odt)");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-lists");
        foreach (string extension in new[] { ".docx", ".odt" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "list-bullets-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-lists/list-bullets-v1{extension} present"); continue; }
            Check(path, extension);
        }
    }

    private static void Check(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-list-bullets-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string Block(string text) => blocks.FirstOrDefault(html => WordFieldsFidelityTest.Strip(html).Contains(text)) ?? "";

        foreach (string size in new[] { "11", "24" })
        {
            string picture = Block("picture " + size);
            Match image = Regex.Match(picture, "<img class=\"word-picture word-picture-bullet\"[^>]*data-word-picture=\"([^\"]+)\"[^>]*>");
            TestReport.Check(image.Success && image.Value.Contains("height:0.9em;") && Regex.IsMatch(image.Value, @"aspect-ratio:(\d+(\.\d+)?)/(\d+(\.\d+)?)")
                             && Math.Abs(Ratio(image.Value) - 2) < 0.05,
                             $"{format}: picture bullet ({size} pt) drawn as a 2:1 picture 0.9 em high (got [{Shorten(picture)}])");
            WordDocumentPicture? bytes = image.Success ? document.ReadPicture(System.Net.WebUtility.HtmlDecode(image.Groups[1].Value)) : null;
            TestReport.Check(bytes is not null && bytes.ContentType == "image/png", $"{format}: picture bullet ({size} pt) resolves to its PNG");
        }
        string symbol = Block("symbol Symbol");
        TestReport.Check(symbol.Contains("\u2022") && !symbol.Contains("\uF0B7") && !symbol.Contains("font-family:'Symbol'"),
                         $"{format}: Symbol bullet recoded to U+2022 (got [{Shorten(symbol)}])");
        string wingdings = Block("symbol Wingdings");
        TestReport.Check(wingdings.Contains("\uF0A7") && wingdings.Contains("Wingdings"), $"{format}: Wingdings bullet keeps its code and font (got [{Shorten(wingdings)}])");
        string styled = Block("styled label");
        Match marker = Regex.Match(styled, "<span class=\"list-marker\"[^>]*>");
        TestReport.Check(marker.Success && marker.Value.Contains("font-weight:bold") && marker.Value.Contains("#FF0000"),
                         $"{format}: the numbered label is bold red (got [{Shorten(styled)}])");
    }

    private static double Ratio(string image)
    {
        Match m = Regex.Match(image, @"aspect-ratio:(\d+(?:\.\d+)?)/(\d+(?:\.\d+)?)");
        return double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) / double.Parse(m.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string Shorten(string html) => html.Length > 400 ? html.Substring(0, 400) + "..." : html;
}
