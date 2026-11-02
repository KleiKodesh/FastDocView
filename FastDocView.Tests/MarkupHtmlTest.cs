using System.Text;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>
/// Loose .html / .htm files: format detection, the encoding order (BOM, then meta charset, then text detection),
/// one block per block-level element, a body-less file, the head's styles kept and scoped, scripts and active
/// content stripped, app links marked like the Word readers', and in-memory / sidecar / 1-7 byte reads agreeing.
/// </summary>
public static class MarkupHtmlTest
{
    private const string Body = """
        <h1 onclick="x()">כותרת ראשית</h1>
        <p class="lead">פסקה ראשונה <b>מודגשת</b></p>
        <script>alert('x')</script>
        <p><a href="otzaria://book/1">ספר</a> <a href="https://e.org/a">web</a> <a href="javascript:alert(1)">bad</a> <a href="#n">here</a></p>
        <iframe src="https://e.org"></iframe><object data="x.swf"></object>
        <h2>סעיף</h2>
        <table><tr><td>א</td><td>ב</td></tr></table>
        <p><img src="a.png" onerror="x()"> סוף</p>
        """;

    public static void Run()
    {
        TestReport.Step("HTML files (.html / .htm)");
        string directory = TestReport.OutputDirectory;
        string full = "<!DOCTYPE html><html><head><meta charset=\"windows-1255\"><title>t</title>"
                      + "<style>body { color: #123456 } p.lead { font-weight: bold } @import url(x.css);</style>"
                      + "<script>var a = '<p>no</p>';</script></head><body>\n" + Body + "\n</body></html>";

        string meta = Path.Combine(directory, "html-meta-1255.html");
        File.WriteAllBytes(meta, Windows1255(full));
        Verify(meta, 1255, "meta charset windows-1255");

        string bom = Path.Combine(directory, "html-bom-utf8.htm");
        File.WriteAllBytes(bom, [0xEF, 0xBB, 0xBF, .. Encoding.UTF8.GetBytes(full)]);   // the BOM beats the meta
        Verify(bom, MarkupTextEncoding.Utf8, "UTF-8 BOM over meta");

        string bare = Path.Combine(directory, "html-no-body.html");
        File.WriteAllBytes(bare, Encoding.UTF8.GetBytes("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n" + Body));
        Verify(bare, MarkupTextEncoding.Utf8, "no <body>, http-equiv charset");

        string detected = Path.Combine(directory, "html-detected.html");
        File.WriteAllBytes(detected, Windows1255("<html><body>\n" + Body + "\n</body></html>"));
        Verify(detected, 1255, "no charset: text detection");
    }

    private static byte[] Windows1255(string text) => MarkupTextEncoding.FromCodePage(1255).GetBytes(text);

    private static void Verify(string path, int codePage, string what)
    {
        TestReport.Check(WordWindowedDocumentOpener.DetectFormat(path) == WordWindowedDocumentOpener.WordFileFormat.Html, $"{what}: HTML detected");
        string? first = null;
        foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar, WordIndexMode.Sidecar })
        {
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache"), ForcedMode = mode };
            using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
            string all = Check(document, codePage, $"{what} ({mode}{(document.Decision.HasCachedIndex ? ", cached" : "")})");
            first ??= all;
            TestReport.Check(all == first, $"{what} ({mode}): same blocks as the first open");
        }
        using var trickle = new TrickleStream(File.ReadAllBytes(path));
        using MarkupWindowedDocument trickled = MarkupWindowedDocument.OpenFromStream(trickle, path, MarkupSourceKind.Html,
            new WordIndexDecision(path, WordIndexMode.InMemory, false, trickle.Length, 0, "test"), null, 0, 0);
        TestReport.Check(Check(trickled, codePage, $"{what} (1-7 byte reads)") == first, $"{what}: 1-7 byte reads give the same blocks");
    }

    private static string Check(MarkupWindowedDocument document, int codePage, string what)
    {
        TestReport.Check(document.Index.CodePage == codePage, $"{what}: encoding {document.Index.CodePage} ({document.Index.EncodingReason}) == {codePage}");
        TestReport.Check(document.IsRightToLeft, $"{what}: right-to-left");
        var blocks = document.RenderBlocks(0, document.BlockCount);
        string all = string.Concat(blocks.Select(b => b.Html + "\n"));
        TestReport.Check(document.BlockCount >= 6, $"{what}: one block per element ({document.BlockCount})");
        TestReport.Check(document.Headings.Count == 2 && document.Headings[0].Text == "כותרת ראשית" && document.Headings[1].Level == 2,
            $"{what}: h1 / h2 outline ({string.Join(" | ", document.Headings.Select(h => h.Level + ":" + h.Text))})");
        TestReport.Check(all.Contains("מודגשת") && all.Contains("סוף") && all.Contains("<td>א</td>"), $"{what}: text and table kept");
        TestReport.Check(!all.Contains("alert") && !all.Contains("onclick") && !all.Contains("onerror") && !all.Contains("<iframe") && !all.Contains("<object")
                         && !all.Contains("javascript"), $"{what}: scripts, handlers, iframes, objects, script URLs removed");
        TestReport.Check(all.Contains("<a class=\"link app-link\" data-app-link=\"1\" href=\"otzaria://book/1\"") || all.Contains("href=\"otzaria://book/1\" class=\"link app-link\" data-app-link=\"1\""),
            $"{what}: otzaria link marked as an app link");
        TestReport.Check(all.Contains("https://e.org/a\" class=\"link\" rel=\"noopener noreferrer\" target=\"_blank\"") && all.Contains("href=\"#n\" class=\"link\">"),
            $"{what}: web link opens outside, #anchor stays in place");
        bool hasHead = document.Index.DocumentCss.Length > 0;
        if (hasHead)
            TestReport.Check(document.Stylesheet.Contains(".word-document { color: #123456 }") && document.Stylesheet.Contains(".word-document p.lead") && !document.Stylesheet.Contains("@import"),
                $"{what}: head styles kept, scoped, @-rules dropped");
        return all;
    }
}
