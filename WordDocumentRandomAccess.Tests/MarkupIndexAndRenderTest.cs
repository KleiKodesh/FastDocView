using System.Diagnostics;
using System.Text;
using WordDocumentRandomAccess.Markup;
using static WordDocumentRandomAccess.Tests.SyntheticMarkupGenerator;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// .txt (read as HTML) and MHTML: synthetic files in every encoding / line ending / transfer encoding,
/// encoding detection, block and heading counts, the tricky prefix blocks, chunk-edge feeds (1-7 bytes
/// per read), in-memory and sidecar opens, random windows, and timings.
/// </summary>
public static class MarkupIndexAndRenderTest
{
    public static void RunText(int count)
    {
        string directory = TestReport.OutputDirectory;
        TestReport.Step($"Synthetic .txt read as HTML, {ExpectedBlockCount(count):N0} blocks");
        var variants = new (string name, int codePage, bool bom, string newline)[]
        {
            ("utf8", 65001, false, "\n"), ("utf8bom", 65001, true, "\r\n"), ("cp1255", 1255, false, "\r\n"),
            ("cp1255cr", 1255, false, "\r"), ("utf16le", 1200, true, "\r\n"), ("utf16le-nobom", 1200, false, "\n"), ("utf16be", 1201, true, "\n"),
        };
        foreach ((string name, int codePage, bool bom, string newline) in variants)
        {
            string path = Path.Combine(directory, $"synthetic-text-{count}-{name}-v2.txt");
            if (!File.Exists(path)) WriteText(path, count, MarkupTextEncoding.FromCodePage(codePage), bom, newline);
            Console.WriteLine($"   — {name} ({new FileInfo(path).Length >> 10} KB)");
            TestReport.Check(WordWindowedDocumentOpener.DetectFormat(path) == WordWindowedDocumentOpener.WordFileFormat.PlainText, $"{name}: .txt detected");

            using (var trickle = new TrickleStream(File.ReadAllBytes(path)))
            using (MarkupWindowedDocument trickled = MarkupWindowedDocument.OpenFromStream(trickle, path, MarkupSourceKind.PlainText,
                       new WordIndexDecision(path, WordIndexMode.InMemory, false, trickle.Length, 0, "test"), null, 0, 0))
            {
                VerifyText(trickled, count, codePage, name + " fed in 1-7 byte reads");
                using MarkupWindowedDocument whole = MarkupWindowedDocument.OpenFromStream(new MemoryStream(File.ReadAllBytes(path)), path, MarkupSourceKind.PlainText,
                       new WordIndexDecision(path, WordIndexMode.InMemory, false, 0, 0, "test"), null, 0, 0);
                TestReport.Check(whole.Index.BlockStarts.SequenceEqual(trickled.Index.BlockStarts) && whole.Index.BlockStackIds.SequenceEqual(trickled.Index.BlockStackIds),
                                 $"{name}: trickled index identical to the one-read index");
            }
            OpenBothModes(path, MarkupSourceKind.PlainText, document => VerifyText(document, count, codePage, name));
        }
    }

    private static void OpenBothModes(string path, MarkupSourceKind kind, Action<MarkupWindowedDocument> verify)
    {
        foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
        {
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache"), ForcedMode = mode };
            string indexPath = policy.IndexPathFor(path);
            if (File.Exists(indexPath)) File.Delete(indexPath);
            using (var document = (MarkupWindowedDocument)TestReport.Timed($"open {mode} (cold)", () => WordWindowedDocumentOpener.Open(path, policy)))
            {
                verify(document);
                TestReport.Check(!document.HasSourceChanged, "source unchanged");
            }
            if (mode != WordIndexMode.Sidecar) continue;
            TestReport.Check(File.Exists(indexPath), "sidecar index written");
            using var reopened = (MarkupWindowedDocument)TestReport.Timed("re-open from cached index", () => WordWindowedDocumentOpener.Open(path, policy));
            TestReport.Check(reopened.Decision.HasCachedIndex, "cached index used");
            verify(reopened);
        }
    }

    private static void VerifyText(MarkupWindowedDocument document, int count, int codePage, string what)
    {
        TestReport.Check(document.Index.CodePage == codePage, $"{what}: encoding {document.Index.CodePage} ({document.Index.EncodingReason}) == {codePage}");
        TestReport.Check(document.BlockCount == ExpectedBlockCount(count), $"{what}: block count {document.BlockCount:N0} == {ExpectedBlockCount(count):N0}");
        TestReport.Check(document.IsRightToLeft, $"{what}: right-to-left");
        IReadOnlyList<WordDocumentBlockHtml> prefix = document.RenderBlocks(0, PrefixBlockCount);
        string H(int i) => prefix[i].Html;
        bool ok = H(0).Contains("<h1>") && H(0).Contains(Title)
                  && H(1).Contains("a &lt; b x&lt;3 &lt;> 5>4") && H(1).Contains(Line)
                  && H(2).Contains("<b>" + Bold) && H(2).Contains("</b></div>")
                  && H(3).StartsWith("<div class=\"text-line\"><b>" + Bold + " end</b>")
                  && H(4).Contains("title='y>z'") && H(4).Contains(Line + " div</div>")
                  && H(5) == "<div class=\"text-line\"><br></div>"
                  && H(6).Contains(After) && !H(6).Contains(Hidden)
                  && H(7).Contains("<h2>" + Title) && H(8).Contains("2</h2>") && H(8).Contains("<h2>")
                  && H(9).Contains(Line + " s") && !H(9).Contains("alert")
                  && H(10).Contains("<p>" + Paragraph + "</p>") && !H(10).Contains("onclick");
        if (!ok) for (int i = 0; i < prefix.Count; i++) Console.WriteLine($"      {i}: {Ascii(prefix[i].Html)}");
        TestReport.Check(ok, $"{what}: stray '<' escaped, open <b> reopened, tag over two lines one block, empty line, comment dropped, script and onclick removed");
        IReadOnlyList<WordDocumentHeading> headings = document.Headings;
        TestReport.Check(headings.Count == ExpectedHeadingCount(count) && headings[0] == new WordDocumentHeading(0, 1, Title + " 1")
                         && headings[1] == new WordDocumentHeading(7, 2, Title + " 2") && headings[2].BlockIndex == PrefixBlockCount && headings[2].Level == 3,
                         $"{what}: headings ({headings.Count}) incl. one spanning two lines");
        var random = new Random(7);
        bool windowsOk = true;
        for (int k = 0; k < 20 && count > 0; k++)
        {
            int start = random.Next(count);
            foreach (WordDocumentBlockHtml block in document.RenderBlocks(PrefixBlockCount + start, 40))
            {
                int i = block.BlockIndex - PrefixBlockCount;
                string expected = i % 1000 == 0 ? $"<h3>{Title} {i}</h3>" : $"{Line} {i} <i>{Paragraph}</i> {i}";
                if (block.Html != $"<div class=\"text-line\">{expected}</div>") { windowsOk = false; Console.WriteLine($"      block {block.BlockIndex}: {Ascii(block.Html)}"); break; }
            }
        }
        TestReport.Check(windowsOk, $"{what}: 20 random windows of 40 lines exact");
    }

    public static void RunMhtml(int count)
    {
        string directory = TestReport.OutputDirectory;
        TestReport.Step($"Synthetic MHTML (Word web archive), {MhtmlExpectedBlockCount(count):N0} blocks");
        foreach (MimeTransferEncoding transfer in new[] { MimeTransferEncoding.QuotedPrintable, MimeTransferEncoding.Base64 })
        {
            string path = Path.Combine(directory, $"synthetic-mhtml-{count}-{transfer}-v1.mht");
            if (!File.Exists(path)) WriteMhtml(path, count, transfer);
            Console.WriteLine($"   — {transfer} ({new FileInfo(path).Length >> 10} KB)");
            TestReport.Check(WordWindowedDocumentOpener.DetectFormat(path) == WordWindowedDocumentOpener.WordFileFormat.WebArchive, $"{transfer}: web archive detected");
            using (var trickle = new TrickleStream(File.ReadAllBytes(path)))
            using (MarkupWindowedDocument trickled = MarkupWindowedDocument.OpenFromStream(trickle, path, MarkupSourceKind.WebArchive,
                       new WordIndexDecision(path, WordIndexMode.InMemory, false, trickle.Length, 0, "test"), null, 0, 0))
                VerifyMhtml(trickled, count, $"{transfer} fed in 1-7 byte reads");
            OpenBothModes(path, MarkupSourceKind.WebArchive, document => VerifyMhtml(document, count, transfer.ToString()));
        }
    }

    private static void VerifyMhtml(MarkupWindowedDocument document, int count, string what)
    {
        TestReport.Check(document.Index.CodePage == 1255, $"{what}: charset windows-1255 from the part header");
        TestReport.Check(document.BlockCount == MhtmlExpectedBlockCount(count), $"{what}: block count {document.BlockCount:N0} == {MhtmlExpectedBlockCount(count):N0}");
        TestReport.Check(document.IsRightToLeft, $"{what}: right-to-left");
        TestReport.Check(document.Stylesheet.Contains(".word-document p.MsoNormal") && !document.Stylesheet.Contains("@page"), $"{what}: Word styles scoped, @page dropped");
        IReadOnlyList<WordDocumentBlockHtml> prefix = document.RenderBlocks(0, MhtmlPrefixBlockCount);
        bool ok = prefix[0].Html.Contains("<h1 dir=RTL>" + Title + " 1</h1>")
                  && prefix[1].Html.StartsWith("<div class=WordSection1 dir=RTL>") && prefix[1].Html.Contains("<b>" + Bold + "</b> a=b")
                  && prefix[2].Kind == WordDocumentBlockKind.Table && prefix[2].Html.Contains(Line)
                  && prefix[3].Html.Contains(MarkupBlockRenderer.ImagePlaceholder) && !prefix[3].Html.Contains("v:shape") && !prefix[3].Html.Contains("<img");
        if (!ok) foreach (WordDocumentBlockHtml block in prefix) Console.WriteLine($"      {block.BlockIndex}: {Ascii(block.Html)}");
        TestReport.Check(ok, $"{what}: heading, paragraph reopening its section div, table block, image placeholder");
        TestReport.Check(document.Headings.Count == MhtmlExpectedHeadingCount(count) && document.Headings[0].Text == Title + " 1", $"{what}: headings {document.Headings.Count}");
        var random = new Random(3);
        bool windowsOk = true;
        for (int k = 0; k < 20 && count > 0; k++)
        {
            int start = random.Next(count);
            foreach (WordDocumentBlockHtml block in document.RenderBlocks(MhtmlPrefixBlockCount + start, 30))
            {
                int i = block.BlockIndex - MhtmlPrefixBlockCount;
                string expected = i % 500 == 0 ? $"<h2 dir=RTL>{Title} {i}</h2>" : $"<p class=MsoNormal dir=RTL>{MhtmlParagraphText(i)}</p>";
                if (!block.Html.Contains(expected)) { windowsOk = false; Console.WriteLine($"      block {block.BlockIndex}: {Ascii(block.Html)}"); break; }
            }
        }
        TestReport.Check(windowsOk, $"{what}: 20 random windows of 30 blocks exact");
    }

    /// <summary>Cold open / reopen / window timings on one file (sidecar forced so the reopen is measured).</summary>
    public static void Profile(string path)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache"), ForcedMode = WordIndexMode.Sidecar };
        string indexPath = policy.IndexPathFor(path);
        if (File.Exists(indexPath)) File.Delete(indexPath);
        var watch = Stopwatch.StartNew();
        int blocks;
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy)) blocks = document.BlockCount;
        double cold = watch.Elapsed.TotalMilliseconds;
        watch.Restart();
        using IWordWindowedDocument reopened = WordWindowedDocumentOpener.Open(path, policy);
        double reopen = watch.Elapsed.TotalMilliseconds;
        var random = new Random(1);
        double worst = 0, total = 0;
        for (int k = 0; k < 20; k++)
        {
            watch.Restart();
            reopened.RenderBlocks(random.Next(Math.Max(1, blocks)), 60);
            double ms = watch.Elapsed.TotalMilliseconds;
            worst = Math.Max(worst, ms); total += ms;
        }
        Console.WriteLine($"   {Path.GetFileName(path).Length} chars name, {new FileInfo(path).Length >> 10} KB, {blocks:N0} blocks: cold {cold:0} ms, reopen {reopen:0.#} ms, window(60) avg {total / 20:0.##} ms worst {worst:0.##} ms");
    }

    private static string Ascii(string text)
    {
        var sb = new StringBuilder();
        foreach (char c in text.Length > 200 ? text.Substring(0, 200) : text) sb.Append(c < 128 ? c.ToString() : c >= 'א' && c <= 'ת' ? "#" : "?");
        return sb.ToString();
    }
}
