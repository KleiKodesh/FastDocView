using System.Text;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>A blank line right after a Torat Emet heading line is dropped in every book (also PutNewLinesAsIs=1); blank lines elsewhere keep their old handling.</summary>
public static class ToratEmetHeadingBlankTest
{
    public static void Run(string directory)
    {
        TestReport.Step("Torat Emet: no blank block after a heading");
        RunGaps(directory);
        foreach (bool isAsIs in new[] { true, false })
        {
            string head = "&UniqueId=1&PutNewLinesAsIs=" + (isAsIs ? "1" : "0") + "&ForcedBookName=X\r\n";
            string text = head + "$ T\r\n\r\n# H2\r\n \r\n\r\nline a\r\n\r\nline b\r\n@ H3\r\nline c\r\n";
            string path = Path.Combine(directory, "torat-emet-heading-blank-" + (isAsIs ? "asis" : "drop") + "-v1.txt");
            File.WriteAllBytes(path, Encoding.ASCII.GetBytes(text));
            Check(path, isAsIs, "small", WordIndexMode.InMemory, directory);
        }

        // boundaries: the blank lines after the heading fall on a fed-piece boundary and on a span anchor
        foreach (int boundary in new[] { MarkupTextEncoding.DetectionSampleBytes, MarkupTextEncoding.DetectionSampleBytes + MarkupRandomAccessIndexBuilder.ChunkSize })
        foreach (int shift in Enumerable.Range(-17, 20))
        {
            var text = new StringBuilder("&UniqueId=1&PutNewLinesAsIs=1&ForcedBookName=X\r\n$ T\r\n");
            while (text.Length < boundary - 80) text.Append("filler line of ascii text\r\n");
            text.Append(new string('f', boundary + shift - text.Length - 2)).Append("\r\n");
            text.Append("# Split heading\r\n\r\n\r\n \r\nafter\r\n\r\nkept\r\n");
            string path = Path.Combine(directory, $"torat-emet-heading-blank-{boundary}-{shift}-v1.txt");
            File.WriteAllBytes(path, Encoding.ASCII.GetBytes(text.ToString()));
            foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
            {
                var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = mode };
                if (File.Exists(policy.IndexPathFor(path))) File.Delete(policy.IndexPathFor(path));
                for (int pass = 0; pass < (mode == WordIndexMode.Sidecar ? 2 : 1); pass++)
                {
                    using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
                    int heading = document.Headings.First(h => h.Text == "Split heading").BlockIndex;
                    IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(heading, Math.Min(7, document.BlockCount - heading));
                    TestReport.Check(blocks[1].Html.Contains("after") && blocks[2].Html.Contains("<br>") && blocks.Count == 4 && blocks[3].Html.Contains("kept"),
                                     $"boundary {boundary}{shift:+0;-0}, {mode}{(pass == 1 ? " reopened" : "")}: heading, text, blank kept between texts ({string.Join("|", blocks.Select(b => System.Text.RegularExpressions.Regex.Replace(b.Html, "<[^>]*>", "").Trim() is { Length: > 0 } t ? t : "BR"))})");
                }
            }
        }
    }

    private static void RunGaps(string directory)
    {
        RunUtf16(directory); RunLabelThenHeading(directory); RunHeadingLast(directory);
    }

    private static void RunUtf16(string directory)
    {
        string path = Path.Combine(directory, "torat-emet-heading-blank-utf16-v1.txt");
        File.WriteAllBytes(path, new byte[] { 0xFF, 0xFE }.Concat(Encoding.Unicode.GetBytes("&UniqueId=1&PutNewLinesAsIs=1&ForcedBookName=X\r\n$ T\r\n\r\n# H\r\n\r\nafter\r\n")).ToArray());
        using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        TestReport.Check(document.ToratEmet is null, $"UTF-16 .txt stays a plain text file (existing rule: UTF-16 is never Torat Emet); {document.BlockCount} blocks, one per line");
    }

    private static string Plain(string html) => System.Text.RegularExpressions.Regex.Replace(html, "<[^>]*>", "").Trim();

    private static void RunLabelThenHeading(string directory)
    {
        foreach (bool isAsIs in new[] { true, false })
        {
            string path = Path.Combine(directory, "torat-emet-heading-blank-label-" + (isAsIs ? "asis" : "drop") + "-v1.txt");
            File.WriteAllText(path, "&UniqueId=1&PutNewLinesAsIs=" + (isAsIs ? "1" : "0") + "&ForcedBookName=X\r\n$ T\r\n! lab\r\n\r\n# H\r\n\r\n\r\nafter\r\n", Encoding.ASCII);
            using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
            int h = document.Headings.First(x => x.Text == "H").BlockIndex;
            TestReport.Check(blocks.Count == h + 2 && Plain(blocks[h + 1].Html) == "after" && blocks[isAsIs ? h - 2 : h - 1].Html.Contains("lab"), $"label, blank, heading, blanks, text (as-is={isAsIs}): nothing blank after the heading ({string.Join("|", blocks.Select(b => Plain(b.Html)))})");
        }
    }

    private static void RunHeadingLast(string directory)
    {
        foreach (string tail in new[] { "# H\r\n\r\n\r\n", "# H\r\n", "# H", "# H\r\n \r\n" })
        {
            string path = Path.Combine(directory, "torat-emet-heading-blank-last-" + tail.Length + "-v1.txt");
            File.WriteAllText(path, "&UniqueId=1&PutNewLinesAsIs=1&ForcedBookName=X\r\n$ T\r\nbody\r\n" + tail, Encoding.ASCII);
            foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
            {
                var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = mode };
                if (File.Exists(policy.IndexPathFor(path))) File.Delete(policy.IndexPathFor(path));
                using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
                IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
                TestReport.Check(document.Headings.Count == 2 && document.Headings[1].BlockIndex == blocks.Count - 1 && blocks[blocks.Count - 1].Html.Contains("<h2>H</h2>"), $"heading as the last line (tail length {tail.Length}, {mode}): last block is the heading ({blocks.Count} blocks)");
            }
        }
    }

    private static void Check(string path, bool isAsIs, string what, WordIndexMode mode, string directory)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = mode };
        using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
        IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
        int h2 = document.Headings.First(h => h.Text == "H2").BlockIndex;
        int h3 = document.Headings.First(h => h.Text == "H3").BlockIndex;
        int title = document.Headings.First(h => h.Text == "T").BlockIndex;
        bool afterHeadings = blocks[title + 1].Html.Contains("te-heading") && blocks[h2 + 1].Html.Contains("line a") && blocks[h3 + 1].Html.Contains("line c");
        // as-is keeps the blank between "line a" and "line b"; the other mode drops it
        bool betweenTexts = isAsIs ? blocks[h2 + 2].Html.Contains("<br>") && blocks[h2 + 3].Html.Contains("line b") : blocks[h2 + 2].Html.Contains("line b");
        TestReport.Check(afterHeadings && betweenTexts, $"{what}, as-is={isAsIs}: no blank block after a heading, other blanks unchanged ({blocks.Count} blocks)");
    }
}
