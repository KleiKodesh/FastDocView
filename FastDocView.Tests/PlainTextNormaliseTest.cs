using System.Text;
using System.Text.RegularExpressions;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>Plain .txt: blank and whitespace-only lines are no blocks, headings keep feeding the contents, heading CSS is scoped to plain text.</summary>
public static class PlainTextNormaliseTest
{
    private static string Plain(string html) => Regex.Replace(html, "<[^>]*>", "").Trim();

    public static void Run()
    {
        TestReport.Step("Plain .txt: blank lines dropped, headings bold by CSS");
        string directory = TestReport.OutputDirectory;
        string text = "\r\n  \r\n<h1 class=\"x\">Title</h1>\r\n\r\n\r\nline a  \r\n \t \r\n\r\nline b\r\n<H2>Sub\r\ntwo</H2>\r\n\r\nend\r\n\r\n\r\n";
        Verify(directory, "plain-blank-v1", Encoding.UTF8.GetBytes(text), new[] { "Title", "line a", "line b", "Sub", "two", "end" }, 2);
        Verify(directory, "plain-blank-utf16-v1", new byte[] { 0xFF, 0xFE }.Concat(Encoding.Unicode.GetBytes(text)).ToArray(), new[] { "Title", "line a", "line b", "Sub", "two", "end" }, 2);
        Verify(directory, "plain-title-only-v1", Encoding.UTF8.GetBytes("<h1>Only</h1>\r\n\r\n"), new[] { "Only" }, 1);
        Verify(directory, "plain-allblank-v1", Encoding.UTF8.GetBytes("\r\n  \r\n"), new string[0], 0);
        Verify(directory, "plain-same-v1", Encoding.UTF8.GetBytes("a\r\nb\r\n c\r\n"), new[] { "a", "b", "c" }, 0);

        // blank runs on fed-piece boundaries and span anchors
        foreach (int boundary in new[] { MarkupTextEncoding.DetectionSampleBytes, MarkupTextEncoding.DetectionSampleBytes + MarkupRandomAccessIndexBuilder.ChunkSize, 3 * MarkupRandomAccessIndex.AnchorSpanBytes })
        foreach (int shift in Enumerable.Range(-9, 14))
        {
            var sb = new StringBuilder("<h1>T</h1>\r\n");
            while (sb.Length < boundary - 80) sb.Append("filler line of ascii text\r\n");
            sb.Append(new string('f', boundary + shift - sb.Length - 2)).Append("\r\n");
            sb.Append("\r\n \r\n\r\n\t\r\n<h3>Split</h3>\r\n\r\n\r\nafter\r\n\r\n\r\nkept\r\n");
            string path = Path.Combine(directory, $"plain-blank-{boundary}-{shift}-v1.txt");
            File.WriteAllBytes(path, Encoding.ASCII.GetBytes(sb.ToString()));
            foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
            {
                var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = mode };
                if (File.Exists(policy.IndexPathFor(path))) File.Delete(policy.IndexPathFor(path));
                for (int pass = 0; pass < (mode == WordIndexMode.Sidecar ? 2 : 1); pass++)
                {
                    using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
                    int heading = document.Headings.First(h => h.Text == "Split").BlockIndex;
                    IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(heading, document.BlockCount - heading);
                    TestReport.Check(blocks.Count == 3 && Plain(blocks[0].Html) == "Split" && Plain(blocks[1].Html) == "after" && Plain(blocks[2].Html) == "kept"
                                     && document.Headings.Count == 2 && document.Headings[0].Level == 1 && document.Headings[1].Level == 3,
                                     $"boundary {boundary}{shift:+0;-0}, {mode}{(pass == 1 ? " reopened" : "")}: heading, after, kept, no empty block");
                }
            }
        }
        // the stylesheet: plain text gets the heading rules, Torat Emet keeps its own
        using (var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(Path.Combine(directory, "plain-same-v1.txt"), new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory }))
            TestReport.Check(document.Stylesheet.Contains("font-weight: 700") && document.Stylesheet.Contains("::after { content: none") && !document.Stylesheet.Contains("te-heading"), "plain text stylesheet: bold headings, no gap, not Torat Emet's");
    }

    private static void Verify(string directory, string name, byte[] bytes, string[] expected, int headings)
    {
        string path = Path.Combine(directory, name + ".txt");
        File.WriteAllBytes(path, bytes);
        foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
        {
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = mode };
            if (File.Exists(policy.IndexPathFor(path))) File.Delete(policy.IndexPathFor(path));
            for (int pass = 0; pass < (mode == WordIndexMode.Sidecar ? 2 : 1); pass++)
            {
                using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
                IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
                string got = string.Join("|", blocks.Select(b => Regex.Replace(b.Html, "<[^>]*>", "").Trim().Replace("\r\n", "\n")));
                string want = string.Join("|", expected.Select(e => e));
                TestReport.Check(blocks.Count == expected.Length && got == want && document.Headings.Count == headings && !blocks.Any(b => b.Html.Contains("<br>") && Plain(b.Html).Length == 0),
                                 $"{name}, {mode}{(pass == 1 ? " reopened" : "")}: {blocks.Count} blocks, {document.Headings.Count} headings ({got})");
            }
        }
    }
}
