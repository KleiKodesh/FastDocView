using System.Text;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Opens real Word files (.docx and .doc) through the format-detecting opener, times the index and
/// a few windows, checks the rendered HTML is non-empty and well-formed-ish, and writes a sample
/// page per file to TestOutput for eyeballing.
/// </summary>
public static class RealDocumentSmokeTest
{
    public static void Run(IEnumerable<string> paths)
    {
        string cacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache");
        foreach (string path in paths)
        {
            TestReport.Step($"Real file: {Path.GetFileName(path)} ({new FileInfo(path).Length >> 10} KB)");
            try
            {
                string indexPath = WordDocumentIndexCacheLocation.IndexPathFor(path, cacheDirectory);
                if (File.Exists(indexPath)) File.Delete(indexPath);
                using IWordWindowedDocument document = TestReport.Timed("open (cold index)", () => WordWindowedDocumentOpener.Open(path, cacheDirectory));
                using IWordWindowedDocument reopened = TestReport.Timed("re-open (cached index)", () => WordWindowedDocumentOpener.Open(path, cacheDirectory));
                Console.WriteLine($"   {document.GetType().Name}: {document.BlockCount:N0} blocks, rtl={document.IsRightToLeft}");
                TestReport.Check(document.BlockCount > 0, "has blocks");

                IReadOnlyList<WordDocumentBlockHtml> first = TestReport.Timed("render first 30 blocks", () => document.RenderBlocks(0, 30));
                int middle = document.BlockCount / 2;
                IReadOnlyList<WordDocumentBlockHtml> mid = TestReport.Timed($"render 30 blocks at {middle}", () => document.RenderBlocks(middle, 30));
                IReadOnlyList<WordDocumentBlockHtml> last = TestReport.Timed("render last 30 blocks", () => document.RenderBlocks(Math.Max(0, document.BlockCount - 30), 30));
                int nonEmpty = first.Concat(mid).Concat(last).Count(b => StripTags(b.Html).Trim().Length > 0);
                TestReport.Check(nonEmpty > 0, $"{nonEmpty} of {first.Count + mid.Count + last.Count} sampled blocks have visible text");
                IReadOnlyList<WordDocumentHeading> headings = TestReport.Timed("headings", () => document.Headings);
                Console.WriteLine($"   {headings.Count} headings" + (headings.Count > 0 ? $", first: {Truncate(headings[0].Text)}" : ""));
                Console.WriteLine($"   first text: {Truncate(StripTags(string.Concat(first.Select(b => b.Html))))}");

                var html = new StringBuilder();
                html.Append("<!doctype html><meta charset=utf-8><style>").Append(document.Stylesheet).Append("</style><div class=word-document dir=")
                    .Append(document.IsRightToLeft ? "rtl" : "ltr").Append('>');
                foreach (WordDocumentBlockHtml block in first.Concat(mid).Concat(last)) html.Append(block.Html);
                html.Append("</div>");
                string sanitizedName = Convert.ToHexString(System.Security.Cryptography.MD5.HashData(Encoding.UTF8.GetBytes(path)))[..8];
                File.WriteAllText(Path.Combine(TestReport.OutputDirectory, $"real-{sanitizedName}{Path.GetExtension(path)}.html"), html.ToString());
            }
            catch (Exception exception)
            {
                TestReport.Check(false, $"opened without exception — {exception.GetType().Name}: {exception.Message}");
            }
        }
    }

    private static string StripTags(string html) => System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", "");
    private static string Truncate(string text) => text.Length > 120 ? text.Substring(0, 120) + "…" : text;
}
