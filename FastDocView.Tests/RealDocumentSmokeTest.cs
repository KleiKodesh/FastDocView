using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Opens real Word files (.docx and .doc) through the format-detecting opener, times the index and
/// a few windows, checks the rendered HTML is non-empty and well-formed-ish, and writes a sample
/// page per file to TestOutput for eyeballing.
/// </summary>
public static class RealDocumentSmokeTest
{
    public static void Run(IEnumerable<string> paths)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        foreach (string path in paths)
        {
            TestReport.Step($"Real file: {Path.GetFileName(path)} ({new FileInfo(path).Length >> 10} KB)");
            try
            {
                string indexPath = policy.IndexPathFor(path);
                if (File.Exists(indexPath)) File.Delete(indexPath);
                WordIndexDecision decision = WordWindowedDocumentOpener.Decide(path, policy);
                Console.WriteLine($"   policy: {decision.Mode} — {decision.Reason}");
                using IWordWindowedDocument document = TestReport.Timed("open (cold)", () => WordWindowedDocumentOpener.Open(path, policy));
                using IWordWindowedDocument reopened = TestReport.Timed("re-open", () => WordWindowedDocumentOpener.Open(path, policy));
                Console.WriteLine($"   {document.GetType().Name}: {document.BlockCount:N0} blocks, rtl={document.IsRightToLeft}");
                TestReport.Check(document.BlockCount > 0, "has blocks");

                IReadOnlyList<WordDocumentBlockHtml> first = TestReport.Timed("render first 30 blocks", () => document.RenderBlocks(0, 30));
                int middle = document.BlockCount / 2;
                IReadOnlyList<WordDocumentBlockHtml> mid = TestReport.Timed($"render 30 blocks at {middle}", () => document.RenderBlocks(middle, 30));
                IReadOnlyList<WordDocumentBlockHtml> last = TestReport.Timed("render last 30 blocks", () => document.RenderBlocks(Math.Max(0, document.BlockCount - 30), 30));
                Console.WriteLine("   window kinds: " + string.Join(",", mid.GroupBy(b => b.Kind).Select(g => g.Key + "=" + g.Count())) + "; list markers: " + mid.Count(b => b.Html.Contains("list-marker")) + "; underline: " + mid.Count(b => b.Html.Contains("underline")));
                foreach (int other in new[] { middle + 2000, middle + 9000, middle / 2 })
                    if (other + 30 < document.BlockCount) TestReport.Timed($"render 30 blocks at {other} (first time)", () => document.RenderBlocks(other, 30));
                // warm (JIT done, caches filled): the steady-state cost of a window, best of 5
                double warm = Enumerable.Range(0, 5).Min(_ => { var watch = System.Diagnostics.Stopwatch.StartNew(); document.RenderBlocks(middle, 30); return watch.Elapsed.TotalMilliseconds; });
                Console.WriteLine($"   render 30 blocks at {middle}, warm (best of 5): {warm:0.#} ms");
                int nonEmpty = first.Concat(mid).Concat(last).Count(b => StripTags(b.Html).Trim().Length > 0);
                TestReport.Check(nonEmpty > 0, $"{nonEmpty} of {first.Count + mid.Count + last.Count} sampled blocks have visible text");
                IReadOnlyList<WordDocumentHeading> headings = TestReport.Timed("headings", () => document.Headings);
                Console.WriteLine($"   {headings.Count} headings" + (headings.Count > 0 ? $", first: {Truncate(headings[0].Text)}" : ""));
                (int footnotes, int endnotes) = document.NoteCounts;
                int notesInSample = first.Concat(mid).Concat(last).Sum(b => b.Notes.Count);
                int referencesInSample = first.Concat(mid).Concat(last).Sum(b => CountOccurrences(b.Html, "class=\"note-reference\""));
                Console.WriteLine($"   notes: {footnotes} footnotes, {endnotes} endnotes; sampled blocks carry {referencesInSample} references and {notesInSample} note bodies");
                TestReport.Check(referencesInSample == notesInSample, "every sampled note reference comes with its note body");
                TestReport.Check(first.Concat(mid).Concat(last).SelectMany(b => b.Notes).All(n => n.Html.Length > 0), "no empty note bodies");
                if (document.NoteReferences.Count > 0)
                {
                    WordDocumentNoteReference firstReference = document.NoteReferences[0];
                    WordDocumentBlockHtml noted = TestReport.Timed("render the first noted block", () => document.RenderBlocks(firstReference.BlockIndex, 1)[0]);
                    TestReport.Check(noted.Notes.Any(n => n.Kind == firstReference.Kind && n.Number == firstReference.Number && n.Html.Length > 0),
                                     $"the block of the first {firstReference.Kind} reference (block {firstReference.BlockIndex}) delivers its body");
                }
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

    private static int CountOccurrences(string text, string value)
    {
        int count = 0;
        for (int at = text.IndexOf(value, StringComparison.Ordinal); at >= 0; at = text.IndexOf(value, at + value.Length, StringComparison.Ordinal)) count++;
        return count;
    }

    private static string StripTags(string html) => System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", "");
    private static string Truncate(string text) => text.Length > 120 ? text.Substring(0, 120) + "…" : text;
}
