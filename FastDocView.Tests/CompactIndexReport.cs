using FastDocView.DocBinary;
using FastDocView.Odt;
using FastDocView.Rtf;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>
/// The numbers of a compact sidecar for any file type: sidecar size with its per-section breakdown, reopen from the
/// sidecar, first window, random single blocks (median / p90) - what NOTES.md's compact-sidecar table reports.
/// </summary>
public static class CompactIndexReport
{
    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"Compact sidecar report: {Path.GetFileName(path)}");
            string cacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-report");
            var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = WordIndexMode.Sidecar };
            string indexPath = policy.IndexPathFor(path);
            if (File.Exists(indexPath)) File.Delete(indexPath);
            using (IWordWindowedDocument first = TestReport.Timed("cold open (index pass)", () => WordWindowedDocumentOpener.Open(path, policy)))
            {
                if (first is OdtWindowedDocument odt) odt.CompleteIndex();
                Console.WriteLine($"   {first.BlockCount:N0} blocks, sidecar {new FileInfo(indexPath).Length:N0} bytes");
                foreach ((string section, long bytes) in Breakdown(first)) Console.WriteLine($"      {bytes,9:N0}  {section}");
            }
            using IWordWindowedDocument document = TestReport.Timed("reopen from sidecar", () => WordWindowedDocumentOpener.Open(path, policy));
            TestReport.Check(document.IsIndexComplete && document.KnownBlockCount == document.BlockCount, "reopened complete from the sidecar");
            TestReport.Timed("first window (40 blocks)", () => document.RenderBlocks(0, 40));
            var random = new Random(17);
            var times = new List<double>();
            for (int i = 0; i < 200; i++)
            {
                int block = random.Next(document.BlockCount);
                var watch = System.Diagnostics.Stopwatch.StartNew();
                document.RenderBlocks(block, 1);
                times.Add(watch.Elapsed.TotalMilliseconds);
            }
            times.Sort();
            Console.WriteLine($"   200 random single blocks (render): median {times[times.Count / 2]:F1} ms, p90 {times[times.Count * 9 / 10]:F1} ms");
        }
    }

    private static IReadOnlyList<(string section, long bytes)> Breakdown(IWordWindowedDocument document) => document switch
    {
        DocBinaryWindowedDocument doc => doc.BlockIndex.SizeBreakdown(),
        OdtWindowedDocument odt => odt.Index.SizeBreakdown(),
        Rtf.RtfWindowedDocument rtf => rtf.Index.SizeBreakdown(),
        Docx.DocxWindowedDocument docx => docx.Index.SizeBreakdown(),
        _ => [],
    };
}
