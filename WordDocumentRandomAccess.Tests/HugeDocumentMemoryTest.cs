using System.Diagnostics;
using WordDocumentRandomAccess.Docx;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// The claim behind the whole library, checked on a document far bigger than the policy threshold:
/// reopening from the cached index is instant, random windows anywhere are milliseconds, and the
/// managed heap stays flat while doing so — no body text is held, like the lines view over SQLite.
/// </summary>
public static class HugeDocumentMemoryTest
{
    public static void Run(int blockCount)
    {
        TestReport.Step($"Huge synthetic docx ({blockCount:N0} blocks): instant reload + flat memory");
        string path = Path.Combine(TestReport.OutputDirectory, $"synthetic-{blockCount}-v3.docx");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        if (!File.Exists(path)) TestReport.Timed("generate", () => SyntheticDocxGenerator.Write(path, blockCount));
        Console.WriteLine($"   file {new FileInfo(path).Length >> 20} MB on disk");

        WordIndexDecision decision = WordWindowedDocumentOpener.Decide(path, policy);
        Console.WriteLine($"   policy: {decision.Mode} — {decision.Reason}");
        TestReport.Check(decision.Mode == WordIndexMode.Sidecar, "policy chooses a sidecar index");

        if (!decision.HasCachedIndex)
        {
            long lastReported = -1;
            var progress = new Progress<WordDocumentIndexProgress>(p =>
            {
                long step = p.BytesDone >> 25;
                if (step != lastReported) { lastReported = step; Console.WriteLine($"      indexing {p.BytesDone >> 20} / {p.BytesTotal >> 20} MB, {p.BlocksFound:N0} blocks"); }
            });
            using IWordWindowedDocument first = TestReport.Timed("first open (builds index)", () => WordWindowedDocumentOpener.Open(path, policy, progress));
        }

        long heapBefore = HeapAfterCollect();
        var watch = Stopwatch.StartNew();
        using var document = (DocxWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
        double reloadMilliseconds = watch.Elapsed.TotalMilliseconds;
        Console.WriteLine($"   reload from cached index: {reloadMilliseconds:0.#} ms, document.xml {document.Index.DocumentEntry.UncompressedSize >> 20} MB, {document.Index.Checkpoints.Count} checkpoints");
        TestReport.Check(reloadMilliseconds < 1500, "reload under 1.5 s (index load, no body read)");
        long heapAfterOpen = HeapAfterCollect();
        Console.WriteLine($"   managed heap: {heapBefore >> 20} MB before → {heapAfterOpen >> 20} MB after open (index: ~32 KB/checkpoint + 24 B/block)");

        var random = new Random(11);
        double worst = 0, total = 0;
        const int windows = 200;
        for (int i = 0; i < windows; i++)
        {
            int first = random.Next(Math.Max(1, document.BlockCount - 40));
            var windowWatch = Stopwatch.StartNew();
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(first, 40);
            double ms = windowWatch.Elapsed.TotalMilliseconds;
            worst = Math.Max(worst, ms); total += ms;
            if (blocks.Count == 0) TestReport.Check(false, $"window at {first} rendered nothing");
        }
        long heapAfterWindows = HeapAfterCollect();
        Console.WriteLine($"   {windows} random windows of 40 blocks: average {total / windows:0.#} ms, worst {worst:0.#} ms");
        Console.WriteLine($"   managed heap after windows: {heapAfterWindows >> 20} MB (process working set {Environment.WorkingSet >> 20} MB)");
        TestReport.Check(worst < 300, "worst random window under 300 ms");
        TestReport.Check(heapAfterWindows - heapAfterOpen < 16 << 20, $"heap grew {(heapAfterWindows - heapAfterOpen) >> 20} MB across {windows} random windows (< 16 MB: nothing retained)");
        TestReport.Check(heapAfterOpen - heapBefore < (document.Index.DocumentEntry.UncompressedSize / 8), "open holds far less than the body (< 1/8 of document.xml)");
    }

    private static long HeapAfterCollect()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        return GC.GetTotalMemory(forceFullCollection: true);
    }
}
