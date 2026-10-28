using System.Diagnostics;
using WordDocumentRandomAccess.Docx;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Notes get the same streaming model as the body: a .docx whose footnotes part is far larger than
/// its body opens without loading the part (the policy weighs body + notes and picks a sidecar),
/// every note is read on demand from its own byte range, and the managed heap stays flat while
/// windows full of notes are rendered all over the document.
/// </summary>
public static class DocxNotesStreamingTest
{
    public static void Run(int blockCount, int noteFillerCharacters)
    {
        TestReport.Step($"Notes streaming: {blockCount:N0} blocks, {noteFillerCharacters:N0} filler characters per note");
        string path = Path.Combine(TestReport.OutputDirectory, $"synthetic-notes-{blockCount}-{noteFillerCharacters}-v1.docx");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        if (!File.Exists(path)) TestReport.Timed("generate", () => SyntheticDocxGenerator.Write(path, blockCount, noteFillerCharacters));
        string indexPath = policy.IndexPathFor(path);
        if (File.Exists(indexPath)) File.Delete(indexPath);

        long footnotesBytes, bodyBytes;
        using (var zip = System.IO.Compression.ZipFile.OpenRead(path))
        {
            footnotesBytes = zip.GetEntry("word/footnotes.xml")!.Length;
            bodyBytes = zip.GetEntry("word/document.xml")!.Length;
        }
        Console.WriteLine($"   file {new FileInfo(path).Length >> 20} MB; document.xml {bodyBytes >> 20} MB, footnotes.xml {footnotesBytes >> 20} MB uncompressed");
        TestReport.Check(footnotesBytes > 100L << 20, "footnotes part is over 100 MB");

        WordIndexDecision decision = WordWindowedDocumentOpener.Decide(path, policy);
        Console.WriteLine($"   policy: {decision.Mode} — {decision.Reason}");
        TestReport.Check(decision.Mode == WordIndexMode.Sidecar, "policy weighs the notes parts too (sidecar, although the body is small)");

        long lastReported = -1;
        var progress = new Progress<WordDocumentIndexProgress>(p =>
        {
            long step = p.BytesDone >> 25;
            if (step != lastReported) { lastReported = step; Console.WriteLine($"      indexing {p.BytesDone >> 20} / {p.BytesTotal >> 20} MB"); }
        });
        using (IWordWindowedDocument first = TestReport.Timed("first open (indexes body + notes)", () => WordWindowedDocumentOpener.Open(path, policy, progress))) { }

        long heapBefore = HeapAfterCollect();
        var watch = Stopwatch.StartNew();
        using var document = (DocxWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
        Console.WriteLine($"   reopen: {watch.Elapsed.TotalMilliseconds:0} ms; footnotes part: {document.Index.FootnotesPart?.NoteIds.Length} notes, {document.Index.FootnotesPart?.Checkpoints.Count} checkpoints");
        long heapAfterOpen = HeapAfterCollect();
        Console.WriteLine($"   heap {heapBefore >> 20} MB → {heapAfterOpen >> 20} MB after open");
        TestReport.Check(heapAfterOpen - heapBefore < footnotesBytes / 8, $"open holds far less than the footnotes part (+{(heapAfterOpen - heapBefore) >> 20} MB vs {footnotesBytes >> 20} MB)");

        var random = new Random(5);
        double worst = 0;
        int notesSeen = 0;
        bool allMatch = true;
        for (int i = 0; i < 60; i++)
        {
            int start = random.Next(Math.Max(1, document.BlockCount - 40));
            var windowWatch = Stopwatch.StartNew();
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(start, 40);
            worst = Math.Max(worst, windowWatch.Elapsed.TotalMilliseconds);
            foreach (WordDocumentBlockHtml block in blocks)
            {
                foreach (WordDocumentNoteHtml note in block.Notes)
                {
                    notesSeen++;
                    if (!note.Html.Contains($"של פסקה {block.BlockIndex}")) { allMatch = false; Console.WriteLine($"   note in block {block.BlockIndex} has the wrong body"); }
                }
            }
        }
        long heapAfterWindows = HeapAfterCollect();
        Console.WriteLine($"   60 random windows: {notesSeen} notes read on demand, worst window {worst:0} ms; heap {heapAfterWindows >> 20} MB (working set {Environment.WorkingSet >> 20} MB)");
        TestReport.Check(notesSeen > 50, "windows delivered notes");
        TestReport.Check(allMatch, "every note body is the note of its own reference");
        TestReport.Check(heapAfterWindows - heapAfterOpen < 32L << 20, $"heap stays flat while reading notes (+{(heapAfterWindows - heapAfterOpen) >> 20} MB, cache bounded)");
    }

    private static long HeapAfterCollect()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        return GC.GetTotalMemory(forceFullCollection: true);
    }
}
