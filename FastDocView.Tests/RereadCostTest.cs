using SeekableDeflate;

namespace FastDocView.Tests;

/// <summary>
/// What re-showing windows costs in inflation (sidecar mode, deflated parts): checkpoint resumes and bytes inflated,
/// from the process-wide counters of <see cref="DeflateRandomAccessStream"/>. Body: 20 windows of 40 blocks read
/// forward from the middle, then paged back and shown forward again. Notes: a jump to the middle note window, then
/// 20 windows of 10 notes paged back. Each block / note is checked to render the same HTML the first time and again.
/// </summary>
public static class RereadCostTest
{
    private const int Windows = 20;
    private const int WindowBlocks = 40;
    private const int WindowNotes = 10;

    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"Re-read cost: {Path.GetFileName(path)}");
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-reread"), ForcedMode = WordIndexMode.Sidecar };
            using (WordWindowedDocumentOpener.Open(path, policy)) { }          // build the sidecar, so the measured open reads it
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);

            int start = Math.Max(0, document.BlockCount / 2 - Windows * WindowBlocks / 2);
            var firstHtml = new Dictionary<int, string>();
            Measure("body: 20 windows forward (first read)", () =>
            {
                for (int w = 0; w < Windows; w++)
                    foreach (WordDocumentBlockHtml block in document.RenderBlocks(start + w * WindowBlocks, WindowBlocks)) firstHtml[block.BlockIndex] = block.Html;
            });
            bool same = true;
            Measure("body: the same 20 windows paged back", () =>
            {
                for (int w = Windows - 1; w >= 0; w--)
                    foreach (WordDocumentBlockHtml block in document.RenderBlocks(start + w * WindowBlocks, WindowBlocks)) same &= firstHtml[block.BlockIndex] == block.Html;
            });
            Measure("body: the same 20 windows forward again", () =>
            {
                for (int w = 0; w < Windows; w++)
                    foreach (WordDocumentBlockHtml block in document.RenderBlocks(start + w * WindowBlocks, WindowBlocks)) same &= firstHtml[block.BlockIndex] == block.Html;
            });
            TestReport.Check(same, "body: re-shown windows render the same HTML");

            int noteCount = document.NoteReferences.Count;
            if (noteCount < WindowNotes * 2) { Console.WriteLine($"   {noteCount} notes: notes skipped"); continue; }
            int noteStart = Math.Max(0, Math.Min(noteCount - WindowNotes, noteCount / 2));
            var firstNotes = new Dictionary<(WordDocumentNoteKind, int), string>();
            Measure("notes: jump to the middle window", () =>
            {
                foreach (WordDocumentNoteHtml note in document.RenderNotes(noteStart, WindowNotes)) firstNotes[(note.Kind, note.Number)] = note.Html;
            });
            Measure("notes: 20 windows paged back", () =>
            {
                for (int w = 1; w <= Windows && noteStart - w * WindowNotes >= 0; w++)
                    foreach (WordDocumentNoteHtml note in document.RenderNotes(noteStart - w * WindowNotes, WindowNotes)) firstNotes[(note.Kind, note.Number)] = note.Html;
            });
            bool sameNotes = true;
            Measure("notes: the jumped-to window again", () =>
            {
                foreach (WordDocumentNoteHtml note in document.RenderNotes(noteStart, WindowNotes)) sameNotes &= firstNotes[(note.Kind, note.Number)] == note.Html;
            });
            TestReport.Check(sameNotes, "notes: the re-shown window renders the same HTML");
        }
    }

    private static void Measure(string what, Action action)
    {
        long resumes = DeflateRandomAccessStream.ProcessCheckpointResumeCount, inflated = DeflateRandomAccessStream.ProcessBytesInflated;
        var watch = System.Diagnostics.Stopwatch.StartNew();
        action();
        watch.Stop();
        Console.WriteLine($"   {what}: {DeflateRandomAccessStream.ProcessCheckpointResumeCount - resumes} checkpoint resumes, "
            + $"{(DeflateRandomAccessStream.ProcessBytesInflated - inflated) / 1024.0 / 1024.0:F2} MB inflated, {watch.Elapsed.TotalMilliseconds:F0} ms");
    }
}
