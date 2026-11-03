using System.Diagnostics;

namespace FastDocView.Tests;

/// <summary>
/// What a first window, a random jump, a notes page and a full search cost in file opens, bytes read and source checks
/// (stats), from the process-wide counters of <see cref="WordDocumentFileStream"/> - for the formats that read the file
/// itself per window (.rtf, .fodt) or reopen it per read (.odt). Sidecar mode, reopened from the sidecar, so the counts
/// are the reader's, not the index pass's.
/// </summary>
public static class FileReadCountTest
{
    private const int WindowBlocks = 30;
    private const int NotePages = 10;
    private const int PageNotes = 10;
    private const string FrequentHebrew = "שבת";

    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"File read counts: {Path.GetFileName(path)}");
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-read-counts"), ForcedMode = WordIndexMode.Sidecar };
            using (WordWindowedDocumentOpener.Open(path, policy)) { }          // build the sidecar, so the measured opens read it
            Measure("open from the sidecar", () => { using (WordWindowedDocumentOpener.Open(path, policy)) { } });

            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
            {
                Measure($"first window ({WindowBlocks} blocks)", () => document.RenderBlocks(0, WindowBlocks));
                int jump = (int)(document.BlockCount * 0.61);
                Measure($"random jump (block {jump:N0})", () => document.RenderBlocks(jump, WindowBlocks));
                Measure("next window after the jump", () => document.RenderBlocks(jump + WindowBlocks, WindowBlocks));
                int notes = document.NoteReferences.Count;
                if (notes >= NotePages * PageNotes)
                {
                    int start = notes / 2;
                    Measure($"notes: {NotePages} pages of {PageNotes} from the middle", () =>
                    {
                        for (int page = 0; page < NotePages; page++) document.RenderNotes(start + page * PageNotes, PageNotes);
                    });
                }
                else Console.WriteLine($"   {notes} notes: notes pages skipped");
            }

            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
            {
                long hits = 0;
                Measure($"full search \"shabbat\" in batches of 500 ({document.BlockCount:N0} blocks)", () =>
                {
                    for (int from = 0; ;)
                    {
                        WordDocumentSearchBatch batch = document.Search(FrequentHebrew, new WordDocumentSearchOptions(MatchCase: false, IncludeNotes: true), from, 500);
                        hits += batch.HitCounts.Sum();
                        if (batch.ReachedEnd) break;
                        from = batch.NextBlock;
                    }
                });
                Console.WriteLine($"      {hits:N0} matches; file {new FileInfo(path).Length / 1024.0 / 1024.0:F2} MB");
            }
        }
    }

    private static void Measure(string what, Action action)
    {
        long opens = WordDocumentFileStream.ProcessOpenCount, bytes = WordDocumentFileStream.ProcessBytesRead, checks = WordDocumentFileStream.ProcessSourceChecks;
        var watch = Stopwatch.StartNew();
        action();
        watch.Stop();
        Console.WriteLine($"   {what}: {WordDocumentFileStream.ProcessOpenCount - opens} opens, {(WordDocumentFileStream.ProcessBytesRead - bytes) / 1024.0 / 1024.0:F3} MB read, "
            + $"{WordDocumentFileStream.ProcessSourceChecks - checks} source checks, {watch.Elapsed.TotalMilliseconds:F0} ms");
    }
}
