using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Notes-only render speed (notes pane pages, the note tooltip): many-notes synthetic .docx / .odt / .fodt / .rtf plus the
/// files given. Per file: the ms for every note in pages of 20 (a fresh open each run, best of a few), the ms for single
/// notes (the tooltip), and a SHA-256 of every note's kind, number, label and HTML - equal before / after a change proves
/// byte-identical output. Prints the extension only (sample names are not ASCII).
/// </summary>
public static class NotesRenderSpeedTest
{
    public static void Run(IEnumerable<string> realPaths, int runs = 3)
    {
        string directory = Path.Combine(TestReport.OutputDirectory, "notes-speed");
        Directory.CreateDirectory(directory);
        string docx = Path.Combine(directory, "many-notes.docx"), odt = Path.Combine(directory, "many-notes.odt");
        string fodt = Path.Combine(directory, "many-notes.fodt"), rtf = Path.Combine(directory, "many-notes.rtf");
        if (!File.Exists(docx)) SyntheticDocxGenerator.Write(docx, 40_000);
        if (!File.Exists(odt)) SyntheticOdtGenerator.WritePackage(odt, 40_000);
        if (!File.Exists(fodt)) SyntheticOdtGenerator.WriteFlat(fodt, 40_000);
        if (!File.Exists(rtf)) SyntheticRtfGenerator.Write(rtf, 40_000);
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache") };
        foreach (string path in new[] { docx, odt, fodt, rtf }.Concat(realPaths))
        {
            string name = Path.GetExtension(path) + (path.StartsWith(directory, StringComparison.Ordinal) ? " synthetic" : " sample");
            TestReport.Step($"Notes-only render speed: {name}");
            long bestPages = long.MaxValue, bestSingles = long.MaxValue, bodyPages = long.MaxValue, bodySingles = long.MaxValue, noteBlocksAsBody = long.MaxValue;
            double coldNote = double.MaxValue, coldBody = double.MaxValue;
            string? hash = null;
            int noteCount = 0;
            for (int run = 0; run < runs; run++)
            {
                using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
                {
                    while (!document.IsIndexComplete) document.ScanAhead(100_000);
                    noteCount = document.NoteReferences.Count;
                    using var sha = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
                    var watch = Stopwatch.StartNew();
                    for (int first = 0; first < noteCount; first += 20)
                        foreach (WordDocumentNoteHtml note in document.RenderNotes(first, 20))
                            sha.AppendData(Encoding.UTF8.GetBytes($"{note.Kind}|{note.Number}|{note.Label}|{note.Html}\n"));
                    bestPages = Math.Min(bestPages, watch.ElapsedMilliseconds);
                    string runHash = Convert.ToHexString(sha.GetHashAndReset())[..16];
                    TestReport.Check(hash is null || hash == runHash, "the same notes every run");
                    hash = runHash;
                }
                using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
                {
                    while (!document.IsIndexComplete) document.ScanAhead(100_000);
                    var random = new Random(7);
                    var watch = Stopwatch.StartNew();
                    for (int i = 0; i < 200 && noteCount > 0; i++) document.RenderNotes(random.Next(noteCount), 1);
                    bestSingles = Math.Min(bestSingles, watch.ElapsedMilliseconds);
                }
                using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
                {
                    // the body side of the gate: the same counts of blocks, in pages of 20 and single random blocks
                    while (!document.IsIndexComplete) document.ScanAhead(100_000);
                    // a body window of 20 blocks where each notes page starts (the same places in the file, the same window count)
                    IReadOnlyList<WordDocumentNoteReference> references = document.NoteReferences;
                    var watch = Stopwatch.StartNew();
                    for (int first = 0; first < references.Count; first += 20) document.RenderBlocks(references[first].BlockIndex, 20);
                    bodyPages = Math.Min(bodyPages, watch.ElapsedMilliseconds);
                    // diagnosis: the notes pages' own blocks through the body path (one call per run of adjacent blocks)
                    watch.Restart();
                    for (int first = 0; first < references.Count; first += 20)
                        for (int i = first; i < Math.Min(references.Count, first + 20);)
                        {
                            int runStart = references[i].BlockIndex, runEnd = runStart;
                            while (i < Math.Min(references.Count, first + 20) && references[i].BlockIndex <= runEnd + 1) runEnd = references[i++].BlockIndex;
                            document.RenderBlocks(runStart, runEnd - runStart + 1);
                        }
                    noteBlocksAsBody = Math.Min(noteBlocksAsBody, watch.ElapsedMilliseconds);
                    var random = new Random(7);
                    watch.Restart();
                    for (int i = 0; i < 200; i++) document.RenderBlocks(random.Next(document.BlockCount), 1);
                    bodySingles = Math.Min(bodySingles, watch.ElapsedMilliseconds);
                }
                // cold: the first note / the first body block right after an open (index from the sidecar)
                using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
                {
                    while (!document.IsIndexComplete) document.ScanAhead(100_000);
                    var watch = Stopwatch.StartNew();
                    document.RenderNotes(document.NoteReferences.Count / 2, 20);
                    coldNote = Math.Min(coldNote, watch.Elapsed.TotalMilliseconds);
                }
                using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
                {
                    while (!document.IsIndexComplete) document.ScanAhead(100_000);
                    var watch = Stopwatch.StartNew();
                    document.RenderBlocks(document.BlockCount / 2, 20);
                    coldBody = Math.Min(coldBody, watch.Elapsed.TotalMilliseconds);
                }
            }
            Console.WriteLine($"  NOTES-SPEED {name}: {noteCount} notes, pages of 20 {bestPages} ms, 200 single notes {bestSingles} ms, hash {hash}");
            Console.WriteLine($"  BODY-SPEED  {name}: 20 blocks at each notes page, {bodyPages} ms, 200 single blocks {bodySingles} ms, the note blocks through the body path {noteBlocksAsBody} ms; cold 20 notes {coldNote:F1} ms vs cold 20 blocks {coldBody:F1} ms");
        }
    }
}
