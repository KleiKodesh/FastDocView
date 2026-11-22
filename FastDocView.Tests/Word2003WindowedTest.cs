using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Word 2003 XML read windowed (Word2003PackageStream): the body is translated block by block from checkpoints, never
/// whole. A synthetic file made from Word's own fixture (test-data/word-xml/word-xml-v1-2003.xml, its body repeated
/// <see cref="Copies"/> times: sections, notes, links, lists, a picture, a table in every copy) spans many translation
/// checkpoints; every window, read in a shuffled order from the sidecar (on demand, built, reopened), must equal the same
/// window of the in-memory read (one continuous translation), notes and search included. With a path
/// (--word2003-big file.xml [file.docx]) the first paint of a big file is timed against its .docx.
/// </summary>
public static class Word2003WindowedTest
{
    private const int Copies = 600;
    private static string Output => Directory.CreateDirectory(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "TestOutput")).FullName;

    public static void Run(IReadOnlyList<string> bigFiles)
    {
        TestReport.Step("Word 2003 XML windowed: blocks translated from checkpoints equal one continuous translation");
        string fixture = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-xml", "word-xml-v1-2003.xml");
        if (!File.Exists(fixture)) { TestReport.Check(false, "test-data/word-xml/word-xml-v1-2003.xml present"); return; }
        string path = Path.Combine(Output, "word2003-windowed-v1-" + Guid.NewGuid().ToString("N") + ".xml");
        string cache = Path.Combine(Output, "word2003-windowed-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            byte[] source = File.ReadAllBytes(fixture);
            int bodyStart = IndexOf(source, "<w:body>") + "<w:body>".Length, bodyEnd = IndexOf(source, "</w:body>");
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            {
                file.Write(source, 0, bodyStart);
                for (int copy = 0; copy < Copies; copy++) file.Write(source, bodyStart, bodyEnd - bodyStart);
                file.Write(source, bodyEnd, source.Length - bodyEnd);
            }

            var inMemory = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = cache };
            var sidecar = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache };
            var onDemand = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache, IndexesOnDemand = true };

            int scansBefore = FastDocView.Docx.Word2003PackageStream.ScanCount;
            WordWindowedDocumentOpener.Decide(path, onDemand);
            List<string> expected, expectedNotes;
            int expectedHits;
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, onDemand))
            {
                TestReport.Check(FastDocView.Docx.Word2003PackageStream.ScanCount - scansBefore == 1, "Word 2003 XML: Decide + Open make one open pass");
                TestReport.Check(!document.IsIndexComplete && document.KnownBlockCount < Copies * 11, $"Word 2003 XML on demand: the first window is shown before the body is translated ({document.KnownBlockCount} blocks known)");
            }
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, inMemory))
            {
                expected = document.RenderBlocks(0, document.BlockCount).Select(block => Normalise(block.Html)).ToList();
                expectedNotes = Enumerable.Range(0, document.NoteReferences.Count).Select(note => Normalise(string.Concat(document.RenderNotes(note, 1).Select(n => n.Html)))).ToList();
                expectedHits = Hits(document, "Last paragraph");
            }
            TestReport.Check(expected.Count == Copies * 11, $"Word 2003 XML: {Copies} copies of the fixture's 11 blocks ({expected.Count})");
            TestReport.Check(expectedNotes.Count == Copies * 2, $"Word 2003 XML: a footnote and an endnote per copy ({expectedNotes.Count})");
            TestReport.Check(expectedHits == Copies, $"Word 2003 XML: search finds every copy's last paragraph ({expectedHits})");

            if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
            foreach ((string name, WordIndexPolicy policy) in new[] { ("on demand", onDemand), ("sidecar reopen", sidecar), ("sidecar built", sidecar) })
            {
                if (name == "sidecar built" && Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                var random = new Random(2003);
                int mismatches = 0, checkedBlocks = 0;
                for (int round = 0; round < 60; round++)
                {
                    int first = random.Next(expected.Count);
                    if (!document.IsIndexComplete) document.ScanAhead(first + 40 - document.KnownBlockCount);
                    IReadOnlyList<WordDocumentBlockHtml> window = document.RenderBlocks(first, 30);
                    for (int i = 0; i < window.Count; i++, checkedBlocks++)
                        if (Normalise(window[i].Html) != expected[first + i]) mismatches++;
                }
                TestReport.Check(mismatches == 0 && checkedBlocks > 0, $"Word 2003 XML ({name}): {checkedBlocks} blocks in shuffled windows equal the continuous translation ({mismatches} differ)");
                if (!document.IsIndexComplete) document.ScanAhead(int.MaxValue / 2);
                int noteMismatches = 0;
                for (int round = 0; round < 40; round++)
                {
                    int note = random.Next(expectedNotes.Count);
                    if (Normalise(string.Concat(document.RenderNotes(note, 1).Select(n => n.Html))) != expectedNotes[note]) noteMismatches++;
                }
                TestReport.Check(noteMismatches == 0 && document.NoteReferences.Count == expectedNotes.Count, $"Word 2003 XML ({name}): notes read at random equal ({noteMismatches} differ)");
                TestReport.Check(Hits(document, "Last paragraph") == expectedHits, $"Word 2003 XML ({name}): the same search hits");
            }
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, sidecar))
                TestReport.Check(Regex.Matches(WordDocumentText.GetText(document), "Last paragraph").Count == Copies, "Word 2003 XML: GetText over the windowed reader has every copy");
        }
        finally
        {
            try { File.Delete(path); } catch (IOException) { }
            try { if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true); } catch (IOException) { }
        }
        foreach (string big in bigFiles) TimeFirstPaint(big);
    }

    /// <summary>First paint (open on demand + the first 60 blocks) and a far window, cold (no sidecar), printed in ms.</summary>
    private static void TimeFirstPaint(string path)
    {
        if (!File.Exists(path)) { TestReport.Check(false, "big file present: " + Path.GetFileName(path)); return; }
        string cache = Path.Combine(Output, "word2003-big-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            var onDemand = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache, IndexesOnDemand = true };
            for (int run = 0; run < 2; run++)
            {
                if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
                FastDocView.Docx.Word2003PackageStream.ForgetLastScan();
                var clock = Stopwatch.StartNew();
                if (path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    using (var file = WordDocumentFileStream.OpenShared(path))
                    {
                        var layout = FastDocView.Docx.Word2003SourceScanner.Scan(file);
                        Console.WriteLine($"   open pass alone: {clock.ElapsedMilliseconds} ms ({layout.HeaderPictures.Count} header pictures, body from byte {layout.BodyStart})");
                    }
                    clock.Restart();
                }
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, onDemand);
                long opened = clock.ElapsedMilliseconds;
                int shown = document.RenderBlocks(0, 60).Count;
                long painted = clock.ElapsedMilliseconds;
                int known = document.KnownBlockCount;
                clock.Restart();
                document.ScanAhead(int.MaxValue / 2);
                long indexed = clock.ElapsedMilliseconds;
                clock.Restart();
                int far = document.BlockCount * 3 / 4;
                int farShown = document.RenderBlocks(far, 30).Count;
                long farMs = clock.ElapsedMilliseconds;
                Console.WriteLine($"   {Path.GetFileName(path)} run {run}: open {opened} ms, first paint ({shown} blocks) {painted} ms with {known} blocks known; rest of the index {indexed} ms ({document.BlockCount} blocks); window at {far}: {farShown} blocks in {farMs} ms; heap {GC.GetTotalMemory(false) >> 20} MB");
                if (run == 1) TestReport.Check(shown > 0 && farShown > 0, $"{Path.GetFileName(path)}: first paint in {painted} ms, far window in {farMs} ms");
            }
        }
        finally { try { if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true); } catch (IOException) { } }
    }

    private static int Hits(IWordWindowedDocument document, string query)
    {
        int hits = 0, from = 0;
        while (true)
        {
            WordDocumentSearchBatch batch = document.Search(query, WordDocumentSearchOptions.Default, from, 500);
            hits += batch.HitCounts.Sum();
            if (batch.ReachedEnd) return hits;
            from = batch.NextBlock;
        }
    }

    /// <summary>Note ids and picture ids carry offsets / part names; everything else must be the same.</summary>
    private static string Normalise(string html)
        => Regex.Replace(Regex.Replace(html, "data-word-picture=\"[^\"]*\"", "data-word-picture"), "data-(?:word-)?note-id=\"[^\"]*\"", "data-note-id");

    private static int IndexOf(byte[] bytes, string text) => bytes.AsSpan().IndexOf(Encoding.ASCII.GetBytes(text));
}
