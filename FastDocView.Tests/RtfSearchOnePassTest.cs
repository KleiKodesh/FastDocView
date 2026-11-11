using System.Diagnostics;
using FastDocView.Rtf;

namespace FastDocView.Tests;

/// <summary>
/// RTF: (1) the resumable span rescan, stepped one block at a time, gives exactly the values of the span parsed in one go;
/// (2) the one-pass search (block bounds and matcher text from the same parse) equals the old path it replaced (each
/// block's range re-parsed from its snapshot, <see cref="RtfWindowedDocument.SearchByBlockReparse"/>): same hit blocks,
/// counts, next block and end for Hebrew plain and pointed, English, mixed, match case on and off,
/// sequential batches of 500 / 37 / 1 and batches that start anywhere. Hebrew is written as \u escapes.
/// </summary>
public static class RtfSearchOnePassTest
{
    private static readonly string[] Queries =
    [
        "שבת",                                   // Hebrew, plain letters
        "שַׁבָּת",           // the same, pointed
        "אבג",                                   // the synthetic file's Hebrew
        "the",                                                  // English
        "The",
        "ו a",                                             // mixed
        " ",
        "qqzzqq",                                               // absent: a full scan
    ];

    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"RTF one pass: {Path.GetFileName(path)}");
            using var document = RtfWindowedDocument.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
            CompareSpans(path, document);
            CompareSearch(path, document);
        }
    }

    private static void CompareSpans(string path, RtfWindowedDocument document)
    {
        RtfRandomAccessIndex index = document.Index;
        WordIndexSpanAnchor[] anchors = index.Anchors;
        var random = new Random(20261103);
        var picked = Enumerable.Range(0, Math.Min(anchors.Length, 6)).Concat(Enumerable.Range(0, 10).Select(_ => random.Next(anchors.Length)))
            .Append(anchors.Length - 1).Distinct().ToList();
        int differences = 0, blocks = 0;
        using FileStream file = WordDocumentFileStream.OpenShared(path);
        foreach (int a in picked)
        {
            long end = a + 1 < anchors.Length ? anchors[a + 1].Position : index.LastBlockEndOffset;
            bool isLast = a + 1 >= anchors.Length;
            RtfRandomAccessIndexBuilder whole = RtfRandomAccessIndexBuilder.ScanSpan(file, index.Tables, anchors[a].Position, index.ScanStateOf(a), end, isLast);
            RtfRandomAccessIndexBuilder stepped = RtfRandomAccessIndexBuilder.ResumableSpan(file, index.Tables, anchors[a].Position, index.ScanStateOf(a), end, isLast);
            for (int local = 0; !stepped.IsComplete; local++)
            {
                stepped.ScanThrough(local);
                file.Position = 0;                   // another reader moved the stream between requests
            }
            bool same = whole.Starts.SequenceEqual(stepped.Starts) && whole.Ends.SequenceEqual(stepped.Ends) && whole.Kinds.SequenceEqual(stepped.Kinds)
                        && whole.TextLengths.SequenceEqual(stepped.TextLengths) && whole.ListLevels.SequenceEqual(stepped.ListLevels);
            for (int i = 0; same && i < whole.Starts.Count; i++)
                same = whole.SnapshotOf(i).AsSpan().SequenceEqual(stepped.SnapshotOf(i)) && whole.LabelOf(i) == stepped.LabelOf(i);
            if (!same) { differences++; Console.WriteLine($"      span {a} differs"); }
            blocks += whole.Starts.Count;
        }
        TestReport.Check(differences == 0, $"{Path.GetFileName(path)}: {picked.Count} spans ({blocks:N0} blocks) - stepped rescan = whole rescan ({differences} differ)");
    }

    private static void CompareSearch(string path, RtfWindowedDocument document)
    {
        var watch = Stopwatch.StartNew();
        int differences = 0, batches = 0;
        long matches = 0;
        var random = new Random(20261103);
        foreach (string query in Queries)
            foreach (bool matchCase in new[] { true, false })
                {
                    var options = new WordDocumentSearchOptions(MatchCase: matchCase);
                    int[] sizes = query == Queries[0] && matchCase ? [500, 37, 1] : [500];
                    foreach (int size in sizes)
                        for (int from = 0, run = 0; ; run++)
                        {
                            (bool same, WordDocumentSearchBatch batch) = Same(document, query, options, from, size);
                            if (!same) differences++;
                            batches++;
                            matches += batch.HitCounts.Sum();
                            if (batch.ReachedEnd || size == 1 && run >= 3000) break;     // one-block batches: the first 3,000
                            from = batch.NextBlock;
                        }
                    for (int i = 0; i < 20; i++)
                    {
                        (bool same, _) = Same(document, query, options, random.Next(document.BlockCount + 1), 1 + random.Next(80));
                        if (!same) differences++;
                        batches++;
                    }
                }
        TestReport.Check(differences == 0, $"{Path.GetFileName(path)}: {batches:N0} batches, {matches:N0} matches - one pass = old path ({differences} batches differ, {watch.ElapsedMilliseconds} ms)");
    }

    private static (bool same, WordDocumentSearchBatch batch) Same(RtfWindowedDocument document, string query, WordDocumentSearchOptions options, int from, int size)
    {
        WordDocumentSearchBatch onePass = document.Search(query, options, from, size);
        WordDocumentSearchBatch reference = document.SearchByBlockReparse(query, options, from, size);
        bool same = onePass.HitBlocks.SequenceEqual(reference.HitBlocks) && onePass.HitCounts.SequenceEqual(reference.HitCounts)
                    && onePass.NextBlock == reference.NextBlock && onePass.ReachedEnd == reference.ReachedEnd;
        if (!same) Console.WriteLine($"      differs: \"{query.Length}\" from {from} size {size} hits {onePass.HitBlocks.Length} vs {reference.HitBlocks.Length}, matches {onePass.HitCounts.Sum()} vs {reference.HitCounts.Sum()}");
        return (same, onePass);
    }
}
