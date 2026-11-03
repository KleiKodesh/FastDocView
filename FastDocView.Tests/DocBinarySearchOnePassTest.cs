using System.Diagnostics;
using FastDocView.DocBinary;

namespace FastDocView.Tests;

/// <summary>
/// The .doc search's one pass (block boundaries and matcher text from the same characters) against the old path it
/// replaced (block ranges from the compact table, each block's text read again): every batch must give the same hit
/// blocks, counts, next block and end - Hebrew plain and pointed, English, mixed, match case on and off, notes on and
/// off, sequential batches of several sizes and batches that start anywhere. Hebrew is written as \u escapes.
/// </summary>
public static class DocBinarySearchOnePassTest
{
    private static readonly string[] SampleFiles =
    [
        "test-data/apache-poi/53379.doc", "test-data/apache-poi/58804.doc", "test-data/apache-poi/Bug33519.doc",
        "test-data/apache-poi/Bug47958.doc", "test-data/apache-poi/Bug53380_1.doc", "test-data/apache-poi/Fuzzed.doc",
        "test-data/apache-poi/GaiaTest.doc", "test-data/apache-poi/Bug61268.doc", "test-data/apache-poi/ob_is.doc",
        "test-data/apache-poi/Bug46220.doc",
    ];

    private static readonly string[] Queries =
    [
        "\u05E9\u05D1\u05EA",                                   // Hebrew, plain letters
        "\u05E9\u05B7\u05C1\u05D1\u05BC\u05B8\u05EA",           // the same, pointed (patah, shin dot, dagesh, qamats)
        "the",                                                  // English
        "The",
        "\u05D5 a",                                             // mixed
        " ",
        "qqzzqq",                                               // absent: a full scan
    ];

    public static void Run(IEnumerable<string> extraPaths)
    {
        TestReport.Step("Search: doc one pass = the old block-range path");
        string root = FindRepositoryRoot();
        IEnumerable<string> paths = SampleFiles.Select(file => Path.Combine(root, file))
            .Concat(extraPaths.Where(path => path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)))
            .Where(File.Exists);
        foreach (string path in paths) Compare(path);
    }

    /// <summary>Only the given files (e.g. a corpus sample); a file that cannot be opened is skipped - the corpus pass owns that check.</summary>
    public static void CompareFiles(IEnumerable<string> paths)
    {
        TestReport.Step("Search: doc one pass = the old block-range path (given files)");
        int skipped = 0;
        foreach (string path in paths)
        {
            DocBinaryWindowedDocument document;
            try { document = DocBinaryWindowedDocument.Open(path); }
            catch (Exception) { skipped++; continue; }
            using (document) Compare(path, document);
        }
        Console.WriteLine($"   {skipped} files skipped (not openable)");
    }

    private static void Compare(string path)
    {
        using var document = (DocBinaryWindowedDocument)WordWindowedDocumentOpener.Open(path);
        Compare(path, document);
    }

    private static void Compare(string path, DocBinaryWindowedDocument document)
    {
        var watch = Stopwatch.StartNew();
        int differences = 0, batches = 0;
        long matches = 0;
        var random = new Random(20261103);
        foreach (string query in Queries)
            foreach (bool matchCase in new[] { true, false })
                foreach (bool notes in new[] { true, false })
                {
                    if (!notes && !matchCase) continue;
                    var options = new WordDocumentSearchOptions(MatchCase: matchCase, IncludeNotes: notes);
                    int[] sizes = query == Queries[0] && matchCase && notes ? [500, 37, 1] : [500];
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
                    // batches that start anywhere (a new cursor each), and one past the end
                    for (int i = 0; i < 20; i++)
                    {
                        (bool same, _) = Same(document, query, options, random.Next(document.BlockCount + 1), 1 + random.Next(80));
                        if (!same) differences++;
                        batches++;
                    }
                }
        TestReport.Check(differences == 0, $"{Path.GetFileName(path)}: {batches:N0} batches, {matches:N0} matches - one pass = old path ({differences} batches differ, {watch.ElapsedMilliseconds} ms)");
    }

    private static (bool same, WordDocumentSearchBatch batch) Same(DocBinaryWindowedDocument document, string query, WordDocumentSearchOptions options, int from, int size)
    {
        WordDocumentSearchBatch onePass = document.Search(query, options, from, size);
        WordDocumentSearchBatch reference = document.SearchByBlockRanges(query, options, from, size);
        bool same = onePass.HitBlocks.SequenceEqual(reference.HitBlocks) && onePass.HitCounts.SequenceEqual(reference.HitCounts)
                    && onePass.NextBlock == reference.NextBlock && onePass.ReachedEnd == reference.ReachedEnd;
        if (!same) Console.WriteLine($"      differs: from {from} size {size} hits {onePass.HitBlocks.Length} vs {reference.HitBlocks.Length}");
        return (same, onePass);
    }

    private static string FindRepositoryRoot()
    {
        string? directory = AppContext.BaseDirectory;
        while (directory is not null && !Directory.Exists(Path.Combine(directory, "test-data"))) directory = Path.GetDirectoryName(directory);
        return directory ?? Directory.GetCurrentDirectory();
    }
}
