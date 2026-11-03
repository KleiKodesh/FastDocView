using System.Text;
using FastDocView;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>
/// The sidecar cache folder's LRU (<see cref="WordDocumentIndexCache"/>): eviction order, the count and size caps, stale
/// and orphan cleanup, files in use left alone, saves and reads racing a trim, and the real open path (record written,
/// last use moved on a validated reopen). All in a scratch folder; the automatic trim is off while it runs.
/// </summary>
public static class WordDocumentIndexCacheTest
{
    private static readonly DateTime BaseTime = new(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);

    public static void Run()
    {
        WordIndexCacheLimits saved = WordDocumentIndexCache.Limits;
        WordDocumentIndexCache.Limits = new WordIndexCacheLimits { TrimsAutomatically = false };
        try
        {
            TestReport.Step("Index cache: eviction order and the count cap");
            EvictionOrderAndCountCap();
            TestReport.Step("Index cache: the size cap");
            SizeCap();
            TestReport.Step("Index cache: stale sidecars (document gone or changed)");
            StaleSidecars();
            TestReport.Step("Index cache: temp files and records left by interrupted writes");
            Orphans();
            TestReport.Step("Index cache: a sidecar held open by another reader is not deleted");
            HeldOpen();
            TestReport.Step("Index cache: saves and reads racing trims");
            Concurrent();
            TestReport.Step("Index cache: the real open path writes the record and marks use");
            RealOpenPath();
        }
        finally { WordDocumentIndexCache.Limits = saved; }
    }

    private static string FreshFolder(string name)
    {
        string directory = Path.Combine(TestReport.OutputDirectory, "index-cache-test", name);
        if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static string Sidecar(string directory, string name, int bytes, int minutesAfterBase)
    {
        string path = Path.Combine(directory, name + WordDocumentIndexCache.SidecarExtension);
        File.WriteAllBytes(path, new byte[bytes]);
        File.SetLastWriteTimeUtc(path, BaseTime.AddMinutes(minutesAfterBase));
        return path;
    }

    private static string Names(string directory) =>
        string.Join(",", Directory.GetFiles(directory, "*" + WordDocumentIndexCache.SidecarExtension).Select(Path.GetFileNameWithoutExtension).OrderBy(n => n, StringComparer.Ordinal));

    private static void EvictionOrderAndCountCap()
    {
        string directory = FreshFolder("count");
        // used in the order c, a, e, b, f, d (minutes 1..6): the three least recent are c, a, e
        foreach ((string name, int minute) in new[] { ("a", 2), ("b", 4), ("c", 1), ("d", 6), ("e", 3), ("f", 5) }) Sidecar(directory, name, 1000, minute);
        WordIndexCacheTrimResult result = WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits { MaxFileCount = 3 });
        TestReport.Check(Names(directory) == "b,d,f" && result.EvictedCount == 3 && result.KeptCount == 3,
                         $"count cap 3 keeps the 3 most recently used (b,d,f): {Names(directory)}, evicted {result.EvictedCount}");
        WordDocumentIndexCache.MarkUsed(Path.Combine(directory, "b" + WordDocumentIndexCache.SidecarExtension), 0);   // b is now the newest
        WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits { MaxFileCount = 1 });
        TestReport.Check(Names(directory) == "b", $"a sidecar marked used survives the next trim: {Names(directory)}");
    }

    private static void SizeCap()
    {
        string directory = FreshFolder("size");
        Sidecar(directory, "old", 3000, 1);
        Sidecar(directory, "mid", 1000, 2);
        Sidecar(directory, "new", 1000, 3);
        WordIndexCacheTrimResult result = WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits { MaxTotalBytes = 2500 });
        TestReport.Check(Names(directory) == "mid,new" && result.KeptBytes == 2000, $"size cap 2,500 B keeps the newest that fit: {Names(directory)} ({result.KeptBytes} B)");
        result = WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits());
        TestReport.Check(result.EvictedCount == 0 && result.KeptCount == 2, "under the default caps nothing goes");
    }

    private static void WriteRecord(string sidecarPath, string documentPath, long length, long ticks) =>
        File.WriteAllText(sidecarPath + ".source", documentPath + "\n" + length + "\n" + ticks + "\n", new UTF8Encoding(false));

    private static void StaleSidecars()
    {
        string directory = FreshFolder("stale");
        string documents = FreshFolder("stale-documents");
        string unchanged = Path.Combine(documents, "unchanged.txt");
        string changed = Path.Combine(documents, "changed.txt");
        File.WriteAllText(unchanged, "one");
        File.WriteAllText(changed, "two");
        foreach ((string name, string document) in new[] { ("unchanged", unchanged), ("changed", changed) })
        {
            var info = new FileInfo(document);
            WriteRecord(Sidecar(directory, name, 100, 1), document, info.Length, info.LastWriteTimeUtc.Ticks);
        }
        File.AppendAllText(changed, " more");                                           // length differs from the record
        WriteRecord(Sidecar(directory, "gone", 100, 1), Path.Combine(documents, "deleted.txt"), 3, 1);
        WriteRecord(Sidecar(directory, "unplugged", 100, 1), Path.Combine(documents, "no-such-folder", "x.txt"), 3, 1);
        Sidecar(directory, "legacy", 100, 1);                                           // no record: LRU only
        WordIndexCacheTrimResult result = WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits());
        TestReport.Check(Names(directory) == "legacy,unchanged,unplugged" && result.StaleCount == 2,
                         $"changed and deleted documents' sidecars go, unchanged / folder-missing / record-less stay: {Names(directory)} (stale {result.StaleCount})");
        TestReport.Check(!File.Exists(Path.Combine(directory, "changed.wordindex.source")) && !File.Exists(Path.Combine(directory, "gone.wordindex.source"))
                         && File.Exists(Path.Combine(directory, "unchanged.wordindex.source")), "a deleted sidecar's record goes with it");
    }

    private static void Orphans()
    {
        string directory = FreshFolder("orphans");
        Sidecar(directory, "kept", 100, 1);
        string oldTemp = Path.Combine(directory, "kept.wordindex.123-abcdef01.tmp");
        string legacyTemp = Path.Combine(directory, "other.wordindex.tmp");
        string freshTemp = Path.Combine(directory, "kept.wordindex.456-abcdef02.tmp");
        string oldRecord = Path.Combine(directory, "lost.wordindex.source");
        string unrelated = Path.Combine(directory, "notes.tmp");
        foreach (string path in new[] { oldTemp, legacyTemp, freshTemp, oldRecord, unrelated }) File.WriteAllText(path, "x");
        foreach (string path in new[] { oldTemp, legacyTemp, oldRecord, unrelated }) File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddHours(-2));
        WordIndexCacheTrimResult result = WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits());
        TestReport.Check(!File.Exists(oldTemp) && !File.Exists(legacyTemp) && !File.Exists(oldRecord) && result.OrphanCount == 3,
                         $"temp files and a record older than an hour with no sidecar go ({result.OrphanCount})");
        TestReport.Check(File.Exists(freshTemp), "a fresh temp file (a write in progress) stays");
        TestReport.Check(File.Exists(unrelated) && Names(directory) == "kept", "files that are not the cache's stay");
    }

    private static void HeldOpen()
    {
        string directory = FreshFolder("held");
        string held = Sidecar(directory, "held", 100, 1);
        Sidecar(directory, "free", 100, 2);
        WordIndexCacheTrimResult result;
        using (new FileStream(held, FileMode.Open, FileAccess.Read, FileShare.Read))   // as File.ReadAllBytes holds it
            result = WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits { MaxFileCount = 0 });
        TestReport.Check(File.Exists(held) && result.SkippedInUseCount == 1 && result.EvictedCount == 1 && Names(directory) == "held",
                         $"a sidecar open in another reader is skipped, the free one evicted: {Names(directory)} (in use {result.SkippedInUseCount})");
        WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits { MaxFileCount = 0 });
        TestReport.Check(!File.Exists(held), "once closed, the next trim takes it");
    }

    private static void Concurrent()
    {
        string directory = FreshFolder("concurrent");
        byte[] payload = Enumerable.Range(0, 64 * 1024).Select(i => (byte)(i * 7)).ToArray();
        string[] paths = Enumerable.Range(0, 8).Select(i => Path.Combine(directory, $"doc{i}{WordDocumentIndexCache.SidecarExtension}")).ToArray();
        foreach (string path in paths) WordDocumentIndexCache.Save(path, stream => stream.Write(payload, 0, payload.Length));
        var errors = new List<string>();
        int torn = 0, reads = 0, saves = 0, trims = 0;
        DateTime until = DateTime.UtcNow.AddSeconds(3);
        // two threads at most (harness rule): one saves and reads like a second app instance, the other trims hard
        var other = new Thread(() =>
        {
            var random = new Random(7);
            while (DateTime.UtcNow < until)
            {
                string path = paths[random.Next(paths.Length)];
                try
                {
                    if (random.Next(2) == 0) { WordDocumentIndexCache.Save(path, stream => stream.Write(payload, 0, payload.Length)); saves++; }
                    else if (File.Exists(path))
                    {
                        byte[] read = File.ReadAllBytes(path);
                        reads++;
                        if (!read.AsSpan().SequenceEqual(payload)) torn++;
                        WordDocumentIndexCache.MarkUsed(path, 0);
                    }
                }
                catch (FileNotFoundException) { }                       // evicted between the check and the read: a cache miss
                catch (IOException) { }                                 // the trim's delete or a save's rename holds it this instant: a cache miss
                catch (Exception exception) { lock (errors) errors.Add(exception.GetType().Name + ": " + exception.Message); }
            }
        });
        other.Start();
        while (DateTime.UtcNow < until)
        {
            try { WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits { MaxFileCount = 3, AbandonedTempAge = TimeSpan.Zero }); trims++; }
            catch (Exception exception) { lock (errors) errors.Add("trim " + exception.GetType().Name + ": " + exception.Message); }
        }
        other.Join();
        Console.WriteLine($"   {trims:N0} trims, {saves:N0} saves, {reads:N0} reads in 3 s");
        TestReport.Check(errors.Count == 0, $"no unexpected exception from trims, saves or reads ({string.Join("; ", errors.Take(3))})");
        TestReport.Check(torn == 0, $"no reader ever saw a partly written sidecar ({torn} torn of {reads})");
        bool whole = paths.Where(File.Exists).All(path => new FileInfo(path).Length == payload.Length);
        TestReport.Check(whole && saves > 0 && trims > 0, "every sidecar left is complete");
    }

    private static void RealOpenPath()
    {
        string directory = FreshFolder("real");
        string documents = FreshFolder("real-documents");
        string document = Path.Combine(documents, "real-open-v1.txt");
        File.WriteAllText(document, string.Join("\n", Enumerable.Range(0, 2000).Select(i => "line " + i)));
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = directory };
        string indexPath = policy.IndexPathFor(document);
        using (MarkupWindowedDocument opened = MarkupWindowedDocument.Open(document, MarkupSourceKind.PlainText, policy)) { }
        string recordPath = indexPath + ".source";
        TestReport.Check(File.Exists(indexPath) && File.Exists(recordPath), "a sidecar save writes the sidecar and its record");
        string[] record = File.Exists(recordPath) ? File.ReadAllText(recordPath).Split('\n') : [];
        TestReport.Check(record.Length >= 3 && string.Equals(record[0], Path.GetFullPath(document), StringComparison.OrdinalIgnoreCase)
                         && record[1] == new FileInfo(document).Length.ToString(), "the record names the document, its length and time");
        TestReport.Check(Directory.GetFiles(directory, "*.tmp").Length == 0, "no temp file left after a save");
        File.SetLastWriteTimeUtc(indexPath, BaseTime);
        using (MarkupWindowedDocument reopened = MarkupWindowedDocument.Open(document, MarkupSourceKind.PlainText, policy))
            TestReport.Check(reopened.BlockCount == 2000, "the reopen used the sidecar");
        TestReport.Check(File.GetLastWriteTimeUtc(indexPath) > BaseTime.AddDays(1), "a validated reopen marks the sidecar used");
        WordIndexCacheTrimResult result = WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits());
        TestReport.Check(result.KeptCount == 1 && result.StaleCount == 0, "the live document's sidecar is not stale");
        File.AppendAllText(document, "\nmore");
        result = WordDocumentIndexCache.Trim(directory, new WordIndexCacheLimits());
        TestReport.Check(result.StaleCount == 1 && !File.Exists(indexPath) && !File.Exists(recordPath), "after the document changes the trim deletes its sidecar and record");
    }
}
