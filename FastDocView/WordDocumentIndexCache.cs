using System.Text;

namespace FastDocView;

/// <summary>
/// How big the sidecar cache folder may grow. Sidecars are 3-60 KB today (compact formats; older layouts up to ~1.5 MB),
/// so the 16 MB / 500 default holds far more documents than one reads in a while; losing an old sidecar only costs one
/// slower first open (decided 2026-11-03). The count cap keeps the folder listing that a trim walks cheap. A host that
/// offers the caps as a setting assigns its own <see cref="WordIndexCacheLimits"/> to <see cref="WordDocumentIndexCache.Limits"/>.
/// </summary>
public sealed class WordIndexCacheLimits
{
    public const long DefaultMaxTotalBytes = 16L << 20;
    public const int DefaultMaxFileCount = 500;

    /// <summary>Total bytes of sidecars kept (records and temp files are tiny and not counted).</summary>
    public long MaxTotalBytes { get; init; } = DefaultMaxTotalBytes;
    /// <summary>Number of sidecars kept.</summary>
    public int MaxFileCount { get; init; } = DefaultMaxFileCount;
    /// <summary>A temp file this old is from an interrupted write (a live write finishes in well under a second).</summary>
    public TimeSpan AbandonedTempAge { get; init; } = TimeSpan.FromHours(1);
    /// <summary>Trim on its own, off the caller's thread: once per process on the first use of the cache, and after a save.</summary>
    public bool TrimsAutomatically { get; init; } = true;
}

/// <summary>What one <see cref="WordDocumentIndexCache.Trim"/> did.</summary>
public sealed record WordIndexCacheTrimResult(int KeptCount, long KeptBytes, int EvictedCount, int StaleCount, int OrphanCount, int SkippedInUseCount);

/// <summary>
/// The sidecar cache folder's housekeeping (file-system hosts only). Every sidecar <c>X.wordindex</c> may have a record
/// <c>X.wordindex.source</c> (the document's full path, length and last-write ticks when the sidecar was saved).
/// <list type="bullet">
/// <item>Last use = the sidecar's last-write time, set on every save and on every open that validates it (NTFS
/// last-access updates are often off, so they are not used). The sidecar's content never depends on that time.</item>
/// <item>Writes are atomic: a temp file unique to the writer (<c>X.wordindex.&lt;pid&gt;-&lt;random&gt;.tmp</c>), then a
/// rename over the old one; a reader sees the old sidecar or the new one, never half of one.</item>
/// <item><see cref="Trim"/> deletes: temp files and records left behind by interrupted writes, sidecars whose
/// document is gone (its folder still there) or changed, then the least recently used until both caps hold.
/// A file another process has open is not deleted (the delete fails on the sharing violation and is skipped); a
/// sidecar used again since the folder was listed is kept.</item>
/// </list>
/// </summary>
public static class WordDocumentIndexCache
{
    public const string SidecarExtension = ".wordindex";
    public const string RecordExtension = ".wordindex.source";
    private const string TempExtension = ".tmp";
    private static readonly TimeSpan TrimAfterSaveInterval = TimeSpan.FromMinutes(10);

    /// <summary>The caps and the automatic trim; the host sets them (before opening documents).</summary>
    public static WordIndexCacheLimits Limits { get; set; } = new();

    private static readonly object Gate = new();
    private static readonly Dictionary<string, DateTime> LastTrimUtc = new(StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> TrimRunning = new(StringComparer.OrdinalIgnoreCase);

    // the documents of the sidecar paths computed lately (IndexPathFor), so Save can write the record without the
    // caller passing the document again; bounded, the oldest goes
    private const int RememberedDocuments = 64;
    private static readonly Dictionary<string, string> DocumentOfIndex = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Queue<string> DocumentOfIndexOrder = new();

    internal static void RememberDocument(string indexPath, string documentFullPath)
    {
        lock (Gate)
        {
            if (DocumentOfIndex.ContainsKey(indexPath)) { DocumentOfIndex[indexPath] = documentFullPath; return; }
            if (DocumentOfIndexOrder.Count >= RememberedDocuments) DocumentOfIndex.Remove(DocumentOfIndexOrder.Dequeue());
            DocumentOfIndex[indexPath] = documentFullPath;
            DocumentOfIndexOrder.Enqueue(indexPath);
        }
    }

    private static string? RememberedDocumentOf(string indexPath)
    {
        lock (Gate) return DocumentOfIndex.TryGetValue(indexPath, out string? path) ? path : null;
    }

    /// <summary>
    /// Saves a sidecar atomically (unique temp file, then rename over the old one) with its record. The sidecar is a
    /// cache: a failure (disk full, the old file held by another process) leaves no temp file and is not an error.
    /// </summary>
    public static bool Save(string indexPath, Action<Stream> write, string? documentPath = null)
    {
        documentPath ??= RememberedDocumentOf(indexPath);
        bool saved;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(indexPath)!);
            if (documentPath is not null && File.Exists(documentPath))
            {
                var info = new FileInfo(documentPath);
                string record = info.FullName + "\n" + info.Length + "\n" + info.LastWriteTimeUtc.Ticks + "\n";
                WriteAtomically(indexPath + RecordExtension.Substring(SidecarExtension.Length), stream =>
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(record);
                    stream.Write(bytes, 0, bytes.Length);
                });
            }
            WriteAtomically(indexPath, write);
            saved = true;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { saved = false; }
        if (saved) ScheduleTrim(Path.GetDirectoryName(indexPath)!, afterSave: true);
        return saved;
    }

    private static void WriteAtomically(string path, Action<Stream> write)
    {
        string temporaryPath = path + "." + CurrentProcessId() + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + TempExtension;
        try
        {
            using (var file = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) write(file);
            try { ReplaceWith(temporaryPath, path); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                // the old file is held this instant (a reader's ReadAllBytes: sharing violation) or pending delete (a trim's
                // delete while another handle with delete sharing is open, e.g. MarkUsed's or a scanner's: the name stays
                // until that handle closes and every open, delete or rename onto it is access denied). Both last
                // microseconds: one retry, then the save is a cache miss
                PauseBeforeRetry();
                ReplaceWith(temporaryPath, path);
            }
        }
        finally
        {
            try { if (File.Exists(temporaryPath)) File.Delete(temporaryPath); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
        }
    }

    private static void ReplaceWith(string temporaryPath, string path)
    {
#if NET
        File.Move(temporaryPath, path, overwrite: true);         // MoveFileEx(REPLACE_EXISTING): one rename
#else
        if (File.Exists(path)) File.Replace(temporaryPath, path, null, ignoreMetadataErrors: true);
        else File.Move(temporaryPath, path);
#endif
    }

    private static void PauseBeforeRetry()
    {
        try { Thread.Sleep(1); }
        catch (PlatformNotSupportedException) { }              // the browser build: retry at once
    }

    /// <summary>
    /// A sidecar's bytes, or null when it cannot be read this instant: gone (evicted, replaced), held by a writer or a
    /// trim's delete (sharing violation, an <see cref="IOException"/>), or pending delete (a delete while another handle
    /// with delete sharing is open leaves the name until that handle closes; opening it is access denied, an
    /// <see cref="UnauthorizedAccessException"/>). Every case is a cache miss, never an error for the caller.
    /// </summary>
    public static byte[]? TryReadSidecar(string indexPath)
    {
        try { return File.ReadAllBytes(indexPath); }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { return null; }
    }

    private static int CurrentProcessId()
    {
#if NET
        return Environment.ProcessId;
#else
        using var process = System.Diagnostics.Process.GetCurrentProcess();
        return process.Id;
#endif
    }

    /// <summary>A sidecar was just validated and is in use: marks it most recently used (one file-time write). Returns <paramref name="index"/>.</summary>
    public static T MarkUsed<T>(string indexPath, T index)
    {
        try { File.SetLastWriteTimeUtc(indexPath, DateTime.UtcNow); }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // held by another process: only the order is off
        ScheduleTrim(Path.GetDirectoryName(indexPath)!, afterSave: false);
        return index;
    }

    /// <summary>
    /// Starts a trim of the folder on a low-priority background thread when one is due: the first use of the folder in this
    /// process, or a save at least 10 minutes after the last trim. Never on the caller's thread.
    /// </summary>
    private static void ScheduleTrim(string directory, bool afterSave)
    {
        WordIndexCacheLimits limits = Limits;
        if (!limits.TrimsAutomatically || string.IsNullOrEmpty(directory)) return;
        lock (Gate)
        {
            bool trimmedBefore = LastTrimUtc.TryGetValue(directory, out DateTime last);
            if (TrimRunning.Contains(directory)) return;
            if (trimmedBefore && !(afterSave && DateTime.UtcNow - last >= TrimAfterSaveInterval)) return;
            LastTrimUtc[directory] = DateTime.UtcNow;
            TrimRunning.Add(directory);
        }
        try
        {
            var thread = new Thread(() =>
            {
                try { Trim(directory, limits); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
                finally { lock (Gate) TrimRunning.Remove(directory); }
            })
            { IsBackground = true, Priority = ThreadPriority.Lowest, Name = "FastDocView index cache trim" };
            thread.Start();
        }
        catch (Exception exception) when (exception is PlatformNotSupportedException or OutOfMemoryException or ThreadStateException)
        {
            lock (Gate) TrimRunning.Remove(directory);          // no threads (browser): the host may call Trim itself
        }
    }

    private sealed class Entry(FileInfo sidecar)
    {
        public FileInfo Sidecar { get; } = sidecar;
        public DateTime LastUseUtc { get; } = sidecar.LastWriteTimeUtc;
        public FileInfo? Record { get; set; }
    }

    /// <summary>Trims the folder now, on the caller's thread (a host may call it on its own schedule). Never throws for files in use.</summary>
    public static WordIndexCacheTrimResult Trim(string? directory = null, WordIndexCacheLimits? limits = null)
    {
        directory ??= WordDocumentIndexCacheLocation.DefaultCacheDirectory;
        limits ??= Limits;
        int stale = 0, orphans = 0, evicted = 0, inUse = 0;
        if (!Directory.Exists(directory)) return new WordIndexCacheTrimResult(0, 0, 0, 0, 0, 0);

        var entries = new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);
        var records = new List<FileInfo>();
        DateTime now = DateTime.UtcNow;
        foreach (FileInfo file in new DirectoryInfo(directory).EnumerateFiles())
        {
            string name = file.Name;
            try
            {
                if (name.EndsWith(TempExtension, StringComparison.OrdinalIgnoreCase) && name.IndexOf(SidecarExtension, StringComparison.OrdinalIgnoreCase) > 0)
                {
                    if (now - file.LastWriteTimeUtc >= limits.AbandonedTempAge) Count(TryDelete(file, null), ref orphans, ref inUse);
                }
                else if (name.EndsWith(RecordExtension, StringComparison.OrdinalIgnoreCase)) records.Add(file);
                else if (name.EndsWith(SidecarExtension, StringComparison.OrdinalIgnoreCase)) entries[file.FullName] = new Entry(file);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // vanished while listing
        }
        foreach (FileInfo record in records)
        {
            string sidecarPath = record.FullName.Substring(0, record.FullName.Length - RecordExtension.Length) + SidecarExtension;
            if (entries.TryGetValue(sidecarPath, out Entry? entry)) entry.Record = record;
            else if (now - record.LastWriteTimeUtc >= limits.AbandonedTempAge) Count(TryDelete(record, null), ref orphans, ref inUse);   // its sidecar's write was interrupted (or the sidecar went)
        }

        var kept = new List<Entry>();
        foreach (Entry entry in entries.Values)
        {
            if (entry.Record is not null && IsStale(entry.Record))
            {
                DeleteResult result = TryDeleteEntry(entry);
                if (result == DeleteResult.Deleted) stale++;
                else { if (result == DeleteResult.InUse) inUse++; kept.Add(entry); }
            }
            else kept.Add(entry);
        }

        kept.Sort((a, b) => b.LastUseUtc.CompareTo(a.LastUseUtc));    // most recent first
        int keptCount = 0;
        long keptBytes = 0;
        foreach (Entry entry in kept)
        {
            long size = SizeOf(entry.Sidecar);
            if (keptCount < limits.MaxFileCount && keptBytes + size <= limits.MaxTotalBytes)
            {
                keptCount++;
                keptBytes += size;
                continue;
            }
            DeleteResult result = TryDeleteEntry(entry);
            if (result == DeleteResult.Deleted) evicted++;
            else
            {
                if (result == DeleteResult.InUse) inUse++;
                keptCount++;
                keptBytes += size;                              // still there: it counts, and older ones go instead
            }
        }
        return new WordIndexCacheTrimResult(keptCount, keptBytes, evicted, stale, orphans, inUse);
    }

    private static long SizeOf(FileInfo file)
    {
        try { return file.Length; }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { return 0; }
    }

    private static void Count(DeleteResult result, ref int deleted, ref int inUse)
    {
        if (result == DeleteResult.Deleted) deleted++;
        else if (result == DeleteResult.InUse) inUse++;
    }

    /// <summary>The record's document is gone (its folder still there, so not an unplugged drive or share) or changed since the save.</summary>
    private static bool IsStale(FileInfo record)
    {
        string[] lines;
        try { lines = File.ReadAllText(record.FullName, Encoding.UTF8).Split('\n'); }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { return false; }
        if (lines.Length < 3 || !long.TryParse(lines[1], out long length) || !long.TryParse(lines[2], out long ticks)) return false;
        string documentPath = lines[0];
        try
        {
            var document = new FileInfo(documentPath);
            if (!document.Exists) return document.Directory is { Exists: true };
            return document.Length != length || document.LastWriteTimeUtc.Ticks != ticks;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) { return false; }
    }

    private enum DeleteResult { Deleted, InUse, UsedAgain }

    private static DeleteResult TryDeleteEntry(Entry entry)
    {
        DeleteResult result = TryDelete(entry.Sidecar, entry.LastUseUtc);
        if (result == DeleteResult.Deleted && entry.Record is not null) TryDelete(entry.Record, null);
        return result;
    }

    /// <summary>
    /// Deletes the file unless it was used again since it was listed (<paramref name="listedLastUseUtc"/>) or another
    /// process holds it: on Windows the delete fails with a sharing violation while any handle without delete sharing
    /// is open (sidecar readers and writers open it that way), and is skipped.
    /// </summary>
    private static DeleteResult TryDelete(FileInfo file, DateTime? listedLastUseUtc)
    {
        try
        {
            if (listedLastUseUtc is { } listed)
            {
                file.Refresh();
                if (!file.Exists) return DeleteResult.UsedAgain;          // replaced or deleted by someone else
                if (file.LastWriteTimeUtc != listed) return DeleteResult.UsedAgain;
            }
            file.Delete();
            return DeleteResult.Deleted;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return DeleteResult.InUse;
        }
    }
}
