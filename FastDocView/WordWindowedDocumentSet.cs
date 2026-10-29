using System.Diagnostics;

namespace FastDocView;

/// <summary>
/// The documents a host keeps open, by id - what a service or dev server needs around
/// <see cref="WordWindowedDocumentOpener"/>, so the host is a few lines:
/// <code>
/// using var documents = new WordWindowedDocumentSet(cacheDirectory);
/// string id = documents.Open(path);
/// var blocks = documents.Use(id, document => document.RenderBlocks(0, 50));
/// </code>
/// Every call on a document runs under that document's lock (a document is single-threaded). A document
/// nobody used for <see cref="IdleCloseAfter"/> is closed; a later call with its id, or a call after the file
/// changed on disk, throws <see cref="WordDocumentClosedException"/> and the host reopens it - quickly, because
/// the sidecar index stays in the cache directory.
/// </summary>
public sealed class WordWindowedDocumentSet : IDisposable
{
    private sealed class OpenDocument(IWordWindowedDocument document)
    {
        public readonly IWordWindowedDocument Document = document;
        public readonly object Gate = new();
        public long LastUsedTicks = Stopwatch.GetTimestamp();
    }

    private readonly Dictionary<string, OpenDocument> _documents = new(StringComparer.Ordinal);
    private readonly object _documentsGate = new();
    private readonly WordIndexPolicy _policy;
    private readonly Timer _idleTimer;

    public TimeSpan IdleCloseAfter { get; }

    /// <param name="cacheDirectory">Where sidecar indexes live (created if missing); shared by every process using it.</param>
    /// <param name="idleCloseAfter">Close a document unused this long (default 10 minutes).</param>
    /// <param name="policy">Index policy; its cache directory is replaced by <paramref name="cacheDirectory"/>.
    /// Default: indexes build on demand, so a first open of a big file returns after its first blocks.</param>
    public WordWindowedDocumentSet(string cacheDirectory, TimeSpan? idleCloseAfter = null, WordIndexPolicy? policy = null)
    {
        Directory.CreateDirectory(cacheDirectory);
        _policy = policy is null
            ? new WordIndexPolicy { CacheDirectory = cacheDirectory, IndexesOnDemand = true }
            : new WordIndexPolicy
            {
                CacheDirectory = cacheDirectory,
                IndexesOnDemand = policy.IndexesOnDemand,
                ForcedMode = policy.ForcedMode,
                SidecarThresholdBytes = policy.SidecarThresholdBytes,
            };
        IdleCloseAfter = idleCloseAfter ?? TimeSpan.FromMinutes(10);
        _idleTimer = new Timer(_ => CloseIdle(), null, TimeSpan.FromMinutes(1), TimeSpan.FromMinutes(1));
    }

    /// <summary>Documents open now.</summary>
    public int Count { get { lock (_documentsGate) return _documents.Count; } }

    /// <summary>The policy every open uses (its <see cref="WordIndexPolicy.IndexPathFor"/> names a file's sidecar).</summary>
    public WordIndexPolicy Policy => _policy;

    /// <summary>Opens the file and returns its new id.</summary>
    /// <param name="forcedMode">Overrides the size decision for this open (null = decide by size).</param>
    /// <param name="progress">Index-building progress, for a host that shows it.</param>
    public string Open(string filePath, WordIndexMode? forcedMode = null, IProgress<WordDocumentIndexProgress>? progress = null)
    {
        WordIndexPolicy policy = forcedMode is null ? _policy : new WordIndexPolicy
        {
            CacheDirectory = _policy.CacheDirectory,
            IndexesOnDemand = _policy.IndexesOnDemand,
            SidecarThresholdBytes = _policy.SidecarThresholdBytes,
            ForcedMode = forcedMode,
        };
        IWordWindowedDocument document = WordWindowedDocumentOpener.Open(filePath, policy, progress);
        string id = Guid.NewGuid().ToString("N");
        lock (_documentsGate) _documents[id] = new OpenDocument(document);
        return id;
    }

    /// <summary>Runs <paramref name="action"/> on the document under its lock.</summary>
    /// <exception cref="WordDocumentClosedException">The id is not open (closed, idle-closed) or the file changed on disk.</exception>
    public T Use<T>(string documentId, Func<IWordWindowedDocument, T> action)
    {
        OpenDocument? open;
        lock (_documentsGate) _documents.TryGetValue(documentId ?? "", out open);
        if (open is null) throw new WordDocumentClosedException("the document is not open");
        lock (open.Gate)
        {
            Interlocked.Exchange(ref open.LastUsedTicks, Stopwatch.GetTimestamp());
            try { return action(open.Document); }
            catch (WordDocumentSourceChangedException exception)
            {
                Close(documentId!);
                throw new WordDocumentClosedException("the file changed on disk", exception);
            }
            catch (ObjectDisposedException exception)
            {
                throw new WordDocumentClosedException("the document was closed", exception);
            }
        }
    }

    public void Close(string documentId)
    {
        OpenDocument? open;
        lock (_documentsGate)
        {
            if (!_documents.TryGetValue(documentId ?? "", out open)) return;
            _documents.Remove(documentId!);
        }
        lock (open.Gate) open.Document.Dispose();
    }

    private void CloseIdle()
    {
        long now = Stopwatch.GetTimestamp();
        long idleTicks = (long)(IdleCloseAfter.TotalSeconds * Stopwatch.Frequency);
        var idle = new List<string>();
        lock (_documentsGate)
        {
            foreach (KeyValuePair<string, OpenDocument> entry in _documents)
                if (now - Interlocked.Read(ref entry.Value.LastUsedTicks) > idleTicks) idle.Add(entry.Key);
        }
        foreach (string id in idle) Close(id);
    }

    public void CloseAll()
    {
        List<string> ids;
        lock (_documentsGate) ids = [.. _documents.Keys];
        foreach (string id in ids) Close(id);
    }

    public void Dispose()
    {
        _idleTimer.Dispose();
        CloseAll();
    }
}

/// <summary>The host must reopen the document: its id is not open (closed or idle-closed) or the file changed on disk.</summary>
public sealed class WordDocumentClosedException : IOException
{
    public WordDocumentClosedException(string message, Exception? inner = null) : base(message, inner) { }
}

/// <summary>Host helpers every service needs on top of <see cref="IWordWindowedDocument"/>.</summary>
public static class WordWindowedDocumentHostExtensions
{
    /// <summary>
    /// Extends an on-demand index by up to <paramref name="maxBlocks"/> blocks, stopping once <paramref name="timeBudget"/>
    /// is spent (checked every <paramref name="stepBlocks"/> blocks), so a host holding the document's lock never makes a
    /// screen fetch wait long. Returns <see cref="IWordWindowedDocument.KnownBlockCount"/>; the host calls again until
    /// <see cref="IWordWindowedDocument.IsIndexComplete"/>.
    /// </summary>
    public static int ScanFor(this IWordWindowedDocument document, int maxBlocks, TimeSpan timeBudget, int stepBlocks = 64)
    {
        int target = (int)Math.Min(int.MaxValue, (long)document.KnownBlockCount + Math.Max(1, maxBlocks));
        long budgetTicks = (long)(timeBudget.TotalSeconds * Stopwatch.Frequency);
        long started = Stopwatch.GetTimestamp();
        while (!document.IsIndexComplete && document.KnownBlockCount < target)
        {
            document.ScanAhead(Math.Min(stepBlocks, target - document.KnownBlockCount));
            if (Stopwatch.GetTimestamp() - started >= budgetTicks) break;
        }
        return document.KnownBlockCount;
    }

    /// <summary>The estimated text length of every block known so far (index-time data, no file access).</summary>
    public static int[] EstimatedTextLengths(this IWordWindowedDocument document)
    {
        var lengths = new int[document.BlockCount];
        for (int index = 0; index < lengths.Length; index++) lengths[index] = document.EstimatedTextLength(index);
        return lengths;
    }
}
