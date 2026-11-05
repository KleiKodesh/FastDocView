namespace FastDocView;

public enum WordIndexMode
{
    /// <summary>Small document: one pass, body kept in memory for this session, nothing written to disk.</summary>
    InMemory,
    /// <summary>Big document: build the checkpoint index once and persist it as a sidecar.</summary>
    Sidecar,
}

/// <summary>What the policy decided for one file, with the numbers behind it — shown in the demo's folder list.</summary>
public sealed record WordIndexDecision(
    string FilePath,
    WordIndexMode Mode,
    bool HasCachedIndex,
    long BodyBytes,
    long EstimatedIndexBytes,
    string Reason);

/// <summary>
/// Decides whether indexing a file gains anything. The checkpoint index pays off only when the
/// document body is large: below the threshold a plain pass is faster than building and loading an
/// index, and keeping the body in memory for the session is cheaper than a sidecar. The host
/// (user settings) overrides the threshold and can force a mode per file or folder.
/// </summary>
public sealed class WordIndexPolicy
{
    // The one-pass open runs at ~40 MB/s of body XML; 8 MB keeps a non-indexed open around 200 ms,
    // which is still "fast preview". Above that the sidecar's instant re-open wins.
    public const long DefaultSidecarThresholdBytes = 8L << 20;

    public static WordIndexPolicy Default { get; } = new();

    /// <summary>Uncompressed text size (.docx: document.xml + footnotes.xml + endnotes.xml; .doc: text bytes) from which a sidecar index is built.</summary>
    public long SidecarThresholdBytes { get; init; } = DefaultSidecarThresholdBytes;

    /// <summary>Forces a mode regardless of size (null = decide by size).</summary>
    public WordIndexMode? ForcedMode { get; init; }

    /// <summary>Where sidecar indexes live (null = <see cref="WordDocumentIndexCacheLocation.DefaultCacheDirectory"/>).</summary>
    public string? CacheDirectory { get; init; }

    /// <summary>
    /// .docx sidecar mode with no cached index: open after the first blocks are scanned and extend
    /// the index as the reader asks for blocks (pulled, single-threaded); the sidecar is saved when
    /// the scan reaches the end. False = build the whole index before the open returns.
    /// </summary>
    public bool IndexesOnDemand { get; init; }

    /// <summary>
    /// A .txt / web archive is scanned at ~50 MB/s (measured on the Torat Emet books, 7 MB = 120-240 ms before the first window), so from 2 MB
    /// its open is on demand (first window after the first blocks, the rest as the reader asks) instead of one blocking pass.
    /// </summary>
    public const long PlainTextSidecarThresholdBytes = 2L << 20;

    public WordIndexDecision Decide(string filePath, long bodyBytes, bool hasCachedIndex, long? thresholdCapBytes = null)
    {
        long threshold = Math.Min(SidecarThresholdBytes, thresholdCapBytes ?? long.MaxValue);
        long estimatedIndexBytes = bodyBytes / 2048 + 8192;     // compact .docx sidecar: ~45 KB per 90 MB; the other formats' per-block layouts are bigger
        if (hasCachedIndex)
            return new WordIndexDecision(filePath, WordIndexMode.Sidecar, true, bodyBytes, estimatedIndexBytes, "index already built");
        if (ForcedMode is { } forced)
            return new WordIndexDecision(filePath, forced, false, bodyBytes, estimatedIndexBytes, "forced by settings");
        return bodyBytes >= threshold
            ? new WordIndexDecision(filePath, WordIndexMode.Sidecar, false, bodyBytes, estimatedIndexBytes, $"text {bodyBytes >> 20} MB ≥ {threshold >> 20} MB")
            : new WordIndexDecision(filePath, WordIndexMode.InMemory, false, bodyBytes, 0, $"text {Math.Max(1, bodyBytes >> 10)} KB < {threshold >> 20} MB — one pass is cheaper than an index");
    }

    public string IndexPathFor(string filePath) => WordDocumentIndexCacheLocation.IndexPathFor(filePath, CacheDirectory);
}
