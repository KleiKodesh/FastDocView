using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.Rtf;

/// <summary>
/// An RTF file opened for windowed rendering behind <see cref="IWordWindowedDocument"/>. RTF is plain bytes, so the
/// file itself is the random-access view (like a flat .fodt): the one index pass records each block's byte range and
/// the group state at its start; a window reads and parses only its blocks' bytes. Small files keep the index in
/// memory for the session; big ones (policy) save it as a sidecar so a reopen reads no RTF at all.
/// </summary>
public sealed class RtfWindowedDocument : IWordWindowedDocument
{
    private readonly RtfRandomAccessIndex _index;
    private readonly Stream _file;
    private readonly RtfBlockHtmlRenderer _renderer;
    private readonly object _readLock = new();
    private readonly int _footnoteCount, _endnoteCount;
    private WordDocumentFileStream? _releasableFile;

    public string FilePath { get; }
    public int BlockCount => _index.BlockCount;
    public bool IsIndexComplete => true;
    public int KnownBlockCount => _index.BlockCount;
    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;
    public bool IsRightToLeft => _index.IsRightToLeft;
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }

    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy)
    {
        bool hasCachedIndex = RtfRandomAccessIndex.TryLoad(policy.IndexPathFor(filePath), filePath) is not null;
        return policy.Decide(filePath, new FileInfo(filePath).Length, hasCachedIndex);
    }

    public static RtfWindowedDocument Open(string filePath, WordIndexPolicy? policy = null,
                                           IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        policy ??= WordIndexPolicy.Default;
        var info = new FileInfo(filePath);
        string indexPath = policy.IndexPathFor(filePath);
        RtfRandomAccessIndex? cached = RtfRandomAccessIndex.TryLoad(indexPath, filePath);
        var file = new WordDocumentFileStream(filePath);
        try
        {
            WordIndexDecision decision = policy.Decide(filePath, file.Length, cached is not null);
            RtfWindowedDocument document = OpenFromStream(file, filePath, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (decision.Mode == WordIndexMode.Sidecar && cached is null)
            {
                try { document._index.Save(indexPath); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // the sidecar is a cache
            }
            document._releasableFile = file;
            file.Release();                                    // no handle between windows (Word must be able to save)
            return document;
        }
        catch
        {
            file.Dispose();
            throw;
        }
    }

    /// <summary>Host-agnostic entry: <paramref name="file"/> must be seekable; the document owns it.</summary>
    public static RtfWindowedDocument OpenFromStream(Stream file, string displayPath, WordIndexDecision decision, RtfRandomAccessIndex? cachedIndex,
                                                     long sourceLength, long sourceLastWriteTicksUtc,
                                                     IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        RtfRandomAccessIndex index = cachedIndex is not null && cachedIndex.MatchesSource(sourceLength, sourceLastWriteTicksUtc)
            ? cachedIndex
            : RtfRandomAccessIndexBuilder.Build(file, sourceLength, sourceLastWriteTicksUtc, progress, cancellation);
        return new RtfWindowedDocument(displayPath, file, index, decision);
    }

    internal RtfRandomAccessIndex Index => _index;

    private RtfWindowedDocument(string filePath, Stream file, RtfRandomAccessIndex index, WordIndexDecision decision)
    {
        FilePath = filePath;
        _file = file;
        _index = index;
        Decision = decision;
        _renderer = new RtfBlockHtmlRenderer(index.Tables);
        foreach (RtfNoteReference note in index.NoteReferences)
            if (note.Kind == WordDocumentNoteKind.Footnote) _footnoteCount++; else _endnoteCount++;
        Stylesheet = WordDocumentHtmlStylesheet.Build(RtfBlockHtmlRenderer.DocumentRunFormatting(index.Tables), IsRightToLeft);
    }

    public int EstimatedTextLength(int blockIndex) => _index.BlockTextLengths[blockIndex];

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                int first = Math.Max(0, firstBlockIndex);
                int end = (int)Math.Min(_index.BlockCount, (long)firstBlockIndex + Math.Max(0, count));
                var result = new List<WordDocumentBlockHtml>(Math.Max(0, end - first));
                for (int blockIndex = first; blockIndex < end; blockIndex++)
                {
                    (string html, IReadOnlyList<WordDocumentNoteHtml> notes) = _renderer.Render(_file, _index, blockIndex);
                    result.Add(new WordDocumentBlockHtml(blockIndex, _index.BlockKinds[blockIndex], html, notes));
                }
                return result;
            }
            finally { _releasableFile?.Release(); }
        }
    }

    public IReadOnlyList<WordDocumentHeading> Headings => _index.Headings;

    public (int footnotes, int endnotes) NoteCounts => (_footnoteCount, _endnoteCount);

    public IReadOnlyList<WordDocumentNoteReference> NoteReferences
        => _noteReferences ??= _index.NoteReferences.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
    private WordDocumentNoteReference[]? _noteReferences;

    public void Dispose() => _file.Dispose();
}
