using FastDocView.Html;

namespace FastDocView.Rtf;

/// <summary>
/// An RTF file opened for windowed rendering behind <see cref="IWordWindowedDocument"/>. RTF is plain bytes, so the
/// file itself is the random-access view (like a flat .fodt): the one index pass records each block's byte range and
/// the group state at its start; a window reads and parses only its blocks' bytes. Small files keep the index in
/// memory for the session; big ones (policy) save it as a sidecar so a reopen reads no RTF at all.
/// </summary>
public sealed class RtfWindowedDocument : IWordWindowedDocument
{
    private readonly RtfRandomAccessIndex _index;
    private readonly RtfCompactBlockTable _table;
    private readonly Stream _file;
    private readonly RtfBlockHtmlRenderer _renderer;
    private readonly object _readLock = new();
    private readonly int _footnoteCount, _endnoteCount;
    private WordDocumentFileStream? _releasableFile;

    public string FilePath { get; }
    public int BlockCount => _index.BlockCount;

    /// <summary>Page text width per section (\paperw / \margl / \margr / \gutter and their per-section forms), from the index.</summary>
    public int? FindBookmark(string name) => _index.Bookmarks.BlockOf(name);     // the index is always complete
    public IReadOnlyList<WordDocumentSection> Sections => _sections ??= WordDocumentPageSetup.FromSectionEnds(_index.SectionLastBlocks, _index.SectionTextWidths);
    private IReadOnlyList<WordDocumentSection>? _sections;
    public bool IsIndexComplete => true;
    public int KnownBlockCount => _index.BlockCount;
    public int ScanAhead(int blocks) => BlockCount;          // always complete
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

    public RtfRandomAccessIndex Index => _index;

    private RtfWindowedDocument(string filePath, Stream file, RtfRandomAccessIndex index, WordIndexDecision decision)
    {
        FilePath = filePath;
        _file = file;
        _index = index;
        _table = new RtfCompactBlockTable(index, file);
        Decision = decision;
        _renderer = new RtfBlockHtmlRenderer(index.Tables);
        _fieldContext = FieldContextFor(filePath);
        _renderer.FieldContext = _fieldContext;
        _table.FieldContext = _fieldContext;
        foreach (RtfNoteReference note in index.NoteReferences)
            if (note.Kind == WordDocumentNoteKind.Footnote) _footnoteCount++; else _endnoteCount++;
        Stylesheet = WordDocumentHtmlStylesheet.Build(RtfBlockHtmlRenderer.DocumentRunFormatting(index.Tables), IsRightToLeft);
    }

    /// <summary>Exact once the block's span has been rescanned (a window near it was read), else its 64-block group's average.</summary>
    public int EstimatedTextLength(int blockIndex)
    {
        lock (_readLock) return _table.KnownTextLength(blockIndex) ?? _index.EstimatedTextLength(blockIndex);
    }

    /// <summary>The block's list label as the index pass counted it, or null - rescanned from its span.</summary>
    public string? ListLabelOf(int blockIndex) => WithTable(() => _table.ListLabelOf(blockIndex));
    public WordDocumentBlockKind BlockKindOf(int blockIndex) => WithTable(() => _table.BlockKindOf(blockIndex));

    private T WithTable<T>(Func<T> lookup)
    {
        lock (_readLock)
        {
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try { return lookup(); }
            finally { _releasableFile?.Release(); }
        }
    }

    /// <summary>
    /// Scans blocks [fromBlock, fromBlock + maxBlocksToScan) in one parse: the span rescan that finds the block bounds feeds
    /// their text to the matcher as it goes (<see cref="RtfCompactBlockTable.Search"/>) - no HTML, no second parse.
    /// Footnotes are groups inside the block, fed when <see cref="WordDocumentSearchOptions.IncludeNotes"/>.
    /// </summary>
    public WordDocumentSearchBatch Search(string query, WordDocumentSearchOptions options, int fromBlock, int maxBlocksToScan)
    {
        var matcher = new WordDocumentSearchMatcher(query, options);
        if (matcher.IsEmpty) return WordDocumentSearchBatch.Nothing;
        fromBlock = Math.Max(0, fromBlock);
        lock (_readLock)
        {
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                int end = (int)Math.Min(_index.BlockCount, (long)fromBlock + Math.Max(1, maxBlocksToScan));
                var search = new RtfSpanSearch(matcher, options.IncludeNotes);
                if (fromBlock < end) _table.Search(fromBlock, end, search);
                return WordDocumentSearchBatch.Finish(search.Hits, search.Counts, Math.Max(end, fromBlock), isIndexComplete: true, _index.BlockCount);
            }
            finally { _releasableFile?.Release(); }
        }
    }

    /// <summary>The per-block reference the one-pass search must equal: each block's range re-parsed from its snapshot (<see cref="RtfSearchTextSink"/>).</summary>
    internal WordDocumentSearchBatch SearchByBlockReparse(string query, WordDocumentSearchOptions options, int fromBlock, int maxBlocksToScan)
    {
        var matcher = new WordDocumentSearchMatcher(query, options);
        if (matcher.IsEmpty) return WordDocumentSearchBatch.Nothing;
        fromBlock = Math.Max(0, fromBlock);
        lock (_readLock)
        {
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                int end = (int)Math.Min(_index.BlockCount, (long)fromBlock + Math.Max(1, maxBlocksToScan));
                var hits = new List<int>();
                var counts = new List<int>();
                _searchSink ??= new RtfSearchTextSink(_index.Tables, _fieldContext);
                for (int blockIndex = fromBlock; blockIndex < end; blockIndex++)
                {
                    matcher.Reset();
                    _searchSink.Feed(_file, _table, blockIndex, matcher, options.IncludeNotes);
                    if (matcher.Count > 0) { hits.Add(blockIndex); counts.Add(matcher.Count); }
                }
                return WordDocumentSearchBatch.Finish(hits, counts, Math.Max(end, fromBlock), isIndexComplete: true, _index.BlockCount);
            }
            finally { _releasableFile?.Release(); }
        }
    }
    private RtfSearchTextSink? _searchSink;
    private readonly Fields.WordFieldContext _fieldContext;
    private bool _isReadingBookmarkText;

    private Fields.WordFieldContext FieldContextFor(string path)
    {
        RtfFieldProperties? properties = null;
        RtfFieldProperties Properties()
        {
            if (properties is not null) return properties;
            long position = _file.Position;
            try { return properties = new RtfFieldProperties(_file); }
            finally { _file.Position = position; }
        }
        long? size = null;
        try { size = File.Exists(path) ? new FileInfo(path).Length : null; } catch (IOException) { } catch (UnauthorizedAccessException) { }
        return new Fields.WordFieldContext
        {
            DocumentProperty = name => Properties().Property(name),
            DocumentDate = name => Properties().Date(name),
            DocumentVariable = name => Properties().Variable(name),
            FileName = Path.GetFileName(path), FilePath = path, FileSize = size,
            BookmarkText = BookmarkText,
        };
    }

    /// <summary>
    /// A REF Word saved without a result: the text after the bookmark's start in the block that holds it (the parser reads
    /// {\*kmkstart} only - RTF 1.5 "Bookmarks" - so the paragraph's rest stands for the bookmarked text); null when there
    /// is no such bookmark.
    /// </summary>
    private string? BookmarkText(string name)
    {
        if (_isReadingBookmarkText || FindBookmark(name) is not { } block) return null;
        _isReadingBookmarkText = true;
        try
        {
            // its own renderer: this runs while the document's renderer is mid-block
            var renderer = new RtfBlockHtmlRenderer(_index.Tables) { FieldContext = _fieldContext };
            long position = _file.Position;
            string html;
            try { html = renderer.Render(_file, _index, _table, block).html; }
            finally { _file.Position = position; }
            string marker = "id=\"" + Html.WordHtmlText.Encode(DocxHtml.DocxBlockHtmlRenderer.BookmarkElementId(name)) + "\"></span>";
            int at = html.IndexOf(marker, StringComparison.Ordinal);
            string after = at < 0 ? html : html.Substring(at + marker.Length);
            return System.Net.WebUtility.HtmlDecode(System.Text.RegularExpressions.Regex.Replace(after, "<[^>]+>", "")).Trim();
        }
        finally { _isReadingBookmarkText = false; }
    }

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try { return RenderBlocksLocked(firstBlockIndex, count); }
            finally { _releasableFile?.Release(); }
        }
    }

    private List<WordDocumentBlockHtml> RenderBlocksLocked(int firstBlockIndex, int count)
    {
        int first = Math.Max(0, firstBlockIndex);
        int end = (int)Math.Min(_index.BlockCount, (long)firstBlockIndex + Math.Max(0, count));
        var result = new List<WordDocumentBlockHtml>(Math.Max(0, end - first));
        for (int blockIndex = first; blockIndex < end; blockIndex++)
        {
            (string html, IReadOnlyList<WordDocumentNoteHtml> notes) = _renderer.Render(_file, _index, _table, blockIndex);
            result.Add(new WordDocumentBlockHtml(blockIndex, _table.BlockKindOf(blockIndex), html, notes));
        }
        return result;
    }

    /// <inheritdoc/>
    /// <remarks>Id "block:ordinal": that block is parsed again (one block's bytes) keeping only that picture's data.</remarks>
    public WordDocumentPicture? ReadPicture(string pictureId)
    {
        string[] parts = (pictureId ?? "").Split(':');
        if (parts.Length != 2 || !int.TryParse(parts[0], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int blockIndex)
            || !int.TryParse(parts[1], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int ordinal)
            || blockIndex >= _index.BlockCount)
            return null;
        lock (_readLock)
        {
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                return _renderer.ReadPicture(_file, _index, _table, blockIndex, ordinal) is { Data: { } data } picture
                    ? WordPictureHtml.Picture(picture.Extension, data) : null;
            }
            finally { _releasableFile?.Release(); }
        }
    }

    /// <summary>
    /// Notes are inline in RTF (footnote groups): only the blocks holding the window's notes are rendered - all of the
    /// window's runs under one lock, one source check and one file open.
    /// </summary>
    public IReadOnlyList<WordDocumentNoteHtml> RenderNotes(int firstNoteIndex, int count)
    {
        lock (_readLock)
        {
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try { return WordDocumentNoteWindow.RenderFromBlocks(NoteReferences, firstNoteIndex, count, RenderBlocksLocked); }
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
