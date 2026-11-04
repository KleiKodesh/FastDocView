using FastDocView.Html;

namespace FastDocView.Rtf;

/// <summary>
/// An RTF file opened for windowed rendering behind <see cref="IWordWindowedDocument"/>. RTF is plain bytes, so the
/// file itself is the random-access view (like a flat .fodt): the one index pass records each block's byte range and
/// the group state at its start; a window reads and parses only its blocks' bytes. Small files keep the index in
/// memory for the session; big ones (policy) save it as a sidecar so a reopen reads no RTF at all.
/// A sidecar first open with <see cref="WordIndexPolicy.IndexesOnDemand"/> returns after the first
/// <see cref="OnDemandFirstBlockCount"/> blocks; <see cref="ScanAhead"/> parses on (the same one pass, paused between calls),
/// windows read the known blocks from the partial index, and the finished index is saved when the pass ends.
/// </summary>
public sealed class RtfWindowedDocument : IWordWindowedDocument
{
    public const int OnDemandFirstBlockCount = RtfRandomAccessIndexBuilder.DirectionSampleBlocks;   // the direction is known by then

    private RtfRandomAccessIndex _indexField;
    private RtfCompactBlockTable _tableField;
    private RtfRandomAccessIndexBuilder? _builder;         // null once the index is complete
    private bool _isIndexStale;                            // the builder knows more blocks than _indexField
    private readonly string? _sidecarPathToSave;
    private readonly Stream _file;
    private readonly RtfBlockHtmlRenderer _renderer;
    private WordDocumentBlockRenderGuard? _blockRenderGuardField;
    private WordDocumentBlockRenderGuard _blockRenderGuard => _blockRenderGuardField ??= new WordDocumentBlockRenderGuard(FilePath);
    private readonly object _readLock = new();
    private int _footnoteCount, _endnoteCount;
    private WordDocumentFileStream? _releasableFile;

    public string FilePath { get; }
    /// <summary>The blocks known so far while an on-demand index is scanning (only grows), exact once complete.</summary>
    public int BlockCount => KnownBlockCount;

    /// <summary>The index of the blocks known so far (a partial index is re-made only after the scan moved on).</summary>
    private RtfRandomAccessIndex _index
    {
        get
        {
            if (_isIndexStale && IsRefreshDue()) RefreshIndex(_builder!.CreatePartialIndex());
            return _indexField;
        }
    }

    private RtfCompactBlockTable _table
    {
        get
        {
            if (_isIndexStale && IsRefreshDue()) RefreshIndex(_builder!.CreatePartialIndex());
            return _tableField;
        }
    }

    // Re-making the partial index copies its tables (O(known blocks)), so it is re-made only when the known blocks doubled or
    // half a second passed - a scan interleaved with window reads stays linear - and at once when a window or search needs blocks
    // past it (CoverKnownBlocks).
    private long _refreshedAtTicks = System.Diagnostics.Stopwatch.GetTimestamp();
    private bool IsRefreshDue() => _builder!.KnownBlockCount >= 2L * Math.Max(1, _indexField.BlockCount)
                                   || System.Diagnostics.Stopwatch.GetTimestamp() - _refreshedAtTicks > System.Diagnostics.Stopwatch.Frequency / 2;

    /// <summary>The partial index covers blocks before <paramref name="end"/> (or every known block).</summary>
    private void CoverKnownBlocks(long end)
    {
        if (_isIndexStale && end > _indexField.BlockCount) RefreshIndex(_builder!.CreatePartialIndex());
    }

    private void RefreshIndex(RtfRandomAccessIndex index)
    {
        _isIndexStale = false;
        _refreshedAtTicks = System.Diagnostics.Stopwatch.GetTimestamp();
        _indexField = index;
        _tableField = new RtfCompactBlockTable(index, _file) { FieldContext = _fieldContextField };
        _lookupTable = null;
        _sections = null;
        _noteReferences = null;
        _footnoteCount = _endnoteCount = 0;
        foreach (RtfNoteReference note in index.NoteReferences)
            if (note.Kind == WordDocumentNoteKind.Footnote) _footnoteCount++; else _endnoteCount++;
    }

    /// <summary>The block a bookmark marks; while the index is still building, the scan goes on only until it is found.</summary>
    public int? FindBookmark(string name)
    {
        while (true)
        {
            if (KnownBookmark(name) is { } block) return block;
            if (IsIndexComplete) return null;
            ScanAhead(4096);
        }
    }

    /// <summary>A bookmark among the blocks known so far (no scanning: field lookups run mid-render).</summary>
    private int? KnownBookmark(string name) => _index.Bookmarks.BlockOf(name);

    /// <summary>Page text width per section (\paperw / \margl / \margr / \gutter and their per-section forms), from the index.</summary>
    public IReadOnlyList<WordDocumentSection> Sections => _sections ??= WordDocumentPageSetup.FromSectionEnds(_index.SectionLastBlocks, _index.SectionTextWidths, _index.SectionColumnLefts);
    private IReadOnlyList<WordDocumentSection>? _sections;
    public bool IsIndexComplete { get { lock (_readLock) return _builder is null; } }
    public int KnownBlockCount { get { lock (_readLock) return _builder?.KnownBlockCount ?? _indexField.BlockCount; } }

    /// <summary>Parses on by up to <paramref name="blocks"/> blocks (index only, nothing rendered); saves the sidecar when the pass ends.</summary>
    public int ScanAhead(int blocks)
    {
        lock (_readLock)
        {
            if (_builder is not { } builder) return _indexField.BlockCount;
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                builder.AdvanceUntilKnown((int)Math.Min(int.MaxValue, (long)builder.KnownBlockCount + Math.Max(1, blocks)));
                if (builder.IsEveryBlockKnown) CompleteIndexLocked();
                else _isIndexStale = true;
                return _builder?.KnownBlockCount ?? _indexField.BlockCount;
            }
            finally { _releasableFile?.Release(); }
        }
    }

    private void CompleteIndexLocked()
    {
        RtfRandomAccessIndex index = _builder!.Complete();
        _builder = null;
        RefreshIndex(index);
        if (_sidecarPathToSave is null) return;
        try { index.Save(_sidecarPathToSave); }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // the sidecar is a cache
    }
    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;
    public bool IsRightToLeft => _isRightToLeft ??= _indexField.IsRightToLeft;     // from the first blocks: the same on a first open and a reopen
    private bool? _isRightToLeft;
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
            bool isOnDemand = decision.Mode == WordIndexMode.Sidecar && cached is null && policy.IndexesOnDemand;
            RtfWindowedDocument document = isOnDemand
                ? OpenOnDemand(file, filePath, decision, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation, indexPath)
                : OpenFromStream(file, filePath, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (!isOnDemand && decision.Mode == WordIndexMode.Sidecar && cached is null)
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
        return new RtfWindowedDocument(displayPath, file, index, null, decision, sidecarPathToSave: null);
    }

    /// <summary>
    /// Returns once the first blocks are known (the header tables precede them, so they are complete); the rest of the index
    /// is parsed by <see cref="ScanAhead"/>. The document's direction and stylesheet come from what the first blocks show.
    /// </summary>
    private static RtfWindowedDocument OpenOnDemand(Stream file, string displayPath, WordIndexDecision decision, long sourceLength, long sourceLastWriteTicksUtc,
                                                    IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation, string sidecarPathToSave)
    {
        // the cancellation stops the open only: later steps run under the host's own calls
        RtfRandomAccessIndexBuilder builder = RtfRandomAccessIndexBuilder.Incremental(file, sourceLength, sourceLastWriteTicksUtc, progress, cancellation);
        builder.AdvanceUntilKnown(OnDemandFirstBlockCount);
        if (builder.IsEveryBlockKnown)
        {
            RtfRandomAccessIndex whole = builder.Complete();
            try { whole.Save(sidecarPathToSave); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
            return new RtfWindowedDocument(displayPath, file, whole, null, decision, sidecarPathToSave: null);
        }
        builder.ClearCancellation();
        return new RtfWindowedDocument(displayPath, file, builder.CreatePartialIndex(), builder, decision, sidecarPathToSave);
    }

    /// <summary>The finished index. Throws while an on-demand open is still scanning.</summary>
    public RtfRandomAccessIndex Index { get { lock (_readLock) return _builder is null ? _indexField : throw new InvalidOperationException("The index is still being built on demand."); } }

    private RtfWindowedDocument(string filePath, Stream file, RtfRandomAccessIndex index, RtfRandomAccessIndexBuilder? builder, WordIndexDecision decision,
                                string? sidecarPathToSave)
    {
        FilePath = filePath;
        _file = file;
        _builder = builder;
        _sidecarPathToSave = sidecarPathToSave;
        Decision = decision;
        _renderer = new RtfBlockHtmlRenderer(index.Tables);
        _indexField = index;
        _fieldContextField = FieldContextFor(filePath);
        _tableField = null!;
        RefreshIndex(index);
        _renderer.FieldContext = _fieldContext;
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
            CoverKnownBlocks(fromBlock + (long)Math.Max(1, maxBlocksToScan));
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                int end = (int)Math.Min(_index.BlockCount, (long)fromBlock + Math.Max(1, maxBlocksToScan));
                var search = new RtfSpanSearch(matcher, options.IncludeNotes) { FieldContext = _fieldContext };
                _fieldContext.CurrentBlock = fromBlock;
                if (fromBlock < end) _table.Search(fromBlock, end, search);
                return WordDocumentSearchBatch.Finish(search.Hits, search.Counts, Math.Max(end, fromBlock), isIndexComplete: _builder is null, _index.BlockCount);
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
            CoverKnownBlocks(fromBlock + (long)Math.Max(1, maxBlocksToScan));
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
                    _fieldContext.CurrentBlock = blockIndex;
                    _searchSink.Feed(_file, _table, blockIndex, matcher, options.IncludeNotes);
                    if (matcher.Count > 0) { hits.Add(blockIndex); counts.Add(matcher.Count); }
                }
                return WordDocumentSearchBatch.Finish(hits, counts, Math.Max(end, fromBlock), isIndexComplete: _builder is null, _index.BlockCount);
            }
            finally { _releasableFile?.Release(); }
        }
    }
    private RtfSearchTextSink? _searchSink;

    /// <summary>
    /// The block table field lookups read (REF, STYLEREF, NOTEREF, list numbers): its own span cache, whose rescans compute
    /// no fields - a lookup runs while the document's own table is mid-rescan, and must not re-enter that rescan.
    /// </summary>
    private RtfCompactBlockTable LookupTable => _lookupTable ??= new RtfCompactBlockTable(_index, _file);
    private RtfCompactBlockTable? _lookupTable;

    private static string PlainHtmlText(string html) => System.Net.WebUtility.HtmlDecode(System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", ""));
    private readonly Fields.WordFieldContext _fieldContextField;
    private Fields.WordFieldContext _fieldContext => _fieldContextField;
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
            BookmarkBlock = KnownBookmark,
            ListLabelOfBlock = block => block >= 0 && block < _index.BlockCount ? LookupTable.ListLabelOf(block) : null,
            ListContextOfBlock = ListContext,
            IsRightToLeftText = IsRightToLeft,
            NoteMarkOfBookmark = name => KnownBookmark(name) is { } block && _index.NoteReferencesInBlock(block) is { Length: > 0 } notes ? notes[0].Label : null,
            StyledParagraph = StyledParagraph,
            Headings = () => _index.Headings,
            IndexEntries = () => _index.IndexEntries.Select(Docx.FieldLookups.ParseEntry).ToList(),
        };
    }

    /// <summary>
    /// A REF Word saved without a result: the text after the bookmark's start in the block that holds it (the parser reads
    /// {\*kmkstart} only - RTF 1.5 "Bookmarks" - so the paragraph's rest stands for the bookmarked text); null when there
    /// is no such bookmark.
    /// </summary>
    /// <summary>
    /// A block's number in full context (REF \w \r): its label, then for each upper level the nearest earlier list paragraph
    /// of that level (the RTF index keeps a level per block, not the list's counters).
    /// </summary>
    private string? ListContext(int block)
    {
        if (block < 0 || block >= _index.BlockCount || LookupTable.ListLabelOf(block) is not { } own) return null;
        int need = LookupTable.ListLevelOf(block);
        if (need < 0) return null;
        int list = LookupTable.ListOf(block);
        var labels = new List<string> { need.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\u0002" + own };
        need--;
        // back to the nearest paragraph of each upper level in the same list (\ls); another list's paragraph ends the walk, and
        // so does ListWalkLimit blocks - a list that starts below level 1 has no upper paragraph to find
        for (int earlier = block - 1; earlier >= 0 && earlier >= block - ListWalkLimit && need >= 0; earlier--)
        {
            int level = LookupTable.ListLevelOf(earlier);
            if (level < 0) continue;
            if (LookupTable.ListOf(earlier) != list) break;
            if (level > need || LookupTable.ListLabelOf(earlier) is not { } label) continue;
            labels.Insert(0, level.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\u0002" + label);
            need = level - 1;
        }
        labels.Insert(0, "#ls" + list.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return string.Join("\u0001", labels);
    }

    /// <summary>STYLEREF: the nearest block whose first paragraph has the named style (\sN from the span scans, the name from \stylesheet).</summary>
    /// <summary>
    /// How far REF \r \w and STYLEREF look from a block (each block read is a span rescan of the RTF, whose index keeps no
    /// per-style or per-list table): a paragraph further away counts as not found.
    /// </summary>
    private const int ListWalkLimit = 4096, StyleSearchLimit = 4096;

    private (int block, string text)? StyledParagraph(string styleName, int from, bool forward)
    {
        var numbers = new HashSet<int>(_index.Tables.Styles.Where(style => string.Equals(style.Value.Name.Trim(), styleName.Trim(), StringComparison.OrdinalIgnoreCase)).Select(style => style.Key));
        if (numbers.Count == 0) return HeadingParagraph(styleName, from, forward);
        int limit = forward ? Math.Min(_index.BlockCount, from + 1 + StyleSearchLimit) : Math.Max(-1, from - StyleSearchLimit);
        for (int block = forward ? from + 1 : Math.Min(from, _index.BlockCount - 1); forward ? block < limit : block > limit; block += forward ? 1 : -1)
        {
            if (!numbers.Contains(LookupTable.StyleOf(block))) continue;
            var renderer = new RtfBlockHtmlRenderer(_index.Tables) { FieldContext = null, ColumnWidthPt = WordDocumentPageSetup.TextWidthPointsAt(Sections, block), ColumnLeftPt = WordDocumentPageSetup.ColumnLeftPointsAt(Sections, block) };
            long position = _file.Position;
            try { return (block, PlainHtmlText(renderer.Render(_file, _index, LookupTable, block).html).Trim()); }
            finally { _file.Position = position; }
        }
        return null;
    }

    private (int block, string text)? HeadingParagraph(string style, int from, bool forward)
    {
        string lower = style.Trim().ToLowerInvariant();
        if (!lower.StartsWith("heading ", StringComparison.Ordinal) || !int.TryParse(lower.Substring(8), out int level)) return null;
        WordDocumentHeading? found = null;
        foreach (WordDocumentHeading heading in _index.Headings)
        {
            if (heading.Level != level) continue;
            if (!forward && heading.BlockIndex <= from) found = heading;
            if (forward && heading.BlockIndex > from) { found = heading; break; }
        }
        return found is { } h ? (h.BlockIndex, h.Text) : null;
    }

    private string? BookmarkText(string name)
    {
        if (_isReadingBookmarkText || KnownBookmark(name) is not { } block) return null;
        _isReadingBookmarkText = true;
        try
        {
            // its own renderer: this runs while the document's renderer is mid-block
            var renderer = new RtfBlockHtmlRenderer(_index.Tables) { FieldContext = _fieldContext, ColumnWidthPt = WordDocumentPageSetup.TextWidthPointsAt(Sections, block), ColumnLeftPt = WordDocumentPageSetup.ColumnLeftPointsAt(Sections, block) };
            long position = _file.Position;
            string html;
            try { html = renderer.Render(_file, _index, LookupTable, block).html; }
            finally { _file.Position = position; }
            string marker = "id=\"" + Html.WordHtmlText.Encode(DocxHtml.DocxBlockHtmlRenderer.BookmarkElementId(name)) + "\"></span>";
            string endMarker = "<span class=\"bookmark-end\" data-bookmark=\"" + Html.WordHtmlText.Encode(name) + "\"></span>";
            int at = html.IndexOf(marker, StringComparison.Ordinal);
            string after = at < 0 ? html : html.Substring(at + marker.Length);
            // the text up to {\*\bkmkend}: in this block, else on through the next blocks (a bookmark over paragraphs)
            var text = new System.Text.StringBuilder();
            for (int next = block + 1; ; next++)
            {
                int end = after.IndexOf(endMarker, StringComparison.Ordinal);
                if (end >= 0) { text.Append(PlainHtmlText(after.Substring(0, end))); break; }
                text.Append(PlainHtmlText(after));
                if (next >= _index.BlockCount || text.Length > 64 * 1024) break;
                long nextPosition = _file.Position;
                try { after = renderer.Render(_file, _index, LookupTable, next).html; }
                finally { _file.Position = nextPosition; }
                text.Append(' ');
            }
            return text.ToString().Trim();
        }
        finally { _isReadingBookmarkText = false; }
    }

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            CoverKnownBlocks(firstBlockIndex + (long)Math.Max(0, count));
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try { return _blockRenderGuard.Render(firstBlockIndex, count, _index.BlockCount, RenderBlocksLocked, _table.BlockKindOf); }
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
            _fieldContext.CurrentBlock = blockIndex;
            _renderer.ColumnWidthPt = WordDocumentPageSetup.TextWidthPointsAt(Sections, blockIndex);
            _renderer.ColumnLeftPt = WordDocumentPageSetup.ColumnLeftPointsAt(Sections, blockIndex);
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
