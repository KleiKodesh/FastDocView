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
                var search = new RtfSpanSearch(matcher, options.IncludeNotes) { FieldContext = _fieldContext };
                _fieldContext.CurrentBlock = fromBlock;
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
                    _fieldContext.CurrentBlock = blockIndex;
                    _searchSink.Feed(_file, _table, blockIndex, matcher, options.IncludeNotes);
                    if (matcher.Count > 0) { hits.Add(blockIndex); counts.Add(matcher.Count); }
                }
                return WordDocumentSearchBatch.Finish(hits, counts, Math.Max(end, fromBlock), isIndexComplete: true, _index.BlockCount);
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
            BookmarkBlock = FindBookmark,
            ListLabelOfBlock = block => block >= 0 && block < _index.BlockCount ? LookupTable.ListLabelOf(block) : null,
            ListContextOfBlock = ListContext,
            IsRightToLeftText = _index.IsRightToLeft,
            NoteMarkOfBookmark = name => FindBookmark(name) is { } block && _index.NoteReferencesInBlock(block) is { Length: > 0 } notes ? notes[0].Label : null,
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
            var renderer = new RtfBlockHtmlRenderer(_index.Tables) { FieldContext = null };
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
        if (_isReadingBookmarkText || FindBookmark(name) is not { } block) return null;
        _isReadingBookmarkText = true;
        try
        {
            // its own renderer: this runs while the document's renderer is mid-block
            var renderer = new RtfBlockHtmlRenderer(_index.Tables) { FieldContext = _fieldContext };
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
            _fieldContext.CurrentBlock = blockIndex;
            _renderer.ColumnWidthPt = WordDocumentPageSetup.TextWidthPointsAt(Sections, blockIndex);
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
