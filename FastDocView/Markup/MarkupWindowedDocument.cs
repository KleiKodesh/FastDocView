using System.Text;
using FastDocView.Html;

namespace FastDocView.Markup;

/// <summary>
/// A plain-text file (read as HTML: tags render, h1-h6 become the outline, one block per line), a
/// Word web archive (MHTML: the HTML part, one block per paragraph-level element) behind the same
/// <see cref="IWordWindowedDocument"/> contract as the Word formats. Like a flat .fodt the file itself
/// is the random-access view: the index pass records block offsets (and, for MHTML, transfer-decoder
/// resume points); a window reads only its own bytes. The index is a sidecar when the file is big.
/// A sidecar first open with <see cref="WordIndexPolicy.IndexesOnDemand"/> returns after the first
/// <see cref="OnDemandFirstBlockCount"/> blocks; <see cref="ScanAhead"/> feeds the same one pass on
/// (<see cref="MarkupIncrementalIndexBuilder"/>), windows read the known blocks from the partial index, and the finished
/// index is saved when the pass ends.
/// </summary>
public sealed class MarkupWindowedDocument : IWordWindowedDocument
{
    /// <summary>Text and web archives carry no page setup: Word's default text width.</summary>
    public int? FindBookmark(string name) => null;     // bookmarks not indexed for this format yet
    /// <summary>Plain text has no pictures; a web archive's pictures are not shown yet (its blocks carry none).</summary>
    public WordDocumentPicture? ReadPicture(string pictureId) => null;
    public IReadOnlyList<WordDocumentSection> Sections => WordDocumentPageSetup.DefaultSections;

    private const int ReadChunkBytes = 1 << 16;

    public const int OnDemandFirstBlockCount = 100;

    private MarkupRandomAccessIndex _indexField;
    private MarkupCompactBlockTable _tableField;
    private MarkupIncrementalIndexBuilder? _builder;       // null once the index is complete
    private bool _isIndexStale;                            // the builder knows more blocks than _indexField
    private readonly string? _sidecarPathToSave;
    private readonly Stream _file;
    private readonly Encoding _encoding;
    private WordDocumentBlockRenderGuard? _blockRenderGuardField;
    private WordDocumentBlockRenderGuard _blockRenderGuard => _blockRenderGuardField ??= new WordDocumentBlockRenderGuard(FilePath);
    private readonly object _readLock = new();
    private WordDocumentFileStream? _releasableFile;

    public string FilePath { get; }
    /// <summary>The blocks known so far while an on-demand index is scanning (only grows), exact once complete.</summary>
    public int BlockCount => KnownBlockCount;
    public bool IsIndexComplete { get { lock (_readLock) return _builder is null; } }
    public int KnownBlockCount { get { lock (_readLock) return _builder?.KnownBlockCount ?? _indexField.BlockCount; } }
    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;
    public bool IsRightToLeft { get; }                     // fixed at open (an on-demand open: from the first 64 KB of text)
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }
    /// <summary>A .txt in the Torat Emet book format: its metadata (block 0 is the hidden head, drawn as a header, when it has metadata); null for any other text.</summary>
    public ToratEmetHead? ToratEmet => _indexField.ToratEmet;
    /// <summary>A Torat Emet book's tab title: its first title heading (full title), else ForcedBookName; null when it has neither (the file name stays).</summary>
    public string? DocumentTitle => _indexField.ToratEmet is { } head ? ToratEmetFormat.TitleOf(head, Headings) : null;

    /// <summary>The finished index. Throws while an on-demand open is still scanning.</summary>
    public MarkupRandomAccessIndex Index { get { lock (_readLock) return _builder is null ? _indexField : throw new InvalidOperationException("The index is still being built on demand."); } }

    /// <summary>The index of the blocks known so far (a partial index is re-made only after the scan moved on).</summary>
    private MarkupRandomAccessIndex _index
    {
        get
        {
            if (_isIndexStale && IsRefreshDue()) RefreshIndex(_builder!.CreatePartialIndex());
            return _indexField;
        }
    }

    private MarkupCompactBlockTable _table
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

    private void RefreshIndex(MarkupRandomAccessIndex index)
    {
        _isIndexStale = false;
        _refreshedAtTicks = System.Diagnostics.Stopwatch.GetTimestamp();
        _indexField = index;
        _tableField = new MarkupCompactBlockTable(index, _file);
    }

    /// <summary>Feeds the index pass on by up to <paramref name="blocks"/> blocks (nothing rendered); saves the sidecar when it ends.</summary>
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
        MarkupRandomAccessIndex index = _builder!.Complete();
        _builder = null;
        RefreshIndex(index);
        if (_sidecarPathToSave is null) return;
        try { index.Save(_sidecarPathToSave); }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // the sidecar is a cache
    }
    public Encoding TextEncoding => _encoding;

    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy, MarkupSourceKind kind)
    {
        bool hasCachedIndex = MarkupRandomAccessIndex.TryLoad(policy.IndexPathFor(filePath), filePath, kind) is not null;
        return policy.Decide(filePath, new FileInfo(filePath).Length, hasCachedIndex, WordIndexPolicy.PlainTextSidecarThresholdBytes);
    }

    public static MarkupWindowedDocument Open(string filePath, MarkupSourceKind kind, WordIndexPolicy? policy = null,
                                              IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        policy ??= WordIndexPolicy.Default;
        var info = new FileInfo(filePath);
        string indexPath = policy.IndexPathFor(filePath);
        MarkupRandomAccessIndex? cached = MarkupRandomAccessIndex.TryLoad(indexPath, filePath, kind);
        var file = new WordDocumentFileStream(filePath);
        try
        {
            WordIndexDecision decision = policy.Decide(filePath, file.Length, cached is not null, WordIndexPolicy.PlainTextSidecarThresholdBytes);
            bool isOnDemand = decision.Mode == WordIndexMode.Sidecar && cached is null && policy.IndexesOnDemand;
            MarkupWindowedDocument document = isOnDemand
                ? OpenOnDemand(file, filePath, kind, decision, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation, indexPath)
                : OpenFromStream(file, filePath, kind, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (!isOnDemand && decision.Mode == WordIndexMode.Sidecar && cached is null)
            {
                try { document._index.Save(indexPath); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // the sidecar is a cache
            }
            document._releasableFile = file;
            file.Release();                                    // no handle between windows
            return document;
        }
        catch
        {
            file.Dispose();
            throw;
        }
    }

    /// <summary>Host-agnostic entry: <paramref name="file"/> must be seekable; the document owns it.</summary>
    public static MarkupWindowedDocument OpenFromStream(Stream file, string displayPath, MarkupSourceKind kind, WordIndexDecision decision,
                                                        MarkupRandomAccessIndex? cachedIndex, long sourceLength, long sourceLastWriteTicksUtc,
                                                        IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        MarkupRandomAccessIndex index = cachedIndex is not null && cachedIndex.Kind == kind && cachedIndex.MatchesSource(sourceLength, sourceLastWriteTicksUtc)
            ? cachedIndex
            : kind switch
            {
                MarkupSourceKind.PlainText => MarkupRandomAccessIndexBuilder.BuildPlainText(file, sourceLength, sourceLastWriteTicksUtc, progress, cancellation, ToratEmetPathRule.Applies(displayPath)),
                _ => MarkupRandomAccessIndexBuilder.BuildWebArchive(file, sourceLength, sourceLastWriteTicksUtc, progress, cancellation),
            };
        return new MarkupWindowedDocument(displayPath, file, index, null, decision, sidecarPathToSave: null);
    }

    /// <summary>
    /// Returns once the first blocks are known (MHTML: and the 64 KB of body text the direction is read from); the rest of the
    /// index is fed by <see cref="ScanAhead"/>.
    /// </summary>
    private static MarkupWindowedDocument OpenOnDemand(Stream file, string displayPath, MarkupSourceKind kind, WordIndexDecision decision, long sourceLength,
                                                       long sourceLastWriteTicksUtc, IProgress<WordDocumentIndexProgress>? progress,
                                                       CancellationToken cancellation, string sidecarPathToSave)
    {
        var builder = new MarkupIncrementalIndexBuilder(kind, file, sourceLength, sourceLastWriteTicksUtc, progress, cancellation,
                                                        isToratEmetByPath: kind == MarkupSourceKind.PlainText && ToratEmetPathRule.Applies(displayPath));
        builder.AdvanceUntilKnown(OnDemandFirstBlockCount);
        while (!builder.HasDirectionSample) builder.AdvanceUntilKnown(builder.KnownBlockCount + 1);
        if (builder.IsEveryBlockKnown)
        {
            MarkupRandomAccessIndex whole = builder.Complete();
            try { whole.Save(sidecarPathToSave); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
            return new MarkupWindowedDocument(displayPath, file, whole, null, decision, sidecarPathToSave: null);
        }
        builder.ClearCancellation();       // the open's cancellation stops the open only; later steps run under the host's own calls
        return new MarkupWindowedDocument(displayPath, file, builder.CreatePartialIndex(), builder, decision, sidecarPathToSave);
    }

    private MarkupWindowedDocument(string filePath, Stream file, MarkupRandomAccessIndex index, MarkupIncrementalIndexBuilder? builder,
                                   WordIndexDecision decision, string? sidecarPathToSave)
    {
        FilePath = filePath;
        _file = file;
        _builder = builder;
        _sidecarPathToSave = sidecarPathToSave;
        _indexField = index;
        _tableField = new MarkupCompactBlockTable(index, file);
        _encoding = MarkupTextEncoding.FromCodePage(index.CodePage);
        Decision = decision;
        IsRightToLeft = index.IsRightToLeft;
        Stylesheet = WordDocumentHtmlStylesheet.Build(new WordRunFormatting(), index.IsRightToLeft)
                     + "\n.word-document .text-line { white-space: pre-wrap; overflow-wrap: anywhere; min-height: 1em; }"
                     + "\n.word-document .image-placeholder { color: #999; font-style: italic; }\n"
                     + (index.ToratEmet is null ? "" : ToratEmetFormat.Stylesheet)
                     + index.DocumentCss;
    }

    public int EstimatedTextLength(int blockIndex)
    {
        lock (_readLock) return _table.KnownTextLength(blockIndex) ?? _index.EstimatedTextLength(blockIndex);
    }

    /// <summary>A block's text range (text coordinates) and the element stack open at its start - rescanned from its span.</summary>
    public long BlockStartOf(int blockIndex) => WithTable(() => _table.BlockStart(blockIndex));
    public long BlockEndOf(int blockIndex) => WithTable(() => _table.BlockEnd(blockIndex));
    public string OpenTagsOf(int blockIndex) => WithTable(() => _table.OpenTags(blockIndex));
    public int StackIdOf(int blockIndex) => WithTable(() => _table.StackId(blockIndex));

    private T WithTable<T>(Func<T> read)
    {
        lock (_readLock)
        {
            try { return read(); }
            finally { _releasableFile?.Release(); }
        }
    }

    /// <summary>
    /// Scans blocks [fromBlock, fromBlock + maxBlocksToScan) with one read of their bytes (an MHTML decode from the first
    /// block's resume point). A .txt line is matched as its characters; an MHTML block is HTML, walked exactly as the
    /// app's search walks a line's HTML (tags skipped, an entity one character). No rendering; nothing kept.
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
                var hits = new List<int>();
                var counts = new List<int>();
                if (fromBlock < end)
                {
                    byte[] text = ReadText(fromBlock, end, out long textStart);
                    for (int block = fromBlock; block < end; block++)
                    {
                        string content = DecodeBlock(text, textStart, block);
                        matcher.Reset();
                        if (_index.ToratEmet is not null && _index.Kind == MarkupSourceKind.PlainText)
                        {
                            if (block > 0 || !_index.ToratEmet.HasHead) matcher.FeedHtml(ToratEmetFormat.WithoutMark(content));     // block 0 is the hidden metadata head
                        }
                        else if (_index.Kind == MarkupSourceKind.PlainText) matcher.Feed(content.AsSpan()); else matcher.FeedHtml(content);
                        if (matcher.Count > 0) { hits.Add(block); counts.Add(matcher.Count); }
                    }
                }
                return WordDocumentSearchBatch.Finish(hits, counts, Math.Max(end, fromBlock), isIndexComplete: _builder is null, _index.BlockCount);
            }
            finally { _releasableFile?.Release(); }
        }
    }

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            CoverKnownBlocks(firstBlockIndex + (long)Math.Max(0, count));
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try { return _blockRenderGuard.Render(firstBlockIndex, count, _index.BlockCount, RenderBlocksLocked, _ => WordDocumentBlockKind.Paragraph); }
            finally { _releasableFile?.Release(); }
        }
    }

    private List<WordDocumentBlockHtml> RenderBlocksLocked(int firstBlockIndex, int count)
    {
        int first = Math.Max(0, firstBlockIndex);
        int last = Math.Min(_index.BlockCount, firstBlockIndex + Math.Max(0, count));     // exclusive
        var result = new List<WordDocumentBlockHtml>(Math.Max(0, last - first));
        if (first >= last) return result;
        byte[] text = ReadText(first, last, out long textStart);
        for (int block = first; block < last; block++)
        {
            string content = DecodeBlock(text, textStart, block);
            string openTags = _table.OpenTags(block);
            string html = _index.Kind != MarkupSourceKind.PlainText ? MarkupBlockRenderer.Render(openTags, content)
                        : _index.ToratEmet is { } toratEmet ? (block == 0 && toratEmet.HasHead ? ToratEmetFormat.RenderHeader(toratEmet, FirstTitle(_index)) : ToratEmetFormat.RenderLine(openTags, content))
                        : MarkupBlockRenderer.RenderLine(openTags, content);
            WordDocumentBlockKind kind = content.TrimStart().StartsWith("<table", StringComparison.OrdinalIgnoreCase) ? WordDocumentBlockKind.Table : WordDocumentBlockKind.Paragraph;
            result.Add(new WordDocumentBlockHtml(block, kind, html, []));
        }
        return result;
    }

    /// <summary>The characters of one block: its bytes, without the note lines of a Torat Emet book that lie in them (the header shows those; the scanner gave them no block).</summary>
    private string DecodeBlock(byte[] text, long textStart, int block)
    {
        long blockStart = _table.BlockStart(block), blockEnd = _table.BlockEnd(block);
        int from = (int)(blockStart - textStart);
        int length = Math.Min((int)(blockEnd - blockStart), text.Length - from);
        if (_index.ToratEmet is not { Notes.Count: > 0 } toratEmet || _index.Kind != MarkupSourceKind.PlainText || blockStart >= toratEmet.Notes[toratEmet.Notes.Count - 1].End)
            return _encoding.GetString(text, from, length);
        var content = new StringBuilder();
        int at = from, end = from + length;
        foreach (ToratEmetNote note in toratEmet.Notes)
        {
            int noteStart = (int)Math.Min(end, Math.Max(at, note.Start - textStart)), noteEnd = (int)Math.Min(end, Math.Max(at, note.End - textStart));
            if (noteEnd == noteStart) continue;
            content.Append(_encoding.GetString(text, at, noteStart - at));
            at = noteEnd;
        }
        return content.Append(_encoding.GetString(text, at, end - at)).ToString();
    }

    private static string? FirstTitle(MarkupRandomAccessIndex index)
    {
        foreach (WordDocumentHeading heading in index.Headings) if (heading.Level == 1) return heading.Text;
        return null;
    }

    /// <summary>
    /// The text bytes of blocks [first, last). The span rescans that find the blocks' bounds pass these very bytes, so they
    /// are scanned in order and copied from what the scans collected (one read). Only when the bounds were already known
    /// (a span scanned earlier and its bytes no longer held) are they read again: a plain read for .txt, a decode from the
    /// first block's resume point for MHTML.
    /// </summary>
    private byte[] ReadText(int first, int last, out long textStart)
    {
        for (int block = first; block < last; block++) _table.ScanThrough(block);   // in order: the collected range stays one piece
        textStart = _table.BlockStart(first);
        int need = checked((int)(_table.BlockEnd(last - 1) - textStart));
        var output = new byte[need];
        if (_table.TryCopyCollected(textStart, output)) return output;
        if (_index.Kind != MarkupSourceKind.WebArchive)
        {
            _file.Position = textStart;
            int total = 0;
            while (total < need)
            {
                int read = _file.Read(output, total, need - total);
                if (read <= 0) break;
                total += read;
            }
            return output;
        }

        long resume = _table.Resume(first);
        int skip = MimeTransferDecoder.SkipOf(resume);
        long encodedAt = MimeTransferDecoder.EncodedOffsetOf(resume);
        var decoder = new MimeTransferDecoder(_index.TransferEncoding);
        var encoded = new byte[ReadChunkBytes];
        var decoded = new byte[need + skip + ReadChunkBytes + 4];
        int produced = 0;
        _file.Position = encodedAt;
        while (produced < need + skip && encodedAt < _index.EncodedEnd)
        {
            int want = (int)Math.Min(encoded.Length, _index.EncodedEnd - encodedAt);
            int read = _file.Read(encoded, 0, want);
            if (read <= 0) break;
            decoder.Feed(encoded, 0, read, encodedAt, decoded, null, ref produced);
            encodedAt += read;
        }
        Buffer.BlockCopy(decoded, skip, output, 0, Math.Max(0, Math.Min(need, produced - skip)));
        return output;
    }

    public IReadOnlyList<WordDocumentHeading> Headings { get { lock (_readLock) return _index.Headings; } }
    public (int footnotes, int endnotes) NoteCounts => (0, 0);
    public IReadOnlyList<WordDocumentNoteReference> NoteReferences => [];
    public IReadOnlyList<WordDocumentNoteHtml> RenderNotes(int firstNoteIndex, int count) => [];

    public void Dispose() => _file.Dispose();
}
