using System.Text;
using FastDocView.Html;

namespace FastDocView.Markup;

/// <summary>
/// A plain-text file (read as HTML: tags render, h1-h6 become the outline, one block per line), a
/// Word web archive (MHTML: the HTML part, one block per paragraph-level element) behind the same
/// <see cref="IWordWindowedDocument"/> contract as the Word formats. Like a flat .fodt the file itself
/// is the random-access view: the index pass records block offsets (and, for MHTML, transfer-decoder
/// resume points); a window reads only its own bytes. The index is a sidecar when the file is big.
/// </summary>
public sealed class MarkupWindowedDocument : IWordWindowedDocument
{
    /// <summary>Text and web archives carry no page setup: Word's default text width.</summary>
    public int? FindBookmark(string name) => null;     // bookmarks not indexed for this format yet
    public IReadOnlyList<WordDocumentSection> Sections => WordDocumentPageSetup.DefaultSections;

    private const int ReadChunkBytes = 1 << 16;

    private readonly MarkupRandomAccessIndex _index;
    private readonly MarkupCompactBlockTable _table;
    private readonly Stream _file;
    private readonly Encoding _encoding;
    private readonly object _readLock = new();
    private WordDocumentFileStream? _releasableFile;

    public string FilePath { get; }
    public int BlockCount => _index.BlockCount;
    public bool IsIndexComplete => true;
    public int KnownBlockCount => _index.BlockCount;
    public int ScanAhead(int blocks) => BlockCount;          // always complete
    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;
    public bool IsRightToLeft => _index.IsRightToLeft;
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }
    public MarkupRandomAccessIndex Index => _index;
    public Encoding TextEncoding => _encoding;

    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy, MarkupSourceKind kind)
    {
        bool hasCachedIndex = MarkupRandomAccessIndex.TryLoad(policy.IndexPathFor(filePath), filePath, kind) is not null;
        return policy.Decide(filePath, new FileInfo(filePath).Length, hasCachedIndex);
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
            WordIndexDecision decision = policy.Decide(filePath, file.Length, cached is not null);
            MarkupWindowedDocument document = OpenFromStream(file, filePath, kind, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (decision.Mode == WordIndexMode.Sidecar && cached is null)
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
                MarkupSourceKind.PlainText => MarkupRandomAccessIndexBuilder.BuildPlainText(file, sourceLength, sourceLastWriteTicksUtc, progress, cancellation),
                _ => MarkupRandomAccessIndexBuilder.BuildWebArchive(file, sourceLength, sourceLastWriteTicksUtc, progress, cancellation),
            };
        return new MarkupWindowedDocument(displayPath, file, index, decision);
    }

    private MarkupWindowedDocument(string filePath, Stream file, MarkupRandomAccessIndex index, WordIndexDecision decision)
    {
        FilePath = filePath;
        _file = file;
        _index = index;
        _table = new MarkupCompactBlockTable(index, file);
        _encoding = MarkupTextEncoding.FromCodePage(index.CodePage);
        Decision = decision;
        Stylesheet = WordDocumentHtmlStylesheet.Build(new WordRunFormatting(), index.IsRightToLeft)
                     + "\n.word-document .text-line { white-space: pre-wrap; overflow-wrap: anywhere; min-height: 1em; }"
                     + "\n.word-document .image-placeholder { color: #999; font-style: italic; }\n"
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
                        long blockStart = _table.BlockStart(block);
                        int from = (int)(blockStart - textStart);
                        int length = (int)(_table.BlockEnd(block) - blockStart);
                        string content = _encoding.GetString(text, from, Math.Min(length, text.Length - from));
                        matcher.Reset();
                        if (_index.Kind == MarkupSourceKind.PlainText) matcher.Feed(content.AsSpan()); else matcher.FeedHtml(content);
                        if (matcher.Count > 0) { hits.Add(block); counts.Add(matcher.Count); }
                    }
                }
                return WordDocumentSearchBatch.Finish(hits, counts, Math.Max(end, fromBlock), isIndexComplete: true, _index.BlockCount);
            }
            finally { _releasableFile?.Release(); }
        }
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
        int last = Math.Min(_index.BlockCount, firstBlockIndex + Math.Max(0, count));     // exclusive
        var result = new List<WordDocumentBlockHtml>(Math.Max(0, last - first));
        if (first >= last) return result;
        byte[] text = ReadText(first, last, out long textStart);
        for (int block = first; block < last; block++)
        {
            long blockStart = _table.BlockStart(block);
            int from = (int)(blockStart - textStart);
            int length = (int)(_table.BlockEnd(block) - blockStart);
            string content = _encoding.GetString(text, from, Math.Min(length, text.Length - from));
            string openTags = _table.OpenTags(block);
            string html = _index.Kind == MarkupSourceKind.PlainText ? MarkupBlockRenderer.RenderLine(openTags, content) : MarkupBlockRenderer.Render(openTags, content);
            WordDocumentBlockKind kind = content.TrimStart().StartsWith("<table", StringComparison.OrdinalIgnoreCase) ? WordDocumentBlockKind.Table : WordDocumentBlockKind.Paragraph;
            result.Add(new WordDocumentBlockHtml(block, kind, html, []));
        }
        return result;
    }

    /// <summary>The text bytes of blocks [first, last): a plain read for .txt, a decode from the first block's resume point for MHTML.</summary>
    private byte[] ReadText(int first, int last, out long textStart)
    {
        textStart = _table.BlockStart(first);
        int need = checked((int)(_table.BlockEnd(last - 1) - textStart));
        var output = new byte[need];
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

    public IReadOnlyList<WordDocumentHeading> Headings => _index.Headings;
    public (int footnotes, int endnotes) NoteCounts => (0, 0);
    public IReadOnlyList<WordDocumentNoteReference> NoteReferences => [];
    public IReadOnlyList<WordDocumentNoteHtml> RenderNotes(int firstNoteIndex, int count) => [];

    public void Dispose() => _file.Dispose();
}
