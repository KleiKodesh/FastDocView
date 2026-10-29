using SeekableDeflate;
using FastDocView.Docx;

namespace FastDocView.Odt;

/// <summary>
/// The index pass of <see cref="OdtRandomAccessIndexBuilder"/>, pulled by the reader instead of run in
/// one go (the OpenDocument twin of <see cref="DocxIncrementalIndexBuilder"/>): it owns the
/// <see cref="OdtBodyBlockScanner"/> and the source of the body — content.xml's inflater for a package,
/// the file itself for a .fodt — and is advanced synchronously, in the caller's thread, one chunk at
/// a time, only as far as a request needs ("until block N is known"). The first open of a big
/// document shows text after a chunk or two; the full pass is paid only when the reader goes to the end.
///
/// A block is known once it has ended (the scanner records its end offset): only then are its byte
/// range, text length, list label and note references final. Notes are inline in the body, so a
/// known block's notes are known too. Same chunks, same checkpoint rule and the same scanner as the
/// one-go build, so <see cref="Complete"/> returns an index identical to it.
/// </summary>
internal sealed class OdtIncrementalIndexBuilder : IOdtBlockTable
{
    private const int FlatChunkSize = OdtRandomAccessIndexBuilder.ChunkSize;   // the one-go flat pass's chunk: same chunks, same index

    private readonly OdtStyleCascade _styles;
    private readonly OdtBodyBlockScanner _scanner;
    private readonly long _sourceLength;
    private readonly long _sourceLastWriteTicksUtc;
    private readonly IProgress<WordDocumentIndexProgress>? _progress;
    private readonly DocxEntryInflationCursor? _contentCursor;     // package
    private readonly Stream? _flatFile;                            // .fodt: shared with the block reader, so every read re-seats its position
    private readonly byte[]? _flatChunk;
    private long _flatPosition;
    private readonly long _totalBytes;
    private long _textLengthSum;
    private int _textLengthSumBlocks;

    public ZipEntryLocation? ContentEntry { get; }
    public IReadOnlyList<DeflateCheckpoint> Checkpoints => (IReadOnlyList<DeflateCheckpoint>?)_contentCursor?.Checkpoints ?? [];
    public bool IsBodyScanned { get; private set; }

    /// <summary>
    /// Every block is known: the body has closed (what follows it in the file is only closing tags) or
    /// the input has ended. <see cref="Complete"/> still reads to the end, so the checkpoints match the one-go build.
    /// </summary>
    public bool IsEveryBlockKnown => IsBodyScanned || _scanner.IsBodyComplete;
    public OdtRandomAccessIndex? CompletedIndex { get; private set; }

    /// <param name="styles">The indexing cascade (styles.xml already added for a package); the scanner adds the sections it captures.</param>
    public OdtIncrementalIndexBuilder(Stream file, bool isFlat, OdtStyleCascade styles, long sourceLength, long sourceLastWriteTicksUtc,
                                      IProgress<WordDocumentIndexProgress>? progress)
    {
        _styles = styles;
        _scanner = new OdtBodyBlockScanner(styles);
        _bookmarks = new WordDocumentBookmarkScanLookup(_scanner.BookmarkNames, _scanner.BookmarkBlocks);
        _sourceLength = sourceLength;
        _sourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        _progress = progress;
        if (isFlat)
        {
            _flatFile = file;
            _flatChunk = new byte[FlatChunkSize];
            _totalBytes = file.Length;
        }
        else
        {
            ContentEntry = OdtPackage.LocateContentEntry(file);
            _contentCursor = new DocxEntryInflationCursor(file, ContentEntry, OdtRandomAccessIndexBuilder.DefaultCheckpointSpanBytes);
            _totalBytes = ContentEntry.UncompressedSize;
        }
    }

    private readonly WordDocumentBookmarkScanLookup _bookmarks;

    /// <summary>A bookmark found so far in a known block, or null.</summary>
    public int? FindKnownBookmark(string name) => _bookmarks.Find(name, KnownBlockCount);

    /// <summary>Scans one more chunk (false once the input is scanned) - lets a lookup advance only as far as it needs.</summary>
    public bool AdvanceOne()
    {
        if (IsBodyScanned) return false;
        AdvanceOneChunk();
        return true;
    }

    private long BytesScanned => _contentCursor?.BytesInflated ?? _flatPosition;

    /// <summary>Blocks that have ended: all but one still being scanned (only the last block can be open).</summary>
    public int KnownBlockCount
    {
        get
        {
            List<long> ends = _scanner.BlockEndOffsets;
            return ends.Count > 0 && ends[ends.Count - 1] < 0 ? ends.Count - 1 : ends.Count;
        }
    }


    /// <summary>Average text length of the known blocks — the stand-in for blocks not reached yet.</summary>
    public int AverageTextLength
    {
        get
        {
            int known = KnownBlockCount;
            if (known == 0) return 0;
            for (; _textLengthSumBlocks < known; _textLengthSumBlocks++) _textLengthSum += _scanner.BlockTextLengths[_textLengthSumBlocks];
            return (int)(_textLengthSum / known);
        }
    }

    /// <summary>The style sections captured so far (font faces, automatic styles: they precede office:body).</summary>
    public System.Xml.Linq.XElement[] StyleSectionElementsForRendering() => _styles.SectionElementsForRendering();

    /// <summary>Scans on until at least <paramref name="blockCount"/> blocks are known or the body ends.</summary>
    public void AdvanceUntilKnown(int blockCount)
    {
        while (!IsBodyScanned && KnownBlockCount < blockCount) AdvanceOneChunk();
    }

    /// <summary>Scans the rest and builds the finished index — identical to the one-go build (idempotent).</summary>
    public OdtRandomAccessIndex Complete()
    {
        if (CompletedIndex is not null) return CompletedIndex;
        while (!IsBodyScanned) AdvanceOneChunk();
        _progress?.Report(new WordDocumentIndexProgress(_totalBytes, _totalBytes, _scanner.BlockStartOffsets.Count));
        CompletedIndex = OdtRandomAccessIndexBuilder.Finish(_scanner, _styles, _sourceLength, _sourceLastWriteTicksUtc, _flatFile is not null,
                                                            ContentEntry, _contentCursor?.Checkpoints ?? []);
        return CompletedIndex;
    }

    private void AdvanceOneChunk()
    {
        int read;
        if (_contentCursor is not null)
        {
            read = _contentCursor.ReadNextChunk();
            if (read > 0) _scanner.Feed(_contentCursor.Chunk, read);
        }
        else
        {
            _flatFile!.Position = _flatPosition;
            read = _flatFile.Read(_flatChunk!, 0, _flatChunk!.Length);
            if (read > 0) { _scanner.Feed(_flatChunk, read); _flatPosition += read; }
        }
        if (read <= 0) { IsBodyScanned = true; return; }
        _progress?.Report(new WordDocumentIndexProgress(BytesScanned, _totalBytes, _scanner.BlockStartOffsets.Count));
    }

    /// <summary>Headings of the known blocks (body order).</summary>
    public IReadOnlyList<WordDocumentHeading> KnownHeadings
    {
        get
        {
            int known = KnownBlockCount;
            List<WordDocumentHeading> headings = _scanner.Headings;
            int count = headings.Count;
            while (count > 0 && headings[count - 1].BlockIndex >= known) count--;
            return count == headings.Count ? headings : headings.GetRange(0, count);
        }
    }

    /// <summary>Sections starting in the known blocks.</summary>
    public IReadOnlyList<WordDocumentSection> KnownSectionStarts
    {
        get
        {
            int known = KnownBlockCount;
            List<WordDocumentSection> sections = _scanner.SectionStarts;
            int count = sections.Count;
            while (count > 1 && sections[count - 1].FirstBlockIndex >= known) count--;
            return sections.GetRange(0, count);
        }
    }

    /// <summary>Note references in the known blocks (body order).</summary>
    public IReadOnlyList<OdtNoteReference> KnownNoteReferences
    {
        get
        {
            int known = KnownBlockCount;
            List<OdtNoteReference> references = _scanner.NoteReferences;
            int count = references.Count;
            while (count > 0 && references[count - 1].BlockIndex >= known) count--;
            return count == references.Count ? references : references.GetRange(0, count);
        }
    }

    // ── IOdtBlockTable over the known blocks ──────────────────────────────────────────

    int IOdtBlockTable.BlockCount => KnownBlockCount;
    public string RootOpenTag => KnownBlockCount > 0 ? _scanner.RootOpenTag : throw new InvalidOperationException("No block scanned yet.");
    public long BlockStartOffset(int blockIndex) => _scanner.BlockStartOffsets[blockIndex];
    public long BlockEndOffset(int blockIndex) => _scanner.BlockEndOffsets[blockIndex];
    public WordDocumentBlockKind BlockKindOf(int blockIndex) => _scanner.BlockKinds[blockIndex];
    public int BlockTextLengthOf(int blockIndex) => _scanner.BlockTextLengths[blockIndex];

    public string? ListLabelOf(int blockIndex)
    {
        int labelIndex = _scanner.BlockListLabelIds[blockIndex];
        return labelIndex >= 0 ? _scanner.ListLabels[labelIndex] : null;
    }

    public string ListLabelLayoutOf(int blockIndex)
    {
        int labelIndex = _scanner.BlockListLabelIds[blockIndex];
        return labelIndex >= 0 ? _scanner.ListLabelLayouts[labelIndex] : "";
    }

    public int ListLevelOf(int blockIndex) => _scanner.BlockListLevels[blockIndex] == 255 ? -1 : _scanner.BlockListLevels[blockIndex];

    public ReadOnlySpan<OdtNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex)
    {
        List<OdtNoteReference> references = _scanner.NoteReferences;
        int start = LowerBound(references, firstBlockIndex), end = LowerBound(references, lastBlockIndex + 1);
        return end > start ? references.GetRange(start, end - start).ToArray() : ReadOnlySpan<OdtNoteReference>.Empty;
    }

    private static int LowerBound(List<OdtNoteReference> references, int blockIndex)
    {
        int low = 0, high = references.Count;
        while (low < high)
        {
            int middle = (low + high) / 2;
            if (references[middle].BlockIndex < blockIndex) low = middle + 1; else high = middle;
        }
        return low;
    }
}
