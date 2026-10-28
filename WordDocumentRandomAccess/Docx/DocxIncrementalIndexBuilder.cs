using WordDocumentRandomAccess.Deflate;
using WordDocumentRandomAccess.Zip;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The sidecar index pass of <see cref="DocxRandomAccessIndexBuilder.Build"/>, pulled by the reader
/// instead of run in one go: it owns the body's inflater + <see cref="DocxBodyBlockScanner"/> and is
/// advanced synchronously, in the caller's thread, one 256 KB chunk at a time, only as far as a
/// request needs ("until block N is known"). The first open of a big document therefore shows text
/// after a chunk or two; the full pass is paid only when the reader goes to the end (or the host
/// calls <see cref="Complete"/>). The notes parts are pulled the same way, as far as the notes
/// referenced so far (<see cref="DocxIncrementalNotesPart"/>).
///
/// While incomplete it serves <see cref="IDocxBlockTable"/> over the blocks known so far — a block
/// is known once the next one has started (or the body has ended): only then are its byte range,
/// text length and heading final. Same chunks, same checkpoints, same scanner as the one-go build,
/// so <see cref="Complete"/> returns an index identical to <see cref="DocxRandomAccessIndexBuilder.Build"/>.
/// </summary>
internal sealed class DocxIncrementalIndexBuilder : IDocxBlockTable
{
    private readonly long _sourceLength;
    private readonly long _sourceLastWriteTicksUtc;
    private readonly IProgress<WordDocumentIndexProgress>? _progress;
    private readonly long _totalBytes;
    private readonly DocxEntryInflationCursor _bodyCursor;
    private readonly DocxBodyBlockScanner _scanner;
    private readonly DocxIncrementalNotesPart? _footnotes;
    private readonly DocxIncrementalNotesPart? _endnotes;
    private bool _isBodyScanned;
    private string? _rootOpenTag;

    public ZipEntryLocation BodyEntry { get; }
    public IReadOnlyList<DeflateCheckpoint> BodyCheckpoints => _bodyCursor.Checkpoints;
    public DocxIncrementalNotesPart? Footnotes => _footnotes;
    public DocxIncrementalNotesPart? Endnotes => _endnotes;

    /// <summary>Set by <see cref="Complete"/>; identical to a one-go build.</summary>
    public DocxRandomAccessIndex? CompletedIndex { get; private set; }

    public DocxIncrementalIndexBuilder(Stream package, DocxDocumentParts parts, long sourceLength, long sourceLastWriteTicksUtc,
                                       IProgress<WordDocumentIndexProgress>? progress = null,
                                       long checkpointSpanBytes = DocxRandomAccessIndexBuilder.DefaultCheckpointSpanBytes)
    {
        _sourceLength = sourceLength;
        _sourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        _progress = progress;
        BodyEntry = DocxRandomAccessIndexBuilder.LocateDocumentEntry(package);
        ZipEntryLocation? footnotes = DocxRandomAccessIndexBuilder.LocateTextEntry(package, DocxRandomAccessIndexBuilder.FootnotesEntryName);
        ZipEntryLocation? endnotes = DocxRandomAccessIndexBuilder.LocateTextEntry(package, DocxRandomAccessIndexBuilder.EndnotesEntryName);
        _totalBytes = BodyEntry.UncompressedSize + (footnotes?.UncompressedSize ?? 0) + (endnotes?.UncompressedSize ?? 0);
        _bodyCursor = new DocxEntryInflationCursor(package, BodyEntry, checkpointSpanBytes);
        _scanner = new DocxBodyBlockScanner(parts.CreateListLabelResolver(), parts.Styles.HeadingLevelOf);
        _footnotes = footnotes is null ? null : new DocxIncrementalNotesPart(package, footnotes, checkpointSpanBytes);
        _endnotes = endnotes is null ? null : new DocxIncrementalNotesPart(package, endnotes, checkpointSpanBytes);
    }

    public bool IsComplete => CompletedIndex is not null;

    /// <summary>Blocks whose range and metadata are final: all but the one still being scanned.</summary>
    public int KnownBlockCount => _scanner.BodyEndOffset >= 0 ? _scanner.BlockStartOffsets.Count : Math.Max(0, _scanner.BlockStartOffsets.Count - 1);

    /// <summary>
    /// The total block count: exact once the body is scanned, otherwise the known count plus the
    /// unscanned body bytes divided by the average block size so far (always more than known).
    /// </summary>
    public int EstimatedBlockCount
    {
        get
        {
            int known = KnownBlockCount;
            if (_isBodyScanned || _scanner.BodyEndOffset >= 0) return known;
            int started = _scanner.BlockStartOffsets.Count;
            if (started == 0) return known + 1;
            double bytesPerBlock = Math.Max(1.0, (double)_bodyCursor.BytesInflated / started);
            long remaining = Math.Max(0, BodyEntry.UncompressedSize - _bodyCursor.BytesInflated);
            return (int)Math.Min(int.MaxValue - 1, Math.Max(known + 1, started + (long)(remaining / bytesPerBlock)));
        }
    }

    /// <summary>Average text length of the blocks known so far — the stand-in for blocks not reached yet.</summary>
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
    private long _textLengthSum;
    private int _textLengthSumBlocks;

    /// <summary>Scans on until at least <paramref name="blockCount"/> blocks are known or the body ends.</summary>
    public void AdvanceUntilKnown(int blockCount)
    {
        while (!_isBodyScanned && KnownBlockCount < blockCount) AdvanceBodyOneChunk();
    }

    /// <summary>Scans the rest of the body and the notes parts and builds the finished index (idempotent).</summary>
    public DocxRandomAccessIndex Complete()
    {
        if (CompletedIndex is not null) return CompletedIndex;
        while (!_isBodyScanned) AdvanceBodyOneChunk();
        DocxNotesPartIndex? footnotesPart = _footnotes?.Complete();
        DocxNotesPartIndex? endnotesPart = _endnotes?.Complete();
        _progress?.Report(new WordDocumentIndexProgress(_totalBytes, _totalBytes, _scanner.BlockStartOffsets.Count));
        CompletedIndex = DocxRandomAccessIndexBuilder.CreateIndex(_sourceLength, _sourceLastWriteTicksUtc, BodyEntry, _scanner,
                                                                  _bodyCursor.Checkpoints, footnotesPart, endnotesPart);
        return CompletedIndex;
    }

    public bool IsBodyScanned => _isBodyScanned;

    private void AdvanceBodyOneChunk()
    {
        int read = _bodyCursor.ReadNextChunk();
        if (read == 0)
        {
            _scanner.Finish();
            if (_scanner.BodyEndOffset < 0) throw new InvalidDataException("word/document.xml has no closing <w:body>.");
            _isBodyScanned = true;
            return;
        }
        _scanner.Feed(_bodyCursor.Chunk, read);
        long done = _bodyCursor.BytesInflated + (_footnotes?.BytesInflated ?? 0) + (_endnotes?.BytesInflated ?? 0);
        _progress?.Report(new WordDocumentIndexProgress(done, _totalBytes, _scanner.BlockStartOffsets.Count));
    }

    /// <summary>Headings of the known blocks (body order).</summary>
    public IReadOnlyList<WordDocumentHeading> KnownHeadings => _scanner.Headings;

    /// <summary>Note references of the known blocks (body order).</summary>
    public IReadOnlyList<DocxNoteReference> KnownNoteReferences => _scanner.NoteReferences;

    // ── IDocxBlockTable over the known blocks ────────────────────────────────────────

    int IDocxBlockTable.BlockCount => KnownBlockCount;

    public string RootOpenTag => _rootOpenTag ??= KnownBlockCount > 0 ? _scanner.RootOpenTag : throw new InvalidOperationException("No block scanned yet.");

    public long BlockStartOffset(int blockIndex) => _scanner.BlockStartOffsets[blockIndex];

    public long BlockEndOffset(int blockIndex)
        => blockIndex + 1 < _scanner.BlockStartOffsets.Count ? _scanner.BlockStartOffsets[blockIndex + 1] : _scanner.BodyEndOffset;

    public WordDocumentBlockKind BlockKindOf(int blockIndex) => _scanner.BlockKinds[blockIndex];

    public int BlockTextLengthOf(int blockIndex) => _scanner.BlockTextLengths[blockIndex];

    public string? StyleIdOf(int blockIndex)
    {
        short tableIndex = _scanner.BlockStyleIds[blockIndex];
        return tableIndex >= 0 ? _scanner.StyleIdTable[tableIndex] : null;
    }

    public string? ListLabelOf(int blockIndex)
    {
        int labelIndex = _scanner.BlockListLabelIds[blockIndex];
        return labelIndex >= 0 ? _scanner.ListLabels[labelIndex] : null;
    }

    public ReadOnlySpan<DocxNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex)
    {
        List<DocxNoteReference> references = _scanner.NoteReferences;
        int start = LowerBound(references, firstBlockIndex), end = LowerBound(references, lastBlockIndex + 1);
        return end > start ? references.GetRange(start, end - start).ToArray() : ReadOnlySpan<DocxNoteReference>.Empty;
    }

    private static int LowerBound(List<DocxNoteReference> references, int blockIndex)
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
