using System.Text;

namespace FastDocView.Markup;

/// <summary>
/// The rescan of one span of a compact <see cref="MarkupRandomAccessIndex"/>, run <b>only as far as a reader needs</b>:
/// the index pass's scanner, seeded with the anchor's snapshot, is fed exactly the pieces the pass fed (.txt: the same
/// 64 KB boundaries; MHTML: 4 KB of encoded bytes decoded from the anchor's resume point), and it pauses as soon as the
/// asked-for block is complete (the next block has started). A window at a span's start therefore scans a few KB, not
/// the whole span. The scan reaches the span's end (the next anchor's position; the last span: the text's end) only
/// when a block there is asked for; then the block count is checked against the index.
/// </summary>
internal sealed class MarkupSpanRescan
{
    private readonly MarkupRandomAccessIndex _index;
    private readonly Stream _file;
    private readonly int _firstBlock;
    private readonly int _blockCount;                  // the span's blocks
    private readonly bool _isLast;
    private readonly long _limit;                      // text position where the next span's scan begins
    private readonly WordIndexByteRing? _collected;    // the text bytes the scan passes, kept for the window that asked (no second read)

    public MarkupBlockScanner Scanner { get; }
    /// <summary>MHTML: each block's decoder resume point, aligned with the scanner's blocks; null for .txt.</summary>
    public List<long>? Resumes { get; }
    /// <summary>The whole span has been scanned: every block's values are final.</summary>
    public bool IsComplete { get; private set; }

    // .txt
    private byte[]? _chunk;
    private long _position;

    // MHTML
    private readonly MimeTransferDecoder? _decoder;
    private long[]? _resumeRing;
    private byte[]? _encoded, _decoded;
    private long[]? _resumePiece;
    private int _skip;
    private long _encodedAt;
    private long _decodedTotal;

    public MarkupSpanRescan(Stream file, MarkupRandomAccessIndex index, int anchorIndex, WordIndexByteRing? collected = null)
    {
        _index = index;
        _file = file;
        _collected = collected;
        MarkupSpanAnchor[] anchors = index.Anchors;
        MarkupSpanAnchor anchor = anchors[anchorIndex];
        _firstBlock = anchor.FirstBlock;
        bool isLastAnchor = anchorIndex + 1 >= anchors.Length;
        _isLast = isLastAnchor && !index.IsPartial;      // a partial index's last span ends like any other: past its last block's end
        _blockCount = (isLastAnchor ? index.BlockCount : anchors[anchorIndex + 1].FirstBlock) - _firstBlock;
        (int codePage, int unitSize, Encoding encoding) = MarkupRandomAccessIndexBuilder.PlainTextEncoding(index.CodePage);   // the same unit rules for MHTML
        bool bigEndian = codePage == MarkupTextEncoding.Utf16BigEndian;
        MarkupBlockMode mode = index.Kind == MarkupSourceKind.WebArchive ? MarkupBlockMode.Elements : MarkupBlockMode.Lines;
        Scanner = anchor.State is null
            ? MarkupBlockScanner.FromStart(mode, unitSize, bigEndian, encoding, index.Stacks, index.StackIds, index.ToratEmet)
            : MarkupBlockScanner.ResumeSpan(mode, unitSize, bigEndian, encoding, index.Stacks, index.StackIds, anchor.State, index.ToratEmet);

        if (index.Kind == MarkupSourceKind.WebArchive)
        {
            _limit = _isLast ? long.MaxValue : isLastAnchor ? index.TextEnd : anchors[anchorIndex + 1].ScanPosition;
            Resumes = anchor.State is null ? [] : [anchor.BlockResume];
            _decoder = new MimeTransferDecoder(index.TransferEncoding);
            _resumeRing = new long[MarkupRandomAccessIndexBuilder.ResumeRingSize];
            _encoded = new byte[MarkupRandomAccessIndexBuilder.FeedPieceBytes];
            _decoded = new byte[MarkupRandomAccessIndexBuilder.FeedPieceBytes + 4];
            _resumePiece = new long[MarkupRandomAccessIndexBuilder.FeedPieceBytes + 4];
            _skip = MimeTransferDecoder.SkipOf(anchor.ScanResume);
            _encodedAt = MimeTransferDecoder.EncodedOffsetOf(anchor.ScanResume);
            _decodedTotal = anchor.ScanPosition;
        }
        else
        {
            _limit = isLastAnchor ? index.TextEnd : anchors[anchorIndex + 1].ScanPosition;
            if (anchor.State is null) Scanner.Begin(index.TextStart);
            _chunk = new byte[Math.Max(MarkupRandomAccessIndexBuilder.ChunkSize, MarkupTextEncoding.DetectionSampleBytes)];   // span 0 starts with the detection sample piece
            _position = anchor.ScanPosition;
        }
        if (Scanner.BlockIndexBase != _firstBlock) throw Changed();
    }

    /// <summary>True when the span's block <paramref name="local"/> has its final values (the next block has started).</summary>
    public bool IsBlockComplete(int local) => IsComplete || local + 1 < Scanner.BlockStartOffsets.Count;

    /// <summary>Scans on until block <paramref name="local"/> of the span is complete.</summary>
    public void ScanThrough(int local)
    {
        if ((uint)local >= (uint)_blockCount) throw new ArgumentOutOfRangeException(nameof(local));
        while (!IsBlockComplete(local))
            if (!(_index.Kind == MarkupSourceKind.WebArchive ? FeedWebArchivePiece() : FeedPlainTextPiece())) Complete();
    }

    private bool FeedPlainTextPiece()
    {
        if (_position >= _limit) return false;
        long pieceEnd = _position < _index.FirstFeedEnd ? _index.FirstFeedEnd : _position + MarkupRandomAccessIndexBuilder.ChunkSize;
        int want = (int)(Math.Min(pieceEnd, _index.TextEnd) - _position);
        if (want <= 0) return false;
        _file.Position = _position;
        int read = MarkupRandomAccessIndexBuilder.ReadFully(_file, _chunk!, want);
        if (read <= 0) return false;
        Scanner.Feed(_chunk!, 0, read, _position);
        _collected?.Retain(_position, _chunk!, 0, read);
        _position += read;
        if (read < want) _position = long.MaxValue;   // the file ended early: nothing more to feed
        return true;
    }

    private bool FeedWebArchivePiece()
    {
        while (true)
        {
            if (_decodedTotal >= _limit || Scanner.IsBodyEnded || _encodedAt >= _index.EncodedEnd) return false;
            int want = (int)Math.Min(_encoded!.Length, _index.EncodedEnd - _encodedAt);
            _file.Position = _encodedAt;
            int read = _file.Read(_encoded, 0, want);
            if (read <= 0) return false;
            int produced = 0;
            _decoder!.Feed(_encoded, 0, read, _encodedAt, _decoded!, _resumePiece, ref produced);
            _encodedAt += read;
            int from = Math.Min(_skip, produced);
            _skip -= from;
            int usable = produced - from;
            if (usable <= 0) continue;
            long[] ring = _resumeRing!;
            for (int k = 0; k < usable; k++) ring[(_decodedTotal + k) & (ring.Length - 1)] = _resumePiece![from + k];
            Scanner.Feed(_decoded!, from, usable, _decodedTotal);
            _collected?.Retain(_decodedTotal, _decoded!, from, usable);
            _decodedTotal += usable;
            MarkupRandomAccessIndexBuilder.MapResumePoints(Scanner, Resumes!, ring, _decodedTotal);
            return true;
        }
    }

    /// <summary>The span's end: the last span finishes its scanner; the block count must match the index.</summary>
    private void Complete()
    {
        if (_isLast)
        {
            Scanner.Finish(_index.TextEnd);
            if (Resumes is not null) while (Resumes.Count > Scanner.BlockStartOffsets.Count) Resumes.RemoveAt(Resumes.Count - 1);
        }
        int found = Scanner.BlockStartOffsets.Count;
        // a span's blocks are complete once the next one has started (or the last span finished)
        if (found < _blockCount || (!_isLast && found == _blockCount) || (_isLast && found != _blockCount) || (Resumes is not null && Resumes.Count < _blockCount))
            throw Changed();
        IsComplete = true;
        _chunk = null; _encoded = null; _decoded = null; _resumePiece = null; _resumeRing = null;   // drop the buffers
    }

    private static InvalidDataException Changed() => new("Markup index: the span's rescan found a different block count (the document changed?).");
}
