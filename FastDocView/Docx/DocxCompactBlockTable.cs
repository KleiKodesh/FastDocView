namespace FastDocView.Docx;

/// <summary>
/// Serves <see cref="IDocxBlockTable"/> from a compact <see cref="DocxRandomAccessIndex"/>: a block's position,
/// kind, style, list label and text length are recovered by rescanning its span — from the span's anchor
/// (a block start at or after a deflate checkpoint) forward, with the same <see cref="DocxBodyBlockScanner"/>
/// and list counter the index pass ran (seeded with the anchor's numbering state), so every value equals
/// the one the pass found. A span is scanned only as far as asked and kept scanning on later requests;
/// a handful of spans stay cached (the ones around the window). The bytes inflated while scanning are kept
/// in a bounded ring (<see cref="RetainedBytes"/>), so a window that was just located is rendered from them
/// without inflating again; sequential windows therefore cost one inflate pass, like before.
/// </summary>
internal sealed class DocxCompactBlockTable : IDocxBlockTable
{
    private const int MaxCachedSpans = 8;
    private const int ChunkSize = 1 << 16;
    private const int RetainedBytes = 1 << 19;          // holds a whole anchor span + a chunk, so a located window is always still in the ring; grows for a search batch up to MaxRetainedBytes
    private const int MaxRetainedBytes = 4 << 20;

    private sealed class Span
    {
        public int Anchor;
        public DocxBodyBlockScanner Scanner = null!;
        public long ScannedTo;
        public bool IsComplete;
        public long LastUse;
    }

    private readonly DocxRandomAccessIndex _index;
    private readonly Stream _view;
    private readonly Func<DocxDocumentParts.DocxListLabeller> _labellerFactory;
    private readonly Dictionary<int, Span> _spans = new();
    private readonly byte[] _chunk = new byte[ChunkSize];
    private long _useCounter;

    // the last bytes read from the view, contiguous: [_ringStart, _ringStart + _ringLength)
    private byte[] _ring = new byte[RetainedBytes];
    private long _ringStart = -1;
    private int _ringLength;

    public DocxCompactBlockTable(DocxRandomAccessIndex index, Stream view, Func<DocxDocumentParts.DocxListLabeller> labellerFactory)
    {
        _index = index;
        _view = view;
        _labellerFactory = labellerFactory;
    }

    public int BlockCount => _index.BlockCount;
    public string RootOpenTag => _index.RootOpenTag;

    private DocxIndexAnchor[] Anchors => _index.Anchors;

    private long SpanEnd(int anchor) => anchor + 1 < Anchors.Length ? Anchors[anchor + 1].Offset : _index.BodyEndOffset;
    private int SpanBlockCount(int anchor) => (anchor + 1 < Anchors.Length ? Anchors[anchor + 1].FirstBlock : BlockCount) - Anchors[anchor].FirstBlock;

    /// <summary>True when the block is the one just before an anchor and its span is not scanned up to it: the anchor's record answers.</summary>
    private bool IsAnchorPredecessor(int blockIndex, out DocxIndexAnchor anchor)
    {
        int next = _index.AnchorIndexOf(blockIndex) + 1;
        if (next < Anchors.Length && Anchors[next].FirstBlock == blockIndex + 1
            && !(_spans.TryGetValue(next - 1, out Span? span) && span.Scanner.BlockStartOffsets.Count > blockIndex - Anchors[next - 1].FirstBlock))
        {
            anchor = Anchors[next];
            return true;
        }
        anchor = null!;
        return false;
    }

    private Span GetSpan(int anchorIndex)
    {
        if (!_spans.TryGetValue(anchorIndex, out Span? span))
        {
            if (_spans.Count >= MaxCachedSpans)
            {
                int oldest = _spans.OrderBy(pair => pair.Value.LastUse).First().Key;
                _spans.Remove(oldest);
            }
            DocxIndexAnchor anchor = Anchors[anchorIndex];
            DocxDocumentParts.DocxListLabeller labeller = _labellerFactory();
            labeller.ImportState(_index.ListStateOf(anchorIndex));
            var scanner = new DocxBodyBlockScanner(labeller.Resolve, headingLevelOfStyle: null);
            scanner.SeedAtBodyLevel(anchor.Offset);
            span = new Span { Anchor = anchorIndex, Scanner = scanner, ScannedTo = anchor.Offset, IsComplete = anchor.Offset >= SpanEnd(anchorIndex) };
            _spans[anchorIndex] = span;
        }
        span.LastUse = ++_useCounter;
        return span;
    }

    /// <summary>Scans the span until block <paramref name="localIndex"/> has started (its offset and kind are known).</summary>
    private Span EnsureStarted(int blockIndex, out int localIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) throw new ArgumentOutOfRangeException(nameof(blockIndex));
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        localIndex = blockIndex - Anchors[anchorIndex].FirstBlock;
        Span span = GetSpan(anchorIndex);
        while (span.Scanner.BlockStartOffsets.Count <= localIndex)
        {
            if (span.IsComplete) throw new InvalidDataException($"Word index: block {blockIndex} not found in its span (the document changed?).");
            ScanMore(span);
        }
        return span;
    }

    /// <summary>Scans until the block's own content is complete (the next block started, or the span ended): style, label and text length are final.</summary>
    private Span EnsureFinished(int blockIndex, out int localIndex)
    {
        Span span = EnsureStarted(blockIndex, out localIndex);
        if (localIndex + 1 < SpanBlockCount(span.Anchor))
            while (span.Scanner.BlockStartOffsets.Count <= localIndex + 1) { if (span.IsComplete) throw new InvalidDataException("Word index: span shorter than its block count."); ScanMore(span); }
        else
            while (!span.IsComplete) ScanMore(span);
        return span;
    }

    private void ScanMore(Span span)
    {
        long end = SpanEnd(span.Anchor);
        int toRead = (int)Math.Min(ChunkSize, end - span.ScannedTo);
        if (toRead <= 0) { span.IsComplete = true; return; }
        if (_view.Position != span.ScannedTo) _view.Position = span.ScannedTo;
        int read = _view.Read(_chunk, 0, toRead);
        if (read <= 0) throw new EndOfStreamException("Part ended inside a span.");
        Retain(span.ScannedTo, _chunk, read);
        span.Scanner.Feed(_chunk, read);
        span.ScannedTo += read;
        if (span.ScannedTo >= end) span.IsComplete = true;
    }

    /// <summary>Appends bytes read from the view to the ring when they continue it, else restarts the ring with them.</summary>
    private void Retain(long offset, byte[] bytes, int count) => Retain(offset, bytes, 0, count);

    private void Retain(long offset, byte[] bytes, int from, int count)
    {
        if (count > _ring.Length)
        {
            if (count > MaxRetainedBytes) { _ringStart = -1; _ringLength = 0; return; }
            var bigger = new byte[Math.Min(MaxRetainedBytes, Math.Max(count, _ring.Length * 2))];
            Buffer.BlockCopy(_ring, 0, bigger, 0, _ringLength);
            _ring = bigger;
        }
        if (_ringStart < 0 || _ringStart + _ringLength != offset) { _ringStart = offset; _ringLength = 0; }
        if (_ringLength + count > _ring.Length && _ring.Length < MaxRetainedBytes)
        {
            // a long contiguous run (a search batch locating thousands of blocks): keep it whole up to the cap, so the batch's bytes are read once
            var bigger = new byte[Math.Min(MaxRetainedBytes, Math.Max(_ringLength + count, _ring.Length * 2))];
            Buffer.BlockCopy(_ring, 0, bigger, 0, _ringLength);
            _ring = bigger;
        }
        if (_ringLength + count > _ring.Length)
        {
            int drop = Math.Max(count, _ring.Length / 2);
            drop = Math.Min(drop, _ringLength);
            Buffer.BlockCopy(_ring, drop, _ring, 0, _ringLength - drop);
            _ringStart += drop;
            _ringLength -= drop;
        }
        Buffer.BlockCopy(bytes, from, _ring, _ringLength, count);
        _ringLength += count;
    }

    public byte[] ReadRange(Stream view, long start, long end)
    {
        var bytes = new byte[end - start];
        int filled = 0;
        if (_ringStart >= 0 && start >= _ringStart && start < _ringStart + _ringLength)
        {
            // the head (often all of it) is in the ring; the rest continues from the ring's end, where the view already is
            filled = (int)Math.Min(bytes.Length, _ringStart + _ringLength - start);
            Buffer.BlockCopy(_ring, (int)(start - _ringStart), bytes, 0, filled);
            if (filled == bytes.Length) return bytes;
            if (view.Position != start + filled) view.Position = start + filled;
            while (filled < bytes.Length)
            {
                int read = view.Read(bytes, filled, bytes.Length - filled);
                if (read <= 0) throw new EndOfStreamException("Part ended inside an element.");
                Retain(start + filled, bytes, filled, read);
                filled += read;
            }
            return bytes;
        }
        view.Position = start;
        while (filled < bytes.Length)
        {
            int read = view.Read(bytes, filled, bytes.Length - filled);
            if (read <= 0) throw new EndOfStreamException("Part ended inside an element.");
            filled += read;
        }
        Retain(start, bytes, 0, bytes.Length);
        return bytes;
    }

    public long BlockStartOffset(int blockIndex)
    {
        if (IsAnchorPredecessor(blockIndex, out DocxIndexAnchor anchor)) return anchor.PreviousBlockOffset;
        Span span = EnsureStarted(blockIndex, out int local);
        return span.Scanner.BlockStartOffsets[local];
    }

    public long BlockEndOffset(int blockIndex)
    {
        if (blockIndex + 1 >= BlockCount) return _index.BodyEndOffset;
        int next = _index.AnchorIndexOf(blockIndex + 1);
        if (Anchors[next].FirstBlock == blockIndex + 1) return Anchors[next].Offset;
        return BlockStartOffset(blockIndex + 1);
    }

    public WordDocumentBlockKind BlockKindOf(int blockIndex)
    {
        if (IsAnchorPredecessor(blockIndex, out DocxIndexAnchor anchor)) return anchor.PreviousBlockKind;
        Span span = EnsureStarted(blockIndex, out int local);
        return span.Scanner.BlockKinds[local];
    }

    public int BlockTextLengthOf(int blockIndex)
    {
        Span span = EnsureFinished(blockIndex, out int local);
        return span.Scanner.BlockTextLengths[local];
    }

    /// <summary>The exact text length when the block's span has been scanned past it, else null (the index's group average stands in).</summary>
    public int? KnownTextLength(int blockIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) return null;
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        if (!_spans.TryGetValue(anchorIndex, out Span? span)) return null;
        int local = blockIndex - Anchors[anchorIndex].FirstBlock;
        bool finished = local + 1 < SpanBlockCount(anchorIndex) ? span.Scanner.BlockStartOffsets.Count > local + 1 : span.IsComplete;
        return finished ? span.Scanner.BlockTextLengths[local] : null;
    }

    public string? StyleIdOf(int blockIndex)
    {
        if (IsAnchorPredecessor(blockIndex, out DocxIndexAnchor anchor)) return anchor.PreviousStyleId;
        Span span = EnsureFinished(blockIndex, out int local);
        short styleIndex = span.Scanner.BlockStyleIds[local];
        return styleIndex >= 0 ? span.Scanner.StyleIdTable[styleIndex] : null;
    }

    public DocxListLabel? ListLabelOf(int blockIndex)
    {
        Span span = EnsureFinished(blockIndex, out int local);
        int labelIndex = span.Scanner.BlockListLabelIds[local];
        return labelIndex >= 0 ? span.Scanner.ListLabels[labelIndex] : null;
    }

    public string? OpenFieldsAt(int blockIndex) => _index.OpenFieldsAt(blockIndex);

    public ReadOnlySpan<DocxNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex)
        => _index.NoteReferencesInBlocks(firstBlockIndex, lastBlockIndex);
}
