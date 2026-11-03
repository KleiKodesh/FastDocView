namespace FastDocView.Odt;

/// <summary>
/// Serves <see cref="IOdtBlockTable"/> from a compact <see cref="OdtRandomAccessIndex"/>: a block's byte range, kind,
/// text length, list level / label / layout are recovered by rescanning its span - from the anchor's block start forward,
/// with a fresh <see cref="OdtBodyBlockScanner"/> seated there (frames + numbering state from the anchor, list styles
/// from the index) - so every value equals the one the index pass found. A span is scanned only as far as asked and kept
/// for later requests; a handful of spans stay cached (LRU). The bytes read while scanning stay in a bounded ring, so
/// the window just located is parsed from them without a second read; sequential windows cost one inflate pass.
/// </summary>
internal sealed class OdtCompactBlockTable : IOdtBlockTable
{
    private const int MaxCachedSpans = 8;
    private const int ChunkSize = 1 << 16;
    private const int RetainedBytes = 1 << 19;
    private const int MaxRetainedBytes = 4 << 20;

    private sealed class Span
    {
        public int Anchor;
        public OdtBodyBlockScanner Scanner = null!;
        public long ScannedTo;
        public bool IsComplete;
        public long LastUse;
    }

    private readonly OdtRandomAccessIndex _index;
    private readonly Stream _view;
    private readonly OdtStyleCascade _scanStyles;
    private readonly Dictionary<int, Span> _spans = new();
    private readonly byte[] _chunk = new byte[ChunkSize];
    private long _useCounter;

    private byte[] _ring = new byte[RetainedBytes];
    private long _ringStart = -1;
    private int _ringLength;

    public OdtCompactBlockTable(OdtRandomAccessIndex index, Stream view)
    {
        _index = index;
        _view = view;
        _scanStyles = new OdtStyleCascade();
        foreach (KeyValuePair<string, OdtListStyle> style in index.UsedListStyles) _scanStyles.AddListStyle(style.Key, style.Value);
    }

    public int BlockCount => _index.BlockCount;
    public string RootOpenTag => _index.RootOpenTag;
    private WordIndexSpanAnchor[] Anchors => _index.Anchors;

    private long SpanEnd(int anchor) => anchor + 1 < Anchors.Length ? Anchors[anchor + 1].Position : _index.LastBlockEndOffset;
    private int SpanBlockCount(int anchor) => (anchor + 1 < Anchors.Length ? Anchors[anchor + 1].FirstBlock : BlockCount) - Anchors[anchor].FirstBlock;

    /// <summary>True when the block is the one just before an anchor and its span is not scanned up to it: the anchor's record answers.</summary>
    private bool IsAnchorPredecessor(int blockIndex, out WordIndexSpanAnchor anchor)
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
            WordIndexSpanAnchor anchor = Anchors[anchorIndex];
            var scanner = new OdtBodyBlockScanner(_scanStyles);
            scanner.SeedAt(_index.RootOpenTag, anchor.Position, _index.ScanStateOf(anchorIndex));
            span = new Span { Anchor = anchorIndex, Scanner = scanner, ScannedTo = anchor.Position, IsComplete = anchor.Position >= SpanEnd(anchorIndex) };
            _spans[anchorIndex] = span;
        }
        span.LastUse = ++_useCounter;
        return span;
    }

    /// <summary>Scans the span until block <paramref name="localIndex"/> has started.</summary>
    private Span EnsureStarted(int blockIndex, out int localIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) throw new ArgumentOutOfRangeException(nameof(blockIndex));
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        localIndex = blockIndex - Anchors[anchorIndex].FirstBlock;
        Span span = GetSpan(anchorIndex);
        while (span.Scanner.BlockStartOffsets.Count <= localIndex)
        {
            if (span.IsComplete) throw new InvalidDataException($"OpenDocument index: block {blockIndex} not found in its span (the document changed?).");
            ScanMore(span);
        }
        return span;
    }

    /// <summary>Scans until the block has ended (its end offset, text length, label and level are final).</summary>
    private Span EnsureFinished(int blockIndex, out int localIndex)
    {
        Span span = EnsureStarted(blockIndex, out localIndex);
        while (span.Scanner.BlockEndOffsets[localIndex] < 0)
        {
            if (span.IsComplete) throw new InvalidDataException("OpenDocument index: a block does not end inside its span (the document changed?).");
            ScanMore(span);
        }
        return span;
    }

    private void ScanMore(Span span)
    {
        long end = SpanEnd(span.Anchor);
        int toRead = (int)Math.Min(ChunkSize, end - span.ScannedTo);
        if (toRead <= 0) { Complete(span); return; }
        if (_view.Position != span.ScannedTo) _view.Position = span.ScannedTo;
        int read = _view.Read(_chunk, 0, toRead);
        if (read <= 0) throw new EndOfStreamException("OpenDocument body ended inside a span.");
        Retain(span.ScannedTo, _chunk, 0, read);
        span.Scanner.Feed(_chunk, read);
        span.ScannedTo += read;
        if (span.ScannedTo >= end) Complete(span);
    }

    private void Complete(Span span)
    {
        span.IsComplete = true;
        if (span.Scanner.BlockStartOffsets.Count != SpanBlockCount(span.Anchor))
            throw new InvalidDataException("OpenDocument index: the span's rescan found a different block count (the document changed?).");
    }

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
            filled = (int)Math.Min(bytes.Length, _ringStart + _ringLength - start);
            Buffer.BlockCopy(_ring, (int)(start - _ringStart), bytes, 0, filled);
            if (filled == bytes.Length) return bytes;
            if (view.Position != start + filled) view.Position = start + filled;
            while (filled < bytes.Length)
            {
                int read = view.Read(bytes, filled, bytes.Length - filled);
                if (read <= 0) throw new EndOfStreamException("OpenDocument body ended inside an element.");
                Retain(start + filled, bytes, filled, read);
                filled += read;
            }
            return bytes;
        }
        view.Position = start;
        while (filled < bytes.Length)
        {
            int read = view.Read(bytes, filled, bytes.Length - filled);
            if (read <= 0) throw new EndOfStreamException("OpenDocument body ended inside an element.");
            filled += read;
        }
        Retain(start, bytes, 0, bytes.Length);
        return bytes;
    }

    public long BlockStartOffset(int blockIndex)
    {
        if (IsAnchorPredecessor(blockIndex, out WordIndexSpanAnchor anchor)) return anchor.PreviousBlockPosition;
        Span span = EnsureStarted(blockIndex, out int local);
        return span.Scanner.BlockStartOffsets[local];
    }

    public long BlockEndOffset(int blockIndex)
    {
        Span span = EnsureFinished(blockIndex, out int local);
        return span.Scanner.BlockEndOffsets[local];
    }

    public WordDocumentBlockKind BlockKindOf(int blockIndex)
    {
        if (IsAnchorPredecessor(blockIndex, out WordIndexSpanAnchor anchor)) return anchor.PreviousBlockKind;
        Span span = EnsureStarted(blockIndex, out int local);
        return span.Scanner.BlockKinds[local];
    }

    public int BlockTextLengthOf(int blockIndex)
    {
        Span span = EnsureFinished(blockIndex, out int local);
        return span.Scanner.BlockTextLengths[local];
    }

    /// <summary>The exact text length when the block's span has been scanned past its end, else null (the index's group average stands in).</summary>
    public int? KnownTextLength(int blockIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) return null;
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        if (!_spans.TryGetValue(anchorIndex, out Span? span)) return null;
        int local = blockIndex - Anchors[anchorIndex].FirstBlock;
        return span.Scanner.BlockEndOffsets.Count > local && span.Scanner.BlockEndOffsets[local] >= 0 ? span.Scanner.BlockTextLengths[local] : null;
    }

    public string? ListLabelOf(int blockIndex)
    {
        Span span = EnsureStarted(blockIndex, out int local);
        int labelIndex = span.Scanner.BlockListLabelIds[local];
        return labelIndex >= 0 ? span.Scanner.ListLabels[labelIndex] : null;
    }

    public string ListLabelLayoutOf(int blockIndex)
    {
        Span span = EnsureStarted(blockIndex, out int local);
        int labelIndex = span.Scanner.BlockListLabelIds[local];
        return labelIndex >= 0 ? span.Scanner.ListLabelLayouts[labelIndex] : "";
    }

    public int ListLevelOf(int blockIndex)
    {
        Span span = EnsureStarted(blockIndex, out int local);
        byte level = span.Scanner.BlockListLevels[local];
        return level == 255 ? -1 : level;
    }

    public ReadOnlySpan<OdtNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex)
        => _index.NoteReferencesInBlocks(firstBlockIndex, lastBlockIndex);
}
