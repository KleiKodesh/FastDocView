namespace FastDocView.Rtf;

/// <summary>
/// Serves block ranges, kinds, start snapshots and list labels from a compact <see cref="RtfRandomAccessIndex"/>: a
/// block's values are recovered by rescanning its whole span (the file bytes from the anchor's block start to the next
/// anchor's) with the same block rules the index pass ran (<see cref="RtfRandomAccessIndexBuilder.ScanSpan"/>), seeded
/// with the anchor's parser snapshot and list state, so every value equals the one the pass found. A span is at most
/// 256 KB of RTF (a few ms); a handful of spans stay cached (LRU).
/// </summary>
internal sealed class RtfCompactBlockTable
{
    private const int MaxCachedSpans = 8;

    private sealed class Span
    {
        public int Anchor;
        public RtfRandomAccessIndexBuilder Scanner = null!;
        public long LastUse;
    }

    private readonly RtfRandomAccessIndex _index;
    private readonly Stream _file;
    private readonly Dictionary<int, Span> _spans = new();
    private long _useCounter;

    public RtfCompactBlockTable(RtfRandomAccessIndex index, Stream file)
    {
        _index = index;
        _file = file;
    }

    public int BlockCount => _index.BlockCount;
    private WordIndexSpanAnchor[] Anchors => _index.Anchors;

    private long SpanEnd(int anchor) => anchor + 1 < Anchors.Length ? Anchors[anchor + 1].Position : _index.LastBlockEndOffset;
    private int SpanBlockCount(int anchor) => (anchor + 1 < Anchors.Length ? Anchors[anchor + 1].FirstBlock : BlockCount) - Anchors[anchor].FirstBlock;

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
            RtfRandomAccessIndexBuilder scanner = RtfRandomAccessIndexBuilder.ScanSpan(_file, _index.Tables, anchor.Position, _index.ScanStateOf(anchorIndex),
                                                                                       SpanEnd(anchorIndex), isLast: anchorIndex + 1 >= Anchors.Length);
            if (scanner.Starts.Count != SpanBlockCount(anchorIndex))
                throw new InvalidDataException("RTF index: the span's rescan found a different block count (the document changed?).");
            span = new Span { Anchor = anchorIndex, Scanner = scanner };
            _spans[anchorIndex] = span;
        }
        span.LastUse = ++_useCounter;
        return span;
    }

    private RtfRandomAccessIndexBuilder SpanOf(int blockIndex, out int localIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) throw new ArgumentOutOfRangeException(nameof(blockIndex));
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        localIndex = blockIndex - Anchors[anchorIndex].FirstBlock;
        return GetSpan(anchorIndex).Scanner;
    }

    public long BlockStartOffset(int blockIndex) => SpanOf(blockIndex, out int local).Starts[local];
    public long BlockEndOffset(int blockIndex) => SpanOf(blockIndex, out int local).Ends[local];
    public WordDocumentBlockKind BlockKindOf(int blockIndex) => SpanOf(blockIndex, out int local).Kinds[local];
    public int BlockTextLengthOf(int blockIndex) => SpanOf(blockIndex, out int local).TextLengths[local];
    /// <summary>The parser's group-state snapshot at the block's first byte (what a window restores before parsing its bytes).</summary>
    public byte[] SnapshotOf(int blockIndex) => SpanOf(blockIndex, out int local).SnapshotOf(local);
    public string? ListLabelOf(int blockIndex) => SpanOf(blockIndex, out int local).LabelOf(local);

    public int ListLevelOf(int blockIndex)
    {
        byte level = SpanOf(blockIndex, out int local).ListLevels[local];
        return level == 255 ? -1 : level;
    }

    /// <summary>The exact text length when the block's span has been scanned, else null (the index's group average stands in).</summary>
    public int? KnownTextLength(int blockIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) return null;
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        return _spans.TryGetValue(anchorIndex, out Span? span) ? span.Scanner.TextLengths[blockIndex - Anchors[anchorIndex].FirstBlock] : null;
    }
}
