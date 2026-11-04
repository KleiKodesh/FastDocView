namespace FastDocView.Rtf;

/// <summary>
/// Serves block ranges, kinds, start snapshots and list labels from a compact <see cref="RtfRandomAccessIndex"/>: a
/// block's values are recovered by rescanning its span (the file bytes from the anchor's block start towards the next
/// anchor's) with the same block rules the index pass ran (<see cref="RtfRandomAccessIndexBuilder.ResumableSpan"/>), seeded
/// with the anchor's parser snapshot and list state, so every value equals the one the pass found. The rescan runs only
/// as far as the asked-for block and resumes from there on the next request (a window at a span's start reads one 64 KB
/// buffer, not the span's 256 KB); a handful of spans stay cached (LRU). A search batch is fed from the same rescan
/// (<see cref="Search"/>), so its blocks are parsed once.
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
        if (!_spans.TryGetValue(anchorIndex, out Span? span)) span = StartSpan(anchorIndex);
        span.LastUse = ++_useCounter;
        return span;
    }

    /// <summary>A fresh (not yet read) rescan of the span, cached in place of any earlier one.</summary>
    private Span StartSpan(int anchorIndex)
    {
        _spans.Remove(anchorIndex);
        if (_spans.Count >= MaxCachedSpans)
        {
            int oldest = _spans.OrderBy(pair => pair.Value.LastUse).First().Key;
            _spans.Remove(oldest);
        }
        WordIndexSpanAnchor anchor = Anchors[anchorIndex];
        RtfRandomAccessIndexBuilder scanner = RtfRandomAccessIndexBuilder.ResumableSpan(_file, _index.Tables, anchor.Position, _index.ScanStateOf(anchorIndex),
                                                                                        SpanEnd(anchorIndex), isLast: anchorIndex + 1 >= Anchors.Length);
        scanner.FieldContext = FieldContext;
        var span = new Span { Anchor = anchorIndex, Scanner = scanner, LastUse = ++_useCounter };
        _spans[anchorIndex] = span;
        return span;
    }

    /// <summary>Scans the span on through <paramref name="localIndex"/>; at the span's end its block count must be the index's.</summary>
    private void ScanThrough(int anchorIndex, RtfRandomAccessIndexBuilder scanner, int localIndex, RtfSpanSearch? search = null)
    {
        try
        {
            if (search is null) scanner.ScanThrough(localIndex);
            else scanner.Search(search, localIndex);
        }
        catch
        {
            _spans.Remove(anchorIndex);            // a rescan that failed part-way is not kept
            throw;
        }
        int expected = SpanBlockCount(anchorIndex);
        if (scanner.IsComplete ? scanner.Starts.Count != expected : scanner.Starts.Count > expected)
        {
            _spans.Remove(anchorIndex);
            throw new InvalidDataException("RTF index: the span's rescan found a different block count (the document changed?).");
        }
    }

    private RtfRandomAccessIndexBuilder SpanOf(int blockIndex, out int localIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) throw new ArgumentOutOfRangeException(nameof(blockIndex));
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        localIndex = blockIndex - Anchors[anchorIndex].FirstBlock;
        RtfRandomAccessIndexBuilder scanner = GetSpan(anchorIndex).Scanner;
        if (localIndex >= scanner.Starts.Count) ScanThrough(anchorIndex, scanner, localIndex);
        return scanner;
    }

    /// <summary>
    /// Feeds blocks [<paramref name="fromBlock"/>, <paramref name="endBlock"/>) to <paramref name="search"/> from the span rescans
    /// themselves: a cached rescan paused right before a block goes on from there (the next batch of a sequential search reads
    /// on where the last one stopped, a rescan short of the block is continued to it); otherwise the span is rescanned afresh from its anchor (the blocks before the batch are
    /// parsed without feeding the matcher). Either way each block is parsed once.
    /// </summary>
    public void Search(int fromBlock, int endBlock, RtfSpanSearch search)
    {
        int block = fromBlock;
        while (block < endBlock)
        {
            int anchorIndex = _index.AnchorIndexOf(block);
            int spanFirst = Anchors[anchorIndex].FirstBlock;
            int local = block - spanFirst;
            int last = Math.Min(endBlock, spanFirst + SpanBlockCount(anchorIndex)) - 1 - spanFirst;
            Span span = _spans.TryGetValue(anchorIndex, out Span? cached) && !cached.Scanner.IsComplete && cached.Scanner.Starts.Count <= local
                ? cached
                : StartSpan(anchorIndex);
            span.LastUse = ++_useCounter;
            if (local > 0 && span.Scanner.Starts.Count < local) ScanThrough(anchorIndex, span.Scanner, local - 1);
            search.SpanFirstBlock = spanFirst;
            ScanThrough(anchorIndex, span.Scanner, last, search);
            block = spanFirst + last + 1;
        }
    }

    public long BlockStartOffset(int blockIndex) => SpanOf(blockIndex, out int local).Starts[local];
    public long BlockEndOffset(int blockIndex) => SpanOf(blockIndex, out int local).Ends[local];
    public WordDocumentBlockKind BlockKindOf(int blockIndex) => SpanOf(blockIndex, out int local).Kinds[local];
    public int BlockTextLengthOf(int blockIndex) => SpanOf(blockIndex, out int local).TextLengths[local];
    /// <summary>The parser's group-state snapshot at the block's first byte (what a window restores before parsing its bytes).</summary>
    public byte[] SnapshotOf(int blockIndex) => SpanOf(blockIndex, out int local).SnapshotOf(local);
    /// <summary>The fields open where a block starts and the field counters there (RtfFieldTracker.Serialize).</summary>
    public string FieldStateOf(int blockIndex) => SpanOf(blockIndex, out int local).FieldStateOf(local);

    /// <summary>What a span search computes field results from (the renderer's context).</summary>
    public Fields.WordFieldContext? FieldContext { get; set; }
    public string? ListLabelOf(int blockIndex) => SpanOf(blockIndex, out int local).LabelOf(local);
    /// <summary>The style number (\sN) of the block's first paragraph; -1 when none.</summary>
    public int StyleOf(int blockIndex) => SpanOf(blockIndex, out int local).BlockStyles[local];
    /// <summary>The list override (\lsN) of the block's first paragraph; -1 when none.</summary>
    public int ListOf(int blockIndex) => SpanOf(blockIndex, out int local).BlockLists[local];

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
        int local = blockIndex - Anchors[anchorIndex].FirstBlock;
        return _spans.TryGetValue(anchorIndex, out Span? span) && local < span.Scanner.TextLengths.Count ? span.Scanner.TextLengths[local] : null;
    }
}
