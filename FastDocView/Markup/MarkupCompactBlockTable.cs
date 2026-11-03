namespace FastDocView.Markup;

/// <summary>
/// Serves block starts, ends, text lengths, open-element stacks and (MHTML) decoder resume points from a compact
/// <see cref="MarkupRandomAccessIndex"/>: a block's values come from rescanning its span with the index pass's own
/// scanner, seeded with the anchor's snapshot (<see cref="MarkupSpanRescan"/>), so every value equals the one the pass
/// found. A rescan runs only as far as the blocks asked for and resumes from there on the next ask. A handful of spans
/// stay cached (LRU); nothing else is kept.
/// </summary>
internal sealed class MarkupCompactBlockTable
{
    private const int MaxCachedSpans = 8;

    private sealed class Span
    {
        public required MarkupSpanRescan Rescan;
        public long LastUse;
    }

    private readonly MarkupRandomAccessIndex _index;
    private readonly Stream _file;
    private readonly Dictionary<int, Span> _spans = new();
    private long _useCounter;

    public MarkupCompactBlockTable(MarkupRandomAccessIndex index, Stream file)
    {
        _index = index;
        _file = file;
    }

    private MarkupSpanAnchor[] Anchors => _index.Anchors;
    private int SpanFirstBlock(int anchor) => Anchors[anchor].FirstBlock;

    private Span GetSpan(int anchorIndex)
    {
        if (!_spans.TryGetValue(anchorIndex, out Span? span))
        {
            if (_spans.Count >= MaxCachedSpans)
            {
                int oldest = _spans.OrderBy(pair => pair.Value.LastUse).First().Key;
                _spans.Remove(oldest);
            }
            span = new Span { Rescan = new MarkupSpanRescan(_file, _index, anchorIndex) };
            _spans[anchorIndex] = span;
        }
        span.LastUse = ++_useCounter;
        return span;
    }

    /// <summary>The scanner of the block's span, scanned at least through that block.</summary>
    private MarkupSpanRescan ScannedThrough(int blockIndex, out int localIndex)
    {
        if ((uint)blockIndex >= (uint)_index.BlockCount) throw new ArgumentOutOfRangeException(nameof(blockIndex));
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        localIndex = blockIndex - SpanFirstBlock(anchorIndex);
        MarkupSpanRescan rescan = GetSpan(anchorIndex).Rescan;
        rescan.ScanThrough(localIndex);
        return rescan;
    }

    public long BlockStart(int blockIndex) => ScannedThrough(blockIndex, out int local).Scanner.BlockStartOffsets[local];

    /// <summary>Where the block ends: the next block's start (the first block of the next span), or the text's end.</summary>
    public long BlockEnd(int blockIndex)
    {
        if (blockIndex + 1 >= _index.BlockCount) return _index.TextEnd;
        int anchorIndex = _index.AnchorIndexOf(blockIndex + 1);
        if (SpanFirstBlock(anchorIndex) == blockIndex + 1) return Anchors[anchorIndex].FirstBlockStart;   // no rescan of the next span
        return BlockStart(blockIndex + 1);
    }

    public int TextLength(int blockIndex) => ScannedThrough(blockIndex, out int local).Scanner.BlockTextLengths[local];
    public string OpenTags(int blockIndex) => _index.Stacks[StackId(blockIndex)];
    public int StackId(int blockIndex) => ScannedThrough(blockIndex, out int local).Scanner.BlockStackIds[local];

    /// <summary>MHTML: the decoder resume point of the block's first decoded byte.</summary>
    public long Resume(int blockIndex)
    {
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        if (SpanFirstBlock(anchorIndex) == blockIndex && Anchors[anchorIndex].State is not null) return Anchors[anchorIndex].BlockResume;
        return ScannedThrough(blockIndex, out int local).Resumes![local];
    }

    /// <summary>The exact text length when the block's span has been scanned through it, else null (the index's group average stands in).</summary>
    public int? KnownTextLength(int blockIndex)
    {
        if ((uint)blockIndex >= (uint)_index.BlockCount) return null;
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        int local = blockIndex - SpanFirstBlock(anchorIndex);
        return _spans.TryGetValue(anchorIndex, out Span? span) && span.Rescan.IsBlockComplete(local) ? span.Rescan.Scanner.BlockTextLengths[local] : null;
    }
}
