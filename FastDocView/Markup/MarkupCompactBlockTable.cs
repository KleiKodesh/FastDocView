namespace FastDocView.Markup;

/// <summary>
/// Serves block starts, ends, text lengths, open-element stacks and (MHTML) decoder resume points from a compact
/// <see cref="MarkupRandomAccessIndex"/>: a block's values come from rescanning its span (at most ~256 KB of text past
/// the anchor, plus the rest of a block that runs longer) with the index pass's own scanner, seeded with the anchor's
/// snapshot, so every value equals the one the pass found. A handful of spans stay cached (LRU); nothing else is kept.
/// </summary>
internal sealed class MarkupCompactBlockTable
{
    private const int MaxCachedSpans = 8;

    private sealed class Span
    {
        public long[] Starts = [];
        public int[] TextLengths = [];
        public int[] StackIds = [];
        public long[] Resumes = [];
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
    private int SpanBlockCount(int anchor) => (anchor + 1 < Anchors.Length ? Anchors[anchor + 1].FirstBlock : _index.BlockCount) - Anchors[anchor].FirstBlock;

    private Span GetSpan(int anchorIndex)
    {
        if (!_spans.TryGetValue(anchorIndex, out Span? span))
        {
            if (_spans.Count >= MaxCachedSpans)
            {
                int oldest = _spans.OrderBy(pair => pair.Value.LastUse).First().Key;
                _spans.Remove(oldest);
            }
            int count = SpanBlockCount(anchorIndex);
            MarkupBlockScanner scanner;
            List<long>? resumes = null;
            if (_index.Kind == MarkupSourceKind.WebArchive) scanner = MarkupRandomAccessIndexBuilder.ScanWebArchiveSpan(_file, _index, anchorIndex, out resumes);
            else scanner = MarkupRandomAccessIndexBuilder.ScanPlainTextSpan(_file, _index, anchorIndex);
            // the span's blocks are complete once the next one has started (or the last span finished)
            bool isLast = anchorIndex + 1 >= Anchors.Length;
            int found = scanner.BlockStartOffsets.Count;
            if (found < count || (!isLast && found == count) || (isLast && found != count) || (resumes is not null && resumes.Count < count)
                || scanner.BlockIndexBase != SpanFirstBlock(anchorIndex))
                throw new InvalidDataException("Markup index: the span's rescan found a different block count (the document changed?).");
            span = new Span
            {
                Starts = scanner.BlockStartOffsets.GetRange(0, count).ToArray(),
                TextLengths = scanner.BlockTextLengths.GetRange(0, count).ToArray(),
                StackIds = scanner.BlockStackIds.GetRange(0, count).ToArray(),
                Resumes = resumes is null ? [] : resumes.GetRange(0, count).ToArray(),
            };
            _spans[anchorIndex] = span;
        }
        span.LastUse = ++_useCounter;
        return span;
    }

    private Span SpanOf(int blockIndex, out int localIndex)
    {
        if ((uint)blockIndex >= (uint)_index.BlockCount) throw new ArgumentOutOfRangeException(nameof(blockIndex));
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        localIndex = blockIndex - SpanFirstBlock(anchorIndex);
        return GetSpan(anchorIndex);
    }

    public long BlockStart(int blockIndex) => SpanOf(blockIndex, out int local).Starts[local];

    /// <summary>Where the block ends: the next block's start (the first block of the next span), or the text's end.</summary>
    public long BlockEnd(int blockIndex)
    {
        if (blockIndex + 1 >= _index.BlockCount) return _index.TextEnd;
        int anchorIndex = _index.AnchorIndexOf(blockIndex + 1);
        if (SpanFirstBlock(anchorIndex) == blockIndex + 1) return Anchors[anchorIndex].FirstBlockStart;   // no rescan of the next span
        return BlockStart(blockIndex + 1);
    }

    public int TextLength(int blockIndex) => SpanOf(blockIndex, out int local).TextLengths[local];
    public string OpenTags(int blockIndex)
    {
        Span span = SpanOf(blockIndex, out int local);
        return _index.Stacks[span.StackIds[local]];
    }
    public int StackId(int blockIndex) => SpanOf(blockIndex, out int local).StackIds[local];

    /// <summary>MHTML: the decoder resume point of the block's first decoded byte.</summary>
    public long Resume(int blockIndex)
    {
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        if (SpanFirstBlock(anchorIndex) == blockIndex && Anchors[anchorIndex].State is not null) return Anchors[anchorIndex].BlockResume;
        return SpanOf(blockIndex, out int local).Resumes[local];
    }

    /// <summary>The exact text length when the block's span has been scanned, else null (the index's group average stands in).</summary>
    public int? KnownTextLength(int blockIndex)
    {
        if ((uint)blockIndex >= (uint)_index.BlockCount) return null;
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        return _spans.TryGetValue(anchorIndex, out Span? span) ? span.TextLengths[blockIndex - SpanFirstBlock(anchorIndex)] : null;
    }
}
