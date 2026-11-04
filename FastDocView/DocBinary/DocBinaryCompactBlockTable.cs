namespace FastDocView.DocBinary;

/// <summary>
/// Serves block positions, kinds and list labels from a compact <see cref="DocBinaryBlockIndex"/>: a block's values
/// are recovered by rescanning its span - from the anchor's CP forward, with the same <see cref="DocBinaryBlockScanner"/>
/// (and list counter, seeded with the anchor's state) the index pass ran, so every value equals the one the pass found.
/// A span is scanned only as far as asked and kept scanning on later requests; a handful of spans stay cached (the ones
/// around the window, LRU). The piece text is read in 64K-character chunks, exactly as the index pass read it.
/// </summary>
internal sealed class DocBinaryCompactBlockTable
{
    private const int MaxCachedSpans = 8;
    private const int ChunkCharacters = 1 << 16;

    private sealed class Span
    {
        public int Anchor;
        public DocBinaryBlockScanner Scanner = null!;
        public int ScannedTo;
        public bool IsComplete;
        public long LastUse;
    }

    private readonly DocBinaryBlockIndex _index;
    private readonly DocBinaryPieceTable _pieces;
    private readonly DocBinaryParagraphHtmlRenderer _renderer;
    private readonly DocBinaryStyleSheet _styles;
    private readonly DocBinaryListDefinitions _lists;
    private readonly Dictionary<int, Span> _spans = new();
    private long _useCounter;

    public DocBinaryCompactBlockTable(DocBinaryBlockIndex index, DocBinaryPieceTable pieces, DocBinaryParagraphHtmlRenderer renderer,
                                      DocBinaryStyleSheet styles, DocBinaryListDefinitions lists)
    {
        _index = index;
        _pieces = pieces;
        _renderer = renderer;
        _styles = styles;
        _lists = lists;
    }

    public int BlockCount => _index.BlockCount;
    private WordIndexSpanAnchor[] Anchors => _index.Anchors;

    private int SpanEndCp(int anchor) => anchor + 1 < Anchors.Length ? (int)Anchors[anchor + 1].Position : _index.MainTextCharacterCount;
    private int SpanBlockCount(int anchor) => (anchor + 1 < Anchors.Length ? Anchors[anchor + 1].FirstBlock : BlockCount) - Anchors[anchor].FirstBlock;

    /// <summary>True when the block is the one just before an anchor and its span is not scanned up to it: the anchor's record answers.</summary>
    private bool IsAnchorPredecessor(int blockIndex, out WordIndexSpanAnchor anchor)
    {
        int next = _index.AnchorIndexOf(blockIndex) + 1;
        if (next < Anchors.Length && Anchors[next].FirstBlock == blockIndex + 1
            && !(_spans.TryGetValue(next - 1, out Span? span) && span.Scanner.BlockCount > blockIndex - Anchors[next - 1].FirstBlock))
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
            var scanner = new DocBinaryBlockScanner(_renderer, _styles, _lists, (int)anchor.Position);
            scanner.ImportListState(_index.ListStateOf(anchorIndex));
            span = new Span { Anchor = anchorIndex, Scanner = scanner, ScannedTo = (int)anchor.Position, IsComplete = false };
            if (span.ScannedTo >= SpanEndCp(anchorIndex)) Complete(span);
            _spans[anchorIndex] = span;
        }
        span.LastUse = ++_useCounter;
        return span;
    }

    private void ScanMore(Span span)
    {
        int end = SpanEndCp(span.Anchor);
        int to = (int)Math.Min((long)span.ScannedTo + ChunkCharacters, end);
        if (to <= span.ScannedTo) { Complete(span); return; }
        _pieces.ScanCharacters(span.ScannedTo, to, (characters, count, chunkCp) => span.Scanner.Feed(characters, count, chunkCp), ChunkCharacters);
        span.ScannedTo = to;
        if (span.ScannedTo >= end) Complete(span);
    }

    /// <summary>The last span ends like the index pass did (unfinished row / paragraph closed); an inner span ends at the next anchor's block start.</summary>
    private void Complete(Span span)
    {
        if (span.Anchor + 1 >= Anchors.Length) span.Scanner.Finish(_index.MainTextCharacterCount);
        else span.Scanner.CloseOpenRow();
        span.IsComplete = true;
        if (span.Scanner.BlockCount != SpanBlockCount(span.Anchor)) throw new InvalidDataException(".doc index: the span's rescan found a different block count (the document changed?).");
    }

    /// <summary>Scans the span until block <paramref name="localIndex"/> has started (its CP and kind are known).</summary>
    private Span EnsureStarted(int blockIndex, out int localIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) throw new ArgumentOutOfRangeException(nameof(blockIndex));
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        localIndex = blockIndex - Anchors[anchorIndex].FirstBlock;
        Span span = GetSpan(anchorIndex);
        while (span.Scanner.BlockCount <= localIndex)
        {
            if (span.IsComplete) throw new InvalidDataException($".doc index: block {blockIndex} not found in its span (the document changed?).");
            ScanMore(span);
        }
        return span;
    }

    public int BlockStartCp(int blockIndex)
    {
        if (IsAnchorPredecessor(blockIndex, out WordIndexSpanAnchor anchor)) return (int)anchor.PreviousBlockPosition;
        Span span = EnsureStarted(blockIndex, out int local);
        return span.Scanner.BlockStartCps[local];
    }

    public int BlockEndCp(int blockIndex)
    {
        if (blockIndex + 1 >= BlockCount) return _index.MainTextCharacterCount;
        int next = _index.AnchorIndexOf(blockIndex + 1);
        if (Anchors[next].FirstBlock == blockIndex + 1) return (int)Anchors[next].Position;
        return BlockStartCp(blockIndex + 1);
    }

    public WordDocumentBlockKind BlockKindOf(int blockIndex)
    {
        if (IsAnchorPredecessor(blockIndex, out WordIndexSpanAnchor anchor)) return anchor.PreviousBlockKind;
        Span span = EnsureStarted(blockIndex, out int local);
        return span.Scanner.BlockKinds[local];
    }

    public string? ListLabelOf(int blockIndex)
    {
        Span span = EnsureStarted(blockIndex, out int local);
        int labelIndex = span.Scanner.BlockListLabelIds[local];
        return labelIndex >= 0 ? span.Scanner.ListLabels[labelIndex] : null;
    }

    /// <summary>The block's list number in full context (REF \w \r), recomputed by the span scan from the anchor's counter state; null when not a list item.</summary>
    public string? ListContextOf(int blockIndex)
    {
        Span span = EnsureStarted(blockIndex, out int local);
        int contextIndex = span.Scanner.BlockListContextIds[local];
        return contextIndex >= 0 ? span.Scanner.ListContexts[contextIndex] : null;
    }

    /// <summary>The exact text length when the block's span has been scanned past it, else null (the index's group average stands in).</summary>
    public int? KnownTextLength(int blockIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) return null;
        if (IsAnchorPredecessor(blockIndex, out WordIndexSpanAnchor anchor)) return (int)(anchor.Position - anchor.PreviousBlockPosition);
        int anchorIndex = _index.AnchorIndexOf(blockIndex);
        if (!_spans.TryGetValue(anchorIndex, out Span? span)) return null;
        int local = blockIndex - Anchors[anchorIndex].FirstBlock;
        if (span.Scanner.BlockCount <= local) return null;
        int end = local + 1 < span.Scanner.BlockCount ? span.Scanner.BlockStartCps[local + 1]
                : local + 1 < SpanBlockCount(anchorIndex) ? -1
                : anchorIndex + 1 < Anchors.Length ? (int)Anchors[anchorIndex + 1].Position
                : span.IsComplete ? _index.MainTextCharacterCount : -1;
        return end < 0 ? null : end - span.Scanner.BlockStartCps[local];
    }

    /// <summary>The block containing a main-text CP.</summary>
    public int BlockAt(int cp)
    {
        if (BlockCount == 0) return 0;
        cp = Math.Max(0, Math.Min(cp, _index.MainTextCharacterCount));
        int anchorIndex = WordIndexSpanAnchor.IndexOfPosition(Anchors, cp);
        Span span = GetSpan(anchorIndex);
        // scan until a block starts past the CP (or the span ends): the block before that start holds the CP
        while (!span.IsComplete && (span.Scanner.BlockCount == 0 || span.Scanner.BlockStartCps[span.Scanner.BlockCount - 1] <= cp)) ScanMore(span);
        int local = DocBinaryBlockIndex.BlockAt(span.Scanner.BlockStartCps, cp);
        return Anchors[anchorIndex].FirstBlock + local;
    }
}
