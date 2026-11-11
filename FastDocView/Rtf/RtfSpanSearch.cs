namespace FastDocView.Rtf;

/// <summary>
/// One search batch fed from the span rescan itself (<see cref="RtfRandomAccessIndexBuilder.Search"/>): the parse that finds
/// the block bounds also hands its text and separators here, so a block is parsed once. The text rules are
/// <see cref="RtfSearchTextSink"/>'s (the per-block reference): paragraph / cell / row / line ends and note bounds are
/// separators, \listtext labels are skipped, footnote text is skipped (the block shows only its mark), a tab is a space.
/// </summary>
internal sealed class RtfSpanSearch(WordDocumentSearchMatcher matcher)
{
    private readonly List<int> _hits = new(), _counts = new();

    public List<int> Hits => _hits;
    public List<int> Counts => _counts;

    /// <summary>The document block index of the span's local block 0 (set per span).</summary>
    public int SpanFirstBlock { get; set; }

    /// <summary>The fields' context: its current block follows the scan (REF / NOTEREF \p).</summary>
    public Fields.WordFieldContext? FieldContext { get; init; }

    public void Text(ReadOnlySpan<char> text, RtfParser parser)
    {
        if (parser.IsInListText || parser.IsInFootnote) return;
        foreach (char c in text)
        {
            if (c == '\t') matcher.Feed(' ');
            else if (c is '\r' or '\n') matcher.Separator();
            else matcher.Feed(c);
        }
    }

    public void Separator() => matcher.Separator();

    /// <summary>Span-local block <paramref name="localIndex"/> ended: its matches are counted and the matcher starts afresh.</summary>
    public void BlockEnded(int localIndex)
    {
        if (FieldContext is not null) FieldContext.CurrentBlock = SpanFirstBlock + localIndex + 1;
        if (matcher.Count > 0) { _hits.Add(SpanFirstBlock + localIndex); _counts.Add(matcher.Count); }
        matcher.Reset();
    }
}
