namespace FastDocView.Rtf;

/// <summary>
/// One search batch fed from the span rescan itself (<see cref="RtfRandomAccessIndexBuilder.Search"/>): the parse that finds
/// the block bounds also hands its text and separators here, so a block is parsed once. The text rules are
/// <see cref="RtfSearchTextSink"/>'s (the per-block reference): paragraph / cell / row / line ends and note bounds are
/// separators, \listtext labels are skipped, footnote text only when the search includes notes, a tab is a space.
/// </summary>
internal sealed class RtfSpanSearch(WordDocumentSearchMatcher matcher, bool includeNotes)
{
    private readonly List<int> _hits = new(), _counts = new();

    public List<int> Hits => _hits;
    public List<int> Counts => _counts;

    /// <summary>The document block index of the span's local block 0 (set per span).</summary>
    public int SpanFirstBlock { get; set; }

    public void Text(ReadOnlySpan<char> text, RtfParser parser)
    {
        if (parser.IsInListText || (parser.IsInFootnote && !includeNotes)) return;
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
        if (matcher.Count > 0) { _hits.Add(SpanFirstBlock + localIndex); _counts.Add(matcher.Count); }
        matcher.Reset();
    }
}
