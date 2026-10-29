namespace FastDocView.Rtf;

/// <summary>
/// The search's sink over the same <see cref="RtfParser"/> the renderer uses, so a block is read exactly as it renders:
/// its text goes to the matcher (no HTML), paragraph / cell / row / line ends are separators, \listtext labels are
/// skipped, and footnote groups are fed (between separators) only when the search includes notes.
/// </summary>
internal sealed class RtfSearchTextSink : RtfSink
{
    private readonly RtfParser _parser;
    private WordDocumentSearchMatcher? _matcher;
    private bool _includeNotes;

    public RtfSearchTextSink(RtfDocumentTables tables) => _parser = new RtfParser(this, tables, readsHeaderTables: false);

    public void Feed(Stream file, RtfRandomAccessIndex index, int blockIndex, WordDocumentSearchMatcher matcher, bool includeNotes)
    {
        _matcher = matcher;
        _includeNotes = includeNotes;
        _parser.RestoreSnapshot(index.States[index.BlockStateIds[blockIndex]]);
        _parser.Run(file, index.BlockStartOffsets[blockIndex], index.BlockEndOffsets[blockIndex]);
        _matcher = null;
    }

    public override void Text(ReadOnlySpan<char> text, in RtfGroupState state)
    {
        if (_parser.IsInListText || (_parser.IsInFootnote && !_includeNotes)) return;
        foreach (char c in text)
        {
            if (c == '\t') _matcher!.Feed(' ');        // \tab arrives as a tab character; a space, as in the other formats
            else if (c is '\r' or '\n') _matcher!.Separator();
            else _matcher!.Feed(c);
        }
    }

    public override void ParagraphEnd(in RtfGroupState state, bool isSectionEnd) => _matcher!.Separator();
    public override void CellEnd(in RtfGroupState state, bool isNested) => _matcher!.Separator();
    public override void RowEnd(in RtfGroupState state, bool isNested) => _matcher!.Separator();
    public override void Break(RtfWord word, in RtfGroupState state) => _matcher!.Separator();
    public override void FootnoteStart() => _matcher!.Separator();
    public override void FootnoteEnd(bool isEndnote) => _matcher!.Separator();
}
