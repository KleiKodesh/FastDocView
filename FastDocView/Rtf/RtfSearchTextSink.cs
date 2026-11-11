namespace FastDocView.Rtf;

/// <summary>
/// The search's sink over the same <see cref="RtfParser"/> the renderer uses, so a block is read exactly as it renders:
/// its text goes to the matcher (no HTML), paragraph / cell / row / line ends are separators, \listtext labels are
/// skipped, text-box text ({\shptxt}) is fed once between separators, and footnote groups are skipped (the block shows only their mark).
/// </summary>
internal sealed class RtfSearchTextSink : RtfSink
{
    private readonly RtfParser _parser;
    private WordDocumentSearchMatcher? _matcher;
    private readonly RtfFieldTracker _fields;

    public RtfSearchTextSink(RtfDocumentTables tables, Fields.WordFieldContext? fieldContext = null)
    {
        _parser = new RtfParser(this, tables, readsHeaderTables: false);
        _fields = new RtfFieldTracker(fieldContext, counts: true);
    }

    public void Feed(Stream file, RtfCompactBlockTable table, int blockIndex, WordDocumentSearchMatcher matcher)
    {
        _matcher = matcher;
        _parser.RestoreSnapshot(table.SnapshotOf(blockIndex));
        _fields.Restore(table.FieldStateOf(blockIndex));
        _parser.Run(file, table.BlockStartOffset(blockIndex), table.BlockEndOffset(blockIndex));
        _matcher = null;
    }

    public override void Text(ReadOnlySpan<char> text, in RtfGroupState state)
    {
        if (_parser.IsInListText || _parser.IsInFootnote) return;
        if (!_fields.IsIdle && !_fields.Shows(text)) return;
        foreach (char c in text)
        {
            if (c == '\t') _matcher!.Feed(' ');        // \tab arrives as a tab character; a space, as in the other formats
            else if (c is '\r' or '\n') _matcher!.Separator();
            else _matcher!.Feed(c);
        }
    }

    // fields: the renderer's rules (RtfFieldTracker) - codes never read, a computed result read where it shows
    public override void FieldStart(int depth) => _fields.Start(depth);
    public override void FieldModifier(RtfWord word) { if (word == RtfWord.FieldLock) _fields.Lock(); else _fields.Private(); }
    public override void FormField(Fields.WordFormFieldState state) => _fields.Form(state);
    public override void FieldInstruction(string instruction) => _fields.Instruction(instruction);
    public override void FieldResultStart() => Feed(_fields.ResultStart(out _));
    public override void GroupClosed(int depth) { if (!_fields.IsIdle) _fields.GroupClosed(depth, (output, _) => Feed(output)); }

    private void Feed(Fields.WordFieldOutput output)
    {
        if (output is not { Text: { } text, IsSearchable: true } || _parser.IsInListText || _parser.IsInFootnote) return;
        _matcher!.Feed(text.AsSpan());
    }

    public override bool WantsMath => true;

    /// <summary>An equation's text, as the renderer shows it (RtfBlockHtmlRenderer.Equation).</summary>
    public override void Equation(byte[] group, in RtfGroupState state)
    {
        if (_parser.IsInListText || _parser.IsInFootnote) return;
        if (!_fields.IsIdle && !_fields.Shows("x".AsSpan())) return;     // inside a code or a replaced result: not shown
        if (RtfOfficeMath.ToOmml(group, state.CodePage, _ => null) is { } equation) DocBinary.DocBinaryParagraphHtmlRenderer.FeedEquationText(equation, _matcher!);
    }

    public override bool WantsObjects => true;

    /// <summary>An embedded chart's texts, as the renderer draws them (RtfBlockHtmlRenderer.Object).</summary>
    public override void Object(byte[] group, in RtfGroupState state)
    {
        if (_parser.IsInListText || _parser.IsInFootnote) return;
        if (!_fields.IsIdle && !_fields.Shows("x".AsSpan())) return;
        if (RtfObjectChart.Read(group) is not { } found) return;
        foreach (string text in found.Chart.Texts()) { _matcher!.Separator(); _matcher.Feed(text.AsSpan()); _matcher.Separator(); }
    }

    public override void ParagraphEnd(in RtfGroupState state, bool isSectionEnd) => _matcher!.Separator();
    public override void CellEnd(in RtfGroupState state, bool isNested) => _matcher!.Separator();
    public override void RowEnd(in RtfGroupState state, bool isNested) => _matcher!.Separator();
    public override void Break(RtfWord word, in RtfGroupState state) => _matcher!.Separator();
    public override void FootnoteStart() => _matcher!.Separator();
    public override void ShapeStart(in RtfGroupState state) => _matcher!.Separator();
    public override void ShapeEnd(RtfShapePlacement shape, in RtfGroupState state) => _matcher!.Separator();
    public override void TextBoxParagraphEnd(in RtfGroupState state) => _matcher!.Separator();
    public override void FootnoteEnd(bool isEndnote) => _matcher!.Separator();
}
