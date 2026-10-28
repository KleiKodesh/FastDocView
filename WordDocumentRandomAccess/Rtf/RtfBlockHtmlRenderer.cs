using System.Globalization;
using System.Text;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.Rtf;

/// <summary>
/// Renders one indexed RTF block to HTML: restores the block's group-state stack, parses only the block's bytes and
/// turns the parser's events into paragraphs (or one table row) with inline formatting. Footnote groups inside the
/// block become a reference &lt;sup class="note-reference"&gt; plus a <see cref="WordDocumentNoteHtml"/> delivered
/// with the block (Word's draft view). Pictures, objects and equations are placeholders.
/// </summary>
internal sealed class RtfBlockHtmlRenderer : RtfSink
{
    private sealed class TableLevel
    {
        public readonly List<string> Cells = new();
        public readonly StringBuilder Cell = new();
    }

    /// <summary>Where paragraphs go: the block (or a note body) and the paragraph being built.</summary>
    private sealed class Flow
    {
        public readonly StringBuilder Output = new();
        public readonly StringBuilder Inline = new();
        public bool HasInline;
        public readonly List<TableLevel> Tables = new();
    }

    private readonly RtfParser _parser;
    private readonly RtfDocumentTables _tables;
    private readonly WordRunFormatting _documentRun;
    private Flow _flow = new();
    private Flow _bodyFlow = new();

    // the open run: text not yet written, and the state it was read in
    private readonly StringBuilder _runText = new();
    private RtfGroupState _runState;
    private StringBuilder? _listTextTarget;

    // per block
    private RtfNoteReference[] _blockNotes = [];
    private int _nextNote;
    private List<WordDocumentNoteHtml>? _renderedNotes;
    private string? _indexListLabel;
    private bool _isFirstParagraph;

    public RtfBlockHtmlRenderer(RtfDocumentTables tables)
    {
        _tables = tables;
        _parser = new RtfParser(this, tables, readsHeaderTables: false);
        _documentRun = DocumentRunFormatting(tables);
    }

    /// <summary>The run formatting the stylesheet gives the whole document (\deff font, 12 pt — \fs default 24).</summary>
    public static WordRunFormatting DocumentRunFormatting(RtfDocumentTables tables)
        => new() { FontName = tables.FontName(tables.DefaultFont), FontSizeHalfPoints = 24 };

    public (string html, IReadOnlyList<WordDocumentNoteHtml> notes) Render(Stream file, RtfRandomAccessIndex index, int blockIndex)
    {
        _bodyFlow = new Flow();
        _flow = _bodyFlow;
        _runText.Clear();
        _listTextTarget = null;
        ReadOnlySpan<RtfNoteReference> notes = index.NoteReferencesInBlock(blockIndex);
        _blockNotes = notes.ToArray();
        _nextNote = 0;
        _renderedNotes = notes.Length == 0 ? null : new List<WordDocumentNoteHtml>(notes.Length);
        _indexListLabel = index.ListLabelOf(blockIndex);
        _isFirstParagraph = true;

        _parser.RestoreSnapshot(index.States[index.BlockStateIds[blockIndex]]);
        _parser.Run(file, index.BlockStartOffsets[blockIndex], index.BlockEndOffsets[blockIndex]);

        // a block that ends without a terminator (the last one) or a row left open
        while (_flow != _bodyFlow) EndNote(isEndnote: false);
        FlushRun();
        if (_flow.HasInline) EmitParagraph(_parser.Current, isSectionEnd: false);
        for (int level = _flow.Tables.Count - 1; level >= 0; level--)
            if (_flow.Tables[level].Cells.Count > 0 || _flow.Tables[level].Cell.Length > 0) CloseRow(level + 1);
        return (_bodyFlow.Output.ToString(), (IReadOnlyList<WordDocumentNoteHtml>?)_renderedNotes ?? []);
    }

    // ── runs ─────────────────────────────────────────────────────────────────────────

    public override void Text(ReadOnlySpan<char> text, in RtfGroupState state)
    {
        if (_runText.Length > 0 && !HasSameCharacterFormatting(_runState, state)) FlushRun();
        _runState = state;
        _runText.Append(text.ToArray());
        if (_listTextTarget is null) _flow.HasInline = true;
    }

    private static bool HasSameCharacterFormatting(in RtfGroupState a, in RtfGroupState b)
        => (a.Flags & RtfStateFlags.CharacterMask) == (b.Flags & RtfStateFlags.CharacterMask)
           && a.Underline == b.Underline && a.Font == b.Font && a.ComplexFont == b.ComplexFont
           && a.FontSize == b.FontSize && a.ComplexFontSize == b.ComplexFontSize
           && a.ForegroundColor == b.ForegroundColor && a.BackgroundColor == b.BackgroundColor && a.ShadingColor == b.ShadingColor
           && a.Highlight == b.Highlight && a.Position == b.Position && a.CharacterSpacing == b.CharacterSpacing;

    private void FlushRun()
    {
        if (_runText.Length == 0) return;
        string text = _runText.ToString();
        _runText.Clear();
        StringBuilder target = _listTextTarget ?? _flow.Inline;
        string css = RunFormatting(_runState, IsComplexScript(text)).ToCss(_documentRun);
        string encoded = WordHtmlText.Encode(text);
        if (css.Length == 0) { target.Append(encoded); return; }
        target.Append("<span style=\"").Append(css).Append("\">").Append(encoded).Append("</span>");
    }

    /// <summary>Hebrew / Arabic text takes the complex-script slot (Word picks the slot by the characters' script).</summary>
    private static bool IsComplexScript(string text)
    {
        foreach (char c in text)
            if (c is >= '\u0590' and <= '\u08FF' or >= '\uFB1D' and <= '\uFDFF' or >= '\uFE70' and <= '\uFEFF') return true;
        return false;
    }

    private WordRunFormatting RunFormatting(in RtfGroupState state, bool isComplex)
    {
        bool complexFont = isComplex && state.Has(RtfStateFlags.ComplexFontSet);
        bool complexSize = isComplex && state.Has(RtfStateFlags.ComplexSizeSet);
        var run = new WordRunFormatting
        {
            IsBold = isComplex && state.Has(RtfStateFlags.ComplexBoldSet) ? state.Has(RtfStateFlags.ComplexBold) : state.Has(RtfStateFlags.Bold),
            IsItalic = isComplex && state.Has(RtfStateFlags.ComplexItalicSet) ? state.Has(RtfStateFlags.ComplexItalic) : state.Has(RtfStateFlags.Italic),
            FontName = _tables.FontName(complexFont ? state.ComplexFont : state.Font),
            FontSizeHalfPoints = complexSize ? state.ComplexFontSize : state.FontSize,
            IsStrikethrough = state.Has(RtfStateFlags.Strike),
            IsSuperscript = state.Has(RtfStateFlags.Superscript),
            IsSubscript = state.Has(RtfStateFlags.Subscript),
            IsHidden = state.Has(RtfStateFlags.Hidden),
            IsAllCaps = state.Has(RtfStateFlags.AllCaps),
            IsSmallCaps = state.Has(RtfStateFlags.SmallCaps),
            IsOutline = state.Has(RtfStateFlags.Outline),
            IsShadow = state.Has(RtfStateFlags.Shadow),
            IsEmboss = state.Has(RtfStateFlags.Emboss),
            IsImprint = state.Has(RtfStateFlags.Imprint),
            PositionHalfPoints = state.Position,
            CharacterSpacingTwips = state.CharacterSpacing,
            ColorHex = _tables.ColorHex(state.ForegroundColor),
            HighlightCssColor = HighlightCss(state.Highlight) ?? HexCss(_tables.ColorHex(state.ShadingColor)) ?? HexCss(_tables.ColorHex(state.BackgroundColor)),
        };
        if (state.Underline != RtfUnderline.None)
        {
            run.IsUnderlined = true;
            run.UnderlineStyleCss = state.Underline switch
            {
                RtfUnderline.Double => "double", RtfUnderline.Dotted => "dotted", RtfUnderline.Dashed => "dashed", RtfUnderline.Wave => "wavy", _ => null,
            };
            run.IsUnderlineThick = state.Underline == RtfUnderline.Thick;
        }
        return run;
    }

    private static string? HexCss(string? hex) => hex is null ? null : "#" + hex;

    /// <summary>\highlightN — the RTF 1.5 spec's fixed palette (LibreOffice's RTF export writes the same ico numbers).</summary>
    private static string? HighlightCss(byte index) => index switch
    {
        1 => "black", 2 => "blue", 3 => "cyan", 4 => "lime", 5 => "magenta", 6 => "red", 7 => "yellow",
        9 => "#000080", 10 => "#008080", 11 => "#008000", 12 => "#800080", 13 => "#800000", 14 => "#808000", 15 => "#808080", 16 => "#c0c0c0",
        _ => null,
    };

    // ── paragraphs ───────────────────────────────────────────────────────────────────

    public override void ParagraphEnd(in RtfGroupState state, bool isSectionEnd)
    {
        if (_listTextTarget is not null) return;                         // a \par inside \listtext: ignored
        FlushRun();
        EmitParagraph(state, isSectionEnd);
    }

    private void EmitParagraph(in RtfGroupState state, bool isSectionEnd)
    {
        bool isInTable = _flow == _bodyFlow && (state.IsInTable || _flow.Tables.Count > 0 && _flow.Tables[0].Cells.Count > 0);
        bool isTopLevel = _flow == _bodyFlow && !isInTable;
        WordParagraphFormatting formatting = ParagraphFormatting(state);
        int outline = isTopLevel ? (state.OutlineLevel is >= 0 and <= 8 ? state.OutlineLevel : state.OutlineLevel == 9 ? -1 : _tables.OutlineLevelOfStyle(state.Style)) : -1;
        string tag = outline is >= 0 and <= 5 ? "h" + (outline + 1) : "p";

        var html = new StringBuilder(_flow.Inline.Length + 64);
        if (state.Has(RtfStateFlags.PageBreakBefore) && isTopLevel) html.Append("<span class=\"page-break\"></span>");
        WordHtmlText.AppendOpenTag(html, tag, formatting.ToCss(), null, formatting.IsRightToLeft);
        if (isTopLevel && _isFirstParagraph && _indexListLabel is { Length: > 0 } label)
            html.Append("<span class=\"list-marker\">").Append(WordHtmlText.Encode(label)).Append("</span> ");
        html.Append(_flow.Inline);
        html.Append("</").Append(tag).Append('>');
        if (isSectionEnd && isTopLevel) html.Append("<hr class=\"section-break\">");
        _flow.Inline.Clear();
        _flow.HasInline = false;
        if (isTopLevel) _isFirstParagraph = false;

        if (isInTable) CellOf(Math.Max(1, (int)state.NestingLevel)).Cell.Append(html);
        else _flow.Output.Append(html);
    }

    private static WordParagraphFormatting ParagraphFormatting(in RtfGroupState state)
    {
        bool rtl = state.Has(RtfStateFlags.RightToLeftParagraph);
        var formatting = new WordParagraphFormatting
        {
            IsRightToLeft = rtl,
            // \ql / \qr are physical sides: Hebrew paragraphs in real files are \rtlpar\qr (start). \lin / \rin are the logical ones.
            Alignment = state.Alignment switch
            {
                RtfAlignment.Left => rtl ? WordParagraphAlignment.End : WordParagraphAlignment.Start,
                RtfAlignment.Right => rtl ? WordParagraphAlignment.Start : WordParagraphAlignment.End,
                RtfAlignment.Center => WordParagraphAlignment.Center,
                RtfAlignment.Justify => WordParagraphAlignment.Justify,
                _ => null,
            },
            LeftIndentTwips = state.Has(RtfStateFlags.HasStartIndent) ? state.StartIndent : rtl ? state.RightIndent : state.LeftIndent,
            RightIndentTwips = state.Has(RtfStateFlags.HasEndIndent) ? state.EndIndent : rtl ? state.LeftIndent : state.RightIndent,
            FirstLineIndentTwips = state.FirstLineIndent,
            SpaceBeforeTwips = state.SpaceBefore,
            SpaceAfterTwips = state.SpaceAfter,
        };
        // \slN: none / 1000 = auto; with \slmult1 a multiple of 240; else negative = exactly, positive = at least (RTF 1.5 spec)
        if (state.LineSpacing != 0 && state.LineSpacing != 1000)
        {
            if (state.Has(RtfStateFlags.LineSpacingMultiple)) { formatting.LineSpacingTwipsOrAuto = Math.Abs(state.LineSpacing); formatting.IsLineSpacingAuto = true; }
            else { formatting.LineSpacingTwipsOrAuto = Math.Abs(state.LineSpacing); formatting.IsLineSpacingAuto = false; }
        }
        return formatting;
    }

    public override void Break(RtfWord word, in RtfGroupState state)
    {
        FlushRun();
        StringBuilder target = _listTextTarget ?? _flow.Inline;
        if (word == RtfWord.Page) target.Append("<span class=\"page-break\"></span>");
        else if (word == RtfWord.Line) target.Append("<br>");
        _flow.HasInline = true;
    }

    public override void Embedded(RtfEmbeddedKind kind, in RtfGroupState state)
    {
        FlushRun();
        (_listTextTarget ?? _flow.Inline).Append(kind switch
        {
            RtfEmbeddedKind.Equation => "<span class=\"unsupported\">[equation]</span>",
            RtfEmbeddedKind.Object => "<span class=\"unsupported\">[object]</span>",
            _ => "<span class=\"unsupported\">[image]</span>",
        });
        _flow.HasInline = true;
    }

    // ── list text (Word's flat label for readers without list support) ───────────────

    public override void ListTextStart()
    {
        FlushRun();
        _listTextTarget = new StringBuilder();
    }

    public override void ListTextEnd()
    {
        FlushRun();
        StringBuilder? label = _listTextTarget;
        _listTextTarget = null;
        if (label is null) return;
        string html = label.ToString().Replace("\t", "");
        if (html.Length == 0) return;
        _flow.Inline.Append("<span class=\"list-marker\">").Append(html).Append("</span> ");
        _flow.HasInline = true;
    }

    // ── tables ───────────────────────────────────────────────────────────────────────

    private TableLevel CellOf(int level)
    {
        while (_flow.Tables.Count < level) _flow.Tables.Add(new TableLevel());
        return _flow.Tables[level - 1];
    }

    public override void CellEnd(in RtfGroupState state, bool isNested)
    {
        if (_flow != _bodyFlow) return;                                   // tables inside notes: text only
        FlushRun();
        int level = isNested ? Math.Max(2, (int)state.NestingLevel) : 1;
        if (_flow.HasInline)
        {
            var inCell = state;
            inCell.Set(RtfStateFlags.InTable, true);
            inCell.NestingLevel = (byte)level;
            EmitParagraph(inCell, isSectionEnd: false);
        }
        TableLevel table = CellOf(level);
        table.Cells.Add(table.Cell.ToString());
        table.Cell.Clear();
    }

    public override void RowEnd(in RtfGroupState state, bool isNested)
    {
        if (_flow != _bodyFlow) return;
        FlushRun();
        int level = isNested ? Math.Max(2, (int)state.NestingLevel) : 1;
        if (_flow.HasInline)
        {
            var inCell = state;
            inCell.Set(RtfStateFlags.InTable, true);
            inCell.NestingLevel = (byte)level;
            EmitParagraph(inCell, isSectionEnd: false);
        }
        CloseRow(level);
    }

    private void CloseRow(int level)
    {
        TableLevel table = CellOf(level);
        if (table.Cell.Length > 0) { table.Cells.Add(table.Cell.ToString()); table.Cell.Clear(); }
        var html = new StringBuilder(256);
        bool isOuter = level == 1;
        WordHtmlText.AppendOpenTag(html, "table", isOuter ? "table-layout:fixed;" : null, isOuter ? "table-row" : null, _parser.IsRowRightToLeft ? true : null);
        html.Append("<tr>");
        IReadOnlyList<int> boundaries = _parser.CellBoundaries;
        IReadOnlyList<short> shading = _parser.CellShading;
        for (int i = 0; i < table.Cells.Count; i++)
        {
            var css = new StringBuilder();
            if (isOuter && i < boundaries.Count)
            {
                int width = boundaries[i] - (i == 0 ? 0 : boundaries[i - 1]);
                if (width > 0) css.Append("width:").Append(WordParagraphFormatting.TwipsToPt(width)).Append("pt;");
                if (_tables.ColorHex(shading[i]) is { } fill) css.Append("background:#").Append(fill).Append(';');
            }
            html.Append("<td");
            if (css.Length > 0) html.Append(" style=\"").Append(css).Append('"');
            html.Append('>').Append(table.Cells[i]).Append("</td>");
        }
        html.Append("</tr></table>");
        table.Cells.Clear();
        if (level == 1) _flow.Output.Append(html);
        else CellOf(level - 1).Cell.Append(html);
    }

    // ── notes ────────────────────────────────────────────────────────────────────────

    public override void FootnoteStart()
    {
        FlushRun();
        _flow = new Flow();
    }

    public override void FootnoteMark(in RtfGroupState state)
    {
        if (_flow == _bodyFlow) return;                                   // the reference is written where the note group ends
        FlushRun();
        RtfNoteReference? reference = _nextNote < _blockNotes.Length ? _blockNotes[_nextNote] : null;
        _flow.Inline.Append("<span class=\"note-mark\">").Append(WordHtmlText.Encode(reference?.Label ?? "")).Append("</span>");
        _flow.HasInline = true;
    }

    public override void FootnoteEnd(bool isEndnote) => EndNote(isEndnote);

    private void EndNote(bool isEndnote)
    {
        FlushRun();
        if (_flow.HasInline) EmitParagraph(_parser.Current, isSectionEnd: false);
        string body = _flow.Output.ToString();
        _flow = _bodyFlow;
        RtfNoteReference? reference = _nextNote < _blockNotes.Length ? _blockNotes[_nextNote++] : null;
        WordDocumentNoteKind kind = reference?.Kind ?? (isEndnote ? WordDocumentNoteKind.Endnote : WordDocumentNoteKind.Footnote);
        int number = reference?.Number ?? 0;
        string label = reference?.Label ?? "*";
        string kindName = kind == WordDocumentNoteKind.Footnote ? "footnote" : "endnote";
        _flow.Inline.Append("<sup class=\"note-reference\" data-note-kind=\"").Append(kindName).Append("\" data-note-number=\"")
            .Append(number.ToString(CultureInfo.InvariantCulture)).Append("\">").Append(number > 0 ? WordHtmlText.Encode(label) : "").Append("</sup>");
        _flow.HasInline = true;
        _renderedNotes?.Add(new WordDocumentNoteHtml(kind, number, label, body));
    }
}
