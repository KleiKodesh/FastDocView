using System.Globalization;
using System.Text;
using FastDocView.Html;
using FastDocView.DocxHtml;

namespace FastDocView.Rtf;

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
        public (int position, string html, bool hasTab)? ListText;   // a \listtext label waiting for its paragraph
        public readonly List<TableLevel> Tables = new();
        public string? OpenLink;            // the <a> tag of a link field whose result is being read in this flow
        public bool IsTextBox;              // a {\shptxt}: its paragraphs are block spans (no <p> inside a <p>)
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

    // fields open here (restored per block from the index), and per open field the flow its link was opened in (null: no link)
    private RtfFieldTracker _fields;
    private readonly List<(int depth, Flow? linkFlow)> _fieldLinks = new();

    /// <summary>What fields Word saved without a result are computed from (the {\info} properties, bookmarks); null: the code only.</summary>
    public Fields.WordFieldContext? FieldContext
    {
        get => _fieldContext;
        set { _fieldContext = value; _fields = new RtfFieldTracker(value, counts: true); }
    }
    private Fields.WordFieldContext? _fieldContext;

    public RtfBlockHtmlRenderer(RtfDocumentTables tables)
    {
        _tables = tables;
        _parser = new RtfParser(this, tables, readsHeaderTables: false);
        _documentRun = DocumentRunFormatting(tables);
        _fields = new RtfFieldTracker(null, counts: true);
    }

    /// <summary>The run formatting the stylesheet gives the whole document (\deff font, 12 pt — \fs default 24).</summary>
    public static WordRunFormatting DocumentRunFormatting(RtfDocumentTables tables)
        => new() { FontName = tables.FontName(tables.DefaultFont), FontSizeHalfPoints = 24 };

    public (string html, IReadOnlyList<WordDocumentNoteHtml> notes) Render(Stream file, RtfRandomAccessIndex index, RtfCompactBlockTable table, int blockIndex)
    {
        _blockIndex = blockIndex;
        _pictureOrdinal = 0;
        _bodyFlow = new Flow();
        _flow = _bodyFlow;
        _runText.Clear();
        _listTextTarget = null;
        ReadOnlySpan<RtfNoteReference> notes = index.NoteReferencesInBlock(blockIndex);
        _blockNotes = notes.ToArray();
        _nextNote = 0;
        _renderedNotes = notes.Length == 0 ? null : new List<WordDocumentNoteHtml>(notes.Length);
        _indexListLabel = table.ListLabelOf(blockIndex);
        _isFirstParagraph = true;
        _fields.Restore(table.FieldStateOf(blockIndex));
        _fieldLinks.Clear();
        for (int i = 0; i < _fields.Depth; i++) _fieldLinks.Add((0, null));
        // a link field whose result runs in from an earlier block stays a link here
        if (_fields.ActiveLink is { } runningLink && _fieldLinks.Count > 0) { _flow.OpenLink = runningLink; _flow.Inline.Append(runningLink); _fieldLinks[_fieldLinks.Count - 1] = (0, _flow); }
        _textBoxOuterFlow = null;
        _textBoxShape = null;

        _parser.RestoreSnapshot(table.SnapshotOf(blockIndex));
        _parser.Run(file, table.BlockStartOffset(blockIndex), table.BlockEndOffset(blockIndex));

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
        if (!_fields.IsIdle && !_fields.Shows(text)) return;                 // a replaced result (page placeholder, recomputed, hidden)
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
           && a.Highlight == b.Highlight && a.Position == b.Position && a.CharacterSpacing == b.CharacterSpacing && a.CharacterScale == b.CharacterScale
           && a.UnderlineColor == b.UnderlineColor && a.EmphasisMark == b.EmphasisMark && a.CharacterPatternColor == b.CharacterPatternColor
           && a.CharacterShading == b.CharacterShading && a.CharacterBorder.Equals(b.CharacterBorder);

    private void FlushRun()
    {
        if (_runText.Length == 0) return;
        string text = _runText.ToString();
        _runText.Clear();
        StringBuilder target = _listTextTarget ?? _flow.Inline;
        // \rtlch marks a right-to-left run (RTF 1.5 spec); inside an \ltrpar paragraph it is an RTL embedding, as w:rtl is
        bool isRunRightToLeft = _runState.RunType is RtfRunType.RightToLeftFirst or RtfRunType.LeftToRightThenRightToLeft;
        string css = RunFormatting(_runState, IsComplexScript(text)).ToCss(_documentRun) + BaselineShiftCss
                     + WordHtmlText.RunDirectionCss(isRunRightToLeft, _runState.Has(RtfStateFlags.RightToLeftParagraph));
        string encoded = WordHtmlText.Encode(text);
        if (css.Length == 0) { WordHtmlText.AppendScaled(target, encoded, _runState.CharacterScale); return; }
        target.Append("<span style=\"").Append(css).Append("\">");
        WordHtmlText.AppendScaled(target, encoded, _runState.CharacterScale);
        target.Append("</span>");
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
            CharacterScalePercent = state.CharacterScale,
            ColorHex = _tables.ColorHex(state.ForegroundColor),
            HighlightCssColor = HighlightCss(state.Highlight),
            // \chcbpat fill + \chcfpat pattern colour at \chshdng hundredths of a percent, mixed as w:shd (WordShadingColor)
            ShadingHex = WordShadingColor.Mix(_tables.ColorHex(state.CharacterPatternColor), _tables.ColorHex(state.ShadingColor), state.CharacterShading / 10),
            UnderlineColorHex = _tables.ColorHex(state.UnderlineColor),
            BorderCss = BorderCss(state.CharacterBorder),
            // the .docx w:em mapping (DocxStyleCascade): dot / under dot = filled dot, comma = filled sesame, circle = open circle
            EmphasisMarkCss = state.EmphasisMark switch { 1 or 4 => "filled dot", 2 => "filled sesame", 3 => "open circle", _ => null },
            IsEmphasisMarkBelow = state.EmphasisMark == 4,
        };
        if (run.HighlightCssColor is null && run.ShadingHex is null) run.HighlightCssColor = HexCss(_tables.ColorHex(state.BackgroundColor));   // \chcbpat (w:shd) wins over \cb
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

    /// <summary>
    /// One RTF border as CSS, as the .docx border path writes it (DocxTableFormatting.BorderCss): \brdrwN twips (LibreOffice
    /// rtfdispatchvalue.cxx:1032-1041 BRDRW -> eighths of a point, the same width), 0.5 pt when not given (the .docx sz default),
    /// \brdrcfN from the colour table, auto = black. Null = no border set, "none" = \brdrnone.
    /// </summary>
    private string? BorderCss(in RtfBorder border)
    {
        if (border.Style == RtfBorderStyle.Absent) return null;
        if (border.Style == RtfBorderStyle.None) return "none";
        double points = border.Width > 0 ? Math.Max(0.25, border.Width / 20.0) : 0.5;
        string style = border.Style switch { RtfBorderStyle.Double => "double", RtfBorderStyle.Dotted => "dotted", RtfBorderStyle.Dashed => "dashed", _ => "solid" };
        return points.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt " + style + " " + (HexCss(_tables.ColorHex(border.Color)) ?? "#000");
    }

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
        bool isInTable = (_flow == _bodyFlow || _flow.IsTextBox) && (state.IsInTable || _flow.Tables.Count > 0 && _flow.Tables[0].Cells.Count > 0);
        bool isTopLevel = _flow == _bodyFlow && !isInTable;
        WordParagraphFormatting formatting = ParagraphFormatting(state);
        int outline = isTopLevel ? (state.OutlineLevel is >= 0 and <= 8 ? state.OutlineLevel : state.OutlineLevel == 9 ? -1 : _tables.OutlineLevelOfStyle(state.Style)) : -1;
        string tag = _flow.IsTextBox ? "span" : outline is >= 0 and <= 5 ? "h" + (outline + 1) : "p";

        // a link field whose result spans paragraphs: closed with each paragraph and reopened in the next
        string? openLink = _flow.OpenLink;
        if (openLink is not null) _flow.Inline.Append("</a>");
        var html = new StringBuilder(_flow.Inline.Length + 64);
        if (state.Has(RtfStateFlags.PageBreakBefore) && isTopLevel) html.Append("<span class=\"page-break\"></span>");
        WordHtmlText.AppendOpenTag(html, tag, formatting.ToCss() + (_flow.IsTextBox ? "display:block;" : ""), _flow.IsTextBox ? "text-box-paragraph" : null, formatting.IsRightToLeft);
        // The label sits at the first-line position and a tab after it reaches the next stop (WordListLabelHtml): the flat
        // \listtext label Word wrote (its \tab = the tab suffix, RTF 1.5 \listtext), or the label counted at index time
        // (its \levelfollow); \leveljc places the label.
        RtfDocumentTables.ListLevelDefinition? listLevel = RtfListNumbering.LevelOf(_tables, state.ListOverride, state.ListLevel);
        WordListLabelAlignment labelAlignment = listLevel?.Justification switch { 1 => WordListLabelAlignment.Center, 2 => WordListLabelAlignment.End, _ => WordListLabelAlignment.Start };
        if (_flow.ListText is { } listText)
        {
            html.Append(_flow.Inline.ToString(0, listText.position));            // net48 has no Append(StringBuilder, int, int)
            WordListLabelHtml.Append(html, listText.html, null, listText.hasTab ? WordListLabelSuffix.Tab : WordListLabelSuffix.Nothing, formatting,
                                     WordListLabelHtml.DefaultTabStopTwips, labelAlignment, isLabelHtml: true);
            html.Append(_flow.Inline.ToString(listText.position, _flow.Inline.Length - listText.position));
            _flow.ListText = null;
        }
        else
        {
            if (isTopLevel && _isFirstParagraph && _indexListLabel is { Length: > 0 } label)
                WordListLabelHtml.Append(html, label, null, listLevel?.Follow switch { 1 => WordListLabelSuffix.Space, 2 => WordListLabelSuffix.Nothing, _ => WordListLabelSuffix.Tab },
                                         formatting, WordListLabelHtml.DefaultTabStopTwips, labelAlignment);
            html.Append(_flow.Inline);
        }
        html.Append("</").Append(tag).Append('>');
        if (isSectionEnd && isTopLevel) html.Append("<hr class=\"section-break\">");
        _flow.Inline.Clear();
        _flow.HasInline = false;
        if (openLink is not null) _flow.Inline.Append(openLink);
        if (isTopLevel) _isFirstParagraph = false;
        _baselineShift = 0;                                            // an ADVANCE up / down shift lasts to the paragraph end

        if (isInTable) CellOf(Math.Max(1, (int)state.NestingLevel)).Cell.Append(html);
        else _flow.Output.Append(html);
    }

    private WordParagraphFormatting ParagraphFormatting(in RtfGroupState state)
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
            // \cbpat fill + \cfpat pattern colour at \shading hundredths of a percent, mixed as the .docx pPr w:shd
            ShadingHex = WordShadingColor.Mix(_tables.ColorHex(state.ParagraphPatternColor), _tables.ColorHex(state.ParagraphBackgroundColor), state.ParagraphShading / 10),
            // \brdrl / \brdrr -> w:pBdr left / right (LibreOffice rtfdispatchflag.cxx:762-795), which the .docx and .doc paths draw as start / end
            BorderTopCss = BorderCss(state.BorderTop),
            BorderBottomCss = BorderCss(state.BorderBottom),
            BorderStartCss = BorderCss(state.BorderLeft),
            BorderEndCss = BorderCss(state.BorderRight),
        };
        // \slN: none / 1000 = auto; with \slmult1 a multiple of 240; else negative = exactly, positive = at least (RTF 1.5 spec)
        if (state.LineSpacing != 0 && state.LineSpacing != 1000)
        {
            if (state.Has(RtfStateFlags.LineSpacingMultiple)) { formatting.LineSpacingTwipsOrAuto = Math.Abs(state.LineSpacing); formatting.IsLineSpacingAuto = true; }
            else { formatting.LineSpacingTwipsOrAuto = Math.Abs(state.LineSpacing); formatting.IsLineSpacingAuto = false; formatting.IsLineSpacingExact = state.LineSpacing < 0; }   // \sl: negative = exactly, positive = at least (RTF 1.5)
        }
        return formatting;
    }

    public override void Break(RtfWord word, in RtfGroupState state)
    {
        FlushRun();
        StringBuilder target = _listTextTarget ?? _flow.Inline;
        if (word == RtfWord.Page) target.Append("<span class=\"page-break\"></span>");
        else if (word == RtfWord.Line)
            target.Append((state.Flags & RtfStateFlags.LineBreakClearMask) switch
            {
                RtfStateFlags.LineBreakClearLeft => "<br style=\"clear:left\">",
                RtfStateFlags.LineBreakClearRight => "<br style=\"clear:right\">",
                RtfStateFlags.LineBreakClearMask => "<br style=\"clear:both\">",
                _ => "<br>",
            });
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

    public override bool WantsPictures => true;
    public override bool WantsPictureData => _capturedPictureOrdinal >= 0;

    private int _blockIndex;
    private int _pictureOrdinal;                 // pictures met in this block so far (their ids count them)
    private int _capturedPictureOrdinal = -1;    // >= 0: ReadPicture wants that picture's bytes
    private RtfPicture? _capturedPicture;

    /// <summary>
    /// A {\pict} in the text: an inline picture at its goal size. A {\shp} with a picture floats as the shape's placement
    /// says (meanings in RtfParser.ReadShape): \shpwr 1 = above and below, 2 / 4 / 5 = beside, 3 = through the text; \shpwr
    /// not given is drawn above and below (a block - our choice, the shape group gives no wrap); \shpwrk 1 / 2 = text only
    /// on that side; \shpwr 4 / 5 along its wrap polygon (else its outline); posh 1 / 4 left, 2 centre, 3 / 5 right, else at \shpleft from the left. Its id is the block and the
    /// picture's place among the block's pictures.
    /// </summary>
    public override void Picture(in RtfPicture picture, in RtfGroupState state)
    {
        string? pictureId = PictureIdOf(picture);
        FlushRun();
        WordPictureLayout layout = new(picture.WidthTwips / 20.0, picture.HeightTwips / 20.0, WordPictureWrap.Inline, WordPictureSide.Left, 0, 0, 0, 0);
        if (picture.Placement is { } shape) layout = ShapeLayout(shape, layout.WidthPt, layout.HeightPt);
        WordPictureHtml.Append(_listTextTarget ?? _flow.Inline, layout, pictureId);
        _flow.HasInline = true;
    }

    /// <summary>The next picture's id (the block and its place among the block's pictures); the one <see cref="ReadPicture"/> asked for is kept.</summary>
    private string? PictureIdOf(in RtfPicture picture)
    {
        int ordinal = _pictureOrdinal++;
        if (ordinal == _capturedPictureOrdinal) _capturedPicture = picture;
        return picture.Extension.Length > 0 && WordPictureHtml.ContentTypeOf(picture.Extension) is not null
            ? _blockIndex.ToString(CultureInfo.InvariantCulture) + ":" + ordinal.ToString(CultureInfo.InvariantCulture)
            : null;
    }

    /// <summary>A floating shape's placement (meanings in RtfParser.OpenShape), size in points.</summary>
    /// <summary>The text width (points) of the section of the block being rendered: the column a floating object picks its nearer edge in (0 = unknown).</summary>
    public double ColumnWidthPt { get; set; }
    /// <summary>Where that column starts from the page's left edge (points; -1 = not known).</summary>
    public double ColumnLeftPt { get; set; } = -1;

    private WordPictureLayout ShapeLayout(RtfShapePlacement shape, double widthPt, double heightPt)
    {
        const double EmuPerPoint = 12700;
        WordPictureWrap wrap = shape.Wrap switch { 2 or 4 or 5 => WordPictureWrap.Square, 3 => WordPictureWrap.None, _ => WordPictureWrap.TopAndBottom };
        WordPictureSide side = shape.HorizontalPosition switch
        {
            1 or 4 => WordPictureSide.Left,
            2 => WordPictureSide.Center,
            3 or 5 => WordPictureSide.Right,
            _ => shape.WrapSide switch { 1 => WordPictureSide.Right, _ => WordPictureSide.Left },
        };
        double offset = shape.HorizontalPosition == 0 ? Math.Max(0, shape.Left / 20.0) : 0;
        // \shpwrk: 1 = text only on the left, 2 = only on the right, 0 / 3 both / largest (rtfdispatchvalue.cxx:1205-1224 -> ST_WrapText)
        WordPictureTextSides sides = shape.WrapSide switch { 1 => WordPictureTextSides.LeftOnly, 2 => WordPictureTextSides.RightOnly, _ => WordPictureTextSides.Both };
        // \shpwr 4 tight / 5 through (Word 97 RTF for Drawing Objects: "5 Wrap through"; Word's through = tight plus the
        // polygon's holes - MS-DOC FSPA wr 5 "wrap tightly, but allow holes") along pWrapPolygonVertices (rtfsdrimport.cxx:856-885).
        // LibreOffice reads 5 as THROUGH = text over the object (rtfdispatchvalue.cxx:1255-1257); Word does not.
        IReadOnlyList<IReadOnlyList<(double X, double Y)>>? contour = shape.Wrap is 4 or 5 ? WordPictureContour.FromWordUnits(shape.WrapPolygon, widthPt, heightPt) : null;
        return new WordPictureLayout(widthPt, heightPt, wrap, side, shape.WrapTop / EmuPerPoint, shape.WrapRight / EmuPerPoint,
                                     shape.WrapBottom / EmuPerPoint, shape.WrapLeft / EmuPerPoint, offset, sides, contour,
                                     ColumnWidthPt, shape.IsPageRelative ? WordPictureOffsetOrigin.Page : shape.IsCharacterRelative ? WordPictureOffsetOrigin.Other : WordPictureOffsetOrigin.Column, ColumnLeftPt);
    }

    /// <summary>WordArt text centred in its font and size (msdffdef.hxx:121-125 gtextUNICODE / gtextSize / gtextFont), known once the shape closed.</summary>
    private static void WrapWordArt(RtfShapePlacement shape)
    {
        foreach (RtfShapePlacement child in shape.Children) WrapWordArt(child);
        if (!shape.IsWordArt || shape.TextHtml is null) return;
        var style = new StringBuilder("display:block;text-align:center;");
        if (shape.WordArtFont is { Length: > 0 } font) style.Append("font-family:'").Append(WordHtmlText.Encode(font.Replace("'", ""))).Append("';");
        if (shape.IsTextEffect && RtfShapeDrawing.FillColorCss(shape) is { } color) style.Append("color:").Append(color).Append(';');
        if (shape.WordArtSizePt > 0) style.Append("font-size:").Append(shape.WordArtSizePt.ToString("0.##", CultureInfo.InvariantCulture)).Append("pt;");
        shape.TextHtml = "<span class=\"word-art\" style=\"" + style + "\">" + shape.TextHtml + "</span>";
    }

    // ── drawing objects ({\shp} / {\shpgrp}) ────────────────────────────────────────

    private Flow? _textBoxOuterFlow;

    public override void ShapeStart(in RtfGroupState state) => FlushRun();

    private RtfShapePlacement? _textBoxShape;

    public override void TextBoxStart(RtfShapePlacement shape)
    {
        FlushRun();
        _textBoxShape = shape;
        _textBoxOuterFlow = _flow;
        _flow = new Flow { IsTextBox = true };
    }

    public override void TextBoxParagraphEnd(in RtfGroupState state)
    {
        if (_listTextTarget is not null) return;
        FlushRun();
        EmitParagraph(state, isSectionEnd: false);
    }

    public override void TextBoxEnd()
    {
        if (_textBoxOuterFlow is null) return;
        FlushRun();
        if (_flow.HasInline) EmitParagraph(_parser.Current, isSectionEnd: false);
        for (int level = _flow.Tables.Count - 1; level >= 0; level--)
            if (_flow.Tables[level].Cells.Count > 0 || _flow.Tables[level].Cell.Length > 0) CloseRow(level + 1);
        string html = _flow.Output.ToString();
        _flow = _textBoxOuterFlow;
        _textBoxOuterFlow = null;
        if (_textBoxShape is not null) _textBoxShape.TextHtml = html;
        _textBoxShape = null;
    }

    /// <summary>
    /// A drawing object: a lone pib picture as before (<see cref="Picture"/>); a plain rectangle text box (shapeType unset, 1 or
    /// 202 = msosptTextBox, not rotated) as a <see cref="WordPictureHtml.OpenTextBox"/> box with the shape's line as its border
    /// and its fill; anything else - other shapes, groups - as <see cref="WordShapeHtml"/>: one frame at the root's placement,
    /// every outline in one SVG, each text box over its shape, group pictures at their box. A group child's box is its relLeft..
    /// relBottom in the group's groupLeft..groupBottom, scaled to the group's own box (LibreOffice rtfsdrimport.cxx:1036-1053).
    /// </summary>
    public override void ShapeEnd(RtfShapePlacement root, in RtfGroupState state)
    {
        FlushRun();
        WrapWordArt(root);
        double width = Math.Max(0, root.Right - root.Left) / 20.0, height = Math.Max(0, root.Bottom - root.Top) / 20.0;
        if (!root.IsGroup && !root.HasText && root.Picture is { } picture)
        {
            Picture(picture with
            {
                WidthTwips = root.Right > root.Left ? root.Right - root.Left : picture.WidthTwips,
                HeightTwips = root.Bottom > root.Top ? root.Bottom - root.Top : picture.HeightTwips,
                Placement = root,
            }, state);
            return;
        }
        StringBuilder target = _listTextTarget ?? _flow.Inline;
        WordPictureLayout layout = ShapeLayout(root, width, height);
        if (!root.IsGroup && root.HasText && root.ShapeType is -1 or 1 or 202 && root.FillType == 0 && root.FillPicture is null && root.RotationDegrees == 0 && !root.FlipHorizontal && !root.FlipVertical)
        {
            WordTextBoxStyle style = RtfShapeDrawing.TextStyle(root);
            style = style with
            {
                BorderColor = RtfShapeDrawing.LineColorCss(root), BorderWidthPt = RtfShapeDrawing.LineWidthPt(root),
                FillColor = RtfShapeDrawing.FillColorCss(root), GrowsWithText = root.FitsShapeToText,
            };
            WordPictureHtml.OpenTextBox(target, layout, style);
            target.Append(root.TextHtml);
            WordPictureHtml.CloseTextBox(target);
            _flow.HasInline = true;
            return;
        }
        var shapes = new List<WordShape>();
        var texts = new List<string>();
        RtfShapeDrawing.Collect(root, 0, 0, width, height, shapes, texts, (in RtfPicture p) => PictureIdOf(p));
        bool inFlow = !root.IsGroup && texts.Count == 1 && root.FitsShapeToText && root.RotationDegrees == 0;
        // a tight wrap without a wrap polygon follows the shape's own outline
        if (root.Wrap is 4 or 5 && layout.Wrap == WordPictureWrap.Square && layout.Contour is null) layout = layout with { Contour = WordShapeHtml.Outlines(shapes) };
        WordShapeHtml.OpenFrame(target, layout, inFlow);
        WordShapeHtml.AppendGeometry(target, shapes, width, height, stretches: inFlow);
        foreach (WordShape shape in shapes)
        {
            if (shape.Geometry == WordShapeGeometry.Picture) { WordShapeHtml.AppendPicture(target, shape); continue; }
            if (shape.TextIndex < 0) continue;
            WordShapeHtml.OpenShapeText(target, shape, inFlow);
            target.Append(texts[shape.TextIndex]);
            WordShapeHtml.CloseShapeText(target);
        }
        WordShapeHtml.CloseFrame(target);
        _flow.HasInline = true;
    }

    /// <summary>The bytes of picture <paramref name="ordinal"/> of block <paramref name="blockIndex"/>: the block parsed again, that picture kept.</summary>
    public RtfPicture? ReadPicture(Stream file, RtfRandomAccessIndex index, RtfCompactBlockTable table, int blockIndex, int ordinal)
    {
        _capturedPictureOrdinal = ordinal;
        _capturedPicture = null;
        try { Render(file, index, table, blockIndex); return _capturedPicture; }
        finally { _capturedPictureOrdinal = -1; _capturedPicture = null; }
    }

    // ── list text (Word's flat label for readers without list support) ───────────────

    // fields and bookmarks

    public override void FieldStart(int depth)
    {
        _fields.Start(depth);
        _fieldLinks.Add((depth, null));
    }

    public override void FieldModifier(RtfWord word)
    {
        if (word == RtfWord.FieldLock) _fields.Lock(); else _fields.Private();
    }

    public override void FormField(Fields.WordFormFieldState state) => _fields.Form(state);

    public override void FieldInstruction(string instruction) => _fields.Instruction(instruction);

    /// <summary>A computed result or placeholder, in the current run's formatting, in the field's link when no link is open.</summary>

    /// <summary>ADVANCE \u / \d: points the paragraph's later runs are raised (negative: lowered), as a relative offset.</summary>
    private double _baselineShift;
    private string BaselineShiftCss => _baselineShift == 0 ? "" : "position:relative;top:" + (-_baselineShift).ToString(System.Globalization.CultureInfo.InvariantCulture) + "pt;";

    private void AppendFieldOutput(Fields.WordFieldOutput output, string? link)
    {
        if (output.IsEmpty || _listTextTarget is not null) return;
        FlushRun();
        _baselineShift += output.BaselineShiftPt;
        if (output.Text is null && output.Html is null) return;
        bool wrap = link is not null && _flow.OpenLink is null;
        if (wrap) _flow.Inline.Append(link);
        if (output.Html is { } markup) _flow.Inline.Append(markup);
        else
        {
            string css = RunFormatting(_parser.Current, IsComplexScript(output.Text!)).ToCss(_documentRun);
            string encoded = WordHtmlText.Encode(output.Text!);
            if (css.Length == 0) _flow.Inline.Append(encoded); else _flow.Inline.Append("<span style=\"").Append(css).Append("\">").Append(encoded).Append("</span>");
        }
        if (wrap) _flow.Inline.Append("</a>");
        _flow.HasInline = true;
    }

    /// <summary>
    /// The result of a HYPERLINK (\l "bookmark" = in-document) or REF / PAGEREF / NOTEREF ... \h field becomes a link, read
    /// with the .docx field rules (DocxBlockHtmlRenderer.HyperlinkFieldTag: Word writes the same instruction text in every
    /// format); only safe targets. A link inside a link renders its text only.
    /// </summary>
    public override void FieldResultStart()
    {
        if (_fields.Depth == 0) return;
        Fields.WordFieldOutput output = _fields.ResultStart(out string? link);
        AppendFieldOutput(output, link);
        if (_listTextTarget is not null || _flow.OpenLink is not null || _fields.Innermost is not { SuppressesResult: false, IsInsideInstruction: false, LinkTag: { } tag }) return;
        if (!_fields.Shows(default)) return;
        FlushRun();
        _flow.Inline.Append(tag);
        _flow.OpenLink = tag;
        _fieldLinks[_fieldLinks.Count - 1] = (_fieldLinks[_fieldLinks.Count - 1].depth, _flow);
    }

    public override void GroupClosed(int depth)
    {
        if (_fields.IsIdle) return;
        int before = _fields.Depth;
        var outputs = new List<(Fields.WordFieldOutput output, string? link)>();
        _fields.GroupClosed(depth, (output, link) => outputs.Add((output, link)));
        for (int ended = before - _fields.Depth; ended > 0 && _fieldLinks.Count > 0; ended--)
        {
            Flow? linkFlow = _fieldLinks[_fieldLinks.Count - 1].linkFlow;
            _fieldLinks.RemoveAt(_fieldLinks.Count - 1);
            if (linkFlow is null) continue;
            FlushRun();
            linkFlow.Inline.Append("</a>");
            linkFlow.OpenLink = null;
        }
        foreach ((Fields.WordFieldOutput output, string? link) in outputs) AppendFieldOutput(output, link);
    }

    /// <summary>{\*\bkmkend name}: an empty marker where the bookmark ends (REF without a result reads up to it).</summary>
    public override void BookmarkEnd(string name)
    {
        FlushRun();
        (_listTextTarget ?? _flow.Inline).Append("<span class=\"bookmark-end\" data-bookmark=\"").Append(WordHtmlText.Encode(name)).Append("\"></span>");
    }

    /// <summary>{\*\bkmkstart name}: the jump target of in-document links.</summary>
    public override void BookmarkStart(string name)
    {
        FlushRun();
        (_listTextTarget ?? _flow.Inline).Append("<span class=\"bookmark\" id=\"").Append(WordHtmlText.Encode(DocxBlockHtmlRenderer.BookmarkElementId(name))).Append("\"></span>");
    }

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
        string text = label.ToString();
        string html = text.Replace("\t", "");
        if (html.Length == 0) return;
        _flow.ListText = (_flow.Inline.Length, html, text.Contains('\t'));     // drawn by EmitParagraph, which knows the indents
        _flow.HasInline = true;
    }

    // ── tables ───────────────────────────────────────────────────────────────────────

    private TableLevel CellOf(int level)
    {
        while (_flow.Tables.Count < level) _flow.Tables.Add(new TableLevel());
        return _flow.Tables[level - 1];
    }

    public override void TextBoxCellEnd(in RtfGroupState state, bool isNested) { if (_listTextTarget is null) CellEnd(state, isNested); }
    public override void TextBoxRowEnd(in RtfGroupState state, bool isNested) { if (_listTextTarget is null) RowEnd(state, isNested); }

    public override void CellEnd(in RtfGroupState state, bool isNested)
    {
        if (_flow != _bodyFlow && !_flow.IsTextBox) return;               // tables inside notes: text only; a text box draws its own
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
        if (_flow != _bodyFlow && !_flow.IsTextBox) return;
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
        // An outer row is as wide as its cells (\cellx are right edges from the column's left, the first cell starts at \trleft -
        // LibreOffice rtfdocumentimpl.cxx: cell width = cellx - previous, first = cellx - trleft; trleft = the table indent);
        // \trqc / \trqr centre it / put it at the right.
        string? rowCss = null;
        if (isOuter)
        {
            IReadOnlyList<int> edges = _parser.CellBoundaries;
            int rowLeft = _parser.RowLeftTwips;
            int rowWidth = edges.Count > 0 ? edges[edges.Count - 1] - rowLeft : 0;
            rowCss = "table-layout:fixed;" + (rowWidth > 0 ? "width:" + WordParagraphFormatting.TwipsToPt(rowWidth) + "pt;" : "")
                   + (_parser.RowAlignment switch
                   {
                       1 => "margin-left:auto;margin-right:auto;",
                       2 => "margin-left:auto;margin-right:0;",
                       // 	rleft: "the leftmost edge of the table with respect to the left edge of its column" (RTF 1.5 spec,
                       // Row Formatting) - a physical left, also in a right-to-left row
                       _ => "margin-left:" + (rowLeft != 0 ? WordParagraphFormatting.TwipsToPt(rowLeft) + "pt" : "0") + ";margin-right:auto;",
                   });
        }
        WordHtmlText.AppendOpenTag(html, "table", rowCss, isOuter ? "table-row" : null, _parser.IsRowRightToLeft ? true : null);
        html.Append("<tr>");
        IReadOnlyList<int> boundaries = _parser.CellBoundaries;
        IReadOnlyList<short> shading = _parser.CellShading;
        for (int i = 0; i < table.Cells.Count; i++)
        {
            var css = new StringBuilder();
            if (isOuter && i < boundaries.Count)
            {
                int width = boundaries[i] - (i == 0 ? _parser.RowLeftTwips : boundaries[i - 1]);
                if (width > 0) css.Append("width:").Append(WordParagraphFormatting.TwipsToPt(width)).Append("pt;");
                if (_tables.ColorHex(shading[i]) is { } fill) css.Append("background:").Append(WordHtmlColor.Css(fill)).Append(';');
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
