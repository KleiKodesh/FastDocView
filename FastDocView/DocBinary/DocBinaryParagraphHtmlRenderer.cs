using System.Text;
using FastDocView.Html;

namespace FastDocView.DocBinary;

/// <summary>
/// Renders .doc text to HTML. A paragraph (CP range): PAPX → style chain + direct paragraph sprms,
/// then CHPX runs across its pieces for character formatting, with Word's control characters
/// mapped (tabs, line breaks, field codes hidden, special characters, note references). A table
/// row (CP range of its cells up to the row-end mark): one &lt;table&gt;&lt;tr&gt; whose cell
/// widths come from the row-end paragraph's sprmTDefTable. A note body: its paragraphs, with the
/// leading auto-number character shown as the note's label.
/// </summary>
public sealed class DocBinaryParagraphHtmlRenderer
{
    private const char CellOrRowMark = '\u0007';
    private const char AutoNumberedNoteMark = '\u0002';

    private readonly DocBinaryPieceTable _pieces;
    private readonly DocBinaryFormattedDiskPages _pages;
    private readonly DocBinaryStyleSheet _styles;
    private readonly DocBinaryNoteReferences? _notes;

    /// <summary>
    /// Table facts of a paragraph's PAPX: inside a table, its depth (sprmPItap), whether it ends an outer row (sprmPFTtp), a
    /// nested cell (sprmPFInnerTableCell) or a nested row (sprmPFInnerTtp), and — on a row end — the cell edges and widths
    /// (twips), the cells (merges, vertical alignment, own borders), right-to-left column order, per-cell shading and the
    /// table's six default borders (top, left, bottom, right, inside horizontal, inside vertical; sprmTTableBorders).
    /// </summary>
    public readonly record struct TableParagraphInfo(bool IsInTable, bool IsRowEnd, int[]? CellWidthsTwips,
                                                     bool IsRightToLeft = false, string?[]? CellShadingHex = null, string?[]? DefaultBorders = null,
                                                     int Justification = 0, int? LeftEdgeTwips = null, int? RowWidthTwips = null,
                                                     int Depth = 1, bool IsInnerCellEnd = false, bool IsInnerRowEnd = false,
                                                     int[]? CellEdgesTwips = null, DocBinaryTableCell[]? Cells = null, int RowHeightTwips = 0);

    public DocBinaryParagraphHtmlRenderer(DocBinaryPieceTable pieces, DocBinaryFormattedDiskPages pages, DocBinaryStyleSheet styles,
                                          DocBinaryNoteReferences? notes = null)
    {
        _pieces = pieces;
        _pages = pages;
        _styles = styles;
        _notes = notes;
    }

    /// <summary>The document's list tables: list paragraphs take their level's indents and label formatting.</summary>
    public DocBinaryListDefinitions Lists { get; set; } = DocBinaryListDefinitions.Empty;

    /// <summary>The default tab stop interval (Dop dxaTab), twips.</summary>
    public int DefaultTabStopTwips { get; set; } = WordListLabelHtml.DefaultTabStopTwips;

    /// <summary>Dop fAutoHyphen: the document hyphenates automatically (LibreOffice ww8scan.cxx WW8Dop byte 5 bit 0x10; ww8par2.cxx sets it on the base style).</summary>
    public bool IsAutoHyphenation { get; set; }

    /// <summary>Whether the paragraph [cpStart, cpEnd) is a drop-cap frame (sprmPDcs fdct 1 or 2, LibreOffice ww8par6.cxx IsDropCap).</summary>
    public bool IsDropCapFrame(int cpStart, int cpEnd) => ResolveParagraph(cpStart, cpEnd).formatting.DropCapLines > 0;
    /// <summary>Bookmarks: each one's start renders as an empty &lt;span class="bookmark" id&gt; (the target of in-document links).</summary>
    public DocBinaryBookmarks Bookmarks { get; set; } = DocBinaryBookmarks.Empty;

    /// <summary>The field tables (flags, fields open at a block start, counters); <see cref="DocBinaryFields.Empty"/> without them.</summary>
    internal DocBinaryFields FieldTable { get; set; } = DocBinaryFields.Empty;

    /// <summary>What fields Word saved without a result are computed from (document properties, bookmarks); null: the code only.</summary>
    public Fields.WordFieldContext? FieldContext { get; set; }

    /// <summary>A form field's FFData at a data-stream location (the 0x01 in its code), or null.</summary>
    public Func<uint, Fields.WordFormFieldState?>? FormFieldData { get; set; }

    /// <summary>The headings (CP, level) - SEQ \s resets at them.</summary>
    public Func<IEnumerable<(int cp, int level)>>? HeadingPositions { get; set; }

    private DocBinaryFieldReader? _fields;                 // the fields open where the story being rendered is

    /// <summary>A field reader for a story starting at <paramref name="cp"/>; <paramref name="restore"/>: rebuild the fields open there (a main-text block).</summary>
    internal DocBinaryFieldReader NewFieldReader(int cp, bool restore)
    {
        bool isMain = cp < _mainTextEnd;
        var reader = new DocBinaryFieldReader(FieldTable, FieldContext,
                                              isMain ? at => FieldTable.CountersBefore(at, _pieces.ReadText, HeadingPositions ?? (() => [])) : null)
        { FormData = FormFieldData, DataLocationAt = DataLocationAt };
        if (restore && isMain) reader.RestoreAt(cp, _pieces.ReadText);
        return reader;
    }

    private int _mainTextEnd = int.MaxValue;

    /// <summary>The main text's character count (ccpText): fields there count and run across blocks.</summary>
    public int MainTextEnd { get => _mainTextEnd; set => _mainTextEnd = value; }

    private bool _isParagraphRightToLeft;   // the paragraph being rendered: a right-to-left run in a left-to-right one is embedded

    public (WordParagraphFormatting formatting, int styleIndex) ResolveParagraph(int cpStart, int cpEnd)
    {
        DocBinaryFormattedDiskPages.PropertyRun? run = ParagraphRunOf(cpStart, cpEnd);
        int styleIndex = 0;
        if (run is { GrpprlLength: >= 2 } papx) styleIndex = BitConverter.ToUInt16(papx.Page, papx.GrpprlStart);
        WordParagraphFormatting style = _styles.ResolveParagraphStyle(styleIndex);
        WordParagraphFormatting formatting = style.Clone();
        (byte[] grpprl, int start, int length)? direct = run is { GrpprlLength: > 2 } directRun
            ? _pages.ExpandHugePapx(directRun.Page, directRun.GrpprlStart + 2, directRun.GrpprlLength - 2) : null;
        byte[]? prm = PrmGrpprlOfParagraph(cpStart, cpEnd);
        void ApplyDirect(WordParagraphFormatting target)
        {
            if (direct is { } d) DocBinaryPropertyModifiers.ApplyParagraphSprms(target, d.grpprl, d.start, d.length);
            // Fast-saved files: the piece's Prm applies after the PAPX ([MS-DOC] 2.4.6.1 step 4).
            if (prm is not null) DocBinaryPropertyModifiers.ApplyParagraphSprms(target, prm, 0, prm.Length);
        }
        ApplyDirect(formatting);
        // A list paragraph: its level's grpprlPapx (indents, the list tab stop) goes between the style and the paragraph's
        // own sprms - LibreOffice ww8par3.cxx applies it at the sprmPIlfo, and Word writes sprms in ascending order, so
        // sprmPIlfo (0x460B) precedes the indent sprms (0x840F, 0x845E, ...) that then override the level.
        if (Lists.LevelOfParagraph(formatting.NumberingId, formatting.NumberingLevel) is { ParagraphGrpprl.Length: > 0 } level)
        {
            formatting = style.Clone();
            DocBinaryPropertyModifiers.ApplyParagraphSprms(formatting, level.ParagraphGrpprl, 0, level.ParagraphGrpprl.Length);
            ApplyDirect(formatting);
        }
        return (formatting, styleIndex);
    }

    /// <summary>The Prm grpprl of the piece holding the paragraph mark, or null (the common case).</summary>
    private byte[]? PrmGrpprlOfParagraph(int cpStart, int cpEnd) => _pieces.PieceOf(Math.Max(cpStart, cpEnd - 1)).PrmGrpprl;

    /// <summary>
    /// sprmPFInTable / sprmPFTtp / sprmPItap / sprmPFInnerTableCell / sprmPFInnerTtp and the row's table sprms of the paragraph
    /// ending at cpEnd. As LibreOffice (ww8par2.cxx:1949-1967): sprmTTableBorders (Word 2000+) wins over sprmTTableBorders80, and
    /// the sprmTSetBrc80 / sprmTSetBrc are applied after sprmTDefTable, in file order.
    /// </summary>
    public TableParagraphInfo ReadTableInfo(int cpStart, int cpEnd)
    {
        DocBinaryFormattedDiskPages.PropertyRun? run = ParagraphRunOf(cpStart, cpEnd);
        bool isInTable = false, isRowEnd = false, isRightToLeft = false, isInnerCellEnd = false, isInnerRowEnd = false;
        int depth = 1, depthChange = 0;
        int[]? widths = null, edges = null;
        DocBinaryTableCell[]? cells = null;
        string?[]? shading = null;
        string?[]? borders = null, borders80 = null;
        List<(byte[] grpprl, int at, int length, bool isBrc80)>? setBorders = null;
        var borderColors = new (byte[] grpprl, int at, int length)?[4];
        int rowHeight = 0;
        int justification = 0, gapHalf = 0;
        int? dxaLeft = null, firstCenter = null, lastCenter = null, preferredWidth = null;
        void ReadSprms(byte[] grpprl, int start, int length)
        {
            foreach ((ushort opcode, int operandAt, int operandLength) in DocBinaryPropertyModifiers.Sprms(grpprl, start, length))
            {
                switch (opcode)
                {
                    case 0x2416: isInTable = grpprl[operandAt] != 0; break;                     // sprmPFInTable
                    case 0x2417: isRowEnd = grpprl[operandAt] != 0; break;                      // sprmPFTtp
                    case 0x244C: isInnerRowEnd = grpprl[operandAt] != 0; break;                 // sprmPFInnerTtp
                    case 0x244B: isInnerCellEnd = grpprl[operandAt] != 0; break;                // sprmPFInnerTableCell
                    case 0x6649: if (operandLength >= 4) depth = BitConverter.ToInt32(grpprl, operandAt); break;          // sprmPItap
                    case 0x664A: if (operandLength >= 4) depthChange += BitConverter.ToInt32(grpprl, operandAt); break;   // sprmPDtap
                    case 0xD608:                                                                // sprmTDefTable
                        widths = CellWidths(grpprl, operandAt, operandLength);
                        if (widths is not null)
                        {
                            firstCenter = BitConverter.ToInt16(grpprl, operandAt + 1);
                            lastCenter = BitConverter.ToInt16(grpprl, operandAt + 1 + widths.Length * 2);
                            edges = new int[widths.Length + 1];
                            for (int i = 0; i < edges.Length; i++) edges[i] = BitConverter.ToInt16(grpprl, operandAt + 1 + i * 2);
                            cells = ReadCells(grpprl, operandAt, operandLength, widths.Length);
                        }
                        break;
                    case 0x548A: case 0x5400: justification = grpprl[operandAt] & 0x3; break;   // sprmTJc / sprmTJc90
                    case 0x9601: dxaLeft = BitConverter.ToInt16(grpprl, operandAt); break;       // sprmTDxaLeft
                    case 0x9602: gapHalf = BitConverter.ToInt16(grpprl, operandAt); break;       // sprmTDxaGapHalf
                    case 0xF614:                                                                 // sprmTTableWidth: ftsWidth, wWidth
                        if (operandLength >= 3 && grpprl[operandAt] == 3) preferredWidth = grpprl[operandAt + 2] * 0x100 + grpprl[operandAt + 1];
                        break;
                    case 0x560B: isRightToLeft = BitConverter.ToUInt16(grpprl, operandAt) != 0; break;   // sprmTFBiDi
                    case 0xD612: shading = CellShading(grpprl, operandAt, operandLength); break;         // sprmTCellShd
                    case 0xD613: borders = TableBorders(grpprl, operandAt, operandLength, 8); break;     // sprmTTableBorders (Brc)
                    case 0xD605: borders80 = TableBorders(grpprl, operandAt, operandLength, 4); break;   // sprmTTableBorders80 (Brc80)
                    case 0x9407: if (operandLength >= 2) rowHeight = BitConverter.ToInt16(grpprl, operandAt); break;      // sprmTDyaRowHeight: > 0 at least, < 0 exact (ww8par2.cxx:3150-3162)
                    // sprmTBrcTopCv / LeftCv / BottomCv / RightCv: a COLORREF per cell for its TC border (Word writes them for colours
                    // outside the 16-colour ico palette; LibreOffice ignores them - ww8par6.cxx:6345-6348 - Word 16 shows them)
                    case 0xD61A: borderColors[DocBinaryTableCell.Top] = (grpprl, operandAt, operandLength); break;
                    case 0xD61B: borderColors[DocBinaryTableCell.Left] = (grpprl, operandAt, operandLength); break;
                    case 0xD61C: borderColors[DocBinaryTableCell.Bottom] = (grpprl, operandAt, operandLength); break;
                    case 0xD61D: borderColors[DocBinaryTableCell.Right] = (grpprl, operandAt, operandLength); break;
                    case 0xD620: (setBorders ??= new()).Add((grpprl, operandAt, operandLength, true)); break;    // sprmTSetBrc80
                    case 0xD62F: (setBorders ??= new()).Add((grpprl, operandAt, operandLength, false)); break;   // sprmTSetBrc
                }
            }
        }
        if (run is { GrpprlLength: > 2 } papx)
        {
            (byte[] grpprl, int start, int length) = _pages.ExpandHugePapx(papx.Page, papx.GrpprlStart + 2, papx.GrpprlLength - 2);
            ReadSprms(grpprl, start, length);
        }
        if (PrmGrpprlOfParagraph(cpStart, cpEnd) is { } prm) ReadSprms(prm, 0, prm.Length);     // a Prm0 can carry sprmPFInTable / sprmPFTtp
        if (cells is not null)
            for (int side = 0; side < 4; side++)
                if (borderColors[side] is { } colors) ApplyBorderColors(cells, side, colors.grpprl, colors.at, colors.length);
        if (cells is not null && setBorders is not null)
            foreach ((byte[] grpprl, int at, int length, bool isBrc80) in setBorders) SetCellBorders(cells, grpprl, at, length, isBrc80);
        // sprmTDxaLeft moves the row so its first cell boundary is dxaLeft - dxaGapHalf (LibreOffice ww8par2.cxx WW8TabDesc:
        // "Use sprmTDxaLeft to adjust the left indent ... Use dxaGapHalf during calculation").
        int? leftEdge = firstCenter is { } first ? (dxaLeft is { } left ? left - gapHalf : first) : null;
        int? rowWidth = preferredWidth ?? (firstCenter is { } a && lastCenter is { } b && b > a ? b - a : null);
        return new TableParagraphInfo(isInTable || isRowEnd || isInnerRowEnd, isRowEnd, widths, isRightToLeft, shading, borders ?? borders80, justification, leftEdge, rowWidth,
                                      Math.Max(1, depth + depthChange), isInnerCellEnd, isInnerRowEnd, edges, cells, rowHeight);
    }

    private DocBinaryFormattedDiskPages.PropertyRun? ParagraphRunOf(int cpStart, int cpEnd)
        => _pages.ParagraphRunAt(_pieces.FileOffsetOf(Math.Max(cpStart, cpEnd - 1)));

    /// <summary>sprmTCellShd: one 10-byte Shd per cell (cvFore, cvBack, ipat) → each cell's background.</summary>
    private static string?[] CellShading(byte[] grpprl, int operandAt, int operandLength)
    {
        var shading = new string?[operandLength / 10];
        for (int cell = 0; cell < shading.Length; cell++)
        {
            int back = operandAt + cell * 10 + 4;
            if (grpprl[back + 3] == 0xFF) continue;                                   // auto
            string hex = $"{grpprl[back]:X2}{grpprl[back + 1]:X2}{grpprl[back + 2]:X2}";
            if (hex != "FFFFFF") shading[cell] = hex;
        }
        return shading;
    }

    /// <summary>sprmTTableBorders(80): six Brc (8 bytes) / Brc80 (4 bytes) - top, left, bottom, right, inside horizontal, inside vertical (LibreOffice ww8par2.cxx ProcessSprmTTableBorders, :1277-1311).</summary>
    private static string?[]? TableBorders(byte[] grpprl, int operandAt, int operandLength, int size)
    {
        if (operandLength < 6 * size) return null;
        var borders = new string?[6];
        for (int i = 0; i < 6; i++) borders[i] = BorderCssOf(grpprl, operandAt + i * size, size) ?? "none";
        return borders;
    }

    /// <summary>A Brc80 (4 bytes) or Brc (8 bytes) as CSS; null for brcType 0 (not set: LibreOffice ww8par2.cxx:2222-2246 takes the table's border then).</summary>
    private static string? BorderCssOf(byte[] grpprl, int at, int size)
    {
        if (at + size > grpprl.Length) return null;
        byte type = grpprl[at + (size == 4 ? 1 : 5)];
        if (type == 0) return null;
        return size == 4 ? DocBinaryPropertyModifiers.Brc80Css(grpprl, at) : DocBinaryPropertyModifiers.BrcCss(grpprl, at, 8);
    }

    /// <summary>The TC80s after the cell edges of a TDefTableOperand (20 bytes each: tcgrf, wWidth, brcTop / Left / Bottom / Right as Brc80); cells not saved stay empty (LibreOffice ReadDef: nFileCols).</summary>
    private static DocBinaryTableCell[] ReadCells(byte[] grpprl, int operandAt, int operandLength, int cellCount)
    {
        var cells = new DocBinaryTableCell[cellCount];
        int first = operandAt + 1 + (cellCount + 1) * 2;
        int saved = Math.Max(0, (operandLength - 1 - (cellCount + 1) * 2) / 20);
        for (int i = 0; i < cellCount; i++)
        {
            var cell = new DocBinaryTableCell();
            int at = first + i * 20;
            if (i < saved && at + 20 <= grpprl.Length)
            {
                cell.Flags = BitConverter.ToUInt16(grpprl, at);
                for (int side = 0; side < 4; side++) cell.Borders[side] = BorderCssOf(grpprl, at + 4 + side * 4, 4);
            }
            cells[i] = cell;
        }
        return cells;
    }

    /// <summary>A sprmTBrc*Cv: per cell a COLORREF (r, g, b, fAuto/0xFF); a cell with its own line on that side takes the colour.</summary>
    private static void ApplyBorderColors(DocBinaryTableCell[] cells, int side, byte[] grpprl, int at, int length)
    {
        for (int i = 0; i < cells.Length && (i + 1) * 4 <= length && at + i * 4 + 4 <= grpprl.Length; i++)
        {
            int color = at + i * 4;
            if (grpprl[color + 3] != 0 || cells[i].Borders[side] is not { } css || css == "none") continue;   // auto / nil: keep the ico colour
            int hash = css.LastIndexOf('#');
            if (hash < 0) continue;
            cells[i].Borders[side] = css.Substring(0, hash) + $"#{grpprl[color]:X2}{grpprl[color + 1]:X2}{grpprl[color + 2]:X2}";
        }
    }

    /// <summary>
    /// sprmTSetBrc80 / sprmTSetBrc: itcFirst, itcLim, the sides (bit 0 top, 1 left, 2 bottom, 3 right), then a Brc80 / Brc put on
    /// cells itcFirst..itcLim-1 (LibreOffice ww8par2.cxx ProcessSprmTSetBRC, :1208-1275).
    /// </summary>
    private static void SetCellBorders(DocBinaryTableCell[] cells, byte[] grpprl, int at, int length, bool isBrc80)
    {
        int size = isBrc80 ? 4 : 8;
        if (length < 3 + size) return;
        int firstCell = grpprl[at], limit = Math.Min((int)grpprl[at + 1], cells.Length);
        int sides = grpprl[at + 2];
        string? css = BorderCssOf(grpprl, at + 3, size);         // brcType 0 clears the side back to the table border (ww8par2.cxx:2226)
        for (int i = firstCell; i < limit; i++)
            for (int side = 0; side < 4; side++)
                if ((sides & (1 << side)) != 0) cells[i].Borders[side] = css;
    }

    /// <summary>TDefTableOperand: itcMac, then itcMac + 1 cell boundaries (rgdxaCenter) → widths.</summary>
    private static int[]? CellWidths(byte[] page, int operandAt, int operandLength)
    {
        if (operandLength < 3) return null;
        int cellCount = page[operandAt];
        if (cellCount <= 0 || operandAt + 1 + (cellCount + 1) * 2 > page.Length) return null;
        var widths = new int[cellCount];
        for (int i = 0; i < cellCount; i++)
            widths[i] = Math.Abs(BitConverter.ToInt16(page, operandAt + 1 + (i + 1) * 2) - BitConverter.ToInt16(page, operandAt + 1 + i * 2));
        return widths;
    }

    /// <param name="previousStyleKey">The style of the paragraph before (contextual spacing).</param>
    /// <param name="noteLabel">Inside a note body: the label its auto-number mark shows.</param>
    /// <param name="listLabel">The paragraph's list label from the index (counted at index time), shown before the text.</param>
    /// <param name="dropCapParagraph">The drop-cap frame paragraph before this one (Word keeps the initial in a paragraph of its own), drawn at its start.</param>
    public string Render(int cpStart, int cpEnd, string? previousStyleKey = null, string? noteLabel = null, string? listLabel = null,
                         (int cpStart, int cpEnd)? dropCapParagraph = null)
    {
        var html = new StringBuilder(512);
        DocBinaryFieldReader? outer = _fields;
        if (_textBoxDepth == 0) _fields = NewFieldReader(cpStart, restore: true);
        AppendParagraph(html, cpStart, cpEnd, previousStyleKey, noteLabel, listLabel, dropCapParagraph);
        if (_textBoxDepth == 0) _fields = outer;
        return html.ToString();
    }

    private string? AppendParagraph(StringBuilder html, int cpStart, int cpEnd, string? previousStyleKey, string? noteLabel, string? listLabel = null,
                                    (int cpStart, int cpEnd)? dropCapParagraph = null)
    {
        (WordParagraphFormatting formatting, int styleIndex) = ResolveParagraph(cpStart, cpEnd);
        formatting.IsHyphenated = IsAutoHyphenation && !formatting.IsAutoHyphenationSuppressed;
        int headingLevel = _styles.HeadingLevelOf(styleIndex);
        string tag = _textBoxDepth > 0 ? "span" : headingLevel is >= 1 and <= 6 ? "h" + headingLevel : "p";

        bool outerRightToLeft = _isParagraphRightToLeft;
        WordHtmlText.AppendOpenTag(html, tag, formatting.ToCss(previousStyleKey) + (_textBoxDepth > 0 ? "display:block;" : ""),
                                   _textBoxDepth > 0 ? "text-box-paragraph" : null, formatting.IsRightToLeft, formatting.HyphenationLanguageAttribute);
        _isParagraphRightToLeft = formatting.IsRightToLeft;
        WordRunFormatting styleBase = _styles.ResolveParagraphStyle(styleIndex).RunFormatting;
        if (listLabel is { Length: > 0 }) AppendListLabel(html, listLabel, formatting, ParagraphMarkFormatting(cpEnd, formatting.RunFormatting, styleBase));
        if (dropCapParagraph is { } dropCap) AppendDropCap(html, dropCap.cpStart, dropCap.cpEnd, formatting.RunFormatting, noteLabel);
        DocBinaryFieldReader fields = _fields ??= NewFieldReader(cpStart, restore: _textBoxDepth == 0);
        _baselineShift = 0;
        if (fields.ReopenLink() is { } openLink) html.Append(openLink);   // a link field that runs on from the paragraph before
        foreach ((WordRunFormatting runFormatting, string text, bool isSpecial, int runCp, uint? picture) in CharacterRuns(cpStart, cpEnd, formatting.RunFormatting, styleBase))
            AppendRunHtml(html, runFormatting, formatting.RunFormatting, text, isSpecial, runCp, fields, noteLabel, picture);
        if (fields.CloseLinkAtParagraphEnd()) html.Append("</a>");    // a link field that goes on past the paragraph
        html.Append("</").Append(tag).Append('>');
        if (_textBoxDepth > 0) _isParagraphRightToLeft = outerRightToLeft;
        return formatting.StyleKey;
    }

    /// <summary>
    /// The label drawn as Word does: the paragraph's character formatting plus the level's grpprlChpx, followed by its
    /// ixchFollow (LVLF, [MS-DOC] 2.9.150; LibreOffice ww8par3.cxx: 0 tab, 1 space, 2 nothing).
    /// </summary>
    private void AppendListLabel(StringBuilder html, string label, WordParagraphFormatting formatting, WordRunFormatting markFormatting)
    {
        DocBinaryListDefinitions.LevelDefinition? level = Lists.LevelOfParagraph(formatting.NumberingId, formatting.NumberingLevel);
        string? css = null;
        WordRunFormatting labelFormatting = markFormatting;
        if (level is { CharacterGrpprl.Length: > 0 })
        {
            labelFormatting = markFormatting.Clone();
            DocBinaryPropertyModifiers.ApplyCharacterSprms(labelFormatting, level.CharacterGrpprl, 0, level.CharacterGrpprl.Length, fonts: _styles.Fonts);
        }
        // Latin or complex-script set by the label's own characters and the paragraph's direction (WordListLabelScript)
        labelFormatting = WordListLabelScript.WithScript(labelFormatting, WordListLabelScript.UsesComplexScript(label, formatting.IsRightToLeft));
        if (level is { CharacterGrpprl.Length: > 0 } && DocBinaryPropertyModifiers.PictureBulletIndexIn(level.CharacterGrpprl, 0, level.CharacterGrpprl.Length) is { } pictureIndex
            && PictureBulletLocation(pictureIndex) is { } location && Pictures.Locate(location) is { } picture)
            label = WordListLabelHtml.PictureLabel(location.ToString(System.Globalization.CultureInfo.InvariantCulture), picture.widthPt, picture.heightPt);
        else if (level is { IsBullet: true })
        {
            (label, bool keepsFont) = WordSymbolFontBullet.Resolve(label, labelFormatting.FontName);
            if (!keepsFont)
            {
                if (ReferenceEquals(labelFormatting, formatting.RunFormatting) || ReferenceEquals(labelFormatting, markFormatting)) labelFormatting = labelFormatting.Clone();
                labelFormatting.FontName = formatting.RunFormatting.FontName;                 // the recoded bullet draws in the text's font
                labelFormatting.ComplexScriptFontName = formatting.RunFormatting.ComplexScriptFontName;
            }
        }
        if (!ReferenceEquals(labelFormatting, formatting.RunFormatting)) css = labelFormatting.ToCss(formatting.RunFormatting);
        if (css is { Length: 0 }) css = null;
        WordListLabelSuffix suffix = level?.FollowCharacter switch { 1 => WordListLabelSuffix.Space, 2 => WordListLabelSuffix.Nothing, _ => WordListLabelSuffix.Tab };
        WordListLabelAlignment alignment = level?.Justification switch { 1 => WordListLabelAlignment.Center, 2 => WordListLabelAlignment.End, _ => WordListLabelAlignment.Start };
        WordListLabelHtml.Append(html, label, css, suffix, formatting, DefaultTabStopTwips, alignment, legacy: level?.Legacy);
    }

    /// <summary>
    /// The paragraph mark's character formatting: Word draws a list label in it (LibreOffice turns on
    /// APPLY_PARAGRAPH_MARK_FORMAT_TO_NUMBERING for .doc, ww8par.cxx:1955), under the level's grpprlChpx.
    /// </summary>
    private WordRunFormatting ParagraphMarkFormatting(int cpEnd, WordRunFormatting paragraphBase, WordRunFormatting styleBase)
    {
        foreach ((WordRunFormatting mark, _, _, _, _) in CharacterRuns(cpEnd - 1, cpEnd, paragraphBase, styleBase)) return mark;
        return paragraphBase;
    }

    private List<uint>? _pictureBullets;

    /// <summary>
    /// Picture bullet <paramref name="index"/>: the index-th inline picture of the paragraph holding the "_PictureBullets"
    /// bookmark, from the bookmark's start on (LibreOffice ww8par.cxx:5302-5334 collects the as-character pictures of that
    /// paragraph at or after the bookmark start, :5343-5348 picks the one the level names). Read once, on the first picture
    /// bullet: one paragraph's runs, its pictures' PICF offsets kept.
    /// </summary>
    private uint? PictureBulletLocation(int index)
    {
        if (_pictureBullets is null)
        {
            _pictureBullets = new List<uint>();
            if (Bookmarks.StartOf("_PictureBullets") is { } start)
            {
                int end = start, total = _pieces.TotalCharacterCount;
                const int Chunk = 1024;
                while (end < total)
                {
                    string text = _pieces.ReadText(end, Math.Min(total, end + Chunk));
                    int mark = text.IndexOf('\r');
                    if (mark >= 0) { end += mark; break; }
                    if (text.Length == 0) { end = total; break; }
                    end += text.Length;
                }
                var empty = new WordRunFormatting();
                foreach ((_, string text, bool isSpecial, _, uint? picture) in CharacterRuns(start, end, empty, empty))
                    if (isSpecial && picture is { } location)
                        foreach (char c in text)
                            if (c == '\u0001') _pictureBullets.Add(location);
            }
        }
        return index < _pictureBullets.Count ? _pictureBullets[index] : null;
    }

    /// <summary>
    /// One outer table row (one block): cells end at cell marks, the row at its row-end mark (the last paragraph). Nested
    /// tables (sprmPItap &gt; 1) are drawn inside their cell, each a whole &lt;table&gt; with rowspans.
    /// <paramref name="cellsMergedBelow"/> (bit i = cell i, from the index) are the cells the next row continues, drawn without
    /// their bottom line; <paramref name="isFirstRow"/> / <paramref name="isLastRow"/> pick the table's outer or inside
    /// horizontal border (LibreOffice ww8par2.cxx:2226-2246).
    /// </summary>
    public string RenderRow(int cpStart, int cpEnd, ulong cellsMergedBelow = 0, bool isFirstRow = true, bool isLastRow = true)
    {
        DocBinaryFieldReader? outer = _fields;
        if (_textBoxDepth == 0) _fields = NewFieldReader(cpStart, restore: true);
        try { return RenderRowWithFields(cpStart, cpEnd, cellsMergedBelow, isFirstRow, isLastRow); }
        finally { if (_textBoxDepth == 0) _fields = outer; }
    }

    /// <summary>A nested table's row while its table is read: the row-end paragraph's facts and its cells' HTML.</summary>
    private readonly record struct NestedRow(TableParagraphInfo Info, List<string> Cells);

    /// <summary>One table depth while a row is read: the cell being filled, the cells of its row so far, its finished rows (nested).</summary>
    private sealed class TableDepth
    {
        public readonly StringBuilder Cell = new();
        public readonly List<string> Cells = new();
        public readonly List<NestedRow> Rows = new();
        public string? PreviousStyleKey;
    }

    private string RenderRowWithFields(int cpStart, int cpEnd, ulong cellsMergedBelow, bool isFirstRow, bool isLastRow)
    {
        string text = _pieces.ReadText(cpStart, cpEnd);
        var paragraphs = new List<(int start, int end, char mark)>();
        int paragraphStart = cpStart;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] is not ('\r' or CellOrRowMark)) continue;
            paragraphs.Add((paragraphStart, cpStart + i + 1, text[i]));
            paragraphStart = cpStart + i + 1;
        }
        if (paragraphs.Count == 0) return Render(cpStart, cpEnd);

        (int lastStart, int lastEnd, _) = paragraphs[paragraphs.Count - 1];
        TableParagraphInfo rowInfo = ReadTableInfo(lastStart, lastEnd);
        int contentCount = rowInfo.IsRowEnd ? paragraphs.Count - 1 : paragraphs.Count;    // a row without its end mark: its last paragraph is content

        // Word 2000+ nesting (LibreOffice ww8par.cxx:2729-2733 TabRowSprm, :3686 tdf#106799): a paragraph's depth is sprmPItap;
        // a nested cell ends at sprmPFInnerTableCell, a nested row at sprmPFInnerTtp (which also ends its cell); an outer cell at
        // its cell mark. A nested table is complete when a paragraph of a shallower depth follows.
        var depths = new List<TableDepth> { new() };
        TableDepth DepthOf(int depth) { while (depths.Count < depth) depths.Add(new TableDepth()); return depths[depth - 1]; }
        void FlushDeeperThan(int depth)
        {
            for (int level = depths.Count; level > depth; level--)
            {
                TableDepth table = depths[level - 1];
                if (table.Cell.Length > 0 || table.Cells.Count > 0) CloseNestedRow(table, null);
                if (table.Rows.Count == 0) continue;
                var nested = new StringBuilder(256);
                AppendNestedTable(nested, table.Rows);
                table.Rows.Clear();
                TableDepth holder = depths[level - 2];
                holder.Cell.Append(nested);
                holder.PreviousStyleKey = null;
            }
        }
        for (int i = 0; i < contentCount; i++)          // the last paragraph is the row-end mark itself
        {
            (int start, int end, char mark) = paragraphs[i];
            TableParagraphInfo info = ReadTableInfo(start, end);
            int depth = info.IsInTable || mark == CellOrRowMark ? info.Depth : 1;
            FlushDeeperThan(depth);
            TableDepth table = DepthOf(depth);
            if (depth > 1 && info.IsInnerRowEnd) { CloseNestedRow(table, info); continue; }
            table.PreviousStyleKey = AppendParagraph(table.Cell, start, end, table.PreviousStyleKey, null);
            bool endsCell = depth > 1 ? info.IsInnerCellEnd : mark == CellOrRowMark;
            if (endsCell) { table.Cells.Add(table.Cell.ToString()); table.Cell.Clear(); table.PreviousStyleKey = null; }
        }
        FlushDeeperThan(1);
        TableDepth row = depths[0];
        if (row.Cell.Length > 0) row.Cells.Add(row.Cell.ToString());

        var html = new StringBuilder(1024);
        // the column order is the table's (sprmTFBiDi - LibreOffice ww8par2.cxx:1871-1872 m_bIsBiDi), not the row-end paragraph's direction
        WordHtmlText.AppendOpenTag(html, "table", "table-layout:fixed;" + RowPlacementCss(rowInfo), "table-row", rowInfo.IsRightToLeft);
        WordTableRowHeight.AppendRowOpen(html, rowInfo.RowHeightTwips);
        AppendRowCells(html, rowInfo, row.Cells, null, cellsMergedBelow, isFirstRow, isLastRow);
        html.Append("</tr></table>");
        return html.ToString();
    }

    /// <summary>A nested row ends: its open cell is closed (a row-end mark ends its cell too) and the row waits for the rest of its table.</summary>
    private static void CloseNestedRow(TableDepth table, TableParagraphInfo? info)
    {
        if (table.Cell.Length > 0) { table.Cells.Add(table.Cell.ToString()); table.Cell.Clear(); }
        table.Rows.Add(new NestedRow(info ?? default, new List<string>(table.Cells)));
        table.Cells.Clear();
        table.PreviousStyleKey = null;
    }

    /// <summary>Cell right edges within Word's tolerance: "corresponding entries in rgdxaCenter ... differ by at most 3" (ww8struc.hxx WW8_TCell, :536).</summary>
    private static bool IsSameEdge(int a, int b) => Math.Abs(a - b) <= 3;

    /// <summary>The right edge of cell i of a row, or null without a definition.</summary>
    private static int? RightEdge(in TableParagraphInfo row, int i) => row.CellEdgesTwips is { } edges && i + 1 < edges.Length ? edges[i + 1] : null;

    /// <summary>
    /// A nested table, all its rows read: one &lt;table&gt;, a cell merged down (fVertRestart, then fVertMerge cells at the same
    /// right edge below; LibreOffice ww8par2.cxx:2582-2620) as one rowspan cell.
    /// </summary>
    private void AppendNestedTable(StringBuilder html, List<NestedRow> rows)
    {
        WordHtmlText.AppendOpenTag(html, "table", "table-layout:fixed;", "nested-table", rows[0].Info.IsRightToLeft ? true : null);
        bool IsContinued(int rowIndex, int edge) => rowIndex < rows.Count && rows[rowIndex].Info.Cells is { } cells
                                                     && Enumerable.Range(0, cells.Length).Any(i => cells[i].ContinuesAbove && RightEdge(rows[rowIndex].Info, i) is { } e && IsSameEdge(e, edge));
        bool HasAbove(int rowIndex, int edge) => rowIndex > 0 && rows[rowIndex - 1].Info.Cells is { } cells
                                                  && Enumerable.Range(0, cells.Length).Any(i => cells[i].IsVerticallyMerged && RightEdge(rows[rowIndex - 1].Info, i) is { } e && IsSameEdge(e, edge));
        for (int r = 0; r < rows.Count; r++)
        {
            TableParagraphInfo info = rows[r].Info;
            int[]? spans = null;
            var covered = new bool[info.Cells?.Length ?? 0];
            if (info.Cells is { } cells)
            {
                spans = new int[cells.Length];
                for (int i = 0; i < cells.Length; i++)
                {
                    if (RightEdge(info, i) is not { } edge) continue;
                    // a continued cell with no merge above it (a table that starts mid-merge) is an ordinary cell
                    if (cells[i].ContinuesAbove && HasAbove(r, edge)) { covered[i] = true; continue; }
                    int span = 1;
                    if (cells[i].IsVerticallyMerged) while (IsContinued(r + span, edge)) span++;
                    spans[i] = span;
                }
            }
            WordTableRowHeight.AppendRowOpen(html, info.RowHeightTwips);
            AppendRowCells(html, info, rows[r].Cells, covered, 0, r == 0, r == rows.Count - 1, spans);
            html.Append("</tr>");
        }
        html.Append("</table>");
    }

    /// <summary>
    /// The &lt;td&gt;s of one row from its definition (LibreOffice ww8par2.cxx WW8TabDesc): widths from the cell edges, fFirstMerged +
    /// fMerged as a colspan (the merged cells' text dropped), fill (sprmTCellShd), vertical alignment (TC vertAlign, :3116-3130),
    /// borders: the cell's own (TC brc*, sprmTSetBrc), else the table's for that side - outer top / bottom on the table's first /
    /// last row, else inside horizontal; outer left / right on the first / last cell, else inside vertical (:2222-2246).
    /// Outer rows draw a cell merged into the one above (fVertMerge) empty, without its top line (class merged-up); nested
    /// tables skip it under the rowspan (<paramref name="covered"/>).
    /// </summary>
    private void AppendRowCells(StringBuilder html, in TableParagraphInfo row, List<string> cellsHtml, bool[]? covered, ulong cellsMergedBelow,
                                bool isFirstRow, bool isLastRow, int[]? rowSpans = null)
    {
        DocBinaryTableCell[]? cells = row.Cells;
        int count = Math.Max(cells?.Length ?? 0, cellsHtml.Count);
        // a row that names no border keeps the stylesheet's cell lines (as before)
        bool hasBorders = row.DefaultBorders is not null || cellsMergedBelow != 0
                          || (cells?.Any(cell => cell.IsVerticallyMerged || cell.Borders.Any(border => border is not null)) ?? false);
        for (int i = 0; i < count; i++)
        {
            DocBinaryTableCell? cell = cells is not null && i < cells.Length ? cells[i] : null;
            if (cell is { ContinuesBefore: true } && i > 0) continue;                                          // drawn by the first cell
            if (covered is not null && rowSpans is not null && i < covered.Length && covered[i]) continue;      // under a rowspan
            bool isMergedUp = rowSpans is null && cell is { ContinuesAbove: true };
            int span = 1;
            while (cells is not null && i + span < cells.Length && cells[i + span].ContinuesBefore) span++;
            int last = i + span - 1;
            var css = new StringBuilder();
            if (row.CellEdgesTwips is { } edges && last + 1 < edges.Length)
            {
                int width = Math.Abs(edges[last + 1] - edges[i]);
                if (width > 0) css.Append("width:").Append(WordParagraphFormatting.TwipsToPt(width)).Append("pt;");
            }
            else if (row.CellWidthsTwips is { } widths && i < widths.Length && widths[i] > 0)
                css.Append("width:").Append(WordParagraphFormatting.TwipsToPt(widths[i])).Append("pt;");
            if (row.CellShadingHex is { } shading && i < shading.Length && shading[i] is { } fill)
                css.Append("background:").Append(WordHtmlColor.Css(fill)).Append(';');
            if (cell is { VerticalAlignment: 1 or 2 } aligned) css.Append("vertical-align:").Append(aligned.VerticalAlignment == 1 ? "middle" : "bottom").Append(';');
            if (hasBorders)
            {
                string?[] table = row.DefaultBorders ?? new string?[6];
                DocBinaryTableCell? lastCell = cells is not null && last < cells.Length ? cells[last] : cell;
                bool isFirstCell = i == 0, isLastCell = last >= count - 1;
                bool isMergedBelow = i < 64 && (cellsMergedBelow & (1UL << i)) != 0;
                string Side(string? own, string? fallback) => own ?? fallback ?? "none";
                css.Append("border-top:").Append(isMergedUp ? "none" : Side(cell?.Borders[DocBinaryTableCell.Top], isFirstRow ? table[0] : table[4])).Append(';');
                css.Append("border-bottom:").Append(isMergedBelow ? "none" : Side(cell?.Borders[DocBinaryTableCell.Bottom], isLastRow ? table[2] : table[4])).Append(';');
                css.Append("border-left:").Append(Side(cell?.Borders[DocBinaryTableCell.Left], isFirstCell ? table[1] : table[5])).Append(';');
                css.Append("border-right:").Append(Side(lastCell?.Borders[DocBinaryTableCell.Right], isLastCell ? table[3] : table[5])).Append(';');
            }
            html.Append("<td");
            if (span > 1) html.Append(" colspan=\"").Append(span.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append('"');
            if (rowSpans is not null && i < rowSpans.Length && rowSpans[i] > 1)
                html.Append(" rowspan=\"").Append(rowSpans[i].ToString(System.Globalization.CultureInfo.InvariantCulture)).Append('"');
            if (isMergedUp) html.Append(" class=\"merged-up\"");
            if (css.Length > 0) html.Append(" style=\"").Append(css).Append('"');
            html.Append('>');
            if (!isMergedUp && i < cellsHtml.Count)
                WordTableRowHeight.AppendCellContent(html, cellsHtml[i], row.RowHeightTwips, rowSpans is not null && i < rowSpans.Length ? rowSpans[i] : 1);
            html.Append("</td>");
        }
    }

    /// <summary>
    /// Where a row sits, as LibreOffice places a .doc table (ww8par2.cxx WW8TabDesc): sprmTJc & 3 = 0 left, 1 centre, 2 right,
    /// 3 centre (aOriArr - physical sides); the row's width is sprmTTableWidth in twips (ftsWidth 3) else its cell boundaries'
    /// span; the first boundary (after sprmTDxaLeft) is the indent from the margin, applied - as LibreOffice's LEFT_AND_WIDTH -
    /// to a left-aligned left-to-right table from the left, and to a right-aligned right-to-left (sprmTFBiDi) table from the
    /// right ("nLeft = text area width - table width - min left").
    /// </summary>
    private static string RowPlacementCss(TableParagraphInfo row)
    {
        var css = new StringBuilder();
        if (row.RowWidthTwips is > 0 and var width) css.Append("width:").Append(WordParagraphFormatting.TwipsToPt(width)).Append("pt;");
        string Indent() => row.LeftEdgeTwips is { } edge and not 0 ? WordParagraphFormatting.TwipsToPt(edge) + "pt" : "0";
        switch (row.Justification)
        {
            case 1: case 3: css.Append("margin-left:auto;margin-right:auto;"); break;
            case 2: css.Append("margin-left:auto;margin-right:").Append(row.IsRightToLeft ? Indent() : "0").Append(';'); break;
            default: css.Append("margin-left:").Append(row.IsRightToLeft ? "0" : Indent()).Append(";margin-right:auto;"); break;
        }
        return css.ToString();
    }

    /// <summary>A footnote/endnote body: its paragraphs; the leading auto-number mark shows <paramref name="label"/>.</summary>
    public string RenderNote(int textStartCp, int textEndCp, string label)
    {
        DocBinaryFieldReader? outer = _fields;
        _fields = NewFieldReader(textStartCp, restore: false);
        try { return RenderNoteWithFields(textStartCp, textEndCp, label); }
        finally { _fields = outer; }
    }

    private string RenderNoteWithFields(int textStartCp, int textEndCp, string label)
    {
        var html = new StringBuilder(256);
        string text = _pieces.ReadText(textStartCp, textEndCp);
        int paragraphStart = textStartCp;
        string? previousStyleKey = null;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] != '\r' && i != text.Length - 1) continue;
            int end = textStartCp + i + 1;
            if (end > paragraphStart) previousStyleKey = AppendParagraph(html, paragraphStart, end, previousStyleKey, label);
            paragraphStart = end;
        }
        return html.ToString();
    }

    /// <summary>
    /// The search's view of [cpStart, cpEnd): the text Word shows (field codes hidden, results kept), a tab a space,
    /// paragraph / cell / row / line / page marks separators, other control characters (note references, pictures) dropped.
    /// </summary>
    /// <summary>Feeds a CP range's text to the matcher, read from the piece table (notes, text boxes; the golden reference for the body's one-pass search).</summary>
    /// <param name="isBlock">A main-text block: the fields open at its start are rebuilt, as the renderer does.</param>
    public void FeedSearchText(int cpStart, int cpEnd, WordDocumentSearchMatcher matcher, bool isBlock = false)
        => FeedSearchCharacters(_pieces.ReadText(cpStart, cpEnd).AsSpan(), cpStart, matcher, isBlock);

    /// <summary>
    /// Feeds one block's (or note's) characters to the matcher through the renderer's field reader: what shows is read (a
    /// cached result once, a computed one instead of a missing / recomputed one), codes, page placeholders and hidden
    /// fields are not; tabs as spaces, marks as separators. The characters' pictures (form-field data) are not known here:
    /// a form field's state comes from its FFData only when rendered - a checkbox is not text either way; a drop-down's
    /// shown entry is searched only through its cached result.
    /// </summary>
    public void FeedSearchCharacters(ReadOnlySpan<char> text, int cpStart, WordDocumentSearchMatcher matcher, bool isBlock)
    {
        DocBinaryFieldReader fields = NewFieldReader(cpStart, restore: isBlock);
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            DocBinaryFieldReader.Step step = fields.Consume(c, cpStart + i, null);
            if (step.Output is { Text: { } shown, IsSearchable: true }) matcher.Feed(shown.AsSpan());
            if (step.IsConsumed) continue;
            if (c >= ' ') matcher.Feed(c);
            else if (c == '\t') matcher.Feed(' ');
            else if (c is '\r' or '\u0007' or '\u000B' or '\u000C' or AutoNumberedNoteMark) matcher.Separator();   // a note's mark is drawn here: a match never crosses it
            else if (c == '\u0001' && DataLocationAt(cpStart + i) is { } location && !IsHiddenAt(cpStart + i, ParagraphEndAfter(text, i, cpStart)))   // hidden: not drawn (AppendRunHtml), so not searched
            {
                if (EquationOf(location) is { } equation) FeedEquationText(equation, matcher);     // a picture of an equation: its text is shown (AppendPicture), so searched
                else if (ObjectIdBefore(cpStart + i) is { } objectId && Charts?.Invoke(objectId) is { } chart && Pictures.Locate(location) is not null)   // drawn only where its preview box is known (AppendPicture)
                    foreach (string chartText in chart.Texts()) { matcher.Separator(); matcher.Feed(chartText.AsSpan()); matcher.Separator(); }
            }
        }
    }

    /// <summary>The CP just after the paragraph mark (or cell mark) at or after <paramref name="i"/> in the fed text; the text's end when there is none.</summary>
    private static int ParagraphEndAfter(ReadOnlySpan<char> text, int i, int cpStart)
    {
        for (int k = i; k < text.Length; k++)
            if (text[k] is '\r' or '\u0007') return cpStart + k + 1;
        return cpStart + text.Length;
    }

    /// <summary>Whether the character at <paramref name="cp"/> is hidden (sprmCFVanish) as the renderer resolves it: paragraph style, character style, direct formatting.</summary>
    private bool IsHiddenAt(int cp, int paragraphEnd)
    {
        // the whole paragraph's range: in a fast-saved file the paragraph's Prm is the one of the piece holding its mark (PrmGrpprlOfParagraph)
        (WordParagraphFormatting paragraph, int styleIndex) = ResolveParagraph(cp, paragraphEnd);
        WordRunFormatting styleBase = _styles.ResolveParagraphStyle(styleIndex).RunFormatting;
        foreach ((WordRunFormatting formatting, _, _, _, _) in CharacterRuns(cp, cp + 1, paragraph.RunFormatting, styleBase)) return formatting.IsHidden;
        return false;
    }

    /// <summary>An equation's text (its m:t, as the renderer shows them) to the matcher.</summary>
    internal static void FeedEquationText(System.Xml.Linq.XElement equation, WordDocumentSearchMatcher matcher)
    {
        matcher.Feed(WordOfficeMathHtml.SearchText(equation).AsSpan());
    }

    /// <summary>Text with no fields to follow (a WordArt string): tabs as spaces, marks as separators.</summary>
    public static void FeedLiteralSearchCharacters(ReadOnlySpan<char> text, WordDocumentSearchMatcher matcher)
    {
        foreach (char c in text)
        {
            if (c >= ' ') matcher.Feed(c);
            else if (c == '\t') matcher.Feed(' ');
            else if (c is '\r' or '\u0007' or '\u000B' or '\u000C') matcher.Separator();
        }
    }

    public string PlainText(int cpStart, int cpEnd)
    {
        var text = new StringBuilder();
        DocBinaryFieldReader fields = NewFieldReader(cpStart, restore: false);
        string characters = _pieces.ReadText(cpStart, cpEnd);
        for (int i = 0; i < characters.Length; i++)
        {
            char c = characters[i];
            DocBinaryFieldReader.Step step = fields.Consume(c, cpStart + i, null);
            if (step.Output.Text is { } shown) text.Append(shown);
            if (step.IsConsumed) continue;
            if (c >= ' ' || c == '\t') text.Append(c == '\t' ? ' ' : c);
        }
        return text.ToString().Trim();
    }

    /// <summary>
    /// The data location (sprmCPicLocation 0x6A03) of the character at <paramref name="cp"/>, whatever sprmCFData / sprmCFOle2 say:
    /// a form field's 0x01 carries sprmCFData with the location of its FFData (LibreOffice ww8par3.cxx:2063-2080 reads the
    /// location the same way, sprm 0x6A03, for form fields); null when none.
    /// </summary>
    internal uint? DataLocationAt(int cp) => CharacterLocationAt(cp, 0x6A03, requiresObject: false);

    /// <summary>
    /// The ObjectPool storage id of an embedded object whose preview picture is the 0x01 at <paramref name="cp"/>: the field
    /// separator 0x14 just before it (Word writes EMBED fields so: 0x13 code 0x14 0x01 0x15) carries sprmCFOle2 and a
    /// sprmCPicLocation that is the storage id (checked on Word 16's .doc: 0x6A03 = 1855890659 = "_1855890659"); LibreOffice
    /// keeps that location as the object's when fOle2 is set (ww8par6.cxx:2951-2989 Read_Obj / Read_PicLoc) and opens
    /// "_" + it (ww8par4.cxx:342). Null when there is none.
    /// </summary>
    internal uint? ObjectIdBefore(int cp)
    {
        if (cp <= 0 || Charts is null || _pieces.ReadText(cp - 1, cp) != "\u0014") return null;
        return CharacterLocationAt(cp - 1, 0x6A03, requiresObject: true);
    }

    private uint? CharacterLocationAt(int cp, ushort sprm, bool requiresObject)
    {
        foreach ((DocBinaryPieceTable.Piece piece, int segmentStart, _) in _pieces.SegmentsOf(cp, cp + 1))
        {
            uint? location = null;
            bool isObject = false;
            if (_pages.CharacterRunAt(piece.FileOffsetOf(segmentStart)) is { GrpprlLength: > 0 } direct)
            {
                location = DocBinaryPropertyModifiers.LocationIn(direct.Page, direct.GrpprlStart, direct.GrpprlLength, sprm);
                isObject = DocBinaryPropertyModifiers.PictureCharacterIn(direct.Page, direct.GrpprlStart, direct.GrpprlLength).isObject;
            }
            if (piece.PrmGrpprl is { } prm)
            {
                if (DocBinaryPropertyModifiers.LocationIn(prm, 0, prm.Length, sprm) is { } prmLocation) location = prmLocation;
                isObject |= DocBinaryPropertyModifiers.PictureCharacterIn(prm, 0, prm.Length).isObject;
            }
            return requiresObject && !isObject ? null : location;
        }
        return null;
    }

    private IEnumerable<(WordRunFormatting formatting, string text, bool isSpecial, int cp, uint? picture)> CharacterRuns(
        int cpStart, int cpEnd, WordRunFormatting paragraphBase, WordRunFormatting styleBase)
    {
        foreach ((DocBinaryPieceTable.Piece piece, int segmentStart, int segmentEnd) in _pieces.SegmentsOf(cpStart, cpEnd))
        {
            int cp = segmentStart;
            while (cp < segmentEnd)
            {
                uint fc = piece.FileOffsetOf(cp);
                DocBinaryFormattedDiskPages.PropertyRun? run = _pages.CharacterRunAt(fc);
                int runEndCp = segmentEnd;
                if (run is { } chpx && chpx.FcEnd > fc)
                    runEndCp = Math.Min(segmentEnd, cp + (int)((chpx.FcEnd - fc) / (uint)piece.BytesPerCharacter));
                if (runEndCp <= cp) runEndCp = cp + 1;

                WordRunFormatting formatting = paragraphBase.Clone();
                bool isSpecial = false;
                char? symbol = null;
                uint? picture = null;
                bool isObject = false;
                if (run is { GrpprlLength: > 0 } direct)
                {
                    int? characterStyle = DocBinaryPropertyModifiers.StyleIndexIn(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                    if (characterStyle is { } index && _styles[index] is { IsParagraphStyle: false } style)
                        DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting, style.CharacterGrpprl, 0, style.CharacterGrpprl.Length, styleBase, _styles.Fonts);
                    DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting, direct.Page, direct.GrpprlStart, direct.GrpprlLength, styleBase, _styles.Fonts);
                    isSpecial = DocBinaryPropertyModifiers.HasSpecialCharacterFlag(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                    symbol = DocBinaryPropertyModifiers.SymbolCharacterIn(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                    (picture, isObject) = DocBinaryPropertyModifiers.PictureCharacterIn(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                }
                if (piece.PrmGrpprl is { } prm)             // fast-saved files: the piece's Prm applies after the CHPX ([MS-DOC] §2.4.6.2)
                {
                    DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting, prm, 0, prm.Length, styleBase, _styles.Fonts);
                    isSpecial |= DocBinaryPropertyModifiers.HasSpecialCharacterFlag(prm, 0, prm.Length);
                    symbol ??= DocBinaryPropertyModifiers.SymbolCharacterIn(prm, 0, prm.Length);
                    (uint? prmPicture, bool prmObject) = DocBinaryPropertyModifiers.PictureCharacterIn(prm, 0, prm.Length);
                    picture = prmPicture ?? picture;
                    isObject |= prmObject;
                }
                string runText = _pieces.ReadText(cp, runEndCp);
                if (symbol is { } symbolCharacter) runText = new string(symbolCharacter, runText.Length);   // sprmCSymbol replaces the placeholder
                yield return (formatting, runText, isSpecial, cp, isObject ? null : picture);
                cp = runEndCp;
            }
        }
    }

    /// <summary>
    /// The drop-cap frame paragraph's runs inside a drop-cap span at the start of the text paragraph (LibreOffice ww8par.cxx
    /// m_bDropCap / m_xPreviousNode joins the two the same way). The runs keep font and color; the size and raised position
    /// Word gives them to fill the frame are dropped (ww8par.cxx: "reset the font size and text position for the dropcap").
    /// </summary>
    private void AppendDropCap(StringBuilder html, int cpStart, int cpEnd, WordRunFormatting paragraphBase, string? noteLabel)
    {
        (WordParagraphFormatting dropCap, int styleIndex) = ResolveParagraph(cpStart, cpEnd);
        WordDropCapHtml.Open(html, dropCap);
        DocBinaryFieldReader? outer = _fields;
        DocBinaryFieldReader fields = _fields = NewFieldReader(cpStart, restore: _textBoxDepth == 0);
        WordRunFormatting styleBase = _styles.ResolveParagraphStyle(styleIndex).RunFormatting;
        foreach ((WordRunFormatting runFormatting, string text, bool isSpecial, int runCp, uint? picture) in CharacterRuns(cpStart, cpEnd, dropCap.RunFormatting, styleBase))
        {
            runFormatting.FontSizeHalfPoints = paragraphBase.FontSizeHalfPoints;
            runFormatting.ComplexScriptFontSizeHalfPoints = paragraphBase.ComplexScriptFontSizeHalfPoints;
            runFormatting.PositionHalfPoints = paragraphBase.PositionHalfPoints;
            AppendRunHtml(html, runFormatting, paragraphBase, text, isSpecial, runCp, fields, noteLabel, picture);
        }
        if (fields.CloseLinkAtParagraphEnd()) html.Append("</a>");
        _fields = outer;
        _baselineShift = 0;
        WordDropCapHtml.Close(html);
    }

    private void AppendRunHtml(StringBuilder html, WordRunFormatting formatting, WordRunFormatting paragraphBase,
                               string text, bool isSpecial, int runCp, DocBinaryFieldReader fieldState, string? noteLabel, uint? picture = null)
    {
        var visible = new StringBuilder(text.Length);
        string css = formatting.ToCss(paragraphBase) + WordHtmlText.RunDirectionCss(formatting.IsRightToLeft, _isParagraphRightToLeft) + BaselineShiftCss;
        void Flush()
        {
            if (visible.Length == 0) return;
            if (css.Length > 0) html.Append("<span style=\"").Append(css).Append("\">");
            WordHtmlText.AppendScaled(html, visible.ToString(), formatting.CharacterScalePercent);
            if (css.Length > 0) html.Append("</span>");
            visible.Clear();
        }
        int nextBookmark = Bookmarks.FirstAtOrAfter(runCp);
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            while (nextBookmark < Bookmarks.StartCps.Length && Bookmarks.StartCps[nextBookmark] == runCp + i)
            {
                Flush();
                html.Append("<span class=\"bookmark\" id=\"").Append(WordHtmlText.Encode(DocxHtml.DocxBlockHtmlRenderer.BookmarkElementId(Bookmarks.Names[nextBookmark++]))).Append("\"></span>");
            }
            DocBinaryFieldReader.Step step = fieldState.Consume(c, runCp + i, isSpecial ? picture : null);
            if (step.IsConsumed)
            {
                if (step.LinkClose) { Flush(); html.Append("</a>"); }
                if (step.LinkOpen is { } linkOpen) { Flush(); html.Append(linkOpen); }
                if (step.Output.BaselineShiftPt != 0)
                {
                    // the text after an ADVANCE \u / \d: a new span from here on
                    Flush();
                    _baselineShift += step.Output.BaselineShiftPt;
                    css = formatting.ToCss(paragraphBase) + WordHtmlText.RunDirectionCss(formatting.IsRightToLeft, _isParagraphRightToLeft) + BaselineShiftCss;
                }
                if (step.Output.Text is not null || step.Output.Html is not null)
                {
                    // a computed result or placeholder, in the run's formatting, in the field's link when no link is open
                    if (step.OutputLink is { } outputLink) { Flush(); html.Append(outputLink); }
                    if (step.Output.Html is { } markup) { Flush(); html.Append(markup); }
                    else visible.Append(WordHtmlText.Encode(step.Output.Text!));
                    if (step.OutputLink is not null) { Flush(); html.Append("</a>"); }
                }
                continue;
            }
            switch (c)
            {
                case '\r': case CellOrRowMark: case '\u000C': case '\u000E': continue;   // paragraph / cell / section / column marks
                case '\t': visible.Append('\t'); continue;
                case '\u000B': visible.Append("<br>"); continue;
                case '\u001E': visible.Append('‑'); continue;
                case '\u001F': visible.Append('­'); continue;
                case '\u0001' when isSpecial && fieldState.IsInShapeResult && !(i + 1 < text.Length && text[i + 1] == '\u0001'): continue;
                case '\u0001' when isSpecial && picture is { } location:
                    Flush();
                    if (!formatting.IsHidden) AppendPicture(html, location, runCp + i);   // hidden text (sprmCFVanish): e.g. the "_PictureBullets" pictures
                    continue;
                case '\u0008' when isSpecial && _textBoxDepth == 0 && FloatingPictures?.At(runCp + i) is { } floating:
                    Flush();
                    WordPictureHtml.Append(html, floating.layout with { ColumnWidthPt = ColumnWidthPt, ColumnLeftPt = ColumnLeftPt }, floating.pictureId);
                    continue;
                case '\u0008' when isSpecial && _textBoxDepth == 0 && FloatingPictures?.Drawing(runCp + i) is { } drawing:
                    Flush();
                    AppendDrawing(html, drawing with { Layout = drawing.Layout with { ColumnWidthPt = ColumnWidthPt, ColumnLeftPt = ColumnLeftPt } });
                    continue;
                case '\u0001': case '\u0008': if (isSpecial) visible.Append("<span class=\"unsupported\">[image]</span>"); continue;
                case AutoNumberedNoteMark:
                    if (isSpecial) AppendNoteMark(visible, runCp + i, noteLabel);
                    continue;
                case '\u0005': continue;                                                  // annotation reference
            }
            if (c < ' ') continue;
            visible.Append(WordHtmlText.Encode(c.ToString()));
        }
        Flush();
    }

    private int _textBoxDepth;

    /// <summary>ADVANCE \u / \d: points the paragraph's later runs are raised (negative: lowered), as a relative offset.</summary>
    private double _baselineShift;
    private string BaselineShiftCss => _baselineShift == 0 ? "" : "position:relative;top:" + (-_baselineShift).ToString(System.Globalization.CultureInfo.InvariantCulture) + "pt;";


    /// <summary>
    /// A drawing object (shapes, a group, text boxes): a plain rectangle text box as <see cref="WordPictureHtml.OpenTextBox"/> with
    /// its line as the border and its fill; anything else as <see cref="WordShapeHtml"/> (one frame at the FSPA placement, the
    /// outlines in one SVG, each text box over its shape, group pictures at their box). A text box's paragraphs are its story's
    /// CP range rendered as block spans; shapes inside a text box are not drawn again (their anchors are in that story).
    /// </summary>
    private void AppendDrawing(StringBuilder html, DocBinaryDrawing drawing)
    {
        _textBoxDepth++;
        try
        {
            if (drawing.IsPlainTextBox)
            {
                WordShape box = drawing.Shapes[0];
                WordTextBoxStyle style = box.TextStyle with
                {
                    BorderColor = box.LineColor, BorderWidthPt = box.LineWidthPt, FillColor = box.FillColor, GrowsWithText = drawing.GrowsWithText,
                };
                WordPictureHtml.OpenTextBox(html, drawing.Layout, style);
                AppendStory(html, drawing.TextRanges[box.TextIndex]);
                WordPictureHtml.CloseTextBox(html);
                return;
            }
            bool inFlow = drawing.Shapes.Count == 1 && drawing.TextRanges.Count == 1 && drawing.GrowsWithText && drawing.Shapes[0].RotationDegrees == 0;
            WordShapeHtml.OpenFrame(html, drawing.Layout, inFlow);
            WordShapeHtml.AppendGeometry(html, drawing.Shapes, drawing.Layout.WidthPt, drawing.Layout.HeightPt, stretches: inFlow);
            foreach (WordShape shape in drawing.Shapes)
            {
                if (shape.Geometry == WordShapeGeometry.Picture) { WordShapeHtml.AppendPicture(html, shape); continue; }
                if (shape.TextIndex < 0) continue;
                WordShapeHtml.OpenShapeText(html, shape, inFlow);
                AppendStory(html, drawing.TextRanges[shape.TextIndex]);
                WordShapeHtml.CloseShapeText(html);
            }
            WordShapeHtml.CloseFrame(html);
        }
        finally { _textBoxDepth--; }
    }

    /// <summary>A text-box story range as paragraphs (each ends at a paragraph mark, the last at the range's end).</summary>
    private void AppendStory(StringBuilder html, DocBinaryShapeText story)
    {
        if (story.Literal is { } literal)
        {
            // WordArt: its own string, in its font and size, centred (a text path fits it to the shape)
            html.Append("<span class=\"text-box-paragraph\" style=\"display:block;text-align:center;white-space:pre-wrap;");
            if (story.FontFamily is { Length: > 0 } font) html.Append("font-family:'").Append(WordHtmlText.Encode(font.Replace("'", ""))).Append("';");
            if (story.Color is { } color) html.Append("color:").Append(color).Append(';');
            if (story.FontSizePt > 0) html.Append("font-size:").Append(story.FontSizePt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)).Append("pt;");
            html.Append("\">").Append(WordHtmlText.Encode(literal)).Append("</span>");
            return;
        }
        (int start, int end) range = (story.Start, story.End);
        DocBinaryFieldReader? outerFields = _fields;
        _fields = NewFieldReader(range.start, restore: false);           // a text box's own fields
        try { AppendStoryText(html, range); }
        finally { _fields = outerFields; }
    }

    private void AppendStoryText(StringBuilder html, (int start, int end) range)
    {
        string text = _pieces.ReadText(range.start, range.end);
        int paragraphStart = range.start;
        string? previousStyleKey = null;
        int? rowStart = null;                     // a table row of the story: its paragraphs up to the row-end mark, drawn as the body draws rows
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] is not ('\r' or CellOrRowMark) && i != text.Length - 1) continue;
            int end = range.start + i + 1;
            if (end > paragraphStart)
            {
                TableParagraphInfo table = ReadTableInfo(paragraphStart, end);
                if (table.IsInTable)
                {
                    rowStart ??= paragraphStart;
                    if (table.IsRowEnd) { html.Append(RenderRow(rowStart.Value, end)); rowStart = null; previousStyleKey = null; }
                }
                else
                {
                    if (rowStart is { } open) { html.Append(RenderRow(open, paragraphStart)); rowStart = null; }
                    previousStyleKey = AppendParagraph(html, paragraphStart, end, previousStyleKey, null);
                }
            }
            paragraphStart = end;
        }
        if (rowStart is { } last) html.Append(RenderRow(last, range.end));
    }

    /// <summary>The Data stream's pictures (inline 0x01 pictures); none for a document without a Data stream.</summary>
    internal DocBinaryPictures Pictures { get; set; } = new(null);

    /// <summary>The main text's floating (0x08) pictures; null when not set.</summary>
    /// <summary>The text width (points) of the section of the block being rendered: the column a floating object picks its nearer edge in (0 = unknown).</summary>
    public double ColumnWidthPt { get; set; }
    /// <summary>Where that column starts from the page's left edge (points; -1 = not known).</summary>
    public double ColumnLeftPt { get; set; } = -1;
    internal DocBinaryFloatingPictures? FloatingPictures { get; set; }

    /// <summary>
    /// An inline picture (a .doc 0x01 picture is always in the line: LibreOffice ww8par.cxx imports it FLY_AS_CHAR) at its PICF
    /// size; its id is the PICF offset (<see cref="IWordWindowedDocument.ReadPicture"/>). A picture we cannot show keeps a box.
    /// </summary>
    private void AppendPicture(StringBuilder html, uint location, int cp = -1)
    {
        if (EquationOf(location) is { } equation) { WordOfficeMathHtml.Append(html, equation, WordOfficeMathHtml.RunColorCss); return; }   // Word's picture of an equation: the equation itself
        if (ObjectIdBefore(cp) is { } objectId && Charts?.Invoke(objectId) is { } chart && Pictures.Locate(location) is { } box)
        {
            // an embedded Excel chart object: drawn from its chart records at the preview picture's size (the preview is not shown)
            WordShapeHtml.OpenFrame(html, new WordPictureLayout(box.widthPt, box.heightPt, WordPictureWrap.Inline, WordPictureSide.Left, 0, 0, 0, 0), growsWithText: false);
            html.Append(chart.Svg(box.widthPt, box.heightPt));
            WordShapeHtml.CloseFrame(html);
            return;
        }
        if (Pictures.Locate(location) is { } picture)
            WordPictureHtml.Append(html, new WordPictureLayout(picture.widthPt, picture.heightPt, WordPictureWrap.Inline, WordPictureSide.Left, 0, 0, 0, 0, Crop: picture.crop),
                                   location.ToString(System.Globalization.CultureInfo.InvariantCulture));
        else html.Append("<span class=\"unsupported\">[image]</span>");
    }

    /// <summary>
    /// The chart of an OLE object's 0x01 (sprmCFOle2 with sprmCPicLocation = its ObjectPool storage id, LibreOffice
    /// ww8par6.cxx:2950-2990 Read_Obj / Read_PicLoc, ww8par4.cxx:342): an embedded Excel chart read on demand; null when the
    /// object is not one (its preview picture is shown).
    /// </summary>
    internal Func<uint, DocxHtml.DocxChart?>? Charts { get; set; }

    /// <summary>The m:oMathPara / m:oMath of a picture's equationXML (<see cref="DocBinaryPictures.EquationXml"/>); null: none or unreadable.</summary>
    internal System.Xml.Linq.XElement? EquationOf(uint location)
    {
        if (Pictures.EquationXml(location) is not { Length: > 0 } xml) return null;
        try
        {
            System.Xml.Linq.XElement root = System.Xml.Linq.XElement.Parse(xml.Trim('\uFEFF', '\0'));   // the property's string ends in NULs
            return root.Descendants().FirstOrDefault(WordOfficeMathHtml.IsMathZone);
        }
        catch (System.Xml.XmlException) { return null; }
    }

    /// <summary>In the main text: a numbered note reference. Inside a note body: the note's own number.</summary>
    private void AppendNoteMark(StringBuilder html, int cp, string? noteLabel)
    {
        if (noteLabel is not null) { html.Append("<span class=\"note-mark\">").Append(WordHtmlText.Encode(noteLabel)).Append("</span>"); return; }
        if (_notes?.At(cp) is not { } reference) { html.Append("<sup class=\"note-reference\">*</sup>"); return; }
        string kind = reference.Kind == WordDocumentNoteKind.Footnote ? "footnote" : "endnote";
        html.Append("<sup class=\"note-reference\" data-note-kind=\"").Append(kind).Append("\" data-note-number=\"").Append(reference.Number).Append("\">")
            .Append(WordHtmlText.Encode(NoteLabel(reference))).Append("</sup>");
    }

    /// <summary>Word's defaults for .doc: footnotes 1, 2, 3; endnotes i, ii, iii.</summary>
    public static string NoteLabel(DocBinaryNoteReferences.NoteReference reference)
        => reference.Number == 0 ? "*" : reference.Kind == WordDocumentNoteKind.Footnote
            ? reference.Number.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : DocxHtml.DocxNumberingDefinitions.FormatNumber(reference.Number, "lowerRoman");

}
