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
    /// Table facts of a paragraph's PAPX: inside a table, ends a row, and — on a row end — the
    /// column widths (twips), right-to-left column order, per-cell shading and the row's border.
    /// </summary>
    public readonly record struct TableParagraphInfo(bool IsInTable, bool IsRowEnd, int[]? CellWidthsTwips,
                                                     bool IsRightToLeft = false, string?[]? CellShadingHex = null, string? BorderCss = null);

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

    /// <summary>sprmPFInTable / sprmPFTtp / sprmPFInnerTtp and sprmTDefTable of the paragraph ending at cpEnd.</summary>
    public TableParagraphInfo ReadTableInfo(int cpStart, int cpEnd)
    {
        DocBinaryFormattedDiskPages.PropertyRun? run = ParagraphRunOf(cpStart, cpEnd);
        bool isInTable = false, isRowEnd = false, isRightToLeft = false;
        int[]? widths = null;
        string?[]? shading = null;
        string? border = null;
        void ReadSprms(byte[] grpprl, int start, int length)
        {
            foreach ((ushort opcode, int operandAt, int operandLength) in DocBinaryPropertyModifiers.Sprms(grpprl, start, length))
            {
                switch (opcode)
                {
                    case 0x2416: isInTable = grpprl[operandAt] != 0; break;                     // sprmPFInTable
                    case 0x2417: case 0x244C: isRowEnd = grpprl[operandAt] != 0; break;         // sprmPFTtp / sprmPFInnerTtp
                    case 0xD608: widths = CellWidths(grpprl, operandAt, operandLength); break;  // sprmTDefTable
                    case 0x560B: isRightToLeft = BitConverter.ToUInt16(grpprl, operandAt) != 0; break;   // sprmTFBiDi
                    case 0xD612: shading = CellShading(grpprl, operandAt, operandLength); break;         // sprmTCellShd
                    case 0xD613: border = TableBorder(grpprl, operandAt, operandLength); break;          // sprmTTableBorders
                }
            }
        }
        if (run is { GrpprlLength: > 2 } papx)
        {
            (byte[] grpprl, int start, int length) = _pages.ExpandHugePapx(papx.Page, papx.GrpprlStart + 2, papx.GrpprlLength - 2);
            ReadSprms(grpprl, start, length);
        }
        if (PrmGrpprlOfParagraph(cpStart, cpEnd) is { } prm) ReadSprms(prm, 0, prm.Length);     // a Prm0 can carry sprmPFInTable / sprmPFTtp
        return new TableParagraphInfo(isInTable || isRowEnd, isRowEnd, widths, isRightToLeft, shading, border);
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

    /// <summary>sprmTTableBorders: six Brc (8 bytes: cv, width in eighths of a point, type) — the top one stands for the row.</summary>
    private static string? TableBorder(byte[] grpprl, int operandAt, int operandLength)
    {
        if (operandLength < 8) return null;
        byte type = grpprl[operandAt + 5];
        if (type is 0 or 0xFF) return "none";
        double width = Math.Max(0.25, grpprl[operandAt + 4] / 8.0);
        string color = grpprl[operandAt + 3] == 0xFF ? "#000" : $"#{grpprl[operandAt]:X2}{grpprl[operandAt + 1]:X2}{grpprl[operandAt + 2]:X2}";
        return $"{width.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt {(type == 3 ? "double" : "solid")} {color}";
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
    public string Render(int cpStart, int cpEnd, string? previousStyleKey = null, string? noteLabel = null, string? listLabel = null)
    {
        var html = new StringBuilder(512);
        AppendParagraph(html, cpStart, cpEnd, previousStyleKey, noteLabel, listLabel);
        return html.ToString();
    }

    private string? AppendParagraph(StringBuilder html, int cpStart, int cpEnd, string? previousStyleKey, string? noteLabel, string? listLabel = null)
    {
        (WordParagraphFormatting formatting, int styleIndex) = ResolveParagraph(cpStart, cpEnd);
        int headingLevel = _styles.HeadingLevelOf(styleIndex);
        string tag = headingLevel is >= 1 and <= 6 ? "h" + headingLevel : "p";

        WordHtmlText.AppendOpenTag(html, tag, formatting.ToCss(previousStyleKey), null, formatting.IsRightToLeft);
        _isParagraphRightToLeft = formatting.IsRightToLeft;
        if (listLabel is { Length: > 0 }) AppendListLabel(html, listLabel, formatting);
        WordRunFormatting styleBase = _styles.ResolveParagraphStyle(styleIndex).RunFormatting;
        var fieldState = new FieldState();
        foreach ((WordRunFormatting runFormatting, string text, bool isSpecial, int runCp) in CharacterRuns(cpStart, cpEnd, formatting.RunFormatting, styleBase))
            AppendRunHtml(html, runFormatting, formatting.RunFormatting, text, isSpecial, runCp, fieldState, noteLabel);
        html.Append("</").Append(tag).Append('>');
        return formatting.StyleKey;
    }

    /// <summary>
    /// The label drawn as Word does: the paragraph's character formatting plus the level's grpprlChpx, followed by its
    /// ixchFollow (LVLF, [MS-DOC] 2.9.150; LibreOffice ww8par3.cxx: 0 tab, 1 space, 2 nothing).
    /// </summary>
    private void AppendListLabel(StringBuilder html, string label, WordParagraphFormatting formatting)
    {
        DocBinaryListDefinitions.LevelDefinition? level = Lists.LevelOfParagraph(formatting.NumberingId, formatting.NumberingLevel);
        string? css = null;
        if (level is { CharacterGrpprl.Length: > 0 })
        {
            WordRunFormatting labelFormatting = formatting.RunFormatting.Clone();
            DocBinaryPropertyModifiers.ApplyCharacterSprms(labelFormatting, level.CharacterGrpprl, 0, level.CharacterGrpprl.Length, fonts: _styles.Fonts);
            css = labelFormatting.ToCss(formatting.RunFormatting);
        }
        WordListLabelSuffix suffix = level?.FollowCharacter switch { 1 => WordListLabelSuffix.Space, 2 => WordListLabelSuffix.Nothing, _ => WordListLabelSuffix.Tab };
        WordListLabelAlignment alignment = level?.Justification switch { 1 => WordListLabelAlignment.Center, 2 => WordListLabelAlignment.End, _ => WordListLabelAlignment.Start };
        WordListLabelHtml.Append(html, label, css, suffix, formatting, DefaultTabStopTwips, alignment);
    }

    /// <summary>One table row: cells end at cell marks, the row at its row-end mark (the last paragraph).</summary>
    public string RenderRow(int cpStart, int cpEnd)
    {
        string text = _pieces.ReadText(cpStart, cpEnd);
        var paragraphs = new List<(int start, int end, bool endsCell)>();
        int paragraphStart = cpStart;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] is not ('\r' or CellOrRowMark)) continue;
            paragraphs.Add((paragraphStart, cpStart + i + 1, text[i] == CellOrRowMark));
            paragraphStart = cpStart + i + 1;
        }
        if (paragraphs.Count == 0) return Render(cpStart, cpEnd);

        (int lastStart, int lastEnd, _) = paragraphs[paragraphs.Count - 1];
        TableParagraphInfo rowInfo = ReadTableInfo(lastStart, lastEnd);
        (WordParagraphFormatting rowFormatting, _) = ResolveParagraph(lastStart, lastEnd);

        var html = new StringBuilder(1024);
        WordHtmlText.AppendOpenTag(html, "table", "table-layout:fixed;", "table-row", rowInfo.IsRightToLeft || rowFormatting.IsRightToLeft);
        html.Append("<tr>");
        int cellIndex = 0;
        bool isCellOpen = false;
        string? previousStyleKey = null;
        for (int i = 0; i < paragraphs.Count - 1; i++)          // the last paragraph is the row-end mark itself
        {
            (int start, int end, bool endsCell) = paragraphs[i];
            if (!isCellOpen)
            {
                var cellCss = new StringBuilder();
                if (rowInfo.CellWidthsTwips is { } widths && cellIndex < widths.Length && widths[cellIndex] > 0)
                    cellCss.Append("width:").Append(WordParagraphFormatting.TwipsToPt(widths[cellIndex])).Append("pt;");
                if (rowInfo.CellShadingHex is { } shading && cellIndex < shading.Length && shading[cellIndex] is { } fill)
                    cellCss.Append("background:#").Append(fill).Append(';');
                if (rowInfo.BorderCss is { } border) cellCss.Append("border:").Append(border).Append(';');
                html.Append("<td");
                if (cellCss.Length > 0) html.Append(" style=\"").Append(cellCss).Append('"');
                html.Append('>');
                isCellOpen = true;
                previousStyleKey = null;
            }
            previousStyleKey = AppendParagraph(html, start, end, previousStyleKey, null);
            if (endsCell) { html.Append("</td>"); isCellOpen = false; cellIndex++; }
        }
        if (isCellOpen) html.Append("</td>");
        html.Append("</tr></table>");
        return html.ToString();
    }

    /// <summary>A footnote/endnote body: its paragraphs; the leading auto-number mark shows <paramref name="label"/>.</summary>
    public string RenderNote(int textStartCp, int textEndCp, string label)
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
    public void FeedSearchText(int cpStart, int cpEnd, WordDocumentSearchMatcher matcher)
    {
        var fieldState = new FieldState();
        foreach (char c in _pieces.ReadText(cpStart, cpEnd))
        {
            if (matcher.Found) return;
            if (fieldState.Consume(c)) continue;
            if (c >= ' ') matcher.Feed(c);
            else if (c == '\t') matcher.Feed(' ');
            else if (c is '\r' or '\u0007' or '\u000B' or '\u000C') matcher.Separator();
        }
    }

    public string PlainText(int cpStart, int cpEnd)
    {
        var text = new StringBuilder();
        var fieldState = new FieldState();
        foreach (char c in _pieces.ReadText(cpStart, cpEnd))
        {
            if (fieldState.Consume(c)) continue;
            if (c >= ' ' || c == '\t') text.Append(c == '\t' ? ' ' : c);
        }
        return text.ToString().Trim();
    }

    private IEnumerable<(WordRunFormatting formatting, string text, bool isSpecial, int cp)> CharacterRuns(
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
                if (run is { GrpprlLength: > 0 } direct)
                {
                    int? characterStyle = DocBinaryPropertyModifiers.StyleIndexIn(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                    if (characterStyle is { } index && _styles[index] is { IsParagraphStyle: false } style)
                        DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting, style.CharacterGrpprl, 0, style.CharacterGrpprl.Length, styleBase, _styles.Fonts);
                    DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting, direct.Page, direct.GrpprlStart, direct.GrpprlLength, styleBase, _styles.Fonts);
                    isSpecial = DocBinaryPropertyModifiers.HasSpecialCharacterFlag(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                    symbol = DocBinaryPropertyModifiers.SymbolCharacterIn(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                }
                if (piece.PrmGrpprl is { } prm)             // fast-saved files: the piece's Prm applies after the CHPX ([MS-DOC] §2.4.6.2)
                {
                    DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting, prm, 0, prm.Length, styleBase, _styles.Fonts);
                    isSpecial |= DocBinaryPropertyModifiers.HasSpecialCharacterFlag(prm, 0, prm.Length);
                    symbol ??= DocBinaryPropertyModifiers.SymbolCharacterIn(prm, 0, prm.Length);
                }
                string runText = _pieces.ReadText(cp, runEndCp);
                if (symbol is { } symbolCharacter) runText = new string(symbolCharacter, runText.Length);   // sprmCSymbol replaces the placeholder
                yield return (formatting, runText, isSpecial, cp);
                cp = runEndCp;
            }
        }
    }

    private void AppendRunHtml(StringBuilder html, WordRunFormatting formatting, WordRunFormatting paragraphBase,
                               string text, bool isSpecial, int runCp, FieldState fieldState, string? noteLabel)
    {
        var visible = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (fieldState.Consume(c)) continue;
            switch (c)
            {
                case '\r': case CellOrRowMark: case '\u000C': case '\u000E': continue;   // paragraph / cell / section / column marks
                case '\t': visible.Append('\t'); continue;
                case '\u000B': visible.Append("<br>"); continue;
                case '\u001E': visible.Append('‑'); continue;
                case '\u001F': visible.Append('­'); continue;
                case '\u0001': case '\u0008': if (isSpecial) visible.Append("<span class=\"unsupported\">[image]</span>"); continue;
                case AutoNumberedNoteMark:
                    if (isSpecial) AppendNoteMark(visible, runCp + i, noteLabel);
                    continue;
                case '\u0005': continue;                                                  // annotation reference
            }
            if (c < ' ') continue;
            visible.Append(WordHtmlText.Encode(c.ToString()));
        }
        if (visible.Length == 0) return;
        string css = formatting.ToCss(paragraphBase) + WordHtmlText.RunDirectionCss(formatting.IsRightToLeft, _isParagraphRightToLeft);
        if (css.Length > 0) html.Append("<span style=\"").Append(css).Append("\">").Append(visible).Append("</span>");
        else html.Append(visible);
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

    /// <summary>
    /// Field codes are hidden, field results shown: 0x13 begins a field, 0x14 separates code from
    /// result, 0x15 ends it. Tracked per nesting level so a nested field inside a result works.
    /// </summary>
    private sealed class FieldState
    {
        private readonly Stack<bool> _isInCode = new();

        /// <summary>True when the character is consumed (a field mark, or inside a field code).</summary>
        public bool Consume(char c)
        {
            switch (c)
            {
                case '\u0013': _isInCode.Push(true); return true;
                case '\u0014': if (_isInCode.Count > 0) { _isInCode.Pop(); _isInCode.Push(false); } return true;
                case '\u0015': if (_isInCode.Count > 0) _isInCode.Pop(); return true;
            }
            foreach (bool isCode in _isInCode) if (isCode) return true;
            return false;
        }
    }
}
