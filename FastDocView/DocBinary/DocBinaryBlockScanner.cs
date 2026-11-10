using FastDocView.Html;

namespace FastDocView.DocBinary;

/// <summary>
/// The block state machine of a .doc's main text: fed the text's characters in CP order, it ends a paragraph at every
/// paragraph / cell / page mark, looks the paragraph's PAPX up once to tell table paragraphs and row ends apart, and
/// emits a block per paragraph or per whole table row, with its kind, list label (counted here,
/// <see cref="DocBinaryListCounter"/>) and heading. The index pass runs it over the whole text; the compact block
/// table runs the very same machine over one span (from an anchor's CP with the anchor's list state), so every value
/// it recovers equals the one the pass found.
/// </summary>
internal sealed class DocBinaryBlockScanner
{
    private const int MaxHeadingTextLength = 300;

    private readonly DocBinaryParagraphHtmlRenderer _renderer;
    private readonly DocBinaryStyleSheet _styles;
    private readonly DocBinaryListCounter _listCounter;
    private readonly Dictionary<string, int> _listLabelIds = new(StringComparer.Ordinal);
    private int _paragraphStart;
    private int _rowStart = -1;                     // first CP of the table row being collected, -1 outside a row

    public List<int> BlockStartCps { get; } = new();
    public List<WordDocumentBlockKind> BlockKinds { get; } = new();
    /// <summary>Per block: an index into <see cref="ListLabels"/>, -1 when the block is not a list item.</summary>
    public List<int> BlockListLabelIds { get; } = new();
    public List<string> ListLabels { get; } = new();
    /// <summary>Per block: an index into <see cref="ListContexts"/> (REF \w \r), -1 when the block is not a list item.</summary>
    public List<int> BlockListContextIds { get; } = new();
    public List<string> ListContexts { get; } = new();
    private readonly Dictionary<string, int> _listContextIds = new();
    public List<WordDocumentHeading> Headings { get; } = new();
    /// <summary>
    /// Outer rows whose cells the next row continues (fVertMerge without fVertRestart below, at the same right edge): the row's
    /// block and a bit per cell. A row is one block, so a cell merged down is drawn as cells without the line between them;
    /// which cell goes on below needs the next row, so it is found here, at index time.
    /// </summary>
    public List<int> MergedRowBlocks { get; } = new();
    public List<ulong> MergedRowCells { get; } = new();
    private DocBinaryParagraphHtmlRenderer.TableParagraphInfo? _previousRow;      // the row block just added, while rows follow each other
    /// <summary>The CP every block of the scan so far ends at (the next block's start) - the last block's end once <see cref="Finish"/> ran.</summary>
    public int ScannedEndCp => _paragraphStart;
    public int BlockCount => BlockStartCps.Count;

    /// <summary>Raised after a block is added, with its index and start CP and whether the list state captured by <see cref="PendingAnchorState"/> is for it.</summary>
    public event Action<int, int>? BlockAdded;

    /// <summary>
    /// The index pass asks for the list state at the start of the next block past a CP (an anchor): the scanner captures it
    /// at the next paragraph start outside a table row, before that paragraph's counter advances, and hands it over with the
    /// block that starts there.
    /// </summary>
    public int CaptureStateAtOrAfterCp { get; set; } = int.MaxValue;
    /// <summary>False for the search's scan: only block boundaries are needed, so no heading text is read.</summary>
    public bool CollectsHeadings { get; init; } = true;
    private (int cp, byte[] state)? _pendingState;
    public byte[]? StateOfLastBlock { get; private set; }

    public DocBinaryBlockScanner(DocBinaryParagraphHtmlRenderer renderer, DocBinaryStyleSheet styles, DocBinaryListDefinitions lists, int startCp = 0)
    {
        _renderer = renderer;
        _styles = styles;
        _listCounter = new DocBinaryListCounter(lists);
        _paragraphStart = startCp;
    }

    /// <summary>Restores the list-numbering state an anchor recorded (what the span's blocks count on from).</summary>
    public void ImportListState(byte[] state) => _listCounter.ImportState(state);

    /// <summary>Feeds one chunk of main-text characters starting at <paramref name="chunkCp"/>.</summary>
    public void Feed(char[] characters, int count, int chunkCp)
    {
        for (int i = 0; i < count; i++)
            if (characters[i] is '\r' or '\u0007' or '\u000C') OnParagraphEnd(chunkCp + i + 1, characters[i]);
    }

    /// <summary>At the main text's end: closes an unfinished row or paragraph; a text with no mark at all is one block.</summary>
    public void Finish(int mainTextCharacterCount)
    {
        CloseOpenRow();
        if (_paragraphStart < mainTextCharacterCount) AddBlock(_paragraphStart, mainTextCharacterCount, WordDocumentBlockKind.Paragraph);
        if (BlockStartCps.Count == 0) AddBlock(0, mainTextCharacterCount, WordDocumentBlockKind.Paragraph);
    }

    /// <summary>
    /// Closes a row that has no row-end mark the way the pass does when the next non-table paragraph starts: an inner span's
    /// rescan stops at the next anchor (that paragraph's start) without seeing it, so the table calls this at the span's end.
    /// </summary>
    public void CloseOpenRow()
    {
        if (_rowStart >= 0) { AddBlock(_rowStart, _paragraphStart, WordDocumentBlockKind.TableRow); _rowStart = -1; _previousRow = null; }
    }

    /// <summary>
    /// This row continues cells of the row before it (fVertMerge without fVertRestart - LibreOffice ww8par2.cxx:2582-2620): the cell
    /// of the row before at the same right edge (within 3 twips, ww8struc.hxx:536) is marked, so that row is drawn without its
    /// bottom line there.
    /// </summary>
    private void NoteCellsMergedDown(in DocBinaryParagraphHtmlRenderer.TableParagraphInfo row)
    {
        if (_previousRow is not { Cells: { } aboveCells, CellEdgesTwips: { } aboveEdges } || row.Cells is not { } cells || row.CellEdgesTwips is not { } edges) return;
        ulong mask = 0;
        for (int i = 0; i < cells.Length && i + 1 < edges.Length; i++)
        {
            if (!cells[i].ContinuesAbove) continue;
            for (int j = 0; j < aboveCells.Length && j + 1 < aboveEdges.Length && j < 64; j++)
                if (aboveCells[j].IsVerticallyMerged && Math.Abs(aboveEdges[j + 1] - edges[i + 1]) <= 3) { mask |= 1UL << j; break; }
        }
        if (mask == 0) return;
        MergedRowBlocks.Add(BlockStartCps.Count - 2);
        MergedRowCells.Add(mask);
    }

    private void AddBlock(int start, int end, WordDocumentBlockKind kind)
    {
        BlockStartCps.Add(start);
        BlockKinds.Add(kind);
        BlockListLabelIds.Add(-1);
        BlockListContextIds.Add(-1);
        StateOfLastBlock = _pendingState is { } pending && pending.cp == start ? pending.state : null;
        if (StateOfLastBlock is not null) _pendingState = null;
        BlockAdded?.Invoke(BlockStartCps.Count - 1, start);
    }

    private void SetListLabel(string label)
    {
        if (!_listLabelIds.TryGetValue(label, out int id)) { id = ListLabels.Count; ListLabels.Add(label); _listLabelIds[label] = id; }
        BlockListLabelIds[BlockListLabelIds.Count - 1] = id;
    }

    private void SetListContext(string context)
    {
        if (!_listContextIds.TryGetValue(context, out int id)) { id = ListContexts.Count; ListContexts.Add(context); _listContextIds[context] = id; }
        BlockListContextIds[BlockListContextIds.Count - 1] = id;
    }

    private void OnParagraphEnd(int end, char mark)
    {
        int start = _paragraphStart;
        _paragraphStart = end;
        if (_rowStart < 0 && start >= CaptureStateAtOrAfterCp && _pendingState is null) _pendingState = (start, _listCounter.ExportState());
        DocBinaryParagraphHtmlRenderer.TableParagraphInfo table = _renderer.ReadTableInfo(start, end);
        (WordParagraphFormatting formatting, int styleIndex) = _renderer.ResolveParagraph(start, end);
        // Every list paragraph advances its counter, cell paragraphs included — a list that runs through a table keeps counting.
        string? label = formatting.NumberingId > 0 ? _listCounter.Next(formatting.NumberingId, formatting.NumberingLevel) : null;
        string? context = label is { Length: > 0 } ? _listCounter.ContextOf(formatting.NumberingId, formatting.NumberingLevel) : null;
        if (table.IsInTable || mark == '\u0007')
        {
            if (_rowStart < 0) _rowStart = start;
            if (table.IsRowEnd)
            {
                AddBlock(_rowStart, end, WordDocumentBlockKind.TableRow);
                _rowStart = -1;
                NoteCellsMergedDown(table);
                _previousRow = table;
            }
            return;                                    // a row is one block: its cell paragraphs' labels are not shown yet
        }
        if (_rowStart >= 0) { AddBlock(_rowStart, start, WordDocumentBlockKind.TableRow); _rowStart = -1; }   // a row without its end mark
        _previousRow = null;
        AddBlock(start, end, WordDocumentBlockKind.Paragraph);
        if (label is { Length: > 0 }) SetListLabel(label);
        if (context is not null) SetListContext(context);
        if (!CollectsHeadings) return;
        int level = _styles.HeadingLevelOf(styleIndex);
        if (level > 0)
        {
            string text = _renderer.PlainText(start, Math.Min(end, start + MaxHeadingTextLength));
            if (text.Length > 0) Headings.Add(new WordDocumentHeading(BlockStartCps.Count - 1, level, text));
        }
    }
}
