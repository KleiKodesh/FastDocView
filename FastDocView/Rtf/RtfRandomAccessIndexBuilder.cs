using System.Text;
using FastDocView.DocxHtml;

namespace FastDocView.Rtf;

/// <summary>
/// The one streaming pass over an RTF file. Blocks: a paragraph ends at \par (or \sect) outside a table; a table
/// row (\trowd … \cell … \row) is one block, like the .doc reader's rows. A block's range starts right after the
/// previous block's terminator (or after the last header group), and the parser's group-state stack at that point
/// is stored with it (interned), so any block renders from its own bytes. List labels, note numbers and headings
/// are settled here; nothing of the text is kept.
/// </summary>
internal sealed class RtfRandomAccessIndexBuilder : RtfSink
{
    private const int HeadingTextLimit = 200;
    private const long MaxRowBytes = 4L << 20;

    private readonly RtfParser _parser;
    private readonly RtfDocumentTables _tables = new();
    private RtfListNumbering? _lists;
    private readonly IProgress<WordDocumentIndexProgress>? _progress;
    private readonly CancellationToken _cancellation;
    private readonly long _totalBytes;

    private readonly List<long> _starts = new(), _ends = new();
    private readonly List<WordDocumentBlockKind> _kinds = new();
    private readonly List<int> _textLengths = new(), _stateIds = new(), _labelIds = new();
    private readonly List<byte> _listLevels = new();
    private readonly List<byte[]> _states = new();
    private readonly Dictionary<ulong, List<int>> _stateIdsByHash = new();
    private byte[] _snapshotBuffer = new byte[1024];
    private readonly List<string> _labels = new();
    private readonly Dictionary<string, int> _labelIdByText = new();
    private readonly List<WordDocumentHeading> _headings = new();
    private readonly List<RtfNoteReference> _notes = new();
    private readonly List<string> _bookmarkNames = new();
    private readonly List<int> _bookmarkBlocks = new();

    // the block being read
    private bool _hasStart, _hasContent, _isRow;
    private long _start;
    private int _stateId;
    private int _textLength;
    private readonly StringBuilder _paragraphText = new();
    private readonly StringBuilder _listText = new();
    private bool _hasListText;
    private int _rowCells;
    private int _rightToLeftParagraphs, _paragraphs;

    // page setup (RTF 1.5 spec defaults: \paperw 12240, \margl / \margr 1800, \gutter 0); a \sectd
    // resets the section's values to the document's
    private int _paperWidth = 12240, _leftMargin = 1800, _rightMargin = 1800, _gutter;
    private int? _sectionPageWidth, _sectionLeftMargin, _sectionRightMargin, _sectionGutter;
    private readonly List<int> _sectionLastBlocks = new(), _sectionTextWidths = new();
    private bool _isSectionEndPending;                          // a \sect inside a table row: the section ends with the row

    // notes
    private bool _hasAutoMark;
    private int _footnoteCount, _endnoteCount;

    // span anchors of the compact index: the block start past every AnchorSpanBytes (and block 0), with the parser snapshot
    // + list state there (captured at TakeStart, committed when that block is added - a re-taken start replaces it)
    private readonly List<(int block, long start, byte[] state)> _anchors = new();
    private (int block, long start, byte[] state)? _pendingAnchor;
    private long _nextAnchorAt;
    private readonly bool _isSpanScan;

    private RtfRandomAccessIndexBuilder(Stream file, IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        _parser = new RtfParser(this, _tables, readsHeaderTables: true);
        _progress = progress;
        _cancellation = cancellation;
        _totalBytes = file.Length;
    }

    /// <summary>A span rescan: the same block rules over the header tables the index stores, from an anchor's parser snapshot and list state.</summary>
    private RtfRandomAccessIndexBuilder(RtfDocumentTables tables, long anchorOffset, byte[] scanState)
    {
        _tables = tables;
        _parser = new RtfParser(this, tables, readsHeaderTables: false);
        _isSpanScan = true;
        (byte[] snapshot, byte[] numbering) = WordIndexScanState.Split(scanState);
        if (snapshot.Length == 0) throw new InvalidDataException("RTF index: an anchor without a parser snapshot.");
        _parser.RestoreSnapshot(snapshot);
        _lists = new RtfListNumbering(tables);
        _lists.ImportState(numbering);
        TakeStart(anchorOffset);
    }

    public static RtfRandomAccessIndex Build(Stream file, long sourceLength, long sourceLastWriteTicksUtc,
                                             IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        var builder = new RtfRandomAccessIndexBuilder(file, progress, cancellation);
        builder._parser.Run(file, 0, file.Length);
        builder.EndBlock(builder._parser.Position, final: true);
        builder.EndSection();
        progress?.Report(new WordDocumentIndexProgress(file.Length, file.Length, builder._starts.Count));
        return builder.CreateIndex(sourceLength, sourceLastWriteTicksUtc);
    }

    /// <summary>
    /// Rescans one span of the compact index: the blocks from the anchor's start up to <paramref name="end"/> (the next anchor's
    /// start, or the last block's end with <paramref name="isLast"/>, which closes a trailing block the way the pass did).
    /// </summary>
    public static RtfRandomAccessIndexBuilder ScanSpan(Stream file, RtfDocumentTables tables, long anchorOffset, byte[] scanState, long end, bool isLast)
    {
        var scanner = new RtfRandomAccessIndexBuilder(tables, anchorOffset, scanState);
        scanner._parser.Run(file, anchorOffset, end);
        if (isLast) scanner.EndBlock(scanner._parser.Position, final: true);
        return scanner;
    }

    // the span's blocks (span-local indexes), read by RtfCompactBlockTable
    public IReadOnlyList<long> Starts => _starts;
    public IReadOnlyList<long> Ends => _ends;
    public IReadOnlyList<WordDocumentBlockKind> Kinds => _kinds;
    public IReadOnlyList<int> TextLengths => _textLengths;
    public IReadOnlyList<byte> ListLevels => _listLevels;
    public byte[] SnapshotOf(int localIndex) => _states[_stateIds[localIndex]];
    public string? LabelOf(int localIndex) => _labelIds[localIndex] >= 0 ? _labels[_labelIds[localIndex]] : null;

    private RtfRandomAccessIndex CreateIndex(long sourceLength, long sourceTicks)
    {
        bool isRightToLeft = _tables.IsRightToLeftDocument || (_paragraphs > 0 && _rightToLeftParagraphs * 2 >= _paragraphs);
        var anchors = new WordIndexSpanAnchor[_anchors.Count];
        byte[] previousState = [];
        for (int i = 0; i < anchors.Length; i++)
        {
            (int block, long start, byte[] state) = _anchors[i];
            anchors[i] = new WordIndexSpanAnchor(block, start, block > 0 ? _starts[block - 1] : -1, block > 0 ? _kinds[block - 1] : WordDocumentBlockKind.Other,
                                                 WordIndexScanState.Delta(previousState, state));
            previousState = state;
        }
        return new RtfRandomAccessIndex(sourceLength, sourceTicks, _tables, isRightToLeft, _starts.Count, _starts.Count > 0 ? _ends[_ends.Count - 1] : 0,
                                        anchors, WordIndexTextLengthGroups.Totals(_textLengths), _headings.ToArray(), _notes.ToArray())
        {
            SectionLastBlocks = _sectionLastBlocks.ToArray(),
            SectionTextWidths = _sectionTextWidths.ToArray(),
            // a bookmark after the last block's end (nothing followed it) marks the last block
            Bookmarks = new WordDocumentBookmarkTable(_bookmarkNames.ToArray(), _bookmarkBlocks.Select(block => Math.Min(block, Math.Max(0, _starts.Count - 1))).ToArray()),
        };
    }

    // ── block boundaries ─────────────────────────────────────────────────────────────

    private void TakeStart() => TakeStart(_parser.Position);

    private void TakeStart(long position)
    {
        _start = position;
        _stateId = InternState();
        _hasStart = true;
        _hasContent = false;
        _isRow = false;
        _textLength = 0;
        _rowCells = 0;
        _paragraphText.Clear();
        _hasListText = false;
        _listText.Clear();
        if (!_isSpanScan && (_starts.Count == 0 || _start >= _nextAnchorAt))
            _pendingAnchor = (_starts.Count, _start, WordIndexScanState.Join(_states[_stateId], _lists?.ExportState() ?? []));
    }

    private void EndBlock(long end, bool final = false, string? listLabel = null, int listLevel = -1)
    {
        if (!_hasStart || (final && !_hasContent)) return;
        int blockIndex = _starts.Count;
        _starts.Add(_start);
        _ends.Add(end);
        _kinds.Add(_isRow ? WordDocumentBlockKind.TableRow : WordDocumentBlockKind.Paragraph);
        _textLengths.Add(_textLength);
        _stateIds.Add(_stateId);
        _listLevels.Add(listLabel is null ? (byte)255 : (byte)Math.Max(0, Math.Min(254, listLevel)));
        _labelIds.Add(listLabel is null ? -1 : InternLabel(listLabel));
        if (_pendingAnchor is { } pending && pending.block == blockIndex)
        {
            _anchors.Add(pending);
            _nextAnchorAt = pending.start + RtfRandomAccessIndex.AnchorSpanBytes;
            _pendingAnchor = null;
        }
        if ((blockIndex & 4095) == 0)
        {
            _cancellation.ThrowIfCancellationRequested();
            _progress?.Report(new WordDocumentIndexProgress(end, _totalBytes, blockIndex));
        }
        if (!final) TakeStart();
        else _hasStart = false;
    }

    /// <summary>The id of the parser's current snapshot; equal stacks share one stored copy (most blocks start alike).</summary>
    private int InternState()
    {
        int size = _parser.SnapshotSize;
        if (_snapshotBuffer.Length < size) _snapshotBuffer = new byte[Math.Max(size, _snapshotBuffer.Length * 2)];
        Span<byte> snapshot = _snapshotBuffer.AsSpan(0, size);
        _parser.WriteSnapshot(snapshot);
        ulong hash = 14695981039346656037;
        foreach (byte b in snapshot) hash = (hash ^ b) * 1099511628211;
        if (!_stateIdsByHash.TryGetValue(hash, out List<int>? candidates)) _stateIdsByHash[hash] = candidates = new List<int>(1);
        foreach (int candidate in candidates)
            if (_states[candidate].AsSpan().SequenceEqual(snapshot)) return candidate;
        int id = _states.Count;
        _states.Add(snapshot.ToArray());
        candidates.Add(id);
        return id;
    }

    private int InternLabel(string label)
    {
        if (!_labelIdByText.TryGetValue(label, out int id))
        {
            id = _labels.Count;
            _labels.Add(label);
            _labelIdByText[label] = id;
        }
        return id;
    }

    // ── parser events ───────────────────────────────────────────────────────────────

    public override bool WantsText
    {
        get
        {
            if (_parser.IsInFootnote) return false;
            if (_parser.IsInListText) return true;
            if (_paragraphText.Length >= HeadingTextLimit) return false;
            ref readonly RtfGroupState state = ref _parser.Current;
            return state.OutlineLevel is >= 0 and <= 8 || OutlineLevelOfStyle(state.Style) >= 0;
        }
    }

    private int _memoStyle = int.MinValue, _memoOutline;

    private int OutlineLevelOfStyle(int style)
    {
        if (style != _memoStyle) { _memoStyle = style; _memoOutline = _tables.OutlineLevelOfStyle(style); }
        return _memoOutline;
    }

    public override void DocumentStarted() => TakeStart();

    public override void GroupClosed(int depth)
    {
        if (depth == 0 && _hasStart && !_hasContent) TakeStart();          // header tables and other root-level groups before the text
    }

    public override void Text(ReadOnlySpan<char> text, in RtfGroupState state)
    {
        if (_parser.IsInListText) { _hasContent = true; _listText.Append(text.ToArray()); return; }
        if (_parser.IsInFootnote) return;
        _hasContent = true;
        _textLength += text.Length;
        int room = HeadingTextLimit - _paragraphText.Length;
        if (room > 0) _paragraphText.Append(text.Slice(0, Math.Min(room, text.Length)).ToArray());
    }

    public override void TextLength(int length, in RtfGroupState state)
    {
        if (_parser.IsInFootnote || _parser.IsInListText) return;
        _hasContent = true;
        _textLength += length;
    }

    public override void ListTextStart() { _hasListText = true; _listText.Clear(); }

    public override void Embedded(RtfEmbeddedKind kind, in RtfGroupState state) { if (!_parser.IsInFootnote) { _hasContent = true; _textLength += 8; } }

    public override void Break(RtfWord word, in RtfGroupState state) { if (!_parser.IsInFootnote) _hasContent = true; }

    public override void PageSetup(RtfWord word, int value)
    {
        switch (word)
        {
            case RtfWord.PaperWidth: _paperWidth = value; break;
            case RtfWord.LeftMargin: _leftMargin = value; break;
            case RtfWord.RightMargin: _rightMargin = value; break;
            case RtfWord.Gutter: _gutter = value; break;
            case RtfWord.SectionDefault: _sectionPageWidth = _sectionLeftMargin = _sectionRightMargin = _sectionGutter = null; break;
            case RtfWord.SectionPageWidth: _sectionPageWidth = value; break;
            case RtfWord.SectionLeftMargin: _sectionLeftMargin = value; break;
            case RtfWord.SectionRightMargin: _sectionRightMargin = value; break;
            case RtfWord.SectionGutter: _sectionGutter = value; break;
        }
    }

    /// <summary>A \sect (or the end of the file) closes the section at the last block with the page setup in effect.</summary>
    private void EndSection()
    {
        if (_starts.Count == 0) return;
        int lastBlock = _starts.Count - 1;
        int width = WordDocumentPageSetup.TextWidth(_sectionPageWidth ?? _paperWidth, _sectionLeftMargin ?? _leftMargin,
                                                    _sectionRightMargin ?? _rightMargin, _sectionGutter ?? _gutter);
        if (_sectionLastBlocks.Count > 0 && _sectionLastBlocks[_sectionLastBlocks.Count - 1] >= lastBlock) return;
        _sectionLastBlocks.Add(lastBlock);
        _sectionTextWidths.Add(width);
    }

    public override void ParagraphEnd(in RtfGroupState state, bool isSectionEnd)
    {
        if (_parser.IsInFootnote || _parser.IsInListText) return;
        _hasContent = true;
        string? label = CountList(state, out bool hasListParagraph);
        // in a row until the row mark; a stray cell mark or in-table flag without one must not swallow the rest of the file
        bool isStray = (!state.IsInTable && _rowCells > 0) || _parser.Position - _start > MaxRowBytes;
        if ((state.IsInTable || _rowCells > 0) && !isStray)
        {
            _isRow = true;
            _paragraphText.Clear();
            _hasListText = false;
            if (isSectionEnd) _isSectionEndPending = true;          // the section ends with this row
            return;
        }
        _paragraphs++;
        if (state.Has(RtfStateFlags.RightToLeftParagraph)) _rightToLeftParagraphs++;
        int outline = state.OutlineLevel is >= 0 and <= 8 ? state.OutlineLevel : state.OutlineLevel == 9 ? -1 : OutlineLevelOfStyle(state.Style);
        if (outline >= 0)
        {
            string text = _paragraphText.ToString().Trim();
            if (text.Length > 0) _headings.Add(new WordDocumentHeading(_starts.Count, outline + 1, text));
        }
        string? storedLabel = _hasListText ? null : label;
        EndBlock(_parser.Position, listLabel: storedLabel, listLevel: hasListParagraph ? state.ListLevel : -1);
        if (isSectionEnd) EndSection();
    }

    /// <summary>Runs the list counter for a list paragraph; returns its label when the index should store one.</summary>
    private string? CountList(in RtfGroupState state, out bool hasListParagraph)
    {
        hasListParagraph = state.ListOverride > 0;
        if (!hasListParagraph) return null;
        _lists ??= new RtfListNumbering(_tables);
        return _lists.Next(state.ListOverride, state.ListLevel);
    }

    public override void CellEnd(in RtfGroupState state, bool isNested)
    {
        if (_parser.IsInFootnote) return;
        _hasContent = true;
        _isRow = true;
        if (!isNested) _rowCells++;
        CountList(state, out _);
        _paragraphText.Clear();
    }

    public override void RowEnd(in RtfGroupState state, bool isNested)
    {
        if (_parser.IsInFootnote || isNested) return;
        _hasContent = true;
        _isRow = true;
        EndBlock(_parser.Position);
        if (_isSectionEndPending) { _isSectionEndPending = false; EndSection(); }
    }

    public override void FootnoteMark(in RtfGroupState state)
    {
        if (!_parser.IsInFootnote) _hasAutoMark = true;
    }

    /// <summary>{\*\bkmkstart name}: the block being read is the bookmark's target (in a note: the block that holds the note).</summary>
    public override void BookmarkStart(string name)
    {
        _bookmarkNames.Add(name);
        _bookmarkBlocks.Add(_starts.Count);
        _hasContent = true;          // the bookmark's group belongs to this block (a root-level group would otherwise move the block start past it)
    }

    public override void FootnoteStart() => _hasContent = true;

    public override void FootnoteEnd(bool isEndnote)
    {
        WordDocumentNoteKind kind = isEndnote ? WordDocumentNoteKind.Endnote : WordDocumentNoteKind.Footnote;
        int number = 0;
        string label = "*";
        if (_hasAutoMark)
        {
            number = isEndnote ? ++_endnoteCount : ++_footnoteCount;
            int shown = number - 1 + (isEndnote ? _tables.EndnoteStart : _tables.FootnoteStart);
            label = FormatNoteNumber(shown, isEndnote ? _tables.EndnoteFormat : _tables.FootnoteFormat);
        }
        _hasAutoMark = false;
        _notes.Add(new RtfNoteReference(_starts.Count, kind, number, label));
    }

    /// <summary>\ftnn* formats; Chicago (*, †, ‡, § then doubled) is from memory, not source-checked.</summary>
    internal static string FormatNoteNumber(int number, string format)
    {
        if (format != "chicago") return DocxNumberingDefinitions.FormatNumber(number, format);
        string[] symbols = ["*", "\u2020", "\u2021", "\u00A7"];
        if (number <= 0) return "*";
        return string.Concat(Enumerable.Repeat(symbols[(number - 1) % 4], (number - 1) / 4 + 1));
    }
}
