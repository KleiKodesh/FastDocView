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
    private CancellationToken _cancellation;

    /// <summary>An on-demand open hands the pass to the document: later steps are not tied to the open's cancellation.</summary>
    public void ClearCancellation() => _cancellation = default;
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

    // fields open at each block start and the SEQ / AUTONUM / LISTNUM counters there (RtfFieldTracker), interned: the
    // renderer and the search restore them, so a field running across blocks and a result-less counted field read right
    private RtfFieldTracker _fields = new(null, counts: true);
    private readonly List<int> _fieldStateIds = new();
    private readonly List<string> _fieldStates = new() { "" };
    private readonly Dictionary<string, int> _fieldStateIdByText = new() { [""] = 0 };
    private int _fieldStateId;

    /// <summary>What a field without a result is computed from, for the text a span search reads (the renderer's context).</summary>
    public Fields.WordFieldContext? FieldContext
    {
        set
        {
            string state = _fields.Serialize();
            _fields = new RtfFieldTracker(value, counts: true);
            _fields.Restore(state);
        }
    }

    /// <summary>The fields open where span-local block <paramref name="localIndex"/> starts (<see cref="RtfFieldTracker.Serialize"/>).</summary>
    public string FieldStateOf(int localIndex) => _fieldStates[_fieldStateIds[localIndex]];

    private int InternFieldState()
    {
        string state = _fields.Serialize();
        if (!_fieldStateIdByText.TryGetValue(state, out int id))
        {
            id = _fieldStates.Count;
            _fieldStates.Add(state);
            _fieldStateIdByText[state] = id;
        }
        return id;
    }

    /// <summary>
    /// An anchor's head: the parser snapshot and the field state, as varint(snapshot length) + snapshot + the field state in
    /// UTF-8 (index format 7).
    /// </summary>
    internal static byte[] AnchorHead(byte[] snapshot, string fieldState)
    {
        var bytes = new List<byte>(snapshot.Length + fieldState.Length + 4);
        IndexBinaryArrays.AppendVarint(bytes, (uint)snapshot.Length);
        bytes.AddRange(snapshot);
        bytes.AddRange(Encoding.UTF8.GetBytes(fieldState));
        return bytes.ToArray();
    }

    internal static (byte[] snapshot, string fieldState) SplitAnchorHead(byte[] head)
    {
        if (head.Length == 0) return ([], "");
        int position = 0;
        int length = (int)IndexBinaryArrays.ReadVarint(head, ref position);
        if (length < 0 || length > head.Length - position) throw new InvalidDataException("RTF index: anchor head damaged.");
        byte[] snapshot = new byte[length];
        Buffer.BlockCopy(head, position, snapshot, 0, length);
        return (snapshot, Encoding.UTF8.GetString(head, position + length, head.Length - position - length));
    }

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
    /// <summary>Blocks the direction majority is counted over (= the on-demand first open's first blocks).</summary>
    internal const int DirectionSampleBlocks = 100;

    // page setup (RTF 1.5 spec defaults: \paperw 12240, \margl / \margr 1800, \gutter 0); a \sectd
    // resets the section's values to the document's
    private int _paperWidth = 12240, _leftMargin = 1800, _rightMargin = 1800, _gutter;
    private int? _sectionPageWidth, _sectionLeftMargin, _sectionRightMargin, _sectionGutter;
    private readonly List<int> _sectionLastBlocks = new(), _sectionTextWidths = new(), _sectionColumnLefts = new();
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
        (byte[] head, byte[] numbering) = WordIndexScanState.Split(scanState);
        (byte[] snapshot, string fieldState) = SplitAnchorHead(head);
        if (snapshot.Length == 0) throw new InvalidDataException("RTF index: an anchor without a parser snapshot.");
        _parser.RestoreSnapshot(snapshot);
        _fields.Restore(fieldState);
        _lists = new RtfListNumbering(tables);
        _lists.ImportState(numbering);
        TakeStart(anchorOffset);
    }

    public static RtfRandomAccessIndex Build(Stream file, long sourceLength, long sourceLastWriteTicksUtc,
                                             IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        RtfRandomAccessIndexBuilder builder = Incremental(file, sourceLength, sourceLastWriteTicksUtc, progress, cancellation);
        builder.AdvanceUntilKnown(int.MaxValue);
        return builder.Complete();
    }

    /// <summary>
    /// The index pass run in steps (an on-demand first open): <see cref="AdvanceUntilKnown"/> parses on until that many blocks
    /// are known, pausing the parser right after a block ends; <see cref="CreatePartialIndex"/> is the index of the blocks known
    /// so far (the same compact form: a window reads them as it reads a finished index); <see cref="Complete"/> the final one.
    /// One parse either way: the same blocks and values as one run to the end.
    /// </summary>
    public static RtfRandomAccessIndexBuilder Incremental(Stream file, long sourceLength, long sourceLastWriteTicksUtc,
                                                          IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
        => new(file, progress, cancellation) { _mainFile = file, _sourceLength = sourceLength, _sourceTicks = sourceLastWriteTicksUtc };

    private Stream? _mainFile;
    private long _sourceLength, _sourceTicks;
    private bool _isMainStarted;

    /// <summary>The index pass reached the end of the file: every block is known.</summary>
    public bool IsEveryBlockKnown { get; private set; }

    /// <summary>Blocks known so far (final values: each has ended).</summary>
    public int KnownBlockCount => _starts.Count;

    /// <summary>Parses on until at least <paramref name="blocks"/> blocks are known or the file ends.</summary>
    public void AdvanceUntilKnown(int blocks)
    {
        _cancellation.ThrowIfCancellationRequested();       // also when the blocks asked for are known already
        while (!IsEveryBlockKnown && _starts.Count < blocks)
        {
            _cancellation.ThrowIfCancellationRequested();
            _pauseAtBlockCount = blocks;
            if (!_isMainStarted) { _isMainStarted = true; _parser.Run(_mainFile!, 0, _totalBytes); }
            else _parser.Continue();
            if (_parser.IsPaused) continue;
            EndBlock(_parser.Position, final: true);
            EndSection();
            _progress?.Report(new WordDocumentIndexProgress(_totalBytes, _totalBytes, _starts.Count));
            IsEveryBlockKnown = true;
            _mainFile = null;
        }
        _pauseAtBlockCount = int.MaxValue;
    }

    /// <summary>The finished index (parses on to the end first when needed).</summary>
    public RtfRandomAccessIndex Complete()
    {
        if (!IsEveryBlockKnown) AdvanceUntilKnown(int.MaxValue);
        return CreateIndex(_sourceLength, _sourceTicks, isPartial: false);
    }

    /// <summary>
    /// The index of the blocks known so far: its last span ends at the last known block's end, so a window reads those blocks
    /// exactly as it will from the final index. Notes and bookmarks of the block still being read are left out; the open
    /// section is closed provisionally at the last known block. Never saved.
    /// </summary>
    public RtfRandomAccessIndex CreatePartialIndex() => CreateIndex(_sourceLength, _sourceTicks, isPartial: !IsEveryBlockKnown);

    /// <summary>
    /// A resumable rescan of one span of the compact index: the blocks from the anchor's start up to <paramref name="end"/> (the
    /// next anchor's start, or the last block's end with <paramref name="isLast"/>, which closes a trailing block the way the pass
    /// did). Nothing is read here; <see cref="ScanThrough"/> parses only as far as the asked-for block, pausing the parser right
    /// after that block ends, and a later call resumes there (the same parse as one run to the end: same blocks, same values).
    /// </summary>
    public static RtfRandomAccessIndexBuilder ResumableSpan(Stream file, RtfDocumentTables tables, long anchorOffset, byte[] scanState, long end, bool isLast)
        => new(tables, anchorOffset, scanState) { _spanFile = file, _spanEnd = end, _isLastSpan = isLast };

    /// <summary>The whole span rescanned in one go (the reference the resumable scan must equal).</summary>
    public static RtfRandomAccessIndexBuilder ScanSpan(Stream file, RtfDocumentTables tables, long anchorOffset, byte[] scanState, long end, bool isLast)
    {
        RtfRandomAccessIndexBuilder scanner = ResumableSpan(file, tables, anchorOffset, scanState, end, isLast);
        scanner.ScanThrough(int.MaxValue - 1);
        return scanner;
    }

    // a resumable span rescan
    private Stream? _spanFile;
    private long _spanEnd;
    private bool _isLastSpan, _isSpanStarted;
    private int _pauseAtBlockCount = int.MaxValue;

    /// <summary>The span has been parsed to its end: every block of it is known.</summary>
    public bool IsComplete { get; private set; }

    /// <summary>Parses on until span-local block <paramref name="localIndex"/> has ended (or the span is complete).</summary>
    public void ScanThrough(int localIndex)
    {
        while (!IsComplete && _starts.Count <= localIndex)
        {
            _pauseAtBlockCount = localIndex + 1;
            if (!_isSpanStarted) { _isSpanStarted = true; _parser.Run(_spanFile!, _start, _spanEnd); }
            else _parser.Continue();
            if (_parser.IsPaused) continue;
            if (_isLastSpan) EndBlock(_parser.Position, final: true);
            IsComplete = true;
            _spanFile = null;
        }
    }

    /// <summary>
    /// Feeds the matcher while the span goes on being scanned: the blocks from the next one to end through span-local block
    /// <paramref name="lastLocalIndex"/> - their text and separators from the same parse that finds their bounds (one pass).
    /// </summary>
    public void Search(RtfSpanSearch search, int lastLocalIndex)
    {
        _search = search;
        _searchFrom = _starts.Count;
        try { ScanThrough(lastLocalIndex); }
        finally { _search = null; }
    }

    private RtfSpanSearch? _search;
    private int _searchFrom;
    private bool IsSearching => _search is not null && _starts.Count >= _searchFrom;

    // the span's blocks (span-local indexes), read by RtfCompactBlockTable
    public IReadOnlyList<long> Starts => _starts;
    public IReadOnlyList<long> Ends => _ends;
    public IReadOnlyList<WordDocumentBlockKind> Kinds => _kinds;
    public IReadOnlyList<int> TextLengths => _textLengths;
    public IReadOnlyList<byte> ListLevels => _listLevels;

    // the paragraph style (\sN) of each block's first paragraph, from the scan itself (span rescans recompute it: no index storage)
    private readonly List<int> _blockStyles = new();
    private int _blockStyle = -1;
    public IReadOnlyList<int> BlockStyles => _blockStyles;
    private readonly List<int> _blockLists = new();
    private int _blockList = -1;
    public IReadOnlyList<int> BlockLists => _blockLists;
    public byte[] SnapshotOf(int localIndex) => _states[_stateIds[localIndex]];
    public string? LabelOf(int localIndex) => _labelIds[localIndex] >= 0 ? _labels[_labelIds[localIndex]] : null;

    private RtfRandomAccessIndex CreateIndex(long sourceLength, long sourceTicks, bool isPartial = false)
    {
        int known = _starts.Count;
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
                                        anchors, WordIndexTextLengthGroups.Totals(_textLengths),
                                        isPartial ? _headings.Where(h => h.BlockIndex < known).ToArray() : _headings.ToArray(),
                                        isPartial ? _notes.Where(n => n.BlockIndex < known).ToArray() : _notes.ToArray())
        {
            SectionLastBlocks = SectionsFor(isPartial, known, _sectionLastBlocks, known - 1),
            SectionTextWidths = SectionsFor(isPartial, known, _sectionTextWidths, CurrentSectionTextWidth()),
            SectionColumnLefts = SectionsFor(isPartial, known, _sectionColumnLefts, WordDocumentPageSetup.ColumnLeft(_sectionLeftMargin ?? _leftMargin, _sectionGutter ?? _gutter)),
            // a bookmark after the last block's end (nothing followed it) marks the last block
            Bookmarks = isPartial
                ? new WordDocumentBookmarkTable(_bookmarkNames.Where((_, i) => _bookmarkBlocks[i] < known).ToArray(), _bookmarkBlocks.Where(block => block < known).ToArray())
                : new WordDocumentBookmarkTable(_bookmarkNames.ToArray(), _bookmarkBlocks.Select(block => Math.Min(block, Math.Max(0, known - 1))).ToArray()),
            IndexEntries = _indexEntries.ToArray(),      // entries of the block being read too: an INDEX lists them as found so far
        };
    }

    private int CurrentSectionTextWidth() => WordDocumentPageSetup.TextWidth(_sectionPageWidth ?? _paperWidth, _sectionLeftMargin ?? _leftMargin,
                                                                            _sectionRightMargin ?? _rightMargin, _sectionGutter ?? _gutter);

    /// <summary>A partial index closes the open section at the last known block (the values in effect there).</summary>
    private int[] SectionsFor(bool isPartial, int known, List<int> closed, int openValue)
    {
        bool isOpen = isPartial && known > 0 && (_sectionLastBlocks.Count == 0 || _sectionLastBlocks[_sectionLastBlocks.Count - 1] < known - 1);
        if (!isOpen) return closed.ToArray();
        var values = new int[closed.Count + 1];
        closed.CopyTo(values);
        values[closed.Count] = openValue;
        return values;
    }

    // ── block boundaries ─────────────────────────────────────────────────────────────

    private void TakeStart() => TakeStart(_parser.Position);

    private void TakeStart(long position)
    {
        _start = position;
        _stateId = InternState();
        _fieldStateId = InternFieldState();
        _hasStart = true;
        _hasContent = false;
        _isRow = false;
        _textLength = 0;
        _blockStyle = -1;
        _blockList = -1;
        _rowCells = 0;
        _paragraphText.Clear();
        _hasListText = false;
        _listText.Clear();
        if (!_isSpanScan && (_starts.Count == 0 || _start >= _nextAnchorAt))
            _pendingAnchor = (_starts.Count, _start, WordIndexScanState.Join(AnchorHead(_states[_stateId], _fieldStates[_fieldStateId]), _lists?.ExportState() ?? []));
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
        _fieldStateIds.Add(_fieldStateId);
        _listLevels.Add(listLabel is null ? (byte)255 : (byte)Math.Max(0, Math.Min(254, listLevel)));
        _blockStyles.Add(_blockStyle);
        _blockLists.Add(_blockList);
        _labelIds.Add(listLabel is null ? -1 : InternLabel(listLabel));
        if (_search is not null && blockIndex >= _searchFrom) _search.BlockEnded(blockIndex);
        if (!final && _starts.Count >= _pauseAtBlockCount) _parser.RequestPause();
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

    public override bool WantsText => IsSearching || IndexWantsText;

    private bool IndexWantsText
    {
        get
        {
            if (_parser.IsInFootnote || _parser.IsInTextBox) return false;
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
        if (!_fields.IsIdle) _fields.GroupClosed(depth, (output, _) => SearchOutput(output));
        if (depth == 0 && _hasStart && !_hasContent) TakeStart();          // header tables and other root-level groups before the text
    }

    // fields: tracked for the block-start state, and for the text a span search reads (the renderer's rules)
    public override void FieldStart(int depth) => _fields.Start(depth);
    public override void FieldModifier(RtfWord word) { if (word == RtfWord.FieldLock) _fields.Lock(); else _fields.Private(); }
    public override void FormField(Fields.WordFormFieldState state) => _fields.Form(state);
    public override void FieldInstruction(string instruction)
    {
        _fields.Instruction(instruction);
        // XE / TA entries for a result-less INDEX / TOA (the full pass only; a field in another field's code is part of that code)
        if (!_isSpanScan && _fields.Depth == 1 && WordFieldInstructionName(instruction) is "XE" or "TA")
            _indexEntries.Add(Math.Max(0, _starts.Count).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\u0001" + instruction);
    }

    private readonly List<string> _indexEntries = new();

    private static string WordFieldInstructionName(string instruction)
    {
        string trimmed = instruction.TrimStart();
        int end = 0;
        while (end < trimmed.Length && char.IsLetter(trimmed[end])) end++;
        return trimmed.Substring(0, end).ToUpperInvariant();
    }
    public override void FieldResultStart() => SearchOutput(_fields.ResultStart(out _));

    private void SearchOutput(Fields.WordFieldOutput output)
    {
        if (IsSearching && output is { Text: { } text, IsSearchable: true }) _search!.Text(text.AsSpan(), _parser);
    }

    public override void Text(ReadOnlySpan<char> text, in RtfGroupState state)
    {
        if (IsSearching)
        {
            if (_fields.IsIdle || _fields.Shows(text)) _search!.Text(text, _parser);
            if (!IndexWantsText) { TextLength(_parser.TextByteCount, state); return; }   // what the index pass saw
        }
        if (_parser.IsInTextBox) return;              // text-box text: searched, not the block's own text (the shape counts as a placeholder)
        if (_parser.IsInListText) { _hasContent = true; _listText.Append(text.ToArray()); return; }
        if (_parser.IsInFootnote) return;
        _hasContent = true;
        _textLength += text.Length;
        int room = HeadingTextLimit - _paragraphText.Length;
        if (room > 0) _paragraphText.Append(text.Slice(0, Math.Min(room, text.Length)).ToArray());
    }

    public override void TextLength(int length, in RtfGroupState state)
    {
        if (_parser.IsInFootnote || _parser.IsInListText || _parser.IsInTextBox) return;
        _hasContent = true;
        _textLength += length;
    }

    public override void ListTextStart() { if (!_parser.IsInTextBox) { _hasListText = true; _listText.Clear(); } }

    /// <summary>A drawing object counts as the placeholder it was before shapes were read (block values unchanged); its text-box text is searched only.</summary>
    public override void ShapeStart(in RtfGroupState state) { if (IsSearching) _search!.Separator(); Embedded(RtfEmbeddedKind.Picture, state); }
    public override void ShapeEnd(RtfShapePlacement shape, in RtfGroupState state) { if (IsSearching) _search!.Separator(); }
    public override void TextBoxParagraphEnd(in RtfGroupState state) { if (IsSearching) _search!.Separator(); }

    public override void Embedded(RtfEmbeddedKind kind, in RtfGroupState state) { if (!_parser.IsInFootnote) { _hasContent = true; _textLength += 8; } }

    public override void Break(RtfWord word, in RtfGroupState state)
    {
        if (IsSearching) _search!.Separator();
        if (!_parser.IsInFootnote) _hasContent = true;
    }

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
        int width = CurrentSectionTextWidth();
        if (_sectionLastBlocks.Count > 0 && _sectionLastBlocks[_sectionLastBlocks.Count - 1] >= lastBlock) return;
        _sectionLastBlocks.Add(lastBlock);
        _sectionTextWidths.Add(width);
        _sectionColumnLefts.Add(WordDocumentPageSetup.ColumnLeft(_sectionLeftMargin ?? _leftMargin, _sectionGutter ?? _gutter));
    }

    public override void ParagraphEnd(in RtfGroupState state, bool isSectionEnd)
    {
        if (IsSearching) _search!.Separator();
        if (_parser.IsInFootnote || _parser.IsInListText) return;
        if (_blockStyle < 0 && !_parser.IsInTextBox) { _blockStyle = state.Style; _blockList = state.ListOverride; }   // the block's (first) paragraph style and list (\ls) - STYLEREF, REF \r \w
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
        // the document's direction is decided from its first DirectionSampleBlocks blocks: the data an on-demand first open
        // has too, so the first open and a reopen from the saved index show the same direction
        if (_starts.Count < DirectionSampleBlocks)
        {
            _paragraphs++;
            if (state.Has(RtfStateFlags.RightToLeftParagraph)) _rightToLeftParagraphs++;
        }
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
        if (IsSearching) _search!.Separator();
        if (_parser.IsInFootnote) return;
        _hasContent = true;
        _isRow = true;
        if (!isNested) _rowCells++;
        CountList(state, out _);
        _paragraphText.Clear();
    }

    public override void RowEnd(in RtfGroupState state, bool isNested)
    {
        if (IsSearching) _search!.Separator();
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

    public override void FootnoteStart()
    {
        if (IsSearching) _search!.Separator();
        _hasContent = true;
    }

    public override void FootnoteEnd(bool isEndnote)
    {
        if (IsSearching) _search!.Separator();
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
