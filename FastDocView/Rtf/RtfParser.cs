using System.Runtime.InteropServices;
using System.Text;

namespace FastDocView.Rtf;

/// <summary>
/// A streaming RTF reader: bytes in, group state + content events out (<see cref="RtfSink"/>). Never builds a
/// document model. Token rules from the RTF 1.5 spec ("RTF syntax", "Conventions of an RTF reader", "Unicode RTF")
/// and LibreOffice rtftokenizer.cxx: a control word is up to 32 ASCII letters, an optional '-' and digits, and a
/// space delimiter that is consumed; a control symbol is one non-letter; \'hh is one byte; \* before an unknown
/// destination skips the group; \binN is N raw bytes; \uN is one UTF-16 unit followed by \ucN skipped "characters"
/// (a byte, a \'hh, or any control word / symbol counts as one; a brace ends the skipping).
/// <para>Two modes. Index mode (<see cref="RtfDocumentTables"/> filled from the header) reads the whole file once;
/// render mode restores a block's saved group-state stack (<see cref="RestoreSnapshot"/>) and reads only that
/// block's bytes, skipping the header tables it is given.</para>
/// </summary>
internal sealed class RtfParser
{
    private const int BufferSize = 1 << 16;
    private const int MaxWordLength = 32;
    private const int MaxDepth = 200;

    private readonly RtfSink _sink;
    private readonly RtfDocumentTables _tables;
    private readonly bool _readsHeaderTables;
    private Stream? _source;
    private readonly byte[] _buffer;
    private int _position, _end;
    private long _bufferOffset, _limit;

    private RtfGroupState[] _stack = new RtfGroupState[32];
    private int _depth = -1;                        // index of the current group; -1 before '{\rtf' / after it closed
    private bool _isIgnorableNext;                  // \* seen
    private int _skipCount;                         // characters still to skip after \uN

    // contexts that end with their group
    private int _footnoteDepth = -1, _listTextDepth = -1, _unicodePairDepth = -1, _nestedTablePropertiesDepth = -1;
    private int _captureDepth = -1;                 // a {\*\fldinst} / {\*\bkmkstart} group whose text is captured
    // a {\*\formfield} inside a field's code: its type, result, default, list entries and default text
    private int _formFieldDepth = -1, _formTextDepth = -1;
    private int _formType, _formResult = 25, _formDefault;
    private bool _formTextIsEntry;
    private readonly List<string> _formEntries = new();
    private string? _formDefaultText;
    private readonly System.Text.StringBuilder _formText = new();
    private readonly StringBuilder _captured = new();
    private bool _isEndnote;

    // drawing objects ({\shp} / {\shpgrp}): the open shapes (outermost first), the {\sn} / {\sv} being read, the open {\shptxt}.
    // A block never ends inside a shape (its \par are text-box paragraph ends), so none of this is in a snapshot.
    private readonly List<RtfShapePlacement> _shapes = new();
    private int _shapePropertyDepth = -1, _textBoxDepth = -1;
    private string? _shapePropertyName;
    private readonly StringBuilder _shapePropertyText = new();

    // table row definition (\trowd ... \cellxN), document-level as in LibreOffice
    private readonly List<RtfCellDefinition> _cells = new();
    private RtfCellDefinition _pendingCell;
    private readonly RtfBorder[] _rowBorders = new RtfBorder[6];      // \trbrdrt / l / b / r / h / v
    // the nested row definition being read ({\*\nesttableprops \trowd ... \cellx ... \nestrow}): Word writes it right before its \nestrow
    private readonly List<RtfCellDefinition> _nestedCells = new();
    private RtfCellDefinition _pendingNestedCell;
    private readonly RtfBorder[] _nestedRowBorders = new RtfBorder[6];
    private bool _isNestedRowRightToLeft;
    private bool _isRowRightToLeft;
    private int _rowLeft;          // \trleft: the row's leftmost edge from the column's, twips
    private byte _rowAlignment;    // \trql 0 / \trqc 1 / \trqr 2

    // pending 8-bit text (decoded in one go: double-byte code pages need the byte pairs together)
    private byte[] _textBytes = new byte[4096];
    private int _textByteCount;
    private ushort _textCodePage;
    private char[] _chars = new char[4096];

    // header capture (index mode)
    private readonly StringBuilder _headerText = new();
    private int _headerDepth = -1;                  // the depth of the table destination (fonttbl, colortbl, ...)
    private int _entryDepth = -1;                   // a style / font entry group
    private int _fontNumber = int.MinValue, _fontCharset = -1, _fontCodePage = -1;
    private bool _isParagraphStyleEntry = true;
    private int _red = -1, _green = -1, _blue = -1;
    private RtfDocumentTables.ListDefinition? _list;
    private List<RtfDocumentTables.ListLevelDefinition>? _listLevels;
    private RtfDocumentTables.ListLevelDefinition? _level;
    private int _listDepth = -1, _levelDepth = -1, _levelTextDepth = -1, _levelNumbersDepth = -1;
    private RtfDocumentTables.ListOverride? _override;
    private int _overrideDepth = -1, _overrideLevelDepth = -1, _overrideLevelIndex = -1;
    private int _overrideStartAt = int.MinValue;

    private readonly byte[] _word = new byte[MaxWordLength + 1];

    public RtfParser(RtfSink sink, RtfDocumentTables tables, bool readsHeaderTables)
    {
        _sink = sink;
        _tables = tables;
        _readsHeaderTables = readsHeaderTables;
        _buffer = new byte[BufferSize];
    }

    public RtfDocumentTables Tables => _tables;

    /// <summary>Absolute file offset of the next unread byte.</summary>
    public long Position => _bufferOffset + _position;

    public int Depth => _depth;
    public ref readonly RtfGroupState Current => ref _stack[_depth < 0 ? 0 : _depth];
    public bool IsInFootnote => _footnoteDepth >= 0;
    public bool IsInListText => _listTextDepth >= 0;
    /// <summary>Inside a shape's {\shptxt}: text-box text (not the block's own text: no heading text, length or list counting).</summary>
    public bool IsInTextBox => _textBoxDepth >= 0;
    public bool IsRowRightToLeft => _isRowRightToLeft;
    public int RowLeftTwips => _rowLeft;
    /// <summary>0 left (\trql, the default), 1 centre (\trqc), 2 right (\trqr).</summary>
    public int RowAlignment => _rowAlignment;
    /// <summary>The row definition's cells (\cellx with what came before each).</summary>
    public IReadOnlyList<RtfCellDefinition> Cells => _cells;
    /// <summary>\trbrdrt / l / b / r / h / v of the row definition.</summary>
    public IReadOnlyList<RtfBorder> RowBorders => _rowBorders;
    /// <summary>The last nested row definition ({\*\nesttableprops}), read before its \nestrow.</summary>
    public IReadOnlyList<RtfCellDefinition> NestedCells => _nestedCells;
    public IReadOnlyList<RtfBorder> NestedRowBorders => _nestedRowBorders;
    public bool IsNestedRowRightToLeft => _isNestedRowRightToLeft;

    /// <summary>Reads <paramref name="source"/> from <paramref name="start"/> up to <paramref name="end"/> (exclusive).</summary>
    public void Run(Stream source, long start, long end)
    {
        _source = source;
        source.Position = start;
        _bufferOffset = start;
        _position = _end = 0;
        _limit = end;
        _isPauseRequested = _isPaused = false;
        Parse();
        if (!_isPaused) FlushText();
    }

    private bool _isPauseRequested, _isPaused;

    /// <summary>The sink asks the run to stop before the next token (<see cref="IsPaused"/>); <see cref="Continue"/> goes on from there.</summary>
    public void RequestPause() => _isPauseRequested = true;

    /// <summary>The last <see cref="Run"/> / <see cref="Continue"/> stopped on a <see cref="RequestPause"/>, not at its end.</summary>
    public bool IsPaused => _isPaused;

    /// <summary>Resumes a paused run exactly where it stopped (the read-ahead buffer is kept; each fill seeks, so the source may be moved meanwhile).</summary>
    public void Continue()
    {
        if (!_isPaused || _source is null) return;
        _isPaused = false;
        Parse();
        if (!_isPaused) FlushText();
    }

    /// <summary>The file bytes behind the text just given to <see cref="RtfSink.Text"/> (what <see cref="RtfSink.TextLength"/> would have said).</summary>
    public int TextByteCount { get; private set; }

    // ── snapshots ─────────────────────────────────────────────────────────────────────

    /// <summary>The group-state stack, row definition and open contexts, as bytes — what a window needs to resume here.</summary>
    public byte[] Snapshot()
    {
        byte[] bytes = new byte[SnapshotSize];
        WriteSnapshot(bytes);
        return bytes;
    }

    private static readonly int CellDefinitionSize = Marshal.SizeOf<RtfCellDefinition>();
    private static readonly int BorderSize = Marshal.SizeOf<RtfBorder>();

    public int SnapshotSize => 4 + (_depth + 1) * RtfGroupState.Size + 4 * 4 + 1 + 4 + 1 + 6 * BorderSize + 4 + _cells.Count * CellDefinitionSize;

    /// <summary>Writes the snapshot into <paramref name="bytes"/> (at least <see cref="SnapshotSize"/> long).</summary>
    public void WriteSnapshot(Span<byte> bytes)
    {
        int frames = _depth + 1;
        int at = 0;
        WriteInt(bytes, ref at, frames);
        MemoryMarshal.AsBytes(_stack.AsSpan(0, frames)).CopyTo(bytes.Slice(at));
        at += frames * RtfGroupState.Size;
        WriteInt(bytes, ref at, _footnoteDepth);
        WriteInt(bytes, ref at, _listTextDepth);
        WriteInt(bytes, ref at, _unicodePairDepth);
        WriteInt(bytes, ref at, _nestedTablePropertiesDepth);
        bytes[at++] = (byte)(_isRowRightToLeft ? 1 : 0);
        WriteInt(bytes, ref at, _rowLeft);
        bytes[at++] = _rowAlignment;
        MemoryMarshal.AsBytes(_rowBorders.AsSpan()).CopyTo(bytes.Slice(at));
        at += 6 * BorderSize;
        WriteInt(bytes, ref at, _cells.Count);
        for (int i = 0; i < _cells.Count; i++)
        {
            RtfCellDefinition cell = _cells[i];
            MemoryMarshal.Write(bytes.Slice(at), ref cell);
            at += CellDefinitionSize;
        }
    }

    public void RestoreSnapshot(byte[] bytes)
    {
        int at = 0;
        int frames = ReadInt(bytes, ref at);
        if (_stack.Length < frames + 8) _stack = new RtfGroupState[frames + 32];
        MemoryMarshal.Cast<byte, RtfGroupState>(bytes.AsSpan(at, frames * RtfGroupState.Size)).CopyTo(_stack);
        at += frames * RtfGroupState.Size;
        _depth = frames - 1;
        _footnoteDepth = ReadInt(bytes, ref at);
        _listTextDepth = ReadInt(bytes, ref at);
        _unicodePairDepth = ReadInt(bytes, ref at);
        _nestedTablePropertiesDepth = ReadInt(bytes, ref at);
        _isRowRightToLeft = bytes[at++] != 0;
        _rowLeft = ReadInt(bytes, ref at);
        _rowAlignment = bytes[at++];
        MemoryMarshal.Cast<byte, RtfBorder>(bytes.AsSpan(at, 6 * BorderSize)).CopyTo(_rowBorders);
        at += 6 * BorderSize;
        int cells = ReadInt(bytes, ref at);
        _cells.Clear();
        for (int i = 0; i < cells; i++)
        {
            _cells.Add(MemoryMarshal.Read<RtfCellDefinition>(bytes.AsSpan(at)));
            at += CellDefinitionSize;
        }
        _nestedCells.Clear(); _pendingNestedCell = default; Array.Clear(_nestedRowBorders, 0, 6); _isNestedRowRightToLeft = false;
        _isIgnorableNext = false;
        _skipCount = 0;
        _textByteCount = 0;
        _captureDepth = -1;
        _captured.Clear();
        _pendingCell = default;
        _isEndnote = false;
        _shapes.Clear();
        _shapePropertyDepth = _textBoxDepth = -1;
        _shapePropertyName = null;
        _shapePropertyText.Clear();
    }

    private static void WriteInt(Span<byte> bytes, ref int at, int value)
    {
        bytes[at++] = (byte)value; bytes[at++] = (byte)(value >> 8); bytes[at++] = (byte)(value >> 16); bytes[at++] = (byte)(value >> 24);
    }

    private static int ReadInt(byte[] bytes, ref int at)
    {
        int value = bytes[at] | bytes[at + 1] << 8 | bytes[at + 2] << 16 | bytes[at + 3] << 24;
        at += 4;
        return value;
    }

    // ── byte input ────────────────────────────────────────────────────────────────────

    private bool Fill()
    {
        _bufferOffset += _end;
        _position = _end = 0;
        long remaining = _limit - _bufferOffset;
        if (remaining <= 0 || _source is null) return false;
        int wanted = (int)Math.Min(_buffer.Length, remaining);
        _source.Position = _bufferOffset;                     // the stream may be shared (a span rescan resumed between two fills)
        int read = 0;
        while (read < wanted)
        {
            int got = _source.Read(_buffer, read, wanted - read);
            if (got <= 0) break;
            read += got;
        }
        _end = read;
        if (read < wanted) _limit = _bufferOffset + read;     // the file is shorter than the index said
        return read > 0;
    }

    private int ReadByte() => _position < _end || Fill() ? _buffer[_position++] : -1;

    private int PeekByte() => _position < _end || Fill() ? _buffer[_position] : -1;

    private void SkipBytes(long count)
    {
        while (count > 0)
        {
            if (_position >= _end && !Fill()) return;
            int step = (int)Math.Min(count, _end - _position);
            _position += step;
            count -= step;
        }
    }

    private static readonly bool[] IsTextStop = BuildStops("\\{}\r\n");
#if NET
    private static readonly System.Buffers.SearchValues<byte> SkipStops = System.Buffers.SearchValues.Create("\\{}"u8);
#endif
    private static readonly bool[] IsSkipStop = BuildStops("\\{}");

    private static bool[] BuildStops(string characters)
    {
        var stops = new bool[256];
        foreach (char c in characters) stops[c] = true;
        return stops;
    }

    // ── the token loop ───────────────────────────────────────────────────────────────

    private void Parse()
    {
        while (true)
        {
            if (_isPauseRequested) { _isPauseRequested = false; _isPaused = true; return; }
            if (_position >= _end && !Fill()) return;
            byte b = _buffer[_position];
            switch (b)
            {
                case (byte)'{':
                    _position++;
                    FlushText();
                    if (_depth >= MaxDepth) { SkipGroupBody(); break; }     // pathological nesting: bounded stack and snapshots
                    Push();
                    break;
                case (byte)'}':
                    _position++;
                    FlushText();
                    if (!Pop()) return;
                    break;
                case (byte)'\\':
                    _position++;
                    ReadControl();
                    break;
                case (byte)'\r':
                case (byte)'\n':
                    _position++;
                    break;
                default:
                {
                    int start = _position;
                    // a plain loop: Hebrew text arrives as short runs between 'hh escapes, too short for a vectorised search
                    bool[] stops = IsTextStop;
                    int end = _end;
                    byte[] buffer = _buffer;
                    int at = start + 1;
                    while (at < end && !stops[buffer[at]]) at++;
                    _position = at;
                    if (_depth >= 0) AddTextBytes(_buffer.AsSpan(start, at - start));
                    break;
                }
            }
        }
    }

    private void Push()
    {
        _skipCount = 0;
        if (_depth + 1 >= _stack.Length) Array.Resize(ref _stack, _stack.Length * 2);
        if (_depth >= 0) _stack[_depth + 1] = _stack[_depth];
        else _stack[0] = RtfGroupState.Initial((short)_tables.DefaultFont, (ushort)_tables.AnsiCodePage);
        _depth++;
        if (_depth == 0) _sink.DocumentStarted();
        ref RtfGroupState state = ref _stack[_depth];
        if (_unicodePairDepth >= 0 && _depth == _unicodePairDepth + 1) state.Set(RtfStateFlags.DropText, true);
        if (_readsHeaderTables && _headerDepth >= 0 && _depth == _headerDepth + 1)
        {
            RtfDestination parent = _stack[_headerDepth].Destination;
            if (parent == RtfDestination.StyleSheet)
            {
                _entryDepth = _depth;
                _isParagraphStyleEntry = true;
                state.Style = 0;
                state.OutlineLevel = -1;
                _headerText.Clear();
            }
            else if (parent == RtfDestination.FontTable)
            {
                _entryDepth = _depth;
            }
        }
    }

    /// <summary>Closes the current group; false when the document group itself closed.</summary>
    private bool Pop()
    {
        _skipCount = 0;
        _isIgnorableNext = false;
        if (_depth < 0) return true;                  // stray '}' after the end
        if (_depth == _footnoteDepth) { _footnoteDepth = -1; _sink.FootnoteEnd(_isEndnote); }
        if (_depth == _listTextDepth) { _listTextDepth = -1; _sink.ListTextEnd(); }
        if (_depth == _unicodePairDepth) _unicodePairDepth = -1;
        if (_depth == _nestedTablePropertiesDepth) _nestedTablePropertiesDepth = -1;
        if (_depth == _formTextDepth)
        {
            _formTextDepth = -1;
            if (_formTextIsEntry) _formEntries.Add(_formText.ToString()); else _formDefaultText = _formText.ToString();
            _formText.Clear();
        }
        if (_depth == _formFieldDepth)
        {
            _formFieldDepth = -1;
            // \ffres 25 = undefined: the default holds ([MS-DOC] FFDataBits, as LibreOffice rtfdispatchvalue.cxx:1449-1455 reads it)
            int result = _formResult == 25 ? _formDefault : _formResult;
            _sink.FormField(_formType switch
            {
                1 => new Fields.WordFormFieldState { IsChecked = result != 0 },
                2 => new Fields.WordFormFieldState { Entries = _formEntries.ToList(), SelectedEntry = result },
                _ => new Fields.WordFormFieldState { DefaultText = _formDefaultText },
            });
        }
        if (_depth == _captureDepth)
        {
            _captureDepth = -1;
            string captured = _captured.ToString();
            _captured.Clear();
            if (_stack[_depth].Destination == RtfDestination.BookmarkStart) { if (captured.Trim().Length > 0) _sink.BookmarkStart(captured.Trim()); }
            else if (_stack[_depth].Destination == RtfDestination.BookmarkEnd) { if (captured.Trim().Length > 0) _sink.BookmarkEnd(captured.Trim()); }
            else _sink.FieldInstruction(captured);
        }
        if (_shapes.Count > 0) CloseShapeGroup();
        if (_readsHeaderTables) CloseHeaderGroup();
        _depth--;
        _sink.GroupClosed(_depth);
        return _depth >= 0;
    }

    private void ReadControl()
    {
        int first = ReadByte();
        if (first < 0) return;
        if (!IsLetter(first))
        {
            if (first == '\'')
            {
                int high = ReadByte();
                if (high < 0) return;
                int value = HexValue(high);
                int next = PeekByte();
                if (next >= 0 && HexValue(next) >= 0) { _position++; value = value * 16 + HexValue(next); }
                if (value < 0) return;                     // 'zz: not a byte (LibreOffice treats junk alike); dropped
                if (_skipCount > 0) { _skipCount--; return; }
                // Hebrew / Arabic text is written as long runs of 'hh: decode the whole run in one go
                Span<byte> run = stackalloc byte[256];
                int count = 0;
                run[count++] = (byte)value;
                while (count < run.Length && _position + 4 <= _end && _buffer[_position] == 0x5C && _buffer[_position + 1] == 0x27
                       && HexValue(_buffer[_position + 2]) is var h and >= 0 && HexValue(_buffer[_position + 3]) is var l and >= 0)
                {
                    run[count++] = (byte)(h * 16 + l);
                    _position += 4;
                }
                AddTextBytes(run.Slice(0, count));
                return;
            }
            FlushText();
            if (_skipCount > 0)
            {
                _skipCount--;
                if (first == '\r' && PeekByte() == '\n') _position++;
                return;
            }
            ControlSymbol(first);
            return;
        }

        int length = 0;
        int c = first;
        while (c >= 0 && IsLetter(c))
        {
            if (length < MaxWordLength) _word[length] = (byte)c;
            length++;
            c = ReadByte();
        }
        if (length > MaxWordLength) length = MaxWordLength;
        bool isNegative = false, hasParameter = false;
        long parameter = 0;
        if (c == '-') { isNegative = true; c = ReadByte(); }
        while (c >= '0' && c <= '9')
        {
            hasParameter = true;
            if (parameter < int.MaxValue) parameter = parameter * 10 + (c - '0');
            c = ReadByte();
        }
        if (isNegative) parameter = -parameter;
        if (c >= 0 && c != ' ') _position--;          // the delimiter belongs to what follows (a space is consumed)
        ReadOnlySpan<byte> name = _word.AsSpan(0, length);
        FlushText();
        RtfWord word = RtfControlWords.Lookup(name);
        if (_skipCount > 0)
        {
            _skipCount--;
            if (word == RtfWord.Binary && parameter > 0) SkipBytes(parameter);
            return;
        }
        Dispatch(word, name, hasParameter, (int)Math.Max(int.MinValue, Math.Min(int.MaxValue, parameter)));
    }

    private static bool IsLetter(int c) => (uint)((c | 0x20) - 'a') <= 'z' - 'a';

    private static int HexValue(int c)
        => c >= '0' && c <= '9' ? c - '0' : c >= 'a' && c <= 'f' ? c - 'a' + 10 : c >= 'A' && c <= 'F' ? c - 'A' + 10 : -1;

    private void ControlSymbol(int symbol)
    {
        switch (symbol)
        {
            case '*': _isIgnorableNext = true; break;
            case '\\': case '{': case '}':
            {
                Span<byte> one = stackalloc byte[1];
                one[0] = (byte)symbol;
                AddTextBytes(one);
                break;
            }
            case '~': AddCharacter('\u00A0'); break;            // nonbreaking space
            case '-': AddCharacter('\u00AD'); break;            // optional hyphen
            case '_': AddCharacter('\u2011'); break;            // nonbreaking hyphen
            case '\r':
                if (PeekByte() == '\n') _position++;
                Dispatch(RtfWord.Paragraph, default, false, 0);   // \<newline> is \par (LibreOffice table)
                break;
            case '\n': Dispatch(RtfWord.Paragraph, default, false, 0); break;
            default: break;                                     // \: index subentry, \| formula
        }
    }

    // ── text ──────────────────────────────────────────────────────────────────────────

    private bool IsTextDestination(in RtfGroupState state)
        => state.Destination is RtfDestination.Body or RtfDestination.Footnote or RtfDestination.ListText or RtfDestination.TextBox && !state.Has(RtfStateFlags.DropText);

    private void AddTextBytes(ReadOnlySpan<byte> bytes)
    {
        if (_depth < 0) return;
        if (_skipCount > 0)
        {
            int skipped = Math.Min(_skipCount, bytes.Length);
            _skipCount -= skipped;
            bytes = bytes.Slice(skipped);
            if (bytes.IsEmpty) return;
        }
        ref RtfGroupState state = ref _stack[_depth];
        if (IsTextDestination(state))
        {
            if (_textByteCount > 0 && _textCodePage != state.CodePage) FlushText();
            _textCodePage = state.CodePage;
            if (_textByteCount + bytes.Length > _textBytes.Length) Array.Resize(ref _textBytes, Math.Max(_textBytes.Length * 2, _textByteCount + bytes.Length));
            bytes.CopyTo(_textBytes.AsSpan(_textByteCount));
            _textByteCount += bytes.Length;
            return;
        }
        if (state.Destination is RtfDestination.ShapePropertyName or RtfDestination.ShapePropertyValue)
        {
            if (_shapePropertyText.Length > 65536) return;          // a long value (vertices) is not one we read
            if (_chars.Length < bytes.Length) _chars = new char[Math.Max(bytes.Length, _chars.Length * 2)];
            int decoded = RtfCodePages.Decode(bytes, state.CodePage, _chars);
            for (int i = 0; i < decoded; i++) if (_chars[i] is not ('\r' or '\n')) _shapePropertyText.Append(_chars[i]);
            return;
        }
        if (state.Destination >= RtfDestination.Shape) return;
        if (state.Destination is RtfDestination.FieldInstruction or RtfDestination.BookmarkStart or RtfDestination.BookmarkEnd or RtfDestination.FormFieldText)
        {
            if (_chars.Length < bytes.Length) _chars = new char[Math.Max(bytes.Length, _chars.Length * 2)];
            int decoded = RtfCodePages.Decode(bytes, state.CodePage, _chars);
            (state.Destination == RtfDestination.FormFieldText ? _formText : _captured).Append(_chars, 0, decoded);
            return;
        }
        if (state.Destination == RtfDestination.FormField) return;
        if (_readsHeaderTables) HeaderText(bytes, state.CodePage);
    }

    private void AddCharacter(char character)
    {
        if (_depth < 0) return;
        FlushText();
        ref RtfGroupState state = ref _stack[_depth];
        if (IsTextDestination(state))
        {
            if (_sink.WantsText) { Span<char> one = stackalloc char[1]; one[0] = character; TextByteCount = 1; _sink.Text(one, state); }
            else _sink.TextLength(1, state);
        }
        else if (state.Destination is RtfDestination.FieldInstruction or RtfDestination.BookmarkStart or RtfDestination.BookmarkEnd) _captured.Append(character);
        else if (state.Destination == RtfDestination.FormFieldText) _formText.Append(character);
        else if (state.Destination is RtfDestination.ShapePropertyName or RtfDestination.ShapePropertyValue) _shapePropertyText.Append(character);
        else if (_readsHeaderTables && _headerDepth >= 0) HeaderCharacter(character);
    }

    private void FlushText()
    {
        if (_textByteCount == 0) return;
        int count = _textByteCount;
        _textByteCount = 0;
        ref RtfGroupState state = ref _stack[_depth < 0 ? 0 : _depth];
        if (!_sink.WantsText) { _sink.TextLength(count, state); return; }
        if (_chars.Length < count) _chars = new char[Math.Max(count, _chars.Length * 2)];
        int decoded = RtfCodePages.Decode(_textBytes.AsSpan(0, count), _textCodePage, _chars);
        TextByteCount = count;
        _sink.Text(_chars.AsSpan(0, decoded), state);
    }

    // ── control words ────────────────────────────────────────────────────────────────

    private void Dispatch(RtfWord word, ReadOnlySpan<byte> name, bool hasParameter, int parameter)
    {
        if (_depth < 0)
        {
            if (word == RtfWord.Rtf) return;
            return;
        }
        if (_isIgnorableNext)
        {
            _isIgnorableNext = false;
            if (word is RtfWord.Unknown or RtfWord.SkippedDestination) { SkipGroup(name); return; }
        }
        ref RtfGroupState state = ref _stack[_depth];
        bool isOn = !hasParameter || parameter != 0;          // toggles: \b on, \b0 off (RTF 1.5 spec)
        bool isComplexSlot = state.RunType is RtfRunType.RightToLeftFirst or RtfRunType.LeftToRightThenRightToLeft;
        switch (word)
        {
            case RtfWord.Unknown:
                if (_shapes.Count > 0 && _textBoxDepth < 0)
                {
                    // \shpbxpage: \shpleft counts from the page edge (LibreOffice rtfdispatchflag.cxx:1229-1232); other shape keywords (\shpz, \shplid, ...) not drawn
                    if (name.SequenceEqual("shpbxpage"u8)) _shapes[_shapes.Count - 1].IsPageRelative = true;
                    return;
                }
                if (_sink.WantsUnhandledWords) _sink.UnhandledWord(name);
                return;
            case RtfWord.SkippedDestination: SkipGroup(name); return;

            // ── header ──
            case RtfWord.Rtf: return;
            case RtfWord.Ansi: if (_readsHeaderTables && _tables.AnsiCodePage == 1252) state.CodePage = 1252; return;
            case RtfWord.Mac: SetDocumentCodePage(10000, ref state); return;
            case RtfWord.Pc: SetDocumentCodePage(437, ref state); return;
            case RtfWord.Pca: SetDocumentCodePage(850, ref state); return;
            case RtfWord.AnsiCodePage: if (hasParameter && parameter > 0) SetDocumentCodePage(parameter, ref state); return;
            case RtfWord.DefaultFont:
                if (!hasParameter) return;
                if (_readsHeaderTables) _tables.DefaultFont = parameter;
                state.Font = state.ComplexFont = (short)parameter;
                return;
            case RtfWord.FontTable: case RtfWord.ColorTable: case RtfWord.StyleSheet: case RtfWord.ListTable: case RtfWord.ListOverrideTable:
                if (!_readsHeaderTables) { SkipGroup(name); return; }
                BeginHeaderTable(word, ref state);
                return;
            case RtfWord.FontCharset:
                if (_headerDepth >= 0 && _stack[_headerDepth].Destination == RtfDestination.FontTable) _fontCharset = parameter;
                return;
            case RtfWord.FontCodePage:
                if (_headerDepth >= 0 && _stack[_headerDepth].Destination == RtfDestination.FontTable) _fontCodePage = parameter;
                return;
            case RtfWord.Red: _red = parameter; return;
            case RtfWord.Green: _green = parameter; return;
            case RtfWord.Blue: _blue = parameter; return;
            case RtfWord.Style: state.Style = (short)parameter; return;
            case RtfWord.CharacterStyle: case RtfWord.SectionStyle: case RtfWord.TableStyle:
                if (_entryDepth >= 0 && _depth >= _entryDepth) _isParagraphStyleEntry = false;
                return;
            case RtfWord.OutlineLevel: state.OutlineLevel = (sbyte)Math.Max(-1, Math.Min(9, parameter)); return;
            case RtfWord.List:
                if (_readsHeaderTables && _headerDepth >= 0 && _stack[_headerDepth].Destination == RtfDestination.ListTable)
                {
                    state.Destination = RtfDestination.List;
                    _list = new RtfDocumentTables.ListDefinition();
                    _listLevels = new List<RtfDocumentTables.ListLevelDefinition>();
                    _listDepth = _depth;
                }
                return;
            case RtfWord.ListId:
                if (_list is not null && _listDepth >= 0 && _overrideDepth < 0) _list.ListId = parameter;
                else if (_override is not null) _override.ListId = parameter;
                return;
            case RtfWord.ListLevel:
                if (_overrideLevelDepth >= 0) { state.Destination = RtfDestination.ListOverrideLevel; return; }
                if (_listLevels is not null && _listDepth >= 0)
                {
                    state.Destination = RtfDestination.ListLevel;
                    _level = new RtfDocumentTables.ListLevelDefinition();
                    _listLevels.Add(_level);
                    _levelDepth = _depth;
                }
                return;
            case RtfWord.LevelNumberFormat: case RtfWord.LevelNumberFormatNew:
                if (_level is not null && _levelDepth >= 0) _level.NumberFormat = parameter;
                return;
            case RtfWord.LevelStartAt:
                if (_overrideLevelDepth >= 0) _overrideStartAt = parameter;
                else if (_level is not null && _levelDepth >= 0) _level.StartAt = parameter;
                return;
            case RtfWord.LevelNoRestart: if (_level is not null && _levelDepth >= 0) _level.NoRestart = isOn; return;
            case RtfWord.LevelLegal: if (_level is not null && _levelDepth >= 0) _level.Legal = isOn; return;
            case RtfWord.LevelJustification: if (_level is not null && _levelDepth >= 0) _level.Justification = parameter; return;
            case RtfWord.LevelFollow: if (_level is not null && _levelDepth >= 0) _level.Follow = parameter; return;
            // \levelold / \levelspace / \levelindent: a Word 6 level (RTF 1.9.1 list table; LibreOffice reads none of them - Word 16 writes
            // them for w:legacy, make-legacy-lists-fixture.ps1)
            case RtfWord.LevelOld: if (_level is not null && _levelDepth >= 0) _level.IsLegacy = isOn; return;
            case RtfWord.LevelSpace: if (_level is not null && _levelDepth >= 0) _level.LegacySpace = parameter; return;
            case RtfWord.LevelIndent: if (_level is not null && _levelDepth >= 0) _level.LegacyIndent = parameter; return;
            // \levelpicture N: the level's bullet is picture N of {\*\listpicture} (LibreOffice rtfdispatchvalue.cxx:130-132 LEVELPICTURE ->
            // lvlPicBulletId; the pictures are numbered from 0 in file order: rtfdocumentimpl.cxx:3853-3864 m_nListPictureId++)
            case RtfWord.LevelPicture: if (_level is not null && _levelDepth >= 0 && hasParameter && parameter >= 0) _level.Picture = parameter; return;
            case RtfWord.ListPicture:
                if (_readsHeaderTables && _headerDepth >= 0 && _stack[_headerDepth].Destination == RtfDestination.ListTable) state.Destination = RtfDestination.ListPicture;
                else SkipGroup(name);
                return;
            // {\*\shppict {\pict ...}} inside {\*\listpicture} (rtfdispatchdestination.cxx:131-133): the picture's place in the file is
            // kept; its bytes are read only when the host asks for the picture
            case RtfWord.ShapePicture when state.Destination == RtfDestination.ListPicture: return;
            case RtfWord.Picture when state.Destination == RtfDestination.ListPicture:
            {
                long offset = Position;
                (RtfPicture body, bool pictureClosed) = ReadPictureBody(keepsData: false);
                _tables.ListPictures.Add(new RtfDocumentTables.ListPictureEntry { Offset = offset, Extension = body.Extension, WidthTwips = body.WidthTwips, HeightTwips = body.HeightTwips });
                if (pictureClosed && !Pop()) _position = _end;
                return;
            }
            case RtfWord.LevelText:
                if (_level is not null && _levelDepth >= 0) { state.Destination = RtfDestination.ListLevelText; _levelTextDepth = _depth; _headerText.Clear(); }
                else SkipGroup(name);
                return;
            case RtfWord.LevelNumbers:
                if (_level is not null && _levelDepth >= 0) { state.Destination = RtfDestination.ListLevelNumbers; _levelNumbersDepth = _depth; _headerText.Clear(); }
                else SkipGroup(name);
                return;
            case RtfWord.ListOverride:
                if (_readsHeaderTables && _headerDepth >= 0 && _stack[_headerDepth].Destination == RtfDestination.ListOverrideTable)
                {
                    state.Destination = RtfDestination.ListOverride;
                    _override = new RtfDocumentTables.ListOverride();
                    _overrideDepth = _depth;
                    _overrideLevelIndex = -1;
                }
                return;
            case RtfWord.ListOverrideLevel:
                if (_override is not null && _overrideDepth >= 0)
                {
                    state.Destination = RtfDestination.ListOverrideLevel;
                    _overrideLevelDepth = _depth;
                    _overrideLevelIndex++;
                    _overrideStartAt = int.MinValue;
                }
                else SkipGroup(name);
                return;
            case RtfWord.ListOverrideStartAt: return;
            case RtfWord.ListOverrideIndex:
                if (_override is not null && _overrideDepth >= 0 && _depth >= _overrideDepth) _override.Index = parameter;
                else state.ListOverride = (short)parameter;
                return;
            case RtfWord.ListLevelIndex: state.ListLevel = (byte)Math.Max(0, Math.Min(8, parameter)); return;
            case RtfWord.FootnoteNumberingArabic: if (_readsHeaderTables) _tables.FootnoteFormat = "decimal"; return;
            case RtfWord.FootnoteNumberingLowerAlpha: if (_readsHeaderTables) _tables.FootnoteFormat = "lowerLetter"; return;
            case RtfWord.FootnoteNumberingUpperAlpha: if (_readsHeaderTables) _tables.FootnoteFormat = "upperLetter"; return;
            case RtfWord.FootnoteNumberingLowerRoman: if (_readsHeaderTables) _tables.FootnoteFormat = "lowerRoman"; return;
            case RtfWord.FootnoteNumberingUpperRoman: if (_readsHeaderTables) _tables.FootnoteFormat = "upperRoman"; return;
            case RtfWord.FootnoteNumberingChicago: if (_readsHeaderTables) _tables.FootnoteFormat = "chicago"; return;
            case RtfWord.EndnoteNumberingArabic: if (_readsHeaderTables) _tables.EndnoteFormat = "decimal"; return;
            case RtfWord.EndnoteNumberingLowerAlpha: if (_readsHeaderTables) _tables.EndnoteFormat = "lowerLetter"; return;
            case RtfWord.EndnoteNumberingUpperAlpha: if (_readsHeaderTables) _tables.EndnoteFormat = "upperLetter"; return;
            case RtfWord.EndnoteNumberingLowerRoman: if (_readsHeaderTables) _tables.EndnoteFormat = "lowerRoman"; return;
            case RtfWord.EndnoteNumberingUpperRoman: if (_readsHeaderTables) _tables.EndnoteFormat = "upperRoman"; return;
            case RtfWord.EndnoteNumberingChicago: if (_readsHeaderTables) _tables.EndnoteFormat = "chicago"; return;
            case RtfWord.FootnoteStart: if (_readsHeaderTables && hasParameter) _tables.FootnoteStart = parameter; return;
            case RtfWord.EndnoteStart: if (_readsHeaderTables && hasParameter) _tables.EndnoteStart = parameter; return;
            case RtfWord.RightToLeftDocument: if (_readsHeaderTables) _tables.IsRightToLeftDocument = true; return;
            // \hyphauto: LibreOffice rtfdocumentimpl.cxx:2359 (autoHyphenation = !bParam || nParam != 0, line 2163); \deflang: rtfdispatchvalue.cxx:909
            case RtfWord.AutoHyphenation: if (_readsHeaderTables) _tables.IsAutoHyphenation = !hasParameter || parameter != 0; return;
            case RtfWord.DefaultLanguage: if (_readsHeaderTables && hasParameter) _tables.DefaultLanguage = (ushort)parameter; return;

            // ── rendered destinations ──
            case RtfWord.Footnote:
                if (_footnoteDepth >= 0 || state.Destination != RtfDestination.Body) { SkipGroup(name); return; }
                state.Destination = RtfDestination.Footnote;
                _footnoteDepth = _depth;
                _isEndnote = false;
                _sink.FootnoteStart();
                return;
            case RtfWord.FootnoteIsEndnote: if (_footnoteDepth >= 0) _isEndnote = true; return;
            // Fields (RTF 1.5 spec "Fields": {\field {\*\fldinst ...} {\fldrslt ...}}): the result shows; the instruction is
            // captured for links (HYPERLINK, REF \h). Bookmarks (spec "Bookmarks": {\*\bkmkstart name}; \bkmkcolf / \bkmkcoll
            // are ignored) are captured as jump targets.
            case RtfWord.Field: if (IsTextDestination(state)) _sink.FieldStart(_depth); return;
            case RtfWord.FieldResult: if (IsTextDestination(state)) _sink.FieldResultStart(); return;
            case RtfWord.FieldLock: case RtfWord.FieldPrivate: if (IsTextDestination(state)) _sink.FieldModifier(word); return;
            case RtfWord.FormField:
                if (state.Destination != RtfDestination.FieldInstruction || _formFieldDepth >= 0) { SkipGroup(name); return; }
                state.Destination = RtfDestination.FormField;
                _formFieldDepth = _depth;
                _formType = 0; _formResult = 25; _formDefault = 0; _formEntries.Clear(); _formDefaultText = null;
                return;
            case RtfWord.FormFieldType: if (state.Destination == RtfDestination.FormField) _formType = parameter; return;
            case RtfWord.FormFieldResult: if (state.Destination == RtfDestination.FormField) _formResult = parameter; return;
            case RtfWord.FormFieldDefaultResult: if (state.Destination == RtfDestination.FormField) _formDefault = parameter; return;
            case RtfWord.FormFieldListEntry: case RtfWord.FormFieldDefaultText:
                if (state.Destination != RtfDestination.FormField || _formTextDepth >= 0) { SkipGroup(name); return; }
                state.Destination = RtfDestination.FormFieldText;
                _formTextDepth = _depth;
                _formTextIsEntry = word == RtfWord.FormFieldListEntry;
                _formText.Clear();
                return;
            case RtfWord.FieldInstruction: case RtfWord.BookmarkStart: case RtfWord.BookmarkEnd:
                if (_captureDepth >= 0 || state.Destination is not (RtfDestination.Body or RtfDestination.Footnote or RtfDestination.TextBox)) { SkipGroup(name); return; }
                state.Destination = word == RtfWord.BookmarkStart ? RtfDestination.BookmarkStart : word == RtfWord.BookmarkEnd ? RtfDestination.BookmarkEnd : RtfDestination.FieldInstruction;
                _captureDepth = _depth;
                _captured.Clear();
                return;
            // {\*\pn ...}: Word 6 paragraph numbering, a destination whose own words describe the number and whose text (none
            // normally) stays in the body - LibreOffice rtfdispatchdestination.cxx:505-507 (PN -> PARAGRAPHNUMBERING, a known
            // destination, so \* does not skip it) and rtfdocumentimpl.cxx:1695-1697 (text in it falls through to the body).
            // Word 16's writer leaves {\*\pn group open for number formats Word 6 had none for (MSONFC 36 and up); skipping
            // the group lost every following paragraph. Word itself refuses those files (5151, "unable to read").
            case RtfWord.ParagraphNumbering:
                if (!IsTextDestination(state)) SkipGroup(name);
                return;
            case RtfWord.ListText: case RtfWord.ParagraphNumberText:
                if (_listTextDepth >= 0 || !IsTextDestination(state)) { SkipGroup(name); return; }
                state.Destination = RtfDestination.ListText;
                _listTextDepth = _depth;
                _sink.ListTextStart();
                return;
            case RtfWord.UnicodePair: _unicodePairDepth = _depth; return;
            case RtfWord.UnicodeDestination: state.Set(RtfStateFlags.DropText, false); return;
            // {\*\shppict {\pict ...}}: Word's picture for readers that know shapes; its \pict is read like a bare one.
            // {\nonshppict} (the copy for older readers) stays skipped (LibreOffice rtfdispatchflag.cxx NONSHPPICT: SKIP).
            case RtfWord.ShapePicture when _sink.WantsPictures && IsTextDestination(state): return;
            case RtfWord.Picture when _sink.WantsPictures && IsTextDestination(state): ReadPicture(state); return;
            case RtfWord.Picture when state.Destination == RtfDestination.ShapePropertyValue && _sink.WantsPictures && _shapePropertyName == "pib"
                                      && _shapes.Count > 0 && _shapes[_shapes.Count - 1].Picture is null:
            {
                (RtfPicture body, bool pictureClosed) = ReadPictureBody(_sink.WantsPictureData);
                _shapes[_shapes.Count - 1].Picture = body;
                if (pictureClosed && !Pop()) _position = _end;
                return;
            }
            // the fill blip ({\sp{\sn fillBlip}{\sv {\pict ...}}}): a pattern's small DIB is always decoded (its pixels draw the
            // pattern); a texture / picture is decoded only when its bytes are asked for
            case RtfWord.Picture when state.Destination == RtfDestination.ShapePropertyValue && _sink.WantsPictures && _shapePropertyName == "fillBlip"
                                      && _shapes.Count > 0 && _shapes[_shapes.Count - 1].FillPicture is null:
            {
                (RtfPicture body, bool pictureClosed) = ReadPictureBody(keepsData: true);
                _shapes[_shapes.Count - 1].FillPicture = _sink.WantsPictureData || body.Extension == ".dib" ? body : body with { Data = null };
                if (pictureClosed && !Pop()) _position = _end;
                return;
            }
            case RtfWord.Shape: case RtfWord.ShapeGroup: OpenShape(word, name, ref state); return;
            case RtfWord.ShapeInstruction:
                if (state.Destination == RtfDestination.Shape) state.Destination = RtfDestination.ShapeInstruction;
                else SkipGroup(name);
                return;
            case RtfWord.ShapeResult: SkipGroup(name); return;    // the copy for readers without shapes: the {\shpinst} is read instead
            case RtfWord.ShapeProperty:
                if (state.Destination == RtfDestination.ShapeInstruction) state.Destination = RtfDestination.ShapeProperty;
                else SkipGroup(name);
                return;
            case RtfWord.ShapePropertyName: case RtfWord.ShapePropertyValue:
                if (state.Destination != RtfDestination.ShapeProperty || _shapePropertyDepth >= 0) { SkipGroup(name); return; }
                state.Destination = word == RtfWord.ShapePropertyName ? RtfDestination.ShapePropertyName : RtfDestination.ShapePropertyValue;
                _shapePropertyDepth = _depth;
                _shapePropertyText.Clear();
                if (word == RtfWord.ShapePropertyName) _shapePropertyName = null;
                return;
            case RtfWord.ShapeText:
            {
                if (state.Destination != RtfDestination.ShapeInstruction || _textBoxDepth >= 0 || _shapes.Count == 0) { SkipGroup(name); return; }
                RtfShapePlacement shape = _shapes[_shapes.Count - 1];
                state.Destination = RtfDestination.TextBox;
                state.Set(RtfStateFlags.DropText, false);
                _textBoxDepth = _depth;
                shape.HasText = true;
                SaveRowDefinition();
                _sink.TextBoxStart(shape);
                return;
            }
            case RtfWord.ShapeLeft: case RtfWord.ShapeTop: case RtfWord.ShapeRight: case RtfWord.ShapeBottom: case RtfWord.ShapeWrap: case RtfWord.ShapeWrapSide:
                if (_shapes.Count > 0 && state.Destination is RtfDestination.Shape or RtfDestination.ShapeInstruction)
                {
                    RtfShapePlacement shape = _shapes[_shapes.Count - 1];
                    switch (word)
                    {
                        case RtfWord.ShapeLeft: shape.Left = parameter; break;
                        case RtfWord.ShapeTop: shape.Top = parameter; break;
                        case RtfWord.ShapeRight: shape.Right = parameter; break;
                        case RtfWord.ShapeBottom: shape.Bottom = parameter; break;
                        case RtfWord.ShapeWrap: shape.Wrap = parameter; break;
                        default: shape.WrapSide = parameter; break;
                    }
                }
                return;
            case RtfWord.Picture: case RtfWord.ShapePicture:
                if (IsTextDestination(state)) _sink.Embedded(RtfEmbeddedKind.Picture, state);
                SkipGroup(name);
                return;
            case RtfWord.ObjectGroup:
                if (IsTextDestination(state)) _sink.Embedded(RtfEmbeddedKind.Object, state);
                SkipGroup(name);
                return;
            case RtfWord.MathGroup:
                if (IsTextDestination(state)) _sink.Embedded(RtfEmbeddedKind.Equation, state);
                SkipGroup(name);
                return;
            case RtfWord.NonShapePicture: SkipGroup(name); return;
            case RtfWord.NestedTableProperties:
                // {\*\nesttableprops ... \nestrow}: the nested row definition and its end mark, no text. Its \trowd / \cellx
                // are not applied (they would replace the outer row cells); its \nestrow is.
                state.Set(RtfStateFlags.DropText, true);
                _nestedTablePropertiesDepth = _depth;
                return;

            // ── unicode ──
            case RtfWord.UnicodeSkipCount: state.UnicodeSkipCount = (byte)Math.Max(0, Math.Min(255, parameter)); return;
            case RtfWord.UnicodeCharacter:
            {
                char character = (char)(parameter < 0 ? parameter + 65536 : parameter);
                AddCharacter(character);
                _skipCount = state.UnicodeSkipCount;
                return;
            }
            case RtfWord.Binary: if (parameter > 0) SkipBytes(parameter); return;

            // ── structure ──
            // page setup: only in the main text (not a footnote, a style definition or another destination)
            case RtfWord.PaperWidth: case RtfWord.LeftMargin: case RtfWord.RightMargin: case RtfWord.Gutter:
            case RtfWord.SectionDefault: case RtfWord.SectionPageWidth: case RtfWord.SectionLeftMargin:
            case RtfWord.SectionRightMargin: case RtfWord.SectionGutter: case RtfWord.RightToLeftGutter: case RtfWord.GutterAtTop:
                if (state.Destination == RtfDestination.Body && !IsInFootnote && (hasParameter || word is RtfWord.SectionDefault or RtfWord.RightToLeftGutter or RtfWord.GutterAtTop))
                    _sink.PageSetup(word, hasParameter ? parameter : 1);
                return;
            case RtfWord.Paragraph: case RtfWord.Section: case RtfWord.Cell: case RtfWord.NestedCell: case RtfWord.Row: case RtfWord.NestedRow:
                // inside a header table (read by the index pass, skipped whole by a window): never a block end, or the two
                // passes would segment differently - e.g. a Shift-JIS trail byte '{' in a style name leaves the \par after it
                // inside the style sheet (LibreOffice's tokenizer opens a group there too, rtftokenizer.cxx:2005)
                if (state.Destination is >= RtfDestination.FontTable and <= RtfDestination.ListOverrideLevel) return;
                if (_shapes.Count > 0)
                {
                    // inside a shape: a text box's paragraph / cell / row ends - its own tables; never the block's
                    if (_textBoxDepth < 0 || !IsTextDestination(state)) return;
                    switch (word)
                    {
                        case RtfWord.Cell: _sink.TextBoxCellEnd(state, isNested: false); return;
                        case RtfWord.NestedCell: _sink.TextBoxCellEnd(state, isNested: true); return;
                        case RtfWord.Row: _sink.TextBoxRowEnd(state, isNested: false); return;
                        case RtfWord.NestedRow: _sink.TextBoxRowEnd(state, isNested: true); return;
                        default: _sink.TextBoxParagraphEnd(state); return;
                    }
                }
                switch (word)
                {
                    case RtfWord.Paragraph: _sink.ParagraphEnd(state, isSectionEnd: false); return;
                    case RtfWord.Section: _sink.ParagraphEnd(state, isSectionEnd: true); return;
                    case RtfWord.Cell: _sink.CellEnd(state, isNested: false); return;
                    case RtfWord.NestedCell: _sink.CellEnd(state, isNested: true); return;
                    case RtfWord.Row: _sink.RowEnd(state, isNested: false); return;
                    default: _sink.RowEnd(state, isNested: true); return;
                }
            case RtfWord.LineBreakClear:
                state.Flags = (state.Flags & ~RtfStateFlags.LineBreakClearMask)
                              | (parameter is >= 1 and <= 3 ? (RtfStateFlags)((uint)parameter << 28) : 0);
                return;
            case RtfWord.Page: case RtfWord.Line: case RtfWord.Column:
                if (IsTextDestination(state)) _sink.Break(word, state);
                return;
            case RtfWord.FootnoteMark: if (IsTextDestination(state)) _sink.FootnoteMark(state); return;
            case RtfWord.Tab: AddCharacter('\t'); return;
            // \emdash ... \rdblquote: LibreOffice rtfdispatchsymbol.cxx inserts cp1252 151, 150, 149, 145-148
            case RtfWord.EmDash: AddCharacter('\u2014'); return;
            case RtfWord.EnDash: AddCharacter('\u2013'); return;
            case RtfWord.Bullet: AddCharacter('\u2022'); return;
            case RtfWord.LeftQuote: AddCharacter('\u2018'); return;
            case RtfWord.RightQuote: AddCharacter('\u2019'); return;
            case RtfWord.LeftDoubleQuote: AddCharacter('\u201C'); return;
            case RtfWord.RightDoubleQuote: AddCharacter('\u201D'); return;
            case RtfWord.EmSpace: AddCharacter('\u2003'); return;
            case RtfWord.EnSpace: AddCharacter('\u2002'); return;
            case RtfWord.QuarterEmSpace: AddCharacter('\u2005'); return;
            case RtfWord.ZeroWidthJoiner: AddCharacter('\u200D'); return;
            case RtfWord.ZeroWidthNonJoiner: AddCharacter('\u200C'); return;
            case RtfWord.ZeroWidthBreakOpportunity: AddCharacter('\u200B'); return;
            case RtfWord.ZeroWidthNoBreak: AddCharacter('\uFEFF'); return;
            case RtfWord.LeftToRightMark: AddCharacter('\u200E'); return;
            case RtfWord.RightToLeftMark: AddCharacter('\u200F'); return;

            // ── paragraph ──
            case RtfWord.ParagraphDefault: state.ResetParagraph(); return;
            case RtfWord.AlignLeft: state.Alignment = RtfAlignment.Left; return;
            case RtfWord.AlignRight: state.Alignment = RtfAlignment.Right; return;
            case RtfWord.AlignCenter: state.Alignment = RtfAlignment.Center; return;
            case RtfWord.AlignJustify: case RtfWord.AlignDistribute: state.Alignment = RtfAlignment.Justify; return;
            case RtfWord.LeftIndent: state.LeftIndent = parameter; return;
            case RtfWord.RightIndent: state.RightIndent = parameter; return;
            case RtfWord.StartIndent: state.StartIndent = parameter; state.Set(RtfStateFlags.HasStartIndent, true); return;
            case RtfWord.EndIndent: state.EndIndent = parameter; state.Set(RtfStateFlags.HasEndIndent, true); return;
            case RtfWord.FirstLineIndent: state.FirstLineIndent = parameter; return;
            case RtfWord.SpaceBefore: state.SpaceBefore = parameter; return;
            case RtfWord.SpaceAfter: state.SpaceAfter = parameter; return;
            case RtfWord.LineSpacing: state.LineSpacing = parameter; return;
            case RtfWord.LineSpacingMultiple: state.Set(RtfStateFlags.LineSpacingMultiple, parameter > 0); return;
            case RtfWord.RightToLeftParagraph: state.Set(RtfStateFlags.RightToLeftParagraph, true); state.RunType = RtfRunType.None; return;
            case RtfWord.LeftToRightParagraph: state.Set(RtfStateFlags.RightToLeftParagraph, false); state.RunType = RtfRunType.None; return;
            case RtfWord.InTable: state.Set(RtfStateFlags.InTable, true); return;
            // \dropcapliN / \dropcaptN (RTF 1.9.1: lines dropped; 1 in text, 2 in margin). LibreOffice reads neither (rtfdocumentimpl.cxx:4236
            // lists LN_CT_FramePr_dropCap in RTFFrame::getSprms but no keyword sets it), so Word is the reference.
            case RtfWord.DropCapLines: state.DropCapLines = (byte)Math.Max(0, Math.Min(hasParameter ? parameter : 0, 255)); return;
            case RtfWord.DropCapType: state.Set(RtfStateFlags.DropCapInMargin, hasParameter && parameter == 2); return;
            case RtfWord.FrameHorizontalDistance: state.FrameHorizontalDistance = (short)Math.Max(0, Math.Min(hasParameter ? parameter : 0, short.MaxValue)); return;   // LibreOffice rtfdispatchvalue.cxx:1515-1525 -> FramePr hSpace
            case RtfWord.ParagraphHyphenation: state.Set(RtfStateFlags.NoAutoHyphenation, hasParameter && parameter == 0); return;
            case RtfWord.Language: if (hasParameter) state.Language = (ushort)parameter; return;   // LibreOffice rtfdispatchvalue.cxx:223
            case RtfWord.NestingLevel: state.NestingLevel = (byte)Math.Max(0, Math.Min(64, hasParameter ? parameter : 1)); return;   // default 1 (LibreOffice table)
            case RtfWord.PageBreakBefore: state.Set(RtfStateFlags.PageBreakBefore, isOn); return;

            // ── table row definition ──
            // A nested row definition ({\*\nesttableprops}) is read into its own lists: LibreOffice rtfdispatchdestination.cxx:192-200
            // ("do not set any properties of outer table at nested table") keeps it apart from the outer row the same way.
            case RtfWord.RowDefault:
                if (_nestedTablePropertiesDepth >= 0)
                {
                    _nestedCells.Clear(); _pendingNestedCell = default; Array.Clear(_nestedRowBorders, 0, 6); _isNestedRowRightToLeft = false;
                    return;
                }
                _cells.Clear(); _pendingCell = default; Array.Clear(_rowBorders, 0, 6); _isRowRightToLeft = false;
                _rowLeft = 0; _rowAlignment = 0;
                return;
            case RtfWord.CellBoundary:
            {
                List<RtfCellDefinition> cells = _nestedTablePropertiesDepth >= 0 ? _nestedCells : _cells;
                ref RtfCellDefinition pending = ref PendingCell;
                pending.Boundary = parameter;
                if (cells.Count < 64) cells.Add(pending);
                pending = default;
                return;
            }
            case RtfWord.CellShading: PendingCell.Shading = ColorIndex(parameter); return;
            case RtfWord.CellShadingPercent: PendingCell.ShadingPercent = (short)Math.Max(0, Math.Min(10000, parameter)); return;
            case RtfWord.CellMergeFirst: PendingCell.Merge |= RtfCellMerge.First; return;
            case RtfWord.CellMergeContinue: PendingCell.Merge |= RtfCellMerge.Continue; return;
            case RtfWord.CellVerticalMergeFirst: PendingCell.Merge |= RtfCellMerge.VerticalFirst; return;
            case RtfWord.CellVerticalMergeContinue: PendingCell.Merge |= RtfCellMerge.VerticalContinue; return;
            case RtfWord.CellAlignTop: PendingCell.VerticalAlignment = 0; return;
            case RtfWord.CellAlignCenter: PendingCell.VerticalAlignment = 1; return;
            case RtfWord.CellAlignBottom: PendingCell.VerticalAlignment = 2; return;
            // \clbrdr* / \trbrdr*: the \brdr* words after them describe that border (LibreOffice rtfdocumentimpl.cxx:196 RTFBorderState::CELL);
            // named without a line style it is no line (\brdrtbl, Word's "no border" in a cell, is read as \brdrnone)
            case RtfWord.CellBorderTop: PendingCell.Top = default; state.BorderTarget = RtfBorderTarget.CellTop; return;
            case RtfWord.CellBorderLeft: PendingCell.Left = default; state.BorderTarget = RtfBorderTarget.CellLeft; return;
            case RtfWord.CellBorderBottom: PendingCell.Bottom = default; state.BorderTarget = RtfBorderTarget.CellBottom; return;
            case RtfWord.CellBorderRight: PendingCell.Right = default; state.BorderTarget = RtfBorderTarget.CellRight; return;
            case RtfWord.RowBorderTop: RowBorder(0) = default; state.BorderTarget = RtfBorderTarget.RowTop; return;
            case RtfWord.RowBorderLeft: RowBorder(1) = default; state.BorderTarget = RtfBorderTarget.RowLeft; return;
            case RtfWord.RowBorderBottom: RowBorder(2) = default; state.BorderTarget = RtfBorderTarget.RowBottom; return;
            case RtfWord.RowBorderRight: RowBorder(3) = default; state.BorderTarget = RtfBorderTarget.RowRight; return;
            case RtfWord.RowBorderHorizontal: RowBorder(4) = default; state.BorderTarget = RtfBorderTarget.RowHorizontal; return;
            case RtfWord.RowBorderVertical: RowBorder(5) = default; state.BorderTarget = RtfBorderTarget.RowVertical; return;
            case RtfWord.RightToLeftRow:
                if (_nestedTablePropertiesDepth >= 0) _isNestedRowRightToLeft = true; else _isRowRightToLeft = true;
                state.RunType = RtfRunType.None;
                return;
            case RtfWord.LeftToRightRow: if (_nestedTablePropertiesDepth >= 0) _isNestedRowRightToLeft = false; else _isRowRightToLeft = false; return;
            // \trleft = the table indent, \trql / \trqc / \trqr = its alignment (LibreOffice rtfdispatchvalue.cxx TRLEFT,
            // rtfdispatchflag.cxx TRQL / TRQC / TRQR -> tblInd / jc); a nested row definition keeps its own (not applied).
            case RtfWord.RowLeft: if (_nestedTablePropertiesDepth < 0) _rowLeft = (int)Math.Max(int.MinValue, Math.Min(int.MaxValue, parameter)); return;
            case RtfWord.RowAlignLeft: if (_nestedTablePropertiesDepth < 0) _rowAlignment = 0; return;
            case RtfWord.RowAlignCenter: if (_nestedTablePropertiesDepth < 0) _rowAlignment = 1; return;
            case RtfWord.RowAlignRight: if (_nestedTablePropertiesDepth < 0) _rowAlignment = 2; return;

            // ── character ──
            case RtfWord.Plain:
                state.ResetCharacter((short)_tables.DefaultFont, _tables.CodePageOfFont(_tables.DefaultFont));
                return;
            case RtfWord.Bold: case RtfWord.AssociatedBold:
                if (isComplexSlot) { state.Set(RtfStateFlags.ComplexBold, isOn); state.Set(RtfStateFlags.ComplexBoldSet, true); }
                else state.Set(RtfStateFlags.Bold, isOn);
                if (_level is not null && _levelDepth >= 0) { if (isComplexSlot) _level.ComplexBold = (sbyte)(isOn ? 1 : 0); else _level.Bold = (sbyte)(isOn ? 1 : 0); }
                return;
            case RtfWord.Italic: case RtfWord.AssociatedItalic:
                if (isComplexSlot) { state.Set(RtfStateFlags.ComplexItalic, isOn); state.Set(RtfStateFlags.ComplexItalicSet, true); }
                else state.Set(RtfStateFlags.Italic, isOn);
                if (_level is not null && _levelDepth >= 0) { if (isComplexSlot) _level.ComplexItalic = (sbyte)(isOn ? 1 : 0); else _level.Italic = (sbyte)(isOn ? 1 : 0); }
                return;
            case RtfWord.Font: case RtfWord.AssociatedFont:
                if (_headerDepth >= 0 && _stack[_headerDepth].Destination == RtfDestination.FontTable) { StartFontEntry(parameter); return; }
                if (word == RtfWord.Font && _level is not null && _levelDepth >= 0) _level.Font = parameter;
                if (isComplexSlot && _level is not null && _levelDepth >= 0) _level.ComplexFont = parameter;
                if (isComplexSlot) { state.ComplexFont = (short)parameter; state.Set(RtfStateFlags.ComplexFontSet, true); }
                else state.Font = (short)parameter;
                if (word == RtfWord.Font) state.CodePage = _tables.CodePageOfFont(parameter);   // LibreOffice: only \f changes the encoding
                return;
            case RtfWord.FontSize: case RtfWord.AssociatedFontSize:
            {
                short size = (short)(hasParameter ? Math.Max(1, Math.Min(3276, parameter)) : 24);
                if (isComplexSlot) { state.ComplexFontSize = size; state.Set(RtfStateFlags.ComplexSizeSet, true); }
                else state.FontSize = size;
                if (_level is not null && _levelDepth >= 0) { if (isComplexSlot) _level.ComplexFontSize = size; else _level.FontSize = size; }
                return;
            }
            case RtfWord.Underline: state.Underline = isOn ? RtfUnderline.Single : RtfUnderline.None; return;
            case RtfWord.UnderlineDouble: state.Underline = isOn ? RtfUnderline.Double : RtfUnderline.None; return;
            case RtfWord.UnderlineDotted: state.Underline = isOn ? RtfUnderline.Dotted : RtfUnderline.None; return;
            case RtfWord.UnderlineDashed: state.Underline = isOn ? RtfUnderline.Dashed : RtfUnderline.None; return;
            case RtfWord.UnderlineWave: state.Underline = isOn ? RtfUnderline.Wave : RtfUnderline.None; return;
            case RtfWord.UnderlineThick: state.Underline = isOn ? RtfUnderline.Thick : RtfUnderline.None; return;
            case RtfWord.UnderlineWords: state.Underline = isOn ? RtfUnderline.Words : RtfUnderline.None; return;
            case RtfWord.UnderlineNone: state.Underline = RtfUnderline.None; return;
            case RtfWord.Strike: case RtfWord.StrikeDouble: state.Set(RtfStateFlags.Strike, isOn); return;
            case RtfWord.ForegroundColor: state.ForegroundColor = (short)parameter; if (_level is not null && _levelDepth >= 0) _level.Color = parameter; return;
            case RtfWord.BackgroundColor: state.BackgroundColor = (short)parameter; return;
            // \chcbpat = the shading fill (LibreOffice rtfdispatchvalue.cxx:920-925 CHCBPAT -> w:shd fill); \chcfpat / \chshdng /
            // \chbg* (RTF 1.5 spec "Character Borders and Shading") are not read by LibreOffice: Word is the reference, mixed as w:shd.
            case RtfWord.CharacterShading: state.ShadingColor = (short)parameter; return;
            case RtfWord.CharacterPatternColor: state.CharacterPatternColor = ColorIndex(parameter); return;
            case RtfWord.CharacterShadingPercent: state.CharacterShading = (short)Math.Max(0, Math.Min(10000, parameter)); return;
            case RtfWord.CharacterShadingPattern: state.CharacterShading = 3333; return;
            case RtfWord.UnderlineColor: state.UnderlineColor = ColorIndex(parameter); return;      // rtfdispatchvalue.cxx:943-948 ULC -> sprmCCvUl
            // rtfdocumentimpl.cxx:2174-2200 ACCNONE / ACCDOT / ACCCOMMA / ACCCIRCLE / ACCUNDERDOT -> w:em; \accdot0 = none
            case RtfWord.EmphasisNone: state.EmphasisMark = 0; return;
            case RtfWord.EmphasisDot: state.EmphasisMark = (byte)(isOn ? 1 : 0); return;
            case RtfWord.EmphasisComma: state.EmphasisMark = (byte)(isOn ? 2 : 0); return;
            case RtfWord.EmphasisCircle: state.EmphasisMark = (byte)(isOn ? 3 : 0); return;
            case RtfWord.EmphasisUnderDot: state.EmphasisMark = (byte)(isOn ? 4 : 0); return;

            // ── borders (rtfdispatchflag.cxx:643-650 BOX, :762-795 BRDRT / L / B / R, :796-802 CHBRDR; modifiers go to the last
            // border named, rtfdocumentimpl.cxx:169-203 putBorderProperty) and paragraph shading (rtfdispatchvalue.cxx:935-941 CBPAT
            // -> w:shd fill; \cfpat / \shading / \bg* are not read by LibreOffice, Word and the RTF 1.5 spec are the reference) ──
            case RtfWord.CharacterBorder: state.CharacterBorder = new RtfBorder { Style = RtfBorderStyle.Single }; state.BorderTarget = RtfBorderTarget.Character; return;
            case RtfWord.ParagraphBorderTop: state.BorderTop = new RtfBorder { Style = RtfBorderStyle.Single }; state.BorderTarget = RtfBorderTarget.Top; return;
            case RtfWord.ParagraphBorderLeft: state.BorderLeft = new RtfBorder { Style = RtfBorderStyle.Single }; state.BorderTarget = RtfBorderTarget.Left; return;
            case RtfWord.ParagraphBorderBottom: state.BorderBottom = new RtfBorder { Style = RtfBorderStyle.Single }; state.BorderTarget = RtfBorderTarget.Bottom; return;
            case RtfWord.ParagraphBorderRight: state.BorderRight = new RtfBorder { Style = RtfBorderStyle.Single }; state.BorderTarget = RtfBorderTarget.Right; return;
            case RtfWord.ParagraphBorderBox:
                state.BorderTop = state.BorderLeft = state.BorderBottom = state.BorderRight = new RtfBorder { Style = RtfBorderStyle.Single };
                state.BorderTarget = RtfBorderTarget.Box;
                return;
            case RtfWord.OtherBorder: state.BorderTarget = RtfBorderTarget.Other; return;
            case RtfWord.BorderWidth: ModifyBorder(ref state, (ref RtfBorder border, int value) => border.Width = (short)Math.Max(0, Math.Min(short.MaxValue, value)), parameter); return;
            case RtfWord.BorderColor: ModifyBorder(ref state, (ref RtfBorder border, int value) => border.Color = ColorIndex(value), parameter); return;
            case RtfWord.BorderSingle: ModifyBorder(ref state, (ref RtfBorder border, int _) => border.Style = RtfBorderStyle.Single, 0); return;
            case RtfWord.BorderDouble: ModifyBorder(ref state, (ref RtfBorder border, int _) => border.Style = RtfBorderStyle.Double, 0); return;
            case RtfWord.BorderDotted: ModifyBorder(ref state, (ref RtfBorder border, int _) => border.Style = RtfBorderStyle.Dotted, 0); return;
            case RtfWord.BorderDashed: ModifyBorder(ref state, (ref RtfBorder border, int _) => border.Style = RtfBorderStyle.Dashed, 0); return;
            case RtfWord.BorderNone: ModifyBorder(ref state, (ref RtfBorder border, int _) => border.Style = RtfBorderStyle.None, 0); return;
            case RtfWord.ParagraphBackgroundColor: state.ParagraphBackgroundColor = ColorIndex(parameter); return;
            case RtfWord.ParagraphPatternColor: state.ParagraphPatternColor = ColorIndex(parameter); return;
            case RtfWord.ParagraphShadingPercent: state.ParagraphShading = (short)Math.Max(0, Math.Min(10000, parameter)); return;
            case RtfWord.ParagraphShadingPattern: state.ParagraphShading = 3333; return;
            case RtfWord.Highlight: state.Highlight = (byte)Math.Max(0, Math.Min(255, parameter)); return;
            case RtfWord.Superscript: state.Set(RtfStateFlags.Superscript, isOn); if (isOn) state.Set(RtfStateFlags.Subscript, false); return;
            case RtfWord.Subscript: state.Set(RtfStateFlags.Subscript, isOn); if (isOn) state.Set(RtfStateFlags.Superscript, false); return;
            case RtfWord.NoSuperSubscript: state.Set(RtfStateFlags.Superscript | RtfStateFlags.Subscript, false); return;
            case RtfWord.RaisedPosition: state.Position = (short)(hasParameter ? parameter : 6); return;      // default 6 (spec)
            case RtfWord.LoweredPosition: state.Position = (short)-(hasParameter ? parameter : 6); return;
            case RtfWord.Hidden: state.Set(RtfStateFlags.Hidden, isOn); return;
            case RtfWord.AllCaps: state.Set(RtfStateFlags.AllCaps, isOn); return;
            case RtfWord.SmallCaps: state.Set(RtfStateFlags.SmallCaps, isOn); return;
            case RtfWord.Outline: state.Set(RtfStateFlags.Outline, isOn); return;
            case RtfWord.Shadow: state.Set(RtfStateFlags.Shadow, isOn); return;
            case RtfWord.Emboss: state.Set(RtfStateFlags.Emboss, isOn); return;
            case RtfWord.Imprint: state.Set(RtfStateFlags.Imprint, isOn); return;
            // \charscalex N: character scale in percent, 1..600 (LibreOffice rtfdispatchvalue.cxx:195 maps it to w:w, DomainMapper.cxx:2495 limits it)
            case RtfWord.CharacterScale: state.CharacterScale = (short)(parameter is >= 1 and <= 600 ? parameter : 100); return;
            case RtfWord.CharacterSpacing: state.CharacterSpacing = (short)Math.Max(short.MinValue, Math.Min(short.MaxValue, parameter)); return;
            case RtfWord.RightToLeftCharacters:
                state.RunType = state.RunType == RtfRunType.LeftToRightFirst ? RtfRunType.LeftToRightThenRightToLeft : RtfRunType.RightToLeftFirst;
                if (_tables.AnsiCodePage == 1255) state.CodePage = 1255;      // LibreOffice: \rtlch with a 1255 default encoding
                return;
            case RtfWord.LeftToRightCharacters:
                state.RunType = state.RunType == RtfRunType.RightToLeftFirst ? RtfRunType.RightToLeftThenLeftToRight : RtfRunType.LeftToRightFirst;
                return;
        }
    }

    /// <summary>A colour-table index; out-of-range values are no colour (not wrapped onto a real entry).</summary>
    private static short ColorIndex(int value) => value is > 0 and <= short.MaxValue ? (short)value : (short)0;

    private delegate void BorderModifier(ref RtfBorder border, int value);

    /// <summary>The cell the next \cellx closes: of the nested row definition inside {\*\nesttableprops}, else of the row's.</summary>
    private ref RtfCellDefinition PendingCell => ref (_nestedTablePropertiesDepth >= 0 ? ref _pendingNestedCell : ref _pendingCell);

    private ref RtfBorder RowBorder(int side) => ref (_nestedTablePropertiesDepth >= 0 ? _nestedRowBorders : _rowBorders)[side];

    /// <summary>A \brdr* modifier on the border(s) named last: all four after \box (rtfdocumentimpl.cxx:169-203 putBorderProperty).</summary>
    private void ModifyBorder(ref RtfGroupState state, BorderModifier modify, int value)
    {
        switch (state.BorderTarget)
        {
            case RtfBorderTarget.CellTop: modify(ref PendingCell.Top, value); break;
            case RtfBorderTarget.CellLeft: modify(ref PendingCell.Left, value); break;
            case RtfBorderTarget.CellBottom: modify(ref PendingCell.Bottom, value); break;
            case RtfBorderTarget.CellRight: modify(ref PendingCell.Right, value); break;
            case >= RtfBorderTarget.RowTop and <= RtfBorderTarget.RowVertical: modify(ref RowBorder(state.BorderTarget - RtfBorderTarget.RowTop), value); break;
            case RtfBorderTarget.Character: modify(ref state.CharacterBorder, value); break;
            case RtfBorderTarget.Top: modify(ref state.BorderTop, value); break;
            case RtfBorderTarget.Left: modify(ref state.BorderLeft, value); break;
            case RtfBorderTarget.Bottom: modify(ref state.BorderBottom, value); break;
            case RtfBorderTarget.Right: modify(ref state.BorderRight, value); break;
            case RtfBorderTarget.Box:
                modify(ref state.BorderTop, value); modify(ref state.BorderLeft, value);
                modify(ref state.BorderBottom, value); modify(ref state.BorderRight, value);
                break;
        }
    }

    private void SetDocumentCodePage(int codePage, ref RtfGroupState state)
    {
        if (_readsHeaderTables) _tables.AnsiCodePage = codePage;
        state.CodePage = (ushort)codePage;
    }

    /// <summary>Skips a group whose '{' was just read without pushing it.</summary>
    private void SkipGroupBody()
    {
        Push();
        SkipGroup(default);
    }

    /// <summary>
    /// The rest of a {\pict} group, decoded: \picwgoal / \pichgoal (twips) scaled by \picscalex / \picscaley percent -
    /// LibreOffice rtfdocumentimpl.cxx resolvePict: the goal size (else \picw / \pich) times scale / 100 - and the picture
    /// data as hex digits (or \binN raw bytes) at the group's own level; nested groups ({\*\blipuid}, {\*\picprop}) are
    /// stepped over. \pngblip / \jpegblip name a type a browser shows (LibreOffice rtfdispatchflag.cxx PNGBLIP / JPEGBLIP);
    /// any other blip (\emfblip, \wmetafile, ...) gives "". Closes the group. The data is only kept when the sink asks for
    /// it (<see cref="RtfSink.WantsPictureData"/>: the host reading one picture) - a rendered block only needs the size.
    /// </summary>
    private void ReadPicture(in RtfGroupState state)
    {
        (RtfPicture picture, bool isClosed) = ReadPictureBody(_sink.WantsPictureData);
        if (isClosed && !Pop()) _position = _end;
        _sink.Picture(picture, state);
    }

    /// <summary>The picture whose {\pict} body starts at <paramref name="offset"/>, with its bytes (a list picture: <see cref="RtfDocumentTables.ListPictures"/>).</summary>
    public RtfPicture ReadPictureAt(Stream source, long offset)
    {
        _source = source;
        source.Position = offset;
        _bufferOffset = offset;
        _position = _end = 0;
        _limit = source.Length;
        return ReadPictureBody(keepsData: true).picture;
    }

    /// <summary>A {\pict} group's body up to its closing brace (consumed; the group is not popped), and whether it closed.</summary>
    private (RtfPicture picture, bool isClosed) ReadPictureBody(bool keepsData)
    {
        int depth = 1, goalWidth = 0, goalHeight = 0, scaleX = 100, scaleY = 100, high = -1;
        bool isClosed = false;
        string extension = "";
        MemoryStream? data = keepsData ? new MemoryStream() : null;
        while (true)
        {
            int b = ReadByte();
            if (b < 0) break;
            if (b == '{') { depth++; continue; }
            if (b == '}') { if (--depth == 0) { isClosed = true; break; } continue; }
            if (b == '\\')
            {
                int c = ReadByte();
                if (c < 0) break;
                if (!IsLetter(c)) continue;
                int length = 0;
                while (c >= 0 && IsLetter(c)) { if (length < MaxWordLength) _word[length] = (byte)c; length++; c = ReadByte(); }
                bool negative = c == '-';
                if (negative) c = ReadByte();
                long parameter = 0;
                while (c >= '0' && c <= '9') { if (parameter < int.MaxValue) parameter = parameter * 10 + (c - '0'); c = ReadByte(); }
                if (c >= 0 && c != ' ') _position--;
                bool isBinary = length == 3 && _word[0] == 'b' && _word[1] == 'i' && _word[2] == 'n';
                if (depth != 1) { if (isBinary && parameter > 0) SkipBytes(parameter); continue; }
                string word = System.Text.Encoding.ASCII.GetString(_word, 0, Math.Min(length, MaxWordLength));
                int value = (int)Math.Min(int.MaxValue, negative ? -parameter : parameter);
                switch (word)
                {
                    case "picwgoal": goalWidth = value; break;
                    case "pichgoal": goalHeight = value; break;
                    case "picscalex": scaleX = value; break;
                    case "picscaley": scaleY = value; break;
                    case "pngblip": extension = ".png"; break;
                    case "jpegblip": extension = ".jpg"; break;
                    case "dibitmap": extension = ".dib"; break;          // a DIB without its file header (RTF 1.5 spec \dibitmapN)
                    case "bin":
                        if (data is null) { SkipBytes(parameter); break; }
                        for (long i = 0; i < parameter; i++) { int raw = ReadByte(); if (raw < 0) break; data.WriteByte((byte)raw); }
                        break;
                }
                continue;
            }
            if (depth != 1 || data is null) continue;
            int digit = HexValue(b);
            if (digit < 0) continue;                                  // white space between the digits
            if (high < 0) high = digit; else { data.WriteByte((byte)(high << 4 | digit)); high = -1; }
        }
        // \picw / \pich are the picture's own extent (pixels for bitmaps), not the shown size: without a goal size the
        // picture keeps its natural size (0 here).
        int shownWidth = goalWidth > 0 ? (int)((long)goalWidth * scaleX / 100) : 0;
        int shownHeight = goalHeight > 0 ? (int)((long)goalHeight * scaleY / 100) : 0;
        return (new RtfPicture(shownWidth, shownHeight, extension, data?.ToArray()), isClosed);
    }

    /// <summary>
    /// A {\shp} or {\shpgrp} (a drawing object: {\shp {\*\shpinst \shpleft... {\sp {\sn name}{\sv value}}... {\shptxt ...}}
    /// {\shprslt ...}}; "Word 97 RTF for Drawing Objects (Shapes)", techshelps.github.io/MSDN/SPECS/newrtf/
    /// word97rtffordrawingobjectsshapes.htm), read by every sink in the one parse so the index, search and renderer see the
    /// same text: the placement words and {\sp} properties go to an <see cref="RtfShapePlacement"/>, a {\shptxt} (which "must
    /// come after all the other properties for the shape (inside the \shpinst destination)", same page) is text-box text
    /// between <see cref="RtfSink.TextBoxStart"/> / <see cref="RtfSink.TextBoxEnd"/> whose \par never end the block, a pib's
    /// {\pict} is read by the renderer, {\shprslt} is skipped. "A {\shpgrp ...} can be substituted for a {\shp ...} at any
    /// place": a {\shp} / {\shpgrp} inside a group's {\shpinst} is its child. Placement words (LibreOffice rtfdispatchvalue.cxx
    /// SHPLEFT ...): \shpwr 1 = NONE, 2 = PARALLEL, 3 = THROUGH + wrapNone, 4 = PARALLEL + wrapTight, 5 = THROUGH; \shpwrk 0
    /// both sides, 1 left, 2 right, 3 largest; posh 1 left, 2 center, 3 right, 4 inside, 5 outside (rtfsdrimport.cxx applyProperty);
    /// dxWrapDist* in EMU. A shape anywhere else (in a text box, a header table) is a placeholder and skipped.
    /// </summary>
    private void OpenShape(RtfWord word, ReadOnlySpan<byte> name, ref RtfGroupState state)
    {
        bool isChild = state.Destination == RtfDestination.ShapeInstruction && _shapes.Count > 0 && _shapes[_shapes.Count - 1].IsGroup;
        if (!isChild && !(IsTextDestination(state) && _shapes.Count == 0))
        {
            if (IsTextDestination(state)) _sink.Embedded(RtfEmbeddedKind.Picture, state);
            SkipGroup(name);
            return;
        }
        var shape = new RtfShapePlacement { IsGroup = word == RtfWord.ShapeGroup, Depth = _depth };
        if (isChild)
        {
            shape.Parent = _shapes[_shapes.Count - 1];
            shape.Parent.Children.Add(shape);
        }
        else _sink.ShapeStart(state);
        _shapes.Add(shape);
        state.Destination = RtfDestination.Shape;
    }

    private void EmitShapeString(RtfShapePlacement shape, string text)
    {
        shape.HasText = true;
        shape.IsWordArt = true;
        shape.WordArtText = text;
        _textBoxDepth = _depth;
        _sink.TextBoxStart(shape);
        ref RtfGroupState state = ref _stack[_depth];
        if (_sink.WantsText) { TextByteCount = text.Length; _sink.Text(text.AsSpan(), state); }
        else _sink.TextLength(text.Length, state);
        _textBoxDepth = -1;
        _sink.TextBoxEnd();
    }

    // A text box's own tables (\trowd ... \cellx, {\*\nesttableprops} inside {\shptxt}) must not replace the row definitions -
    // outer and nested - of the table the shape sits in, nor be read into a nested definition the shape sits inside.
    private (RtfCellDefinition[] cells, RtfBorder[] borders, RtfCellDefinition pending, bool rightToLeft, int left, byte alignment,
             RtfCellDefinition[] nestedCells, RtfBorder[] nestedBorders, RtfCellDefinition nestedPending, bool nestedRightToLeft, int nestedDepth)? _outerRow;

    private void SaveRowDefinition()
    {
        _outerRow = (_cells.ToArray(), (RtfBorder[])_rowBorders.Clone(), _pendingCell, _isRowRightToLeft, _rowLeft, _rowAlignment,
                     _nestedCells.ToArray(), (RtfBorder[])_nestedRowBorders.Clone(), _pendingNestedCell, _isNestedRowRightToLeft, _nestedTablePropertiesDepth);
        _nestedTablePropertiesDepth = -1;              // the text box's own \trowd starts its own row, not the nested definition's
    }

    private void RestoreRowDefinition()
    {
        if (_outerRow is not { } row) return;
        _cells.Clear(); _cells.AddRange(row.cells);
        row.borders.CopyTo(_rowBorders, 0);
        (_pendingCell, _isRowRightToLeft, _rowLeft, _rowAlignment) = (row.pending, row.rightToLeft, row.left, row.alignment);
        _nestedCells.Clear(); _nestedCells.AddRange(row.nestedCells);
        row.nestedBorders.CopyTo(_nestedRowBorders, 0);
        (_pendingNestedCell, _isNestedRowRightToLeft, _nestedTablePropertiesDepth) = (row.nestedPending, row.nestedRightToLeft, row.nestedDepth);
        _outerRow = null;
    }

    /// <summary>A group inside a shape closes: a property name / value, a text box, or a shape itself.</summary>
    private void CloseShapeGroup()
    {
        if (_depth == _shapePropertyDepth)
        {
            _shapePropertyDepth = -1;
            string text = _shapePropertyText.ToString().Trim();
            _shapePropertyText.Clear();
            if (_stack[_depth].Destination == RtfDestination.ShapePropertyName) _shapePropertyName = text;
            else if (_shapePropertyName is { } property)
            {
                RtfShapePlacement shape = _shapes[_shapes.Count - 1];
                shape.SetProperty(property, text);
                // gtextUNICODE: the shape's text as one string (LibreOffice rtfsdrimport.cxx:442-443, 949-957 sets it as the shape's
                // text; LibreOffice's own RTF export writes a custom shape's text this way) - given to the sinks as a one-paragraph
                // text box, so it renders and is searched like {\shptxt} text
                if (property == "gtextUNICODE" && !shape.HasText && text.Length > 0) EmitShapeString(shape, text);
            }
        }
        if (_depth == _textBoxDepth) { _textBoxDepth = -1; _sink.TextBoxEnd(); RestoreRowDefinition(); }
        RtfShapePlacement last = _shapes[_shapes.Count - 1];
        if (_depth != last.Depth) return;
        _shapes.RemoveAt(_shapes.Count - 1);
        _shapePropertyName = null;
        if (_shapes.Count == 0) _sink.ShapeEnd(last, _stack[_depth - 1 < 0 ? 0 : _depth - 1]);
    }

    /// <summary>Reads on until the group depth drops to <paramref name="target"/> (the closing brace consumed); returns the new depth.</summary>
    private int SkipToDepth(int target)
    {
        int depth = target + 1;
        while (depth > target)
        {
            int b = ReadByte();
            if (b < 0) return depth;
            if (b == '{') depth++;
            else if (b == '}') depth--;
            else if (b == '\\')
            {
                int c = ReadByte();
                if (c != 'b') continue;
                // \binN: N raw bytes that may hold braces
                if (ReadByte() != 'i') { _position--; continue; }
                if (ReadByte() != 'n') { _position--; continue; }
                long parameter = 0;
                int d = ReadByte();
                while (d >= '0' && d <= '9') { parameter = parameter * 10 + (d - '0'); d = ReadByte(); }
                if (d >= 0 && d != ' ') _position--;
                SkipBytes(parameter);
            }
        }
        return depth;
    }


    /// <summary>Skips to the end of the current group (braces counted, escapes and \binN data stepped over) and closes it.</summary>
    private void SkipGroup(ReadOnlySpan<byte> name)
    {
        _sink.DestinationSkipped(name);
        int depth = 1;
        bool[] stops = IsSkipStop;
        while (true)
        {
            if (_position >= _end && !Fill()) return;
            byte[] buffer = _buffer;
            int at = _position, end = _end;
#if NET
            int found = buffer.AsSpan(at, end - at).IndexOfAny(SkipStops);
            at = found < 0 ? end : at + found;
#else
            while (at < end && !stops[buffer[at]]) at++;
#endif
            _position = at;
            if (at >= end) continue;
            byte b = buffer[at];
            _position++;
            if (b == '{') { depth++; continue; }
            if (b == '}')
            {
                if (--depth == 0) { if (!Pop()) _position = _end; return; }
                continue;
            }
            // backslash: a control symbol is one byte; a word may be \binN
            int c = ReadByte();
            if (c < 0) return;
            if (!IsLetter(c)) continue;
            int length = 0;
            while (c >= 0 && IsLetter(c)) { if (length < MaxWordLength) _word[length] = (byte)c; length++; c = ReadByte(); }
            bool isBinary = length == 3 && _word[0] == 'b' && _word[1] == 'i' && _word[2] == 'n';
            long parameter = 0;
            if (c == '-') c = ReadByte();
            while (c >= '0' && c <= '9') { if (parameter < int.MaxValue) parameter = parameter * 10 + (c - '0'); c = ReadByte(); }
            if (c >= 0 && c != ' ') _position--;
            if (isBinary && parameter > 0) SkipBytes(parameter);
        }
    }

    // ── header tables (index mode) ────────────────────────────────────────────────────

    private void BeginHeaderTable(RtfWord word, ref RtfGroupState state)
    {
        state.Destination = word switch
        {
            RtfWord.FontTable => RtfDestination.FontTable,
            RtfWord.ColorTable => RtfDestination.ColorTable,
            RtfWord.StyleSheet => RtfDestination.StyleSheet,
            RtfWord.ListTable => RtfDestination.ListTable,
            _ => RtfDestination.ListOverrideTable,
        };
        _headerDepth = _depth;
        _headerText.Clear();
        _red = _green = _blue = -1;
        _fontNumber = int.MinValue;
        if (word == RtfWord.ColorTable) _tables.Colors.Clear();
    }

    private void StartFontEntry(int number)
    {
        CommitFont();
        _fontNumber = number;
        _fontCharset = -1;
        _fontCodePage = -1;
        _headerText.Clear();
    }

    private void CommitFont()
    {
        if (_fontNumber == int.MinValue) return;
        string rawName = _headerText.ToString().Trim();
        (string name, int suffixCodePage) = RtfCodePages.SplitFontNameSuffix(rawName);
        int codePage = _fontCodePage > 0 ? _fontCodePage : -1;
        if (codePage < 0 && _fontCharset >= 0)
        {
            int fromCharset = RtfCodePages.CodePageOfCharset(_fontCharset);
            codePage = fromCharset == 0 ? _tables.AnsiCodePage : fromCharset;      // charset 1 = the default (\ansicpg)
        }
        if (suffixCodePage > 0) codePage = suffixCodePage;
        if (!_tables.Fonts.ContainsKey(_fontNumber))
            _tables.Fonts[_fontNumber] = new RtfDocumentTables.FontEntry { Name = name, Charset = _fontCharset, CodePage = codePage };
        _fontNumber = int.MinValue;
        _headerText.Clear();
    }

    private void HeaderText(ReadOnlySpan<byte> bytes, ushort codePage)
    {
        if (_headerDepth < 0 || _depth < _headerDepth) return;
        RtfDestination destination = _stack[_depth].Destination;
        Span<char> decoded = stackalloc char[4];
        Span<byte> one = stackalloc byte[1];
        foreach (byte b in bytes)
        {
            if (b == ';' && destination is RtfDestination.FontTable or RtfDestination.ColorTable or RtfDestination.StyleSheet or RtfDestination.ListLevelText)
            {
                HeaderEntryEnd(destination);
                continue;
            }
            if (destination is RtfDestination.ListLevelNumbers) { _headerText.Append((char)b); continue; }
            if (destination is RtfDestination.FontTable or RtfDestination.StyleSheet or RtfDestination.ListLevelText)
            {
                one[0] = b;
                int count = RtfCodePages.Decode(one, b < 0x80 ? RtfCodePages.Windows1252 : codePage, decoded);
                for (int i = 0; i < count; i++) _headerText.Append(decoded[i]);
            }
        }
    }

    private void HeaderCharacter(char character)
    {
        if (_depth < _headerDepth) return;
        RtfDestination destination = _stack[_depth].Destination;
        if (destination is RtfDestination.FontTable or RtfDestination.StyleSheet or RtfDestination.ListLevelText) _headerText.Append(character);
    }

    private void HeaderEntryEnd(RtfDestination destination)
    {
        switch (destination)
        {
            case RtfDestination.FontTable:
                CommitFont();
                break;
            case RtfDestination.ColorTable:
                _tables.Colors.Add(_red < 0 && _green < 0 && _blue < 0 ? null
                    : $"{Math.Max(0, _red) & 0xFF:X2}{Math.Max(0, _green) & 0xFF:X2}{Math.Max(0, _blue) & 0xFF:X2}");
                _red = _green = _blue = -1;
                break;
            case RtfDestination.StyleSheet:
                if (_entryDepth >= 0 && _isParagraphStyleEntry)
                {
                    ref RtfGroupState state = ref _stack[_depth];
                    if (!_tables.Styles.ContainsKey(state.Style))
                        _tables.Styles[state.Style] = new RtfDocumentTables.StyleEntry { Name = _headerText.ToString().Trim(), OutlineLevel = state.OutlineLevel };
                }
                _headerText.Clear();
                _isParagraphStyleEntry = false;              // one name per entry
                break;
            case RtfDestination.ListLevelText:
                _headerText.Append(';');                     // kept: ';' may be literal text inside the counted length
                break;
        }
    }

    private void CloseHeaderGroup()
    {
        if (_depth == _levelTextDepth)
        {
            if (_level is not null)
            {
                string raw = _headerText.ToString();
                if (raw.Length > 0)
                {
                    int length = raw[0];
                    _level.Text = length < raw.Length ? raw.Substring(1, length) : raw.Substring(1).TrimEnd(';');
                }
            }
            _levelTextDepth = -1;
            _headerText.Clear();
        }
        else if (_depth == _levelNumbersDepth)
        {
            if (_level is not null)
            {
                string raw = _headerText.ToString();
                var positions = new List<byte>();
                foreach (char c in raw) if (c != ';' && c > 0 && c < 256) positions.Add((byte)c);
                _level.NumberPositions = positions.ToArray();
            }
            _levelNumbersDepth = -1;
            _headerText.Clear();
        }
        else if (_depth == _levelDepth) { _levelDepth = -1; _level = null; }
        else if (_depth == _listDepth)
        {
            if (_list is not null && _listLevels is not null)
            {
                _list.Levels = _listLevels.ToArray();
                _tables.Lists[_list.ListId] = _list;
            }
            _list = null; _listLevels = null; _listDepth = -1;
        }
        else if (_depth == _overrideLevelDepth)
        {
            if (_override is not null && _overrideStartAt != int.MinValue && _overrideLevelIndex is >= 0 and < 9)
                _override.StartAtByLevel[_overrideLevelIndex] = _overrideStartAt;
            _overrideLevelDepth = -1;
        }
        else if (_depth == _overrideDepth)
        {
            if (_override is not null && _override.Index != 0) _tables.ListOverrides[_override.Index] = _override;
            _override = null; _overrideDepth = -1;
        }
        else if (_depth == _entryDepth)
        {
            if (_headerDepth >= 0 && _stack[_headerDepth].Destination == RtfDestination.FontTable) CommitFont();
            _entryDepth = -1;
        }
        else if (_depth == _headerDepth)
        {
            if (_stack[_headerDepth].Destination == RtfDestination.FontTable) CommitFont();
            _headerDepth = -1;
        }
    }
}
