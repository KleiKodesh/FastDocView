using System.Runtime.InteropServices;
using System.Text;

namespace WordDocumentRandomAccess.Rtf;

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
    private bool _isEndnote;

    // table row definition (\trowd ... \cellxN), document-level as in LibreOffice
    private readonly List<int> _cellBoundaries = new();
    private readonly List<short> _cellShading = new();
    private short _pendingCellShading;
    private bool _isRowRightToLeft;

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
    public bool IsRowRightToLeft => _isRowRightToLeft;
    public IReadOnlyList<int> CellBoundaries => _cellBoundaries;
    public IReadOnlyList<short> CellShading => _cellShading;

    /// <summary>Reads <paramref name="source"/> from <paramref name="start"/> up to <paramref name="end"/> (exclusive).</summary>
    public void Run(Stream source, long start, long end)
    {
        _source = source;
        source.Position = start;
        _bufferOffset = start;
        _position = _end = 0;
        _limit = end;
        Parse();
        FlushText();
    }

    // ── snapshots ─────────────────────────────────────────────────────────────────────

    /// <summary>The group-state stack, row definition and open contexts, as bytes — what a window needs to resume here.</summary>
    public byte[] Snapshot()
    {
        byte[] bytes = new byte[SnapshotSize];
        WriteSnapshot(bytes);
        return bytes;
    }

    public int SnapshotSize => 4 + (_depth + 1) * RtfGroupState.Size + 4 * 4 + 1 + 4 + _cellBoundaries.Count * 6;

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
        WriteInt(bytes, ref at, _cellBoundaries.Count);
        for (int i = 0; i < _cellBoundaries.Count; i++)
        {
            WriteInt(bytes, ref at, _cellBoundaries[i]);
            bytes[at++] = (byte)_cellShading[i]; bytes[at++] = (byte)(_cellShading[i] >> 8);
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
        int cells = ReadInt(bytes, ref at);
        _cellBoundaries.Clear(); _cellShading.Clear();
        for (int i = 0; i < cells; i++)
        {
            _cellBoundaries.Add(ReadInt(bytes, ref at));
            _cellShading.Add((short)(bytes[at] | bytes[at + 1] << 8)); at += 2;
        }
        _isIgnorableNext = false;
        _skipCount = 0;
        _textByteCount = 0;
        _pendingCellShading = 0;
        _isEndnote = false;
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
        => state.Destination is RtfDestination.Body or RtfDestination.Footnote or RtfDestination.ListText && !state.Has(RtfStateFlags.DropText);

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
        if (_readsHeaderTables) HeaderText(bytes, state.CodePage);
    }

    private void AddCharacter(char character)
    {
        if (_depth < 0) return;
        FlushText();
        ref RtfGroupState state = ref _stack[_depth];
        if (IsTextDestination(state))
        {
            if (_sink.WantsText) { Span<char> one = stackalloc char[1]; one[0] = character; _sink.Text(one, state); }
            else _sink.TextLength(1, state);
        }
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

            // ── rendered destinations ──
            case RtfWord.Footnote:
                if (_footnoteDepth >= 0 || state.Destination != RtfDestination.Body) { SkipGroup(name); return; }
                state.Destination = RtfDestination.Footnote;
                _footnoteDepth = _depth;
                _isEndnote = false;
                _sink.FootnoteStart();
                return;
            case RtfWord.FootnoteIsEndnote: if (_footnoteDepth >= 0) _isEndnote = true; return;
            case RtfWord.Field: case RtfWord.FieldResult: return;          // transparent: \fldinst is skipped, the result shows
            case RtfWord.ListText: case RtfWord.ParagraphNumberText:
                if (_listTextDepth >= 0 || !IsTextDestination(state)) { SkipGroup(name); return; }
                state.Destination = RtfDestination.ListText;
                _listTextDepth = _depth;
                _sink.ListTextStart();
                return;
            case RtfWord.UnicodePair: _unicodePairDepth = _depth; return;
            case RtfWord.UnicodeDestination: state.Set(RtfStateFlags.DropText, false); return;
            case RtfWord.Picture: case RtfWord.ShapePicture: case RtfWord.Shape:
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
            case RtfWord.Paragraph: _sink.ParagraphEnd(state, isSectionEnd: false); return;
            case RtfWord.Section: _sink.ParagraphEnd(state, isSectionEnd: true); return;
            case RtfWord.Cell: _sink.CellEnd(state, isNested: false); return;
            case RtfWord.NestedCell: _sink.CellEnd(state, isNested: true); return;
            case RtfWord.Row: _sink.RowEnd(state, isNested: false); return;
            case RtfWord.NestedRow: _sink.RowEnd(state, isNested: true); return;
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
            case RtfWord.NestingLevel: state.NestingLevel = (byte)Math.Max(0, Math.Min(64, hasParameter ? parameter : 1)); return;   // default 1 (LibreOffice table)
            case RtfWord.PageBreakBefore: state.Set(RtfStateFlags.PageBreakBefore, isOn); return;

            // ── table row definition ──
            case RtfWord.RowDefault:
                if (_nestedTablePropertiesDepth >= 0) return;
                _cellBoundaries.Clear(); _cellShading.Clear(); _pendingCellShading = 0; _isRowRightToLeft = false;
                return;
            case RtfWord.CellBoundary:
                if (_nestedTablePropertiesDepth >= 0) return;
                if (_cellBoundaries.Count < 64) { _cellBoundaries.Add(parameter); _cellShading.Add(_pendingCellShading); }
                _pendingCellShading = 0;
                return;
            case RtfWord.CellShading: if (_nestedTablePropertiesDepth >= 0) return; _pendingCellShading = (short)parameter; return;
            case RtfWord.RightToLeftRow: _isRowRightToLeft = true; state.RunType = RtfRunType.None; return;
            case RtfWord.LeftToRightRow: _isRowRightToLeft = false; return;

            // ── character ──
            case RtfWord.Plain:
                state.ResetCharacter((short)_tables.DefaultFont, _tables.CodePageOfFont(_tables.DefaultFont));
                return;
            case RtfWord.Bold: case RtfWord.AssociatedBold:
                if (isComplexSlot) { state.Set(RtfStateFlags.ComplexBold, isOn); state.Set(RtfStateFlags.ComplexBoldSet, true); }
                else state.Set(RtfStateFlags.Bold, isOn);
                return;
            case RtfWord.Italic: case RtfWord.AssociatedItalic:
                if (isComplexSlot) { state.Set(RtfStateFlags.ComplexItalic, isOn); state.Set(RtfStateFlags.ComplexItalicSet, true); }
                else state.Set(RtfStateFlags.Italic, isOn);
                return;
            case RtfWord.Font: case RtfWord.AssociatedFont:
                if (_headerDepth >= 0 && _stack[_headerDepth].Destination == RtfDestination.FontTable) { StartFontEntry(parameter); return; }
                if (isComplexSlot) { state.ComplexFont = (short)parameter; state.Set(RtfStateFlags.ComplexFontSet, true); }
                else state.Font = (short)parameter;
                if (word == RtfWord.Font) state.CodePage = _tables.CodePageOfFont(parameter);   // LibreOffice: only \f changes the encoding
                return;
            case RtfWord.FontSize: case RtfWord.AssociatedFontSize:
            {
                short size = (short)(hasParameter ? Math.Max(1, Math.Min(3276, parameter)) : 24);
                if (isComplexSlot) { state.ComplexFontSize = size; state.Set(RtfStateFlags.ComplexSizeSet, true); }
                else state.FontSize = size;
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
            case RtfWord.ForegroundColor: state.ForegroundColor = (short)parameter; return;
            case RtfWord.BackgroundColor: state.BackgroundColor = (short)parameter; return;
            case RtfWord.CharacterShading: state.ShadingColor = (short)parameter; return;
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
