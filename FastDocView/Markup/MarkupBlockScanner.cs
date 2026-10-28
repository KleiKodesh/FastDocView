using System.Net;
using System.Text;

namespace FastDocView.Markup;

public enum MarkupBlockMode : byte
{
    /// <summary>.txt read as HTML: one block per line; a line break inside a tag or comment does not end the block.</summary>
    Lines = 0,
    /// <summary>HTML body (MHTML): a block starts at each paragraph-level element outside tables (p, h1-6, table, div, lists...).</summary>
    Elements = 1,
}

/// <summary>
/// The index pass over HTML-ish text without decoding it: a byte-level (or UTF-16 code-unit) state
/// machine fed chunk by chunk that records block start offsets, and per block the stack of elements
/// open at its start (as their original start tags, interned), so a window can be parsed on its own.
/// Also records headings (h1-h6 text), text length per block and, in <see cref="MarkupBlockMode.Elements"/>,
/// the &lt;style&gt; text of the head. Only ASCII code units drive the state, so single-byte code pages and
/// UTF-8 are scanned as raw bytes (vectorised search over text runs); UTF-16 is scanned per 2-byte unit.
/// </summary>
public sealed class MarkupBlockScanner
{
    public const int MaximumTagUnits = 2048;              // a '<' that does not close within this many code units (incl. '<' and '>') is text
    public const int MaximumCommentBytes = 64 * 1024;
    public const int LineBlockSplitBytes = 64 * 1024;     // a longer line is split at a space (bounded windows)
    public const int LineBlockHardLimitBytes = 256 * 1024; // ... or anywhere outside a tag once this long (no spaces at all)
    private const int MaximumStackDepth = 32;
    private const int MaximumHeadingBytes = 400;
    private const int MaximumStyleBytes = 512 * 1024;

    private static readonly HashSet<string> VoidElements = new(StringComparer.Ordinal)
        { "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr" };
    internal static readonly HashSet<string> RawTextElements = new(StringComparer.Ordinal) { "script", "style", "title", "textarea", "xmp" };
    private static readonly HashSet<string> BlockStartElements = new(StringComparer.Ordinal)
        { "p", "h1", "h2", "h3", "h4", "h5", "h6", "table", "div", "ul", "ol", "li", "dl", "dt", "dd", "blockquote", "pre", "hr", "center", "address" };
    private static readonly HashSet<string> ClosesOpenParagraph = new(StringComparer.Ordinal)
        { "p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "dl", "table", "pre", "blockquote", "hr", "center", "address" };
    private static readonly HashSet<string> NeverOnStack = new(StringComparer.Ordinal) { "html", "head", "body" };

    private enum State : byte { Text, TagStart, EndTagStart, Tag, Bang, BangDash, Comment, Bogus, RawText }

    private readonly MarkupBlockMode _mode;
    private readonly int _unitSize;
    private readonly bool _bigEndian;
    private readonly Encoding _encoding;
    private readonly bool _isUtf8;

    // outputs
    public List<long> BlockStartOffsets { get; } = [];
    public List<int> BlockTextLengths { get; } = [];
    public List<int> BlockStackIds { get; } = [];
    public List<string> Stacks { get; } = [""];          // id 0 = nothing open
    public List<WordDocumentHeading> Headings { get; } = [];
    public bool IsBodyStarted => _isInBody;
    public bool IsBodyEnded { get; private set; }

    private State _state = State.Text;
    private readonly List<byte> _tagBytes = new(256);
    private long _tagStartOffset;
    private int _lastTagByteCount;
    private int _commentLength, _commentDashes;
    private bool _isQuoted; private int _quoteUnit; private int _lastNonSpaceUnit;
    private string _rawName = "";
    private int _rawMatch; private long _rawEndTagOffset;
    private readonly List<byte> _styleBytes = [];
    private bool _isCapturingStyle;

    private readonly List<(string name, string tag)> _stack = [];
    private int _unpushedDepth;                          // elements past MaximumStackDepth: their end tags are absorbed first
    private bool _isStackDirty;
    private int _stackId;
    private readonly Dictionary<string, int> _stackIds = new(StringComparer.Ordinal) { [""] = 0 };

    private bool _isInBody;
    private int _tableDepth;
    private bool _blockHasContent;
    private long _blockStartOffset;
    private bool _pendingCarriageReturn;
    private readonly bool _isDoubleByte; private bool _isDoubleByteTrailNext;
    private int _carry = -1; private long _carryOffset;

    private int _headingLevel; private int _headingBlock = -1; private readonly List<byte> _headingBytes = [];

    public MarkupBlockScanner(MarkupBlockMode mode, int unitSize, bool bigEndian, Encoding encoding)
    {
        _mode = mode; _unitSize = unitSize; _bigEndian = bigEndian; _encoding = encoding;
        _isUtf8 = encoding.CodePage == 65001;
        _isDoubleByte = encoding.CodePage is 932 or 936 or 949 or 950 or 51949;
        if (mode == MarkupBlockMode.Lines) { _isInBody = true; }
    }

    /// <summary>Starts the first block (Lines mode: at the first text byte after any BOM).</summary>
    public void Begin(long offset)
    {
        if (_mode == MarkupBlockMode.Lines) StartBlock(offset);
    }

    /// <summary>The head's &lt;style&gt; text (Elements mode), bounded.</summary>
    public string StyleText => _styleBytes.Count == 0 ? "" : _encoding.GetString(_styleBytes.ToArray());

    public void Feed(byte[] buffer, int offset, int count, long absoluteOffset)
    {
        if (IsBodyEnded) return;
        if (_unitSize == 1) { FeedBytes(buffer, offset, count, absoluteOffset); return; }
        int i = offset, end = offset + count;
        if (_carry >= 0 && i < end)
        {
            byte first = (byte)_carry; byte second = buffer[i];
            _carry = -1;
            ProcessUnit(UnitOf(first, second), _carryOffset, first, second);
            i++;
        }
        for (; i + 1 < end; i += 2)
        {
            if (IsBodyEnded) return;
            ProcessUnit(UnitOf(buffer[i], buffer[i + 1]), absoluteOffset + (i - offset), buffer[i], buffer[i + 1]);
        }
        if (i < end) { _carry = buffer[i]; _carryOffset = absoluteOffset + (i - offset); }
    }

    private int UnitOf(byte first, byte second) => _bigEndian ? (first << 8) | second : (second << 8) | first;

    private void FeedBytes(byte[] buffer, int offset, int count, long absoluteOffset)
    {
        int i = offset, end = offset + count;
        while (i < end)
        {
            if (IsBodyEnded) return;
            if (_state == State.Text && !_pendingCarriageReturn)
            {
                // vectorised skip over a text run: only '<' and line breaks matter here
                ReadOnlySpan<byte> rest = buffer.AsSpan(i, end - i);
                bool mustWatchSpaces = _mode == MarkupBlockMode.Lines && absoluteOffset + (i - offset) - _blockStartOffset >= LineBlockSplitBytes;
                int found = mustWatchSpaces ? 0                                     // long line: per byte until it splits at a space
                          : _mode == MarkupBlockMode.Lines ? rest.IndexOfAny((byte)'<', (byte)'\n', (byte)'\r')
                          : rest.IndexOf((byte)'<');
                int skipped = found < 0 ? rest.Length : found;
                if (skipped > 0)
                {
                    AddText(rest.Slice(0, skipped));
                    i += skipped;
                    if (found < 0) break;
                }
            }
            byte b = buffer[i];
            ProcessUnit(b, absoluteOffset + (i - offset), b, 0);
            i++;
        }
    }

    private void AddText(ReadOnlySpan<byte> text)
    {
        if (BlockTextLengths.Count > 0) BlockTextLengths[BlockTextLengths.Count - 1] += text.Length / _unitSize;
        if (!_blockHasContent)
            foreach (byte b in text) if (b > 0x20) { _blockHasContent = true; break; }
        if (_headingBlock >= 0 && _headingBytes.Count < MaximumHeadingBytes)
            for (int k = 0; k < text.Length && _headingBytes.Count < MaximumHeadingBytes; k++) _headingBytes.Add(text[k]);
    }

    private void AppendRaw(List<byte> target, byte first, byte second)
    {
        target.Add(first);
        if (_unitSize == 2) target.Add(second);
    }

    private void ProcessUnit(int unit, long offset, byte first, byte second)
    {
        if (_pendingCarriageReturn)
        {
            _pendingCarriageReturn = false;
            if (unit == '\n') { AddTextUnit(unit, first, second); EndLine(offset + _unitSize); return; }
            EndLine(offset);
        }
        switch (_state)
        {
            case State.Text:
                if (unit == '<') { _state = State.TagStart; BeginTag(offset, first, second); return; }
                AddTextUnit(unit, first, second);
                if (_mode == MarkupBlockMode.Lines)
                {
                    if (unit == '\n') EndLine(offset + _unitSize);
                    else if (unit == '\r') _pendingCarriageReturn = true;
                    else if (unit == ' ' && offset - _blockStartOffset >= LineBlockSplitBytes) EndLine(offset + _unitSize);
                    else if (offset + _unitSize - _blockStartOffset >= LineBlockHardLimitBytes && IsCharacterEnd(unit)) EndLine(offset + _unitSize);
                }
                return;

            case State.TagStart:
                if (IsAsciiLetter(unit)) { _state = State.Tag; AppendRaw(_tagBytes, first, second); _isQuoted = false; _lastNonSpaceUnit = unit; return; }
                if (unit == '/') { _state = State.EndTagStart; AppendRaw(_tagBytes, first, second); return; }
                if (unit == '!') { _state = State.Bang; AppendRaw(_tagBytes, first, second); return; }
                if (unit == '?') { _state = State.Bogus; AppendRaw(_tagBytes, first, second); _commentLength = 0; return; }
                AbandonTag(); ProcessUnit(unit, offset, first, second); return;

            case State.EndTagStart:
                if (IsAsciiLetter(unit)) { _state = State.Tag; AppendRaw(_tagBytes, first, second); _isQuoted = false; _lastNonSpaceUnit = unit; return; }
                AbandonTag(); ProcessUnit(unit, offset, first, second); return;

            case State.Bang:
                AppendRaw(_tagBytes, first, second);
                if (unit == '-') { _state = State.BangDash; return; }
                _state = State.Bogus; _commentLength = 0;
                if (unit == '>') CompleteBogus();
                return;

            case State.BangDash:
                if (unit == '-') { _state = State.Comment; _commentLength = 0; _commentDashes = 0; _tagBytes.Clear(); return; }
                AppendRaw(_tagBytes, first, second);
                _state = State.Bogus; _commentLength = 0;
                if (unit == '>') CompleteBogus();
                return;

            case State.Comment:
                if (unit == '>' && _commentDashes >= 2) { _state = State.Text; _blockHasContent = true; return; }
                _commentDashes = unit == '-' ? _commentDashes + 1 : 0;
                if (++_commentLength > MaximumCommentBytes) { _state = State.Text; }   // unterminated: the rest is text again
                return;

            case State.Bogus:
                if (unit == '>') { CompleteBogus(); return; }
                if (++_commentLength > MaximumCommentBytes) _state = State.Text;
                return;

            case State.Tag:
                if (_tagBytes.Count / _unitSize >= MaximumTagUnits) { AbandonTag(); ProcessUnit(unit, offset, first, second); return; }
                AppendRaw(_tagBytes, first, second);
                if (_isQuoted) { if (unit == _quoteUnit) _isQuoted = false; }
                else if (unit == '>') { CompleteTag(); return; }
                else if ((unit == '"' || unit == '\'') && _lastNonSpaceUnit == '=') { _isQuoted = true; _quoteUnit = unit; }
                if (unit > ' ') _lastNonSpaceUnit = unit;
                return;

            case State.RawText:
                ProcessRawTextUnit(unit, offset, first, second);
                return;
        }
    }

    private void AddTextUnit(int unit, byte first, byte second)
    {
        // Double-byte code pages (Shift_JIS, GBK, UHC/EUC-KR, Big5): a byte >= 0x81 that is not a
        // trail starts a two-byte character; its trail (always >= 0x40) must stay in the same block.
        if (_isDoubleByte) _isDoubleByteTrailNext = !_isDoubleByteTrailNext && IsDoubleByteLead(unit);
        if (BlockTextLengths.Count > 0) BlockTextLengths[BlockTextLengths.Count - 1]++;
        if (unit > ' ') _blockHasContent = true;
        if (_headingBlock >= 0 && _headingBytes.Count < MaximumHeadingBytes) AppendRaw(_headingBytes, first, second);
    }

    /// <summary>True when a block may end right after this unit: not inside a UTF-8 sequence or a UTF-16 surrogate pair.</summary>
    private bool IsCharacterEnd(int unit)
    {
        if (_unitSize == 2) return unit < 0xD800 || unit > 0xDBFF;
        if (_isDoubleByte) return !_isDoubleByteTrailNext;
        return unit < 0x80 || !_isUtf8;
    }

    /// <summary>Lead-byte ranges: 932 81-9F and E0-FC; 936, 949/51949, 950 81-FE (Windows code page tables).</summary>
    private bool IsDoubleByteLead(int unit) =>
        _encoding.CodePage == 932 ? (unit >= 0x81 && unit <= 0x9F) || (unit >= 0xE0 && unit <= 0xFC) : unit >= 0x81 && unit <= 0xFE;

    private static bool IsAsciiLetter(int unit) => (unit | 0x20) >= 'a' && (unit | 0x20) <= 'z';

    private void BeginTag(long offset, byte first, byte second)
    {
        _tagBytes.Clear();
        _tagStartOffset = offset;
        AppendRaw(_tagBytes, first, second);
    }

    /// <summary>A '<' that turned out not to start a tag: its bytes are text (the renderer escapes it the same way).</summary>
    private void AbandonTag()
    {
        _state = State.Text;
        if (BlockTextLengths.Count > 0) BlockTextLengths[BlockTextLengths.Count - 1]++;
        _blockHasContent = true;
        if (_tagBytes.Count <= _unitSize) { _tagBytes.Clear(); return; }
        // the '<' is text; everything after it is scanned again, as the renderer does (it may hold line breaks or real tags)
        byte[] rest = _tagBytes.GetRange(_unitSize, _tagBytes.Count - _unitSize).ToArray();
        long restOffset = _tagStartOffset + _unitSize;
        _tagBytes.Clear();
        for (int k = 0; k + _unitSize <= rest.Length; k += _unitSize)
        {
            byte a = rest[k], b = _unitSize == 2 ? rest[k + 1] : (byte)0;
            ProcessUnit(_unitSize == 2 ? UnitOf(a, b) : a, restOffset + k, a, b);
        }
    }

    private void CompleteBogus()
    {
        _state = State.Text;
        _tagBytes.Clear();
    }

    private void ProcessRawTextUnit(int unit, long offset, byte first, byte second)
    {
        if (_isCapturingStyle && _styleBytes.Count < MaximumStyleBytes) AppendRaw(_styleBytes, first, second);
        string pattern = _rawName;          // matched as "</" + name
        int lower = unit is >= 'A' and <= 'Z' ? unit | 0x20 : unit;
        int expected = _rawMatch == 0 ? '<' : _rawMatch == 1 ? '/' : pattern[_rawMatch - 2];
        if (lower == expected)
        {
            if (_rawMatch == 0) _rawEndTagOffset = offset;
            _rawMatch++;
            if (_rawMatch == pattern.Length + 2)
            {
                if (_isCapturingStyle) _styleBytes.RemoveRange(Math.Max(0, _styleBytes.Count - _rawMatch * _unitSize), Math.Min(_styleBytes.Count, _rawMatch * _unitSize));
                _isCapturingStyle = false;
                _tagBytes.Clear();
                foreach (char c in "</" + pattern) AppendRaw(_tagBytes, _unitSize == 2 && _bigEndian ? (byte)0 : (byte)c, _unitSize == 2 && _bigEndian ? (byte)c : (byte)0);
                _tagStartOffset = _rawEndTagOffset;
                _state = State.Tag; _isQuoted = false; _lastNonSpaceUnit = 'a';
                _rawMatch = 0;
            }
            return;
        }
        _rawMatch = lower == '<' ? 1 : 0;
        if (_rawMatch == 1) _rawEndTagOffset = offset;
        if (_mode == MarkupBlockMode.Lines)
        {
            if (BlockTextLengths.Count > 0) BlockTextLengths[BlockTextLengths.Count - 1]++;
            if (unit == '\n') EndLine(offset + _unitSize);
            else if (unit == '\r') _pendingCarriageReturn = true;
        }
    }

    // ── tags ─────────────────────────────────────────────────────────────────────────

    private void CompleteTag()
    {
        _state = State.Text;
        string tag = _encoding.GetString(_tagBytes.ToArray());
        _lastTagByteCount = _tagBytes.Count;
        _tagBytes.Clear();
        (string name, bool isEnd, bool isSelfClosing) = ParseTag(tag);
        if (name.Length == 0) return;
        if (isEnd) EndTag(name);
        else StartTag(name, tag, isSelfClosing);
    }

    internal static (string name, bool isEnd, bool isSelfClosing) ParseTag(string tag)
    {
        int i = 1;
        bool isEnd = tag.Length > 1 && tag[1] == '/';
        if (isEnd) i = 2;
        int start = i;
        while (i < tag.Length && (char.IsLetterOrDigit(tag[i]) || tag[i] == ':' || tag[i] == '-' || tag[i] == '_')) i++;
        string name = tag.Substring(start, i - start).ToLowerInvariant();
        bool isSelfClosing = tag.Length >= 2 && tag[tag.Length - 2] == '/';
        return (name, isEnd, isSelfClosing);
    }

    private void StartTag(string name, string tag, bool isSelfClosing)
    {
        if (_mode == MarkupBlockMode.Elements)
        {
            if (!_isInBody)
            {
                if (name == "body") { _isInBody = true; StartBlock(_tagStartOffset + _lastTagByteCount); }
                else if (name == "style" && !isSelfClosing) { _rawName = name; _rawMatch = 0; _state = State.RawText; _isCapturingStyle = true; }
                else if (RawTextElements.Contains(name) && !isSelfClosing) { _rawName = name; _rawMatch = 0; _state = State.RawText; }
                return;
            }
            if (BlockStartElements.Contains(name) && _tableDepth == 0 && _blockHasContent)
            {
                PopForStart(name);
                StartBlock(_tagStartOffset);
            }
            if (name == "table") _tableDepth++;
        }
        if (_mode == MarkupBlockMode.Lines || !BlockStartElements.Contains(name)) _blockHasContent = true;   // an empty block absorbs wrapper starts (div, p) instead of becoming an empty block
        if (NeverOnStack.Contains(name)) return;
        if (name.Length == 2 && name[0] == 'h' && name[1] >= '1' && name[1] <= '6' && _headingBlock < 0 && !isSelfClosing)
        {
            _headingLevel = name[1] - '0';
            _headingBlock = BlockStartOffsets.Count - 1;
            _headingBytes.Clear();
        }
        if (VoidElements.Contains(name) || isSelfClosing) return;
        PopForStart(name);
        if (_stack.Count < MaximumStackDepth) { _stack.Add((name, tag)); _isStackDirty = true; } else _unpushedDepth++;
        if (RawTextElements.Contains(name)) { _rawName = name; _rawMatch = 0; _state = State.RawText; }
    }

    /// <summary>The implied end tags HTML applies before a start tag (a new p closes an open p, a cell closes a cell...).</summary>
    private void PopForStart(string name)
    {
        if (_stack.Count == 0) return;
        if (name == "li") PopThrough(["li"], ["ul", "ol", "table", "td", "th"]);
        else if (name is "dt" or "dd") PopThrough(["dt", "dd"], ["dl", "table", "td", "th"]);
        else if (name is "td" or "th") PopThrough(["td", "th"], ["tr", "table"]);
        else if (name == "tr") PopThrough(["tr"], ["table", "tbody", "thead", "tfoot"]);
        else if (name == "option" && _stack[_stack.Count - 1].name == "option") { _stack.RemoveAt(_stack.Count - 1); _isStackDirty = true; }
        if (ClosesOpenParagraph.Contains(name) || name is "li" or "dt" or "dd") PopThrough(["p"], ParagraphScopeBoundaries);
    }

    private static readonly string[] ParagraphScopeBoundaries = ["table", "td", "th", "caption", "button", "object", "applet", "marquee", "template"];

    /// <summary>Pops through the nearest open element named in <paramref name="targets"/>, unless a boundary element sits above it.</summary>
    private void PopThrough(string[] targets, string[] boundaries)
    {
        for (int i = _stack.Count - 1; i >= 0; i--)
        {
            string open = _stack[i].name;
            if (Array.IndexOf(targets, open) >= 0) { _stack.RemoveRange(i, _stack.Count - i); _isStackDirty = true; return; }
            if (Array.IndexOf(boundaries, open) >= 0) return;
        }
    }

    private void EndTag(string name)
    {
        if (_mode == MarkupBlockMode.Elements)
        {
            if (!_isInBody) return;
            if (name == "body" || name == "html") { IsBodyEnded = true; return; }
            if (name == "table" && _tableDepth > 0) _tableDepth--;
        }
        _blockHasContent = true;
        if (_headingBlock >= 0 && name.Length == 2 && name[0] == 'h' && name[1] - '0' == _headingLevel) FinishHeading();
        if (_unpushedDepth > 0) { _unpushedDepth--; return; }
        for (int i = _stack.Count - 1; i >= 0; i--)
        {
            if (_stack[i].name != name) continue;
            _stack.RemoveRange(i, _stack.Count - i);
            _isStackDirty = true;
            return;
        }
    }

    // ── blocks ───────────────────────────────────────────────────────────────────────

    private void EndLine(long nextBlockOffset) => StartBlock(nextBlockOffset);

    private void StartBlock(long offset)
    {
        if (_headingBlock >= 0 && _mode == MarkupBlockMode.Elements) FinishHeading();   // a .txt heading may run over several lines
        if (BlockStartOffsets.Count > 0 && BlockStartOffsets[BlockStartOffsets.Count - 1] == offset) return;
        BlockStartOffsets.Add(offset);
        BlockTextLengths.Add(0);
        BlockStackIds.Add(CurrentStackId());
        _blockStartOffset = offset;
        _blockHasContent = false;
    }

    private int CurrentStackId()
    {
        if (!_isStackDirty) return _stackId;
        _isStackDirty = false;
        var key = new StringBuilder();
        foreach ((_, string tag) in _stack) key.Append(tag);
        string text = key.ToString();
        if (!_stackIds.TryGetValue(text, out _stackId))
        {
            _stackId = Stacks.Count;
            Stacks.Add(text);
            _stackIds.Add(text, _stackId);
        }
        return _stackId;
    }

    private void FinishHeading()
    {
        string text = WebUtility.HtmlDecode(_encoding.GetString(_headingBytes.ToArray()));
        text = string.Join(" ", text.Split([' ', '\t', '\r', '\n', ' '], StringSplitOptions.RemoveEmptyEntries));
        if (text.Length > 0) Headings.Add(new WordDocumentHeading(_headingBlock, _headingLevel, text));
        _headingBlock = -1;
        _headingBytes.Clear();
    }

    /// <summary>Ends the pass at <paramref name="endOffset"/> (end of text); drops a trailing empty block.</summary>
    public void Finish(long endOffset)
    {
        _pendingCarriageReturn = false;
        if (_headingBlock >= 0) FinishHeading();
        if (BlockStartOffsets.Count > 0 && BlockStartOffsets[BlockStartOffsets.Count - 1] >= endOffset)
        {
            BlockStartOffsets.RemoveAt(BlockStartOffsets.Count - 1);
            BlockTextLengths.RemoveAt(BlockTextLengths.Count - 1);
            BlockStackIds.RemoveAt(BlockStackIds.Count - 1);
        }
        EndOffset = endOffset;
    }

    /// <summary>Where the body ends: <see cref="Finish"/>'s offset, or the &lt;/body&gt; tag's start in Elements mode.</summary>
    public long EndOffset { get; private set; }
    public long BodyEndTagOffset => _tagStartOffset;
}
