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
    internal static readonly HashSet<string> ClosesOpenParagraph = new(StringComparer.Ordinal)
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
    /// <summary>
    /// Per block, a web archive's selector state of the elements open at its start (<see cref="CurrentStructure"/>; "" without a
    /// sheet that needs it). Not interned (each block's differs): the index keeps it only at its anchors, a span rescan rebuilds
    /// the rest, so it costs no memory per block.
    /// </summary>
    internal List<WebArchiveStructure> BlockStructures { get; } = [];
    public List<string> Stacks { get; private set; } = [""];   // id 0 = nothing open
    /// <summary>The document-wide index of <see cref="BlockStartOffsets"/>[0] (blocks before it were handed off and removed).</summary>
    public int BlockIndexBase { get; private set; }
    public List<WordDocumentHeading> Headings { get; } = [];
    /// <summary>LastLevelIndex 0: tilde headings kept aside; <see cref="Finish"/> adds them back when the contents would otherwise hold nothing besides the title (a book is never un-navigable).</summary>
    private readonly List<WordDocumentHeading> _hiddenTildeHeadings = [];
    public bool IsBodyStarted => _isInBody;


    private static readonly HashSet<string> HeadElements = new(StringComparer.Ordinal)
        { "html", "head", "meta", "link", "base", "title", "style", "script", "noscript", "template" };
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

    private readonly List<(string name, string tag, WebArchiveMatch? state)> _stack = [];
    // a web archive whose sheet has sibling combinators or child-index pseudo-classes: each open element's resolution state
    // (WebArchiveMatch), written into the interned stacks so a block resolves its selectors without the blocks before it
    private WebArchiveStylesheet? _structureSheet;
    private WebArchiveMatch? _rootState;
    private WebArchiveStructure? _lastStructure;
    private bool _isRescan;                              // a span rescan: it never records element ends (the index pass did)
    private bool _isRootEnded;
    private bool _isTopUnmarked = true;                  // the innermost open element may still be without content (MarkContent)
    /// <summary>
    /// The index pass over a sheet with later-sibling rules: the ends of the elements that span blocks, as they end (no whole-document
    /// pass: each is known once its end tag - or an implied end - is read).
    /// </summary>
    internal WebArchiveElementEnds ElementEnds { get; } = new();
    private bool RecordsEnds => !_isRescan && _structureSheet is { HasLaterSiblingRules: true };
    private bool _isStructureChanged;                    // an element opened or the stack changed since _lastStructure
    private int _unpushedDepth;                          // elements past MaximumStackDepth: their end tags are absorbed first
    private bool _isStackDirty;
    private int _stackId;
    private Dictionary<string, int> _stackIds = new(StringComparer.Ordinal) { [""] = 0 };
    private bool _isStackTableShared;                    // a span rescan: every stack it meets was interned by the index pass

    private bool _isInBody;
    private int _tableDepth;
    private bool _blockHasContent;
    private long _blockStartOffset;
    private bool _pendingCarriageReturn;
    private readonly bool _isDoubleByte; private bool _isDoubleByteTrailNext;
    private int _carry = -1; private long _carryOffset;

    private int _headingLevel; private int _headingBlock = -1; private readonly List<byte> _headingBytes = [];

    // Torat Emet text (.txt book files): the metadata head (text before _toratEmetHeadEnd) is ONE block, a line that starts with a mark and a
    // space is a heading (its text recorded without the mark), and only the formatting whitelist is kept on the element stack.
    private readonly long _toratEmetHeadEnd = -1;            // >= 0: Torat Emet text; where its head ends
    private readonly ToratEmetNote[] _notes = [];            // the note lines (hidden: no block, no text), in text order
    private int _noteCursor;                                 // the first note not yet passed
    private long _noteSkipEnd;                               // bytes before this offset belong to a note line and are dropped (0: not inside one)
    private readonly bool _isDroppingBlankLines;             // PutNewLinesAsIs=0: blank and whitespace-only lines are not blocks
    private readonly bool _isTildeInContents;                // LastLevelIndex >= 1 (or no metadata): a tilde line is a table-of-contents entry
    private int _markState;                                  // 0 none, 1 at a line start, 2 mark read - a space must follow
    private int _markLevel;
    private bool _isMarkHeading;
    private bool _isHeadingLine;                                // Torat Emet: the open block is a heading line (carried across span points: a rescan resumes mid-line)
    private bool _isAfterHeading;                                // Torat Emet: the open block follows a heading line; blank lines there are dropped even with PutNewLinesAsIs=1

    public MarkupBlockScanner(MarkupBlockMode mode, int unitSize, bool bigEndian, Encoding encoding, ToratEmetHead? toratEmet = null)
    {
        _mode = mode; _unitSize = unitSize; _bigEndian = bigEndian; _encoding = encoding;
        if (toratEmet is not null && mode == MarkupBlockMode.Lines && unitSize == 1)
        {
            _toratEmetHeadEnd = toratEmet.HeadEndOffset;
            _notes = toratEmet.Notes.ToArray();
            _isDroppingBlankLines = !toratEmet.IsNewLinesAsIs;
            _isTildeInContents = !toratEmet.HasHead || toratEmet.LastLevelIndex >= 1;
        }
        else if (mode == MarkupBlockMode.Lines) _isDroppingBlankLines = true;      // plain text: blank and whitespace-only lines are no blocks
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
            if (_state == State.Text && !_pendingCarriageReturn && _labelState < 2 && _noteSkipEnd == 0)
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
        if (_isTopUnmarked && _structureSheet is not null && text.Length > 0) MarkContent();
        if (!_blockHasContent)
            foreach (byte b in text) if (b > 0x20) { _blockHasContent = true; break; }
        int skip = _markState != 0 ? ConsumeMarkPrefix(text) : 0;
        if (_headingBlock >= 0 && _headingBytes.Count < MaximumHeadingBytes)
            for (int k = skip; k < text.Length && _headingBytes.Count < MaximumHeadingBytes; k++) _headingBytes.Add(text[k]);
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
        if (offset < _noteSkipEnd)
        {   // inside a note line (the header shows it): dropped whole, line break included
            if (offset + _unitSize >= _noteSkipEnd) _noteSkipEnd = 0;
            return;
        }
        if (_labelState >= 2 && ResolveLabel(unit, offset, first, second)) return;
        switch (_state)
        {
            case State.Text:
                if (unit == '<' && (_toratEmetHeadEnd < 0 || offset >= _toratEmetHeadEnd)) { _state = State.TagStart; BeginTag(offset, first, second); return; }   // the Torat Emet head is metadata: its '<' is text
                AddTextUnit(unit, first, second);
                if (_mode == MarkupBlockMode.Lines)
                {
                    if (unit == '\n') EndLine(offset + _unitSize);
                    else if (unit == '\r') _pendingCarriageReturn = true;
                    else if (unit == ' ' && offset - _blockStartOffset >= LineBlockSplitBytes) EndLine(offset + _unitSize, false);
                    else if (offset + _unitSize - _blockStartOffset >= LineBlockHardLimitBytes && IsCharacterEnd(unit)) EndLine(offset + _unitSize, false);
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
        if (_isTopUnmarked && _structureSheet is not null) MarkContent();
        if (unit > ' ') _blockHasContent = true;
        if (_markState != 0 && ConsumeMarkPrefix([first]) == 1) return;      // the mark and its space are not part of the heading's text
        if (_headingBlock >= 0 && _headingBytes.Count < MaximumHeadingBytes) AppendRaw(_headingBytes, first, second);
    }

    /// <summary>Torat Emet: reads a line's mark and the space after it from the start of <paramref name="text"/> (a piece may end between them);
    /// returns how many bytes belong to them. A heading mark starts the line's heading.</summary>
    private int ConsumeMarkPrefix(ReadOnlySpan<byte> text)
    {
        int used = 0;
        while (_markState != 0 && used < text.Length)
        {
            byte b = text[used];
            if (_markState == 1)
            {
                _markLevel = b < 0x80 ? ToratEmetFormat.LevelOfMark(b) : 0;
                if (_markLevel == 0) { _markState = 0; break; }
                _markState = 2; used++;
            }
            else
            {
                _markState = 0;
                if (b != (byte)' ') break;
                used++;
                if (_markLevel < 0) _labelState = 1;                      // an exclamation-mark line: a label that may merge with the next line
                else if (_markLevel > 0 && _headingBlock < 0)
                {
                    _headingLevel = _markLevel;
                    _headingBlock = BlockIndexBase + BlockStartOffsets.Count - 1;
                    _headingBytes.Clear();
                    _isMarkHeading = true;
                }
                if (_markLevel > 0)
                {
                    _isHeadingLine = true;
                }
            }
        }
        return used;
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
        _markState = 0;
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
                if (name == "body") { _isInBody = true; BeginStructure(WebArchiveStylesheet.ParseShared(WebArchiveStylesheet.Normalize(StyleText))); StartBlock(_tagStartOffset + _lastTagByteCount); }
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
        if (_toratEmetHeadEnd >= 0 && !ToratEmetFormat.FormattingTags.Contains(name)) return;               // Torat Emet: only the formatting whitelist is kept open
        if (NeverOnStack.Contains(name)) return;
        if (name.Length == 2 && name[0] == 'h' && name[1] >= '1' && name[1] <= '6' && _headingBlock < 0 && !isSelfClosing)
        {
            _headingLevel = name[1] - '0';
            _headingBlock = BlockIndexBase + BlockStartOffsets.Count - 1;
            _headingBytes.Clear();
        }
        if (VoidElements.Contains(name) || isSelfClosing) { OpenStructure(name, tag); return; }
        PopForStart(name);
        WebArchiveMatch? state = OpenStructure(name, tag);
        if (_stack.Count < MaximumStackDepth) { _stack.Add((name, tag, state)); _isStackDirty = _isStructureChanged = _isTopUnmarked = true; } else _unpushedDepth++;
        if (RawTextElements.Contains(name)) { _rawName = name; _rawMatch = 0; _state = State.RawText; }
    }

    /// <summary>Starts the element state of a web archive's body when its sheet needs it (the html and body as the renderer synthesizes them).</summary>
    private void BeginStructure(WebArchiveStylesheet sheet)
    {
        if (!sheet.NeedsStructure) return;
        _structureSheet = sheet;
        WebArchiveMatch html = sheet.Open(null, new WebArchiveElement("html", null, [], false, "<html>"));
        _rootState = sheet.Open(html, new WebArchiveElement("body", null, [], false, "<body>"));
    }

    /// <summary>Counts a body element under the open element (or body) and returns its state; null without a structure sheet or past the stack depth.</summary>
    private WebArchiveMatch? OpenStructure(string name, string tag)
    {
        if (_structureSheet is null || _unpushedDepth > 0) return null;
        WebArchiveMatch? parent = _stack.Count > 0 ? _stack[_stack.Count - 1].state : _rootState;
        if (parent is null) return null;
        _isStructureChanged = _isTopUnmarked = true;
        return _structureSheet.Open(parent, WebArchiveElement.FromTag(name, tag), sourcesOnly: true);
    }

    /// <summary>
    /// The selector state of body and of each open element, joined by '|' (body first; an entry without a state is empty) - what
    /// <see cref="MarkupBlockRenderer.WithStructure"/> puts into a block's open tags. "" without a structure sheet.
    /// </summary>
    public string CurrentStructure() => CurrentStructureSnapshot().Text;

    /// <summary>
    /// <see cref="CurrentStructure"/> as a snapshot that later elements do not change, its text written only when asked (a span
    /// anchor, a rendered block). The previous one is reused while no element opened and the stack did not change.
    /// </summary>
    private WebArchiveStructure CurrentStructureSnapshot()
    {
        if (_structureSheet is null || _rootState is null) return WebArchiveStructure.Empty;
        if (_lastStructure is not null && !_isStructureChanged) return _lastStructure;
        var entries = new WebArchiveMatch?[_stack.Count + 1];
        entries[0] = _rootState.Freeze();
        for (int i = 0; i < _stack.Count; i++) entries[i + 1] = _stack[i].state?.Freeze();
        _isStructureChanged = false;
        return _lastStructure = new WebArchiveStructure(entries);
    }

    /// <summary>Text inside the innermost open element (or body): it is not :empty.</summary>
    private void MarkContent()
    {
        WebArchiveMatch? top = _stack.Count > 0 ? _stack[_stack.Count - 1].state : _rootState;
        if (_unpushedDepth > 0) return;                      // past the stack depth: no state to mark
        if (top is not null) top.HasContent = true;
        _isTopUnmarked = false;
    }

    /// <summary>The open elements from <paramref name="from"/> up end here: the ones that span blocks leave their end facts.</summary>
    private void RecordEnds(int from)
    {
        if (!RecordsEnds) return;
        for (int i = Math.Max(0, from); i < _stack.Count; i++)
            if (_stack[i].state is { SpansBlocks: true } state && ElementPath(i) is { } path) ElementEnds.Add(path, state.EndFacts());
    }

    /// <summary>The body ends (its end tag or the text's end): every open element ends with it.</summary>
    private void EndRoot()
    {
        if (_isRootEnded || !_isInBody) return;
        _isRootEnded = true;
        RecordEnds(0);
        if (RecordsEnds && _rootState is { SpansBlocks: true } root) ElementEnds.Add("", root.EndFacts());
    }

    /// <summary>The path of the open element at <paramref name="depth"/>: "/" and its child index under body, then under each open ancestor; null where a state is missing.</summary>
    private string? ElementPath(int depth)
    {
        var path = new StringBuilder(depth * 3 + 3);
        for (int i = 0; i <= depth; i++)
        {
            if (_stack[i].state is not { } state) return null;
            path.Append('/').Append(state.Index);
        }
        return path.ToString();
    }

    /// <summary>The implied end tags HTML applies before a start tag (a new p closes an open p, a cell closes a cell...).</summary>
    private void PopForStart(string name)
    {
        if (_stack.Count == 0) return;
        if (name == "li") PopThrough(["li"], ["ul", "ol", "table", "td", "th"]);
        else if (name is "dt" or "dd") PopThrough(["dt", "dd"], ["dl", "table", "td", "th"]);
        else if (name is "td" or "th") PopThrough(["td", "th"], ["tr", "table"]);
        else if (name == "tr") PopThrough(["tr"], ["table", "tbody", "thead", "tfoot"]);
        else if (name == "option" && _stack[_stack.Count - 1].name == "option") { RecordEnds(_stack.Count - 1); _stack.RemoveAt(_stack.Count - 1); _isStackDirty = _isStructureChanged = _isTopUnmarked = true; }
        if (ClosesOpenParagraph.Contains(name) || name is "li" or "dt" or "dd") PopThrough(["p"], ParagraphScopeBoundaries);
    }

    private static readonly string[] ParagraphScopeBoundaries = ["table", "td", "th", "caption", "button", "object", "applet", "marquee", "template"];

    /// <summary>Pops through the nearest open element named in <paramref name="targets"/>, unless a boundary element sits above it.</summary>
    private void PopThrough(string[] targets, string[] boundaries)
    {
        for (int i = _stack.Count - 1; i >= 0; i--)
        {
            string open = _stack[i].name;
            if (Array.IndexOf(targets, open) >= 0) { RecordEnds(i); _stack.RemoveRange(i, _stack.Count - i); _isStackDirty = _isStructureChanged = _isTopUnmarked = true; return; }
            if (Array.IndexOf(boundaries, open) >= 0) return;
        }
    }

    private void EndTag(string name)
    {
        if (_mode == MarkupBlockMode.Elements)
        {
            if (!_isInBody) return;
            if (name == "body" || name == "html") { EndRoot(); IsBodyEnded = true; return; }
            if (name == "table" && _tableDepth > 0) _tableDepth--;
        }
        _blockHasContent = true;
        if (_toratEmetHeadEnd >= 0 && !ToratEmetFormat.FormattingTags.Contains(name)) return;
        if (_headingBlock >= 0 && !_isMarkHeading && name.Length == 2 && name[0] == 'h' && name[1] - '0' == _headingLevel) FinishHeading();
        if (_unpushedDepth > 0) { _unpushedDepth--; _isTopUnmarked = true; return; }
        for (int i = _stack.Count - 1; i >= 0; i--)
        {
            if (_stack[i].name != name) continue;
            RecordEnds(i);
            _stack.RemoveRange(i, _stack.Count - i);
            _isStackDirty = _isStructureChanged = _isTopUnmarked = true;
            return;
        }
    }

    // ── blocks ───────────────────────────────────────────────────────────────────────

    /// <summary>A real line break after a Torat Emet label line (an exclamation mark and a space) does not start a block yet: the next line's first bytes decide (see <see cref="ResolveLabel"/>). A split inside a long line never merges.</summary>
    private void EndLine(long nextBlockOffset, bool isLineBreak = true)
    {
        if (isLineBreak && _notes.Length > 0) nextBlockOffset = SkipNotes(nextBlockOffset);
        if (isLineBreak && _labelState == 1) { _labelState = _isDroppingBlankLines ? 4 : 2; _labelNextStart = nextBlockOffset; _labelAtLineStart = true; return; }
        if (isLineBreak && _labelState == 4) return;                       // the label absorbs blank lines until its text (or a heading) shows
        if (isLineBreak && (_isDroppingBlankLines || _isAfterHeading) && !_blockHasContent && (BlockIndexBase + BlockStartOffsets.Count > 1 || _toratEmetHeadEnd < 0) && _labelState == 0 && nextBlockOffset >= _toratEmetHeadEnd)
        {   // PutNewLinesAsIs=0: a blank or whitespace-only line is no block; the open block restarts at the next line (the line before it absorbs the blank bytes)
            BlockStartOffsets[BlockStartOffsets.Count - 1] = nextBlockOffset;
            BlockTextLengths[BlockTextLengths.Count - 1] = 0;
            _blockStartOffset = nextBlockOffset;
            BlockStackIds[BlockStackIds.Count - 1] = CurrentStackId();     // tags of the dropped line stay open
            if (_toratEmetHeadEnd >= 0) _markState = 1;
            return;
        }
        StartBlock(nextBlockOffset);
    }

    /// <summary>Where the next block starts when a line starts at <paramref name="offset"/>: after every note line that starts there (the bytes of a note are dropped as they are fed).</summary>
    private long SkipNotes(long offset)
    {
        while (_noteCursor < _notes.Length && _notes[_noteCursor].End <= offset) _noteCursor++;
        while (_noteCursor < _notes.Length && _notes[_noteCursor].Start == offset)
        {
            offset = _notes[_noteCursor].End;
            _noteSkipEnd = offset;
            _noteCursor++;
        }
        return offset;
    }

    // Torat Emet label lines: 0 not in one, 1 inside the label line, 2 its line break was read and the next line's first byte is awaited,
    // 3 that byte is a mark character and is held until the byte after it shows whether it is a heading mark (a mark and a space).
    private int _labelState;
    private long _labelNextStart, _labelHeldOffset;
    private bool _labelAtLineStart;
    private byte _labelHeld;

    /// <summary>
    /// A label line merges with the next line (one block) unless the label is the last line, the next line is blank, or it starts with a mark and a space
    /// (a heading, or another label). Called for every unit while a decision is pending; returns true when the unit was only held back.
    /// </summary>
    private bool ResolveLabel(int unit, long offset, byte first, byte second)
    {
        if (_labelState == 4)
        {   // blank lines are noise here: the label waits for the next line that has text (state 4 only when blank lines are dropped)
            if (unit == '\r' || unit == '\n') { _labelAtLineStart = true; return false; }
            if (_labelAtLineStart && unit < 0x80 && ToratEmetFormat.LevelOfMark(unit) != 0) { _labelState = 3; _labelHeld = first; _labelHeldOffset = offset; return true; }
            if (unit == ' ' || unit == '\t') { _labelAtLineStart = false; return false; }
            _labelState = 0; return false;
        }
        if (_labelState == 2)
        {
            if (unit == '\r' || unit == '\n') { _labelState = 0; StartBlock(_labelNextStart); return false; }   // blank next line: the label stays alone
            if (unit < 0x80 && ToratEmetFormat.LevelOfMark(unit) != 0) { _labelState = 3; _labelHeld = first; _labelHeldOffset = offset; return true; }
            _labelState = 0; return false;                                                                                  // merged: the block goes on
        }
        byte held = _labelHeld; long heldOffset = _labelHeldOffset;
        _labelState = 0;
        if (unit == ' ') StartBlock(heldOffset);                                                                       // a mark and a space: a heading line
        ProcessUnit(held, heldOffset, held, 0);                                                                             // the held byte, now where it belongs
        return false;
    }

    private void StartBlock(long offset)
    {
        _labelState = 0;
        if (_toratEmetHeadEnd >= 0 && BlockStartOffsets.Count > 0 && offset < _toratEmetHeadEnd) return;   // the metadata head is one block
        bool isMarkedHeadingEnd = _isHeadingLine;
        _isHeadingLine = false;
        if (_headingBlock >= 0 && (_mode == MarkupBlockMode.Elements || _isMarkHeading)) FinishHeading();   // a .txt heading may run over several lines; a marked one is one line
        if (BlockStartOffsets.Count > 0 && BlockStartOffsets[BlockStartOffsets.Count - 1] == offset) return;
        _isAfterHeading = isMarkedHeadingEnd && _toratEmetHeadEnd >= 0;
        if (_toratEmetHeadEnd >= 0)
        {
            _markState = offset >= _toratEmetHeadEnd ? 1 : 0;
            if (offset == _toratEmetHeadEnd) { _stack.Clear(); _unpushedDepth = 0; _isStackDirty = _isStructureChanged = _isTopUnmarked = true; }   // tags the head left open do not reach the text
        }
        if (RecordsEnds && BlockStartOffsets.Count + BlockIndexBase > 0)
        {
            _rootState!.SpansBlocks = true;                                   // a block starts inside each open element
            foreach ((_, _, WebArchiveMatch? open) in _stack) if (open is not null) open.SpansBlocks = true;
        }
        BlockStartOffsets.Add(offset);
        BlockTextLengths.Add(0);
        BlockStackIds.Add(CurrentStackId());
        BlockStructures.Add(CurrentStructureSnapshot());
        _blockStartOffset = offset;
        _blockHasContent = false;
    }

    private int CurrentStackId()
    {
        if (!_isStackDirty) return _stackId;
        _isStackDirty = false;
        var key = new StringBuilder();
        foreach ((_, string tag, _) in _stack) key.Append(tag);
        string text = key.ToString();
        if (!_stackIds.TryGetValue(text, out _stackId))
        {
            if (_isStackTableShared) throw new InvalidDataException("Markup index: a span rescan met an element stack the index pass did not (the document changed?).");
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
               bool isHiddenTilde = _isMarkHeading && _headingLevel == 4 && !_isTildeInContents;     // LastLevelIndex 0: the lowest level is not in the index
        if (text.Length > 0 && isHiddenTilde) _hiddenTildeHeadings.Add(new WordDocumentHeading(_headingBlock, _headingLevel, text));
        else if (text.Length > 0) Headings.Add(new WordDocumentHeading(_headingBlock, _headingLevel, text));
        _isMarkHeading = false;
        _headingBlock = -1;
        _headingBytes.Clear();
    }

    /// <summary>Ends the pass at <paramref name="endOffset"/> (end of text); drops a trailing empty block.</summary>
    public void Finish(long endOffset)
    {
        if (_pendingCarriageReturn && _toratEmetHeadEnd >= 0) EndLine(endOffset);        // Torat Emet: a last lone CR ends its line like a LF (no trailing empty block)
        _pendingCarriageReturn = false;
        if (_labelState == 3) { _labelState = 0; ProcessUnit(_labelHeld, _labelHeldOffset, _labelHeld, 0); }   // a label then one mark byte at the very end: that byte joins the label
        _labelState = 0;
        if (_headingBlock >= 0) FinishHeading();
        EndRoot();
        if (_hiddenTildeHeadings.Count > 0 && !Headings.Any(heading => heading.Level > 1))
        {
            Headings.AddRange(_hiddenTildeHeadings);
            Headings.Sort((a, b) => a.BlockIndex.CompareTo(b.BlockIndex));
        }
        _hiddenTildeHeadings.Clear();
        if (BlockStartOffsets.Count > 0 && BlockStartOffsets[BlockStartOffsets.Count - 1] >= endOffset)
        {
            BlockStartOffsets.RemoveAt(BlockStartOffsets.Count - 1);
            BlockTextLengths.RemoveAt(BlockTextLengths.Count - 1);
            BlockStackIds.RemoveAt(BlockStackIds.Count - 1);
            BlockStructures.RemoveAt(BlockStructures.Count - 1);
        }
        EndOffset = endOffset;
    }

    /// <summary>Where the body ends: <see cref="Finish"/>'s offset, or the &lt;/body&gt; tag's start in Elements mode.</summary>
    public long EndOffset { get; private set; }
    public long BodyEndTagOffset => _tagStartOffset;

    // -- compact index: hand-off of finished blocks, span snapshots ---------------------------------------

    /// <summary>Removes the first <paramref name="count"/> blocks (they are finished: only the last block still grows).</summary>
    public void RemoveFirstBlocks(int count)
    {
        if (count <= 0) return;
        if (count >= BlockStartOffsets.Count) throw new ArgumentOutOfRangeException(nameof(count), "the current block stays");
        BlockStartOffsets.RemoveRange(0, count);
        BlockTextLengths.RemoveRange(0, count);
        BlockStackIds.RemoveRange(0, count);
        BlockStructures.RemoveRange(0, count);
        BlockIndexBase += count;
    }

    /// <summary>
    /// True between two fed pieces when the scanner holds nothing but text-level state (no tag, comment, raw text,
    /// pending CR or half a code unit), inside the body with a block open: a <see cref="ExportSpanState"/> taken here
    /// lets a fresh scanner continue exactly as this one would.
    /// </summary>
    public bool IsAtSpanPoint => _state == State.Text && !_pendingCarriageReturn && _carry < 0 && _labelState == 0 && _noteSkipEnd == 0 && _markState != 2 && !IsToratEmetMarkLost && _isInBody && !IsBodyEnded
                                 && BlockStartOffsets.Count > 0;

    /// <summary>Torat Emet: a tag opened at a line start cleared the mark state; a resumed scanner could not know that, so no span point here.</summary>
    private bool IsToratEmetMarkLost => _toratEmetHeadEnd >= 0 && _markState == 0 && BlockStartOffsets.Count > 0
        && BlockStartOffsets[BlockStartOffsets.Count - 1] >= _toratEmetHeadEnd && BlockTextLengths[BlockTextLengths.Count - 1] == 0;

    /// <summary>What a span rescan needs to continue from here (see <see cref="IsAtSpanPoint"/>); headings are not carried.</summary>
    public MarkupScanState ExportSpanState()
    {
        if (!IsAtSpanPoint) throw new InvalidOperationException("not at a span point");
        int last = BlockStartOffsets.Count - 1;
        _isStackDirty = true;
        int liveStackId = CurrentStackId();
        var tagLengths = new int[_stack.Count];
        for (int i = 0; i < _stack.Count; i++) tagLengths[i] = _stack[i].tag.Length;
        return new MarkupScanState(BlockIndexBase + last, BlockStartOffsets[last], BlockStackIds[last], BlockTextLengths[last],
                                   liveStackId, tagLengths, _unpushedDepth, _tableDepth, _blockHasContent, _isDoubleByteTrailNext, _isAfterHeading, _isHeadingLine,
                                   BlockStructures[last].Text, CurrentStructure());
    }

    /// <summary>A scanner that continues from <paramref name="state"/> with the index pass's interned stacks.</summary>
    public static MarkupBlockScanner ResumeSpan(MarkupBlockMode mode, int unitSize, bool bigEndian, Encoding encoding,
                                                List<string> stacks, Dictionary<string, int> stackIds, MarkupScanState state, ToratEmetHead? toratEmet = null,
                                                string documentCss = "")
    {
        var scanner = new MarkupBlockScanner(mode, unitSize, bigEndian, encoding, toratEmet) { _isRescan = true };
        scanner.ShareStackTable(stacks, stackIds);
        if ((uint)state.LiveStackId >= (uint)stacks.Count || (uint)state.BlockStackId >= (uint)stacks.Count) throw new InvalidDataException("Markup index: anchor stack out of range.");
        string live = stacks[state.LiveStackId];
        int at = 0;
        if (mode == MarkupBlockMode.Elements && documentCss.Length > 0) scanner.BeginStructure(WebArchiveStylesheet.ParseShared(documentCss));
        string[]? structure = null;
        if (scanner._structureSheet is not null)
        {
            // the body's state first, then each open element's (from the anchor: the index pass's own)
            structure = state.LiveStructure.Split('|');
            if (structure.Length != state.TagLengths.Length + 1 || WebArchiveMatch.Deserialize(structure[0], null) is not { } root)
                throw new InvalidDataException("Markup index: anchor structure damaged.");
            scanner._rootState = root;
        }
        foreach (int length in state.TagLengths)
        {
            if (length < 2 || length > live.Length - at) throw new InvalidDataException("Markup index: anchor stack damaged.");
            string tag = live.Substring(at, length);
            at += length;
            string name = ParseTag(tag).name;
            WebArchiveMatch? entryState = null;
            if (structure is not null && structure[scanner._stack.Count + 1].Length > 0)
            {
                WebArchiveMatch? parent = scanner._stack.Count > 0 ? scanner._stack[scanner._stack.Count - 1].state : scanner._rootState;
                entryState = WebArchiveMatch.Deserialize(structure[scanner._stack.Count + 1], parent) ?? throw new InvalidDataException("Markup index: anchor structure damaged.");
            }
            scanner._stack.Add((name, tag, entryState));
        }
        if (at != live.Length) throw new InvalidDataException("Markup index: anchor stack damaged.");
        scanner._stackId = state.LiveStackId;
        scanner._isStackDirty = false;
        scanner._unpushedDepth = state.UnpushedDepth;
        scanner._tableDepth = state.TableDepth;
        scanner._isInBody = true;
        scanner.BlockIndexBase = state.Block;
        scanner.BlockStartOffsets.Add(state.BlockStart);
        scanner.BlockTextLengths.Add(state.BlockTextLength);
        scanner.BlockStackIds.Add(state.BlockStackId);
        scanner.BlockStructures.Add(new WebArchiveStructure(state.BlockStructure));
        scanner._blockStartOffset = state.BlockStart;
        scanner._blockHasContent = state.BlockHasContent;
        scanner._isDoubleByteTrailNext = state.IsDoubleByteTrailNext;
        scanner._isAfterHeading = state.IsAfterHeading;
        scanner._isHeadingLine = state.IsHeadingLine;
        // Torat Emet: a block that has read no text yet is at a line start, where a mark (and a label) is still to come
        if (scanner._toratEmetHeadEnd >= 0 && state.BlockTextLength == 0 && state.BlockStart >= scanner._toratEmetHeadEnd) scanner._markState = 1;
        return scanner;
    }

    /// <summary>A rescan from the very start (span 0) with the index pass's interned stacks.</summary>
    public static MarkupBlockScanner FromStart(MarkupBlockMode mode, int unitSize, bool bigEndian, Encoding encoding,
                                              List<string> stacks, Dictionary<string, int> stackIds, ToratEmetHead? toratEmet = null)
    {
        var scanner = new MarkupBlockScanner(mode, unitSize, bigEndian, encoding, toratEmet) { _isRescan = true };
        scanner.ShareStackTable(stacks, stackIds);
        return scanner;
    }

    private void ShareStackTable(List<string> stacks, Dictionary<string, int> stackIds)
    {
        Stacks = stacks;
        _stackIds = stackIds;
        _isStackTableShared = true;
    }
}

/// <summary>A <see cref="MarkupBlockScanner"/> snapshot at a span point: the open block (document index, start, stack id,
/// text length so far) and the text-level state (element stack as an interned stack + its tags' lengths, depths, flags).</summary>
public sealed record MarkupScanState(int Block, long BlockStart, int BlockStackId, int BlockTextLength, int LiveStackId, int[] TagLengths,
                                     int UnpushedDepth, int TableDepth, bool BlockHasContent, bool IsDoubleByteTrailNext, bool IsAfterHeading = false, bool IsHeadingLine = false,
                                     string BlockStructure = "", string LiveStructure = "");
