using System.Text;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Streaming byte-level scan of word/document.xml. It never builds a DOM: it tracks element depth
/// and records the uncompressed byte offset where each child of &lt;w:body&gt; starts, the body's
/// end, the root element's opening tag (its namespace declarations are needed to parse a block in
/// isolation), the paragraph style id of each paragraph (for headings) and a rough text length
/// (for scroll-height estimates). Fed in arbitrary chunks; state carries across chunk boundaries.
/// </summary>
internal sealed class DocxBodyBlockScanner
{
    private enum State { Text, TagOpen, TagName, TagBody, TagQuote, Comment, ProcessingInstruction, CData }

    private const int MaxRootTagBytes = 1 << 16;

    private State _state = State.Text;
    private long _offset;                       // uncompressed offset of the next byte to be fed
    private long _tagStartOffset;
    private int _depth;
    private int _bodyDepth = -1;                // depth inside <w:body>; -1 before/after the body
    private bool _isClosingTag;
    private bool _previousWasSlash;
    private byte _quoteChar;
    private int _markerMatched;                 // progress through "-->" / "?>" / "]]>"

    private readonly StringBuilder _tagName = new(32);
    private bool _isTagNameComplete;
    private readonly StringBuilder _tagAttributes = new();
    private bool _isCapturingAttributes;
    private readonly MemoryStream _rootOpenTag = new();
    private bool _isRootTagDone;
    private bool _isInsideTextRun;              // inside <w:t>…</w:t>

    public List<long> BlockStartOffsets { get; } = new();
    public List<WordDocumentBlockKind> BlockKinds { get; } = new();
    public List<int> BlockTextLengths { get; } = new();
    public List<short> BlockStyleIds { get; } = new();           // index into StyleIdTable, -1 = none
    public List<string> StyleIdTable { get; } = new();
    public long BodyEndOffset { get; private set; } = -1;
    public string RootOpenTag => Encoding.UTF8.GetString(_rootOpenTag.ToArray());

    private readonly Dictionary<string, short> _styleIdLookup = new(StringComparer.Ordinal);

    public void Feed(byte[] chunk, int count)
    {
        for (int i = 0; i < count; i++)
        {
            byte b = chunk[i];
            if (!_isRootTagDone && _rootOpenTag.Length < MaxRootTagBytes && _state != State.Text && _depth == 0
                && _state != State.ProcessingInstruction && _state != State.Comment)
                _rootOpenTag.WriteByte(b);

            switch (_state)
            {
                case State.Text:
                    if (b == (byte)'<')
                    {
                        _state = State.TagOpen;
                        _tagStartOffset = _offset;
                        if (!_isRootTagDone && _depth == 0) { _rootOpenTag.SetLength(0); _rootOpenTag.WriteByte(b); }
                    }
                    else if (_isInsideTextRun && BlockTextLengths.Count > 0) BlockTextLengths[BlockTextLengths.Count - 1]++;
                    break;

                case State.TagOpen:
                    _tagName.Clear();
                    _isTagNameComplete = false;
                    _isClosingTag = false;
                    _previousWasSlash = false;
                    _isCapturingAttributes = false;
                    _tagAttributes.Clear();
                    if (b == (byte)'!') { _state = State.Comment; _markerMatched = 0; }        // <!-- or <![CDATA[ (decided below)
                    else if (b == (byte)'?') { _state = State.ProcessingInstruction; _markerMatched = 0; }
                    else if (b == (byte)'/') { _isClosingTag = true; _state = State.TagName; }
                    else { _state = State.TagName; _tagName.Append((char)b); }
                    break;

                case State.TagName:
                    if (b == (byte)'>' || b == (byte)'/' || b <= (byte)' ')
                    {
                        OnTagNameComplete();
                        _state = State.TagBody;
                        goto case State.TagBody;
                    }
                    if (_tagName.Length < 64) _tagName.Append((char)b);
                    break;

                case State.TagBody:
                    if (b == (byte)'"' || b == (byte)'\'') { _quoteChar = b; _state = State.TagQuote; }
                    else if (b == (byte)'>') { OnTagEnd(); _state = State.Text; }
                    if (_isCapturingAttributes && b != (byte)'>') _tagAttributes.Append((char)b);
                    _previousWasSlash = b == (byte)'/';
                    break;

                case State.TagQuote:
                    if (_isCapturingAttributes) _tagAttributes.Append((char)b);
                    if (b == _quoteChar) _state = State.TagBody;
                    break;

                case State.Comment:
                    // Entered on "<!": "<![CDATA[" ends with "]]>", "<!--" ends with "-->"; both end on '>'
                    // after their marker. The 7th byte tells CDATA ("[") from a comment ("-").
                    if (_markerMatched == 0 && b == (byte)'[') { _state = State.CData; _markerMatched = 0; break; }
                    if (b == (byte)'-') _markerMatched = Math.Min(2, _markerMatched + 1);
                    else if (b == (byte)'>' && _markerMatched == 2) _state = State.Text;
                    else _markerMatched = 0;
                    break;

                case State.CData:
                    if (b == (byte)']') _markerMatched = Math.Min(2, _markerMatched + 1);
                    else if (b == (byte)'>' && _markerMatched == 2) _state = State.Text;
                    else _markerMatched = 0;
                    break;

                case State.ProcessingInstruction:
                    if (b == (byte)'?') _markerMatched = 1;
                    else if (b == (byte)'>' && _markerMatched == 1) _state = State.Text;
                    else _markerMatched = 0;
                    break;
            }
            _offset++;
        }
    }

    private void OnTagNameComplete()
    {
        if (_isTagNameComplete) return;
        _isTagNameComplete = true;
        string name = _tagName.ToString();
        if (_isClosingTag) return;
        // Only one tag's attributes are ever interesting: a body-level paragraph's <w:pStyle>.
        _isCapturingAttributes = _bodyDepth >= 0 && _depth == _bodyDepth + 2 && name == "w:pStyle";
    }

    private void OnTagEnd()
    {
        string name = _tagName.ToString();
        bool isSelfClosing = _previousWasSlash;

        if (_isClosingTag)
        {
            if (name == "w:t") _isInsideTextRun = false;
            _depth--;
            if (name == "w:body" && _bodyDepth >= 0) { BodyEndOffset = _tagStartOffset; _bodyDepth = -1; }
            return;
        }

        if (_depth == 0) _isRootTagDone = true;

        if (_bodyDepth >= 0 && _depth == _bodyDepth) OnBodyChildStarted(name);
        else if (_isCapturingAttributes) OnParagraphStyle(_tagAttributes.ToString());

        if (name == "w:body" && _bodyDepth < 0) _bodyDepth = _depth + 1;
        if (name == "w:t" && !isSelfClosing) _isInsideTextRun = true;
        if (!isSelfClosing) _depth++;
        _isCapturingAttributes = false;
    }

    private void OnBodyChildStarted(string name)
    {
        BlockStartOffsets.Add(_tagStartOffset);
        BlockKinds.Add(name switch
        {
            "w:p" => WordDocumentBlockKind.Paragraph,
            "w:tbl" => WordDocumentBlockKind.Table,
            "w:sdt" => WordDocumentBlockKind.ContentControl,
            _ => WordDocumentBlockKind.Other,
        });
        BlockTextLengths.Add(0);
        BlockStyleIds.Add(-1);
    }

    private void OnParagraphStyle(string attributes)
    {
        int valueStart = attributes.IndexOf("w:val=", StringComparison.Ordinal);
        if (valueStart < 0 || BlockStyleIds.Count == 0) return;
        valueStart += 7;                                                   // past w:val=" (or ')
        int valueEnd = attributes.IndexOfAny(['"', '\''], valueStart);
        if (valueEnd < 0) return;
        string styleId = attributes.Substring(valueStart, valueEnd - valueStart);
        if (!_styleIdLookup.TryGetValue(styleId, out short tableIndex))
        {
            tableIndex = (short)StyleIdTable.Count;
            StyleIdTable.Add(styleId);
            _styleIdLookup[styleId] = tableIndex;
        }
        BlockStyleIds[BlockStyleIds.Count - 1] = tableIndex;
    }
}
