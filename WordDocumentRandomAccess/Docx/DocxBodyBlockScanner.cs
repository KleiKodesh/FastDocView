using System.Text;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Streaming byte-level scan of word/document.xml. It never builds a DOM: it tracks element depth
/// and records the uncompressed byte offset where each child of &lt;w:body&gt; starts, the body's
/// end, the root element's opening tag (its namespace declarations are needed to parse a block in
/// isolation), each paragraph's style id (for headings), its list label (counted here, because a
/// windowed renderer never sees the earlier items) and a rough text length (for scroll-height
/// estimates). Fed in arbitrary chunks; state carries across chunk boundaries.
/// </summary>
internal sealed class DocxBodyBlockScanner
{
    private enum State { Text, TagOpen, TagName, TagBody, TagQuote, Comment, ProcessingInstruction, CData }

    private const int MaxRootTagBytes = 1 << 16;

    private readonly DocxListLabelResolver? _listLabelResolver;

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
    private string? _capturedAttributeTag;      // which tag's attributes are being captured (w:pStyle / w:numId / w:ilvl)
    private readonly MemoryStream _rootOpenTag = new();
    private bool _isRootTagDone;
    private bool _isInsideTextRun;              // inside <w:t>…</w:t>

    // the current body-level paragraph's pPr, resolved into a label when </w:pPr> closes
    private string? _currentStyleId;
    private int _currentNumId = -1;
    private int _currentLevel;

    public List<long> BlockStartOffsets { get; } = new();
    public List<WordDocumentBlockKind> BlockKinds { get; } = new();
    public List<int> BlockTextLengths { get; } = new();
    public List<short> BlockStyleIds { get; } = new();           // index into StyleIdTable, -1 = none
    public List<string> StyleIdTable { get; } = new();
    public List<int> BlockListLabelIds { get; } = new();         // index into ListLabels, -1 = not a list item
    public List<string> ListLabels { get; } = new();
    public long BodyEndOffset { get; private set; } = -1;
    public string RootOpenTag => Encoding.UTF8.GetString(_rootOpenTag.ToArray());

    private readonly Dictionary<string, short> _styleIdLookup = new(StringComparer.Ordinal);

    public DocxBodyBlockScanner(DocxListLabelResolver? listLabelResolver = null)
    {
        _listLabelResolver = listLabelResolver;
    }

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
                    _capturedAttributeTag = null;
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
                    if (_capturedAttributeTag is not null && b != (byte)'>') _tagAttributes.Append((char)b);
                    _previousWasSlash = b == (byte)'/';
                    break;

                case State.TagQuote:
                    if (_capturedAttributeTag is not null) _tagAttributes.Append((char)b);
                    if (b == _quoteChar) _state = State.TagBody;
                    break;

                case State.Comment:
                    // Entered on "<!": "<![CDATA[" ends with "]]>", "<!--" ends with "-->"; both end on '>'
                    // after their marker. The byte after '!' tells CDATA ("[") from a comment ("-").
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
        if (_isClosingTag || _bodyDepth < 0) return;
        string name = _tagName.ToString();
        // Only a body-level paragraph's pPr children are interesting: <w:pStyle> (depth body+2) and
        // <w:numPr>'s <w:numId>/<w:ilvl> (depth body+3).
        if (_depth == _bodyDepth + 2 && name == "w:pStyle") _capturedAttributeTag = name;
        else if (_depth == _bodyDepth + 3 && name is "w:numId" or "w:ilvl") _capturedAttributeTag = name;
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
            else if (name == "w:pPr" && _bodyDepth >= 0 && _depth == _bodyDepth + 1) OnParagraphPropertiesClosed();
            return;
        }

        if (_depth == 0) _isRootTagDone = true;

        if (_bodyDepth >= 0 && _depth == _bodyDepth) OnBodyChildStarted(name);
        else if (_capturedAttributeTag is not null) OnCapturedAttribute(_capturedAttributeTag, _tagAttributes.ToString());

        if (name == "w:body" && _bodyDepth < 0) _bodyDepth = _depth + 1;
        if (name == "w:t" && !isSelfClosing) _isInsideTextRun = true;
        if (!isSelfClosing) _depth++;
        _capturedAttributeTag = null;
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
        BlockListLabelIds.Add(-1);
        _currentStyleId = null;
        _currentNumId = -1;
        _currentLevel = 0;
    }

    private void OnCapturedAttribute(string tag, string attributes)
    {
        string? value = AttributeValue(attributes, "w:val=");
        if (value is null || BlockStyleIds.Count == 0) return;
        switch (tag)
        {
            case "w:pStyle":
                _currentStyleId = value;
                if (!_styleIdLookup.TryGetValue(value, out short tableIndex))
                {
                    tableIndex = (short)StyleIdTable.Count;
                    StyleIdTable.Add(value);
                    _styleIdLookup[value] = tableIndex;
                }
                BlockStyleIds[BlockStyleIds.Count - 1] = tableIndex;
                break;
            case "w:numId": if (int.TryParse(value, out int numId)) _currentNumId = numId; break;
            case "w:ilvl": if (int.TryParse(value, out int level)) _currentLevel = level; break;
        }
    }

    private void OnParagraphPropertiesClosed()
    {
        if (_listLabelResolver is null || BlockListLabelIds.Count == 0) return;
        string? label = _listLabelResolver(_currentStyleId, _currentNumId, _currentLevel);
        if (label is null) return;
        BlockListLabelIds[BlockListLabelIds.Count - 1] = ListLabels.Count;
        ListLabels.Add(label);
    }

    private static string? AttributeValue(string attributes, string attributeWithEquals)
    {
        int valueStart = attributes.IndexOf(attributeWithEquals, StringComparison.Ordinal);
        if (valueStart < 0) return null;
        valueStart += attributeWithEquals.Length + 1;                                  // past the opening quote
        int valueEnd = attributes.IndexOfAny(['"', '\''], valueStart);
        return valueEnd < 0 ? null : attributes.Substring(valueStart, valueEnd - valueStart);
    }
}
