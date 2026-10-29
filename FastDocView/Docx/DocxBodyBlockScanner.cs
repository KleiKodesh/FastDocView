using System.Text;

namespace FastDocView.Docx;

/// <summary>
/// Streaming byte-level scan of word/document.xml. It never builds a DOM: it tracks element depth
/// and records the uncompressed byte offset where each child of &lt;w:body&gt; starts, the body's
/// end, the root element's opening tag (its namespace declarations are needed to parse a block in
/// isolation), each paragraph's style id (for headings), its list label (counted here, because a
/// windowed renderer never sees the earlier items), heading text, note references and a rough
/// text length (for scroll-height estimates). Fed in arbitrary chunks; state carries across chunk
/// boundaries. Text and tag bodies are skipped with vectorised span searches, and tag names are
/// compared as bytes against the dozen names that matter — no string per tag.
/// </summary>
internal sealed class DocxBodyBlockScanner
{
    private enum State { Text, TagOpen, TagName, TagBody, TagQuote, Comment, ProcessingInstruction, CData, DocumentType }

    private enum KnownTag : byte { Other, Body, Paragraph, Table, ContentControl, Text, ParagraphProperties, ParagraphStyle, NumberingId, NumberingLevel, FootnoteReference, EndnoteReference, FieldCharacter, FieldInstruction, DeletedContent }

    private const int MaxRootTagBytes = 1 << 16;
    private const int MaxHeadingTextBytes = 600;
    private const int MaxTagNameBytes = 64;

    private static readonly (byte[] name, KnownTag tag)[] KnownTags =
    [
        ("w:p"u8.ToArray(), KnownTag.Paragraph), ("w:t"u8.ToArray(), KnownTag.Text), ("w:pPr"u8.ToArray(), KnownTag.ParagraphProperties),
        ("w:tbl"u8.ToArray(), KnownTag.Table), ("w:sdt"u8.ToArray(), KnownTag.ContentControl), ("w:body"u8.ToArray(), KnownTag.Body),
        ("w:pStyle"u8.ToArray(), KnownTag.ParagraphStyle), ("w:numId"u8.ToArray(), KnownTag.NumberingId), ("w:ilvl"u8.ToArray(), KnownTag.NumberingLevel),
        ("w:footnoteReference"u8.ToArray(), KnownTag.FootnoteReference), ("w:endnoteReference"u8.ToArray(), KnownTag.EndnoteReference),
        ("w:del"u8.ToArray(), KnownTag.DeletedContent), ("w:moveFrom"u8.ToArray(), KnownTag.DeletedContent),
        ("w:fldChar"u8.ToArray(), KnownTag.FieldCharacter), ("w:instrText"u8.ToArray(), KnownTag.FieldInstruction),
    ];

    private readonly DocxListLabelResolver? _listLabelResolver;
    private readonly Func<string, int>? _headingLevelOfStyle;

    // heading text of the current body-level paragraph, when its style is a heading
    private int _currentHeadingLevel;
    private readonly MemoryStream _headingText = new();

    // note references: running auto-number per kind (footnote, endnote)
    private readonly int[] _noteAutoNumbers = new int[2];

    private State _state = State.Text;
    private long _offset;                       // uncompressed offset of the next byte to be fed
    private long _tagStartOffset;
    private int _depth;
    private int _bodyDepth = -1;                // depth inside <w:body>; -1 before/after the body
    private bool _isClosingTag;
    private bool _previousWasSlash;
    private byte _quoteChar;
    private int _markerMatched;                 // progress through "-->" / "?>" / "]]>"
    private readonly DocxDocumentTypeDeclarationSkipper _documentType = new();
    private bool _isFirstByteAfterBang;          // the byte after "<!" decides comment / CDATA / DOCTYPE

    private readonly byte[] _tagName = new byte[MaxTagNameBytes];
    private int _tagNameLength;
    private KnownTag _tag;
    private bool _isTagNameComplete;
    private readonly MemoryStream _tagAttributes = new();
    private bool _isCapturingAttributes;        // only for pStyle / numId / ilvl / note references
    private readonly MemoryStream _rootOpenTag = new();
    private bool _isRootTagDone;
    private bool _isInsideTextRun;              // inside <w:t>…</w:t>
    private bool _isInsideFieldInstruction;     // inside <w:instrText>…</w:instrText>
    private int _deletedContentDepth;           // open w:del / w:moveFrom elements: fields inside them are not shown, so not followed

    // complex fields (w:fldChar begin / separate / end) open at the current position: instruction bytes + whether separated
    private readonly List<(MemoryStream instruction, bool isSeparated)> _openFields = new();
    private readonly Dictionary<string, int> _openFieldStackLookup = new(StringComparer.Ordinal);

    /// <summary>Per block: index into <see cref="OpenFieldStacks"/> of the fields open where the block starts, -1 = none.</summary>
    public List<int> BlockOpenFieldIds { get; } = new();
    /// <summary>Distinct stacks of open fields, outermost first: each "S" or "I" (separated or still in its instruction) + the instruction, joined by U+0001.</summary>
    public List<string> OpenFieldStacks { get; } = new();

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
    public List<DocxListLabel> ListLabels { get; } = new();
    public List<WordDocumentHeading> Headings { get; } = new();
    public List<DocxNoteReference> NoteReferences { get; } = new();
    public long BodyEndOffset { get; private set; } = -1;
    public string RootOpenTag => Encoding.UTF8.GetString(_rootOpenTag.ToArray());

    private readonly Dictionary<string, short> _styleIdLookup = new(StringComparer.Ordinal);

    public DocxBodyBlockScanner(DocxListLabelResolver? listLabelResolver = null, Func<string, int>? headingLevelOfStyle = null)
    {
        _listLabelResolver = listLabelResolver;
        _headingLevelOfStyle = headingLevelOfStyle;
    }

    /// <summary>Call once after the last chunk: closes the last block's heading.</summary>
    public void Finish() => CloseHeading();

    public void Feed(byte[] chunk, int count)
    {
        var span = new ReadOnlySpan<byte>(chunk, 0, count);
        int position = 0;
        while (position < count)
        {
            switch (_state)
            {
                case State.Text:
                {
                    int next = span.Slice(position).IndexOf((byte)'<');
                    int textLength = next < 0 ? count - position : next;
                    if (textLength > 0) OnText(span.Slice(position, textLength));
                    position += textLength;
                    if (next < 0) break;
                    _tagStartOffset = _offset + position;
                    int consumed = TryProcessWholeTag(span.Slice(position + 1), chunk, position + 1);
                    if (consumed > 0) { position += 1 + consumed; break; }
                    _state = State.TagOpen;
                    if (!_isRootTagDone && _depth == 0) { _rootOpenTag.SetLength(0); _rootOpenTag.WriteByte((byte)'<'); }
                    position++;
                    break;
                }

                case State.TagOpen:
                {
                    byte b = span[position++];
                    CaptureRootTagByte(b);
                    _tagNameLength = 0;
                    _isTagNameComplete = false;
                    _isClosingTag = false;
                    _previousWasSlash = false;
                    _isCapturingAttributes = false;
                    _tagAttributes.SetLength(0);
                    if (b == (byte)'!') { _state = State.Comment; _markerMatched = 0; _isFirstByteAfterBang = true; }        // <!-- or <![CDATA[ (decided below)
                    else if (b == (byte)'?') { _state = State.ProcessingInstruction; _markerMatched = 0; }
                    else if (b == (byte)'/') { _isClosingTag = true; _state = State.TagName; }
                    else { _state = State.TagName; _tagName[_tagNameLength++] = b; }
                    break;
                }

                case State.TagName:
                {
                    ReadOnlySpan<byte> rest = span.Slice(position);
                    int end = IndexOfNameEnd(rest);
                    int take = end < 0 ? rest.Length : end;
                    int room = Math.Min(take, MaxTagNameBytes - _tagNameLength);
                    rest.Slice(0, room).CopyTo(_tagName.AsSpan(_tagNameLength));
                    _tagNameLength += room;
                    if (!_isRootTagDone && _depth == 0) _rootOpenTag.Write(chunk, position, take);
                    position += take;
                    if (end < 0) break;
                    OnTagNameComplete();
                    _state = State.TagBody;
                    break;
                }

                case State.TagBody:
                {
                    ReadOnlySpan<byte> rest = span.Slice(position);
                    int stop = rest.IndexOfAny((byte)'>', (byte)'"', (byte)'\'');
                    int take = stop < 0 ? rest.Length : stop;
                    if (take > 0)
                    {
                        if (_isCapturingAttributes) _tagAttributes.Write(chunk, position, take);
                        if (!_isRootTagDone && _depth == 0) _rootOpenTag.Write(chunk, position, take);
                        _previousWasSlash = rest[take - 1] == (byte)'/';
                        position += take;
                    }
                    if (stop < 0) break;
                    byte b = span[position++];
                    if (b == (byte)'>')
                    {
                        CaptureRootTagByte(b);
                        OnTagEnd();
                        _state = State.Text;
                    }
                    else
                    {
                        if (_isCapturingAttributes) _tagAttributes.WriteByte(b);
                        CaptureRootTagByte(b);
                        _quoteChar = b;
                        _previousWasSlash = false;
                        _state = State.TagQuote;
                    }
                    break;
                }

                case State.TagQuote:
                {
                    ReadOnlySpan<byte> rest = span.Slice(position);
                    int close = rest.IndexOf(_quoteChar);
                    int take = close < 0 ? rest.Length : close + 1;
                    if (_isCapturingAttributes) _tagAttributes.Write(chunk, position, take);
                    if (!_isRootTagDone && _depth == 0) _rootOpenTag.Write(chunk, position, take);
                    position += take;
                    if (close >= 0) _state = State.TagBody;
                    break;
                }

                case State.Comment:
                {
                    // Entered on "<!": "<![CDATA[" ends with "]]>", "<!--" ends with "-->"; both end on '>'
                    // after their marker. The byte after '!' tells CDATA ("[") from a comment ("-") and
                    // from a <!DOCTYPE …> ("D"), whose internal subset may hold '>' and "-->".
                    byte b = span[position++];
                    bool isFirst = _isFirstByteAfterBang;
                    _isFirstByteAfterBang = false;
                    if (isFirst && b == (byte)'[') { _state = State.CData; _markerMatched = 0; break; }
                    if (isFirst && b == (byte)'D') { _state = State.DocumentType; _documentType.Reset(); break; }
                    if (b == (byte)'-') _markerMatched = Math.Min(2, _markerMatched + 1);
                    else if (b == (byte)'>' && _markerMatched == 2) _state = State.Text;
                    else _markerMatched = 0;
                    break;
                }

                case State.CData:
                {
                    byte b = span[position++];
                    if (b == (byte)']') _markerMatched = Math.Min(2, _markerMatched + 1);
                    else if (b == (byte)'>' && _markerMatched == 2) _state = State.Text;
                    else _markerMatched = 0;
                    break;
                }

                case State.DocumentType:
                    if (_documentType.Feed(span[position++])) _state = State.Text;
                    break;

                case State.ProcessingInstruction:
                {
                    byte b = span[position++];
                    if (b == (byte)'?') _markerMatched = 1;
                    else if (b == (byte)'>' && _markerMatched == 1) _state = State.Text;
                    else _markerMatched = 0;
                    break;
                }
            }
        }
        _offset += count;
    }

    /// <summary>
    /// The fast path: an element tag that lies whole in this chunk is handled with one search for
    /// '>' instead of a state-machine step per byte. Returns the bytes consumed after '<', or 0 to
    /// fall back to the state machine (tag crosses the chunk edge, is a comment / PI / CDATA, or has
    /// unbalanced quotes — a '>' inside an attribute value would otherwise be taken for the end).
    /// </summary>
    private int TryProcessWholeTag(ReadOnlySpan<byte> afterOpen, byte[] chunk, int chunkPosition)
    {
        int end = afterOpen.IndexOf((byte)'>');
        if (end <= 0) return 0;
        ReadOnlySpan<byte> tag = afterOpen.Slice(0, end);
        byte first = tag[0];
        if (first is (byte)'!' or (byte)'?') return 0;
        if (tag.IndexOfAny((byte)'"', (byte)0x27) >= 0)
        {
            int doubleQuotes = 0, singleQuotes = 0;
            foreach (byte b in tag) { if (b == (byte)'"') doubleQuotes++; else if (b == (byte)0x27) singleQuotes++; }
            if ((doubleQuotes & 1) != 0 || (singleQuotes & 1) != 0) return 0;
        }

        _isClosingTag = first == (byte)'/';
        int nameStart = _isClosingTag ? 1 : 0;
        int nameEndRelative = IndexOfNameEnd(tag.Slice(nameStart));
        int nameEnd = nameEndRelative < 0 ? tag.Length : nameStart + nameEndRelative;
        ReadOnlySpan<byte> name = tag.Slice(nameStart, nameEnd - nameStart);
        _tagNameLength = Math.Min(name.Length, MaxTagNameBytes);
        name.Slice(0, _tagNameLength).CopyTo(_tagName);
        _previousWasSlash = tag[tag.Length - 1] == (byte)'/';
        _isTagNameComplete = false;
        _isCapturingAttributes = false;
        OnTagNameComplete();

        if (!_isRootTagDone && _depth == 0)
        {
            _rootOpenTag.SetLength(0);
            _rootOpenTag.WriteByte((byte)'<');
            _rootOpenTag.Write(chunk, chunkPosition, Math.Min(end + 1, MaxRootTagBytes));
        }
        if (_isCapturingAttributes)
        {
            _tagAttributes.SetLength(0);
            _tagAttributes.Write(chunk, chunkPosition + nameEnd, tag.Length - nameEnd);
        }
        OnTagEnd();
        return end + 1;
    }

    private static int IndexOfNameEnd(ReadOnlySpan<byte> rest)
    {
        for (int i = 0; i < rest.Length; i++)
        {
            byte b = rest[i];
            if (b == (byte)'>' || b == (byte)'/' || b <= (byte)' ') return i;
        }
        return -1;
    }

    private void CaptureRootTagByte(byte b)
    {
        if (!_isRootTagDone && _depth == 0 && _rootOpenTag.Length < MaxRootTagBytes) _rootOpenTag.WriteByte(b);
    }

    private void OnText(ReadOnlySpan<byte> text)
    {
        if (_isInsideFieldInstruction && _openFields.Count > 0 && !_openFields[_openFields.Count - 1].isSeparated)
            foreach (byte b in text) _openFields[_openFields.Count - 1].instruction.WriteByte(b);      // net48 MemoryStream has no span overload
        if (!_isInsideTextRun || BlockTextLengths.Count == 0) return;
        BlockTextLengths[BlockTextLengths.Count - 1] += text.Length;
        if (_currentHeadingLevel > 0 && _headingText.Length < MaxHeadingTextBytes)
        {
            ReadOnlySpan<byte> kept = text.Slice(0, (int)Math.Min(text.Length, MaxHeadingTextBytes - _headingText.Length));
            foreach (byte b in kept) _headingText.WriteByte(b);       // net48 MemoryStream has no span overload; ≤ 600 bytes
        }
    }

    private void OnTagNameComplete()
    {
        if (_isTagNameComplete) return;
        _isTagNameComplete = true;
        _tag = Classify(_tagName.AsSpan(0, _tagNameLength));
        if (_isClosingTag || _bodyDepth < 0) return;
        // A body-level paragraph's pPr children: <w:pStyle> (depth body+2) and <w:numPr>'s
        // <w:numId>/<w:ilvl> (depth body+3); and note references anywhere inside the body.
        _isCapturingAttributes = _tag switch
        {
            KnownTag.ParagraphStyle => _depth == _bodyDepth + 2,
            KnownTag.NumberingId or KnownTag.NumberingLevel => _depth == _bodyDepth + 3,
            KnownTag.FootnoteReference or KnownTag.EndnoteReference or KnownTag.FieldCharacter => _depth > _bodyDepth,
            _ => false,
        };
    }

    private static KnownTag Classify(ReadOnlySpan<byte> name)
    {
        if (name.Length < 3 || name[0] != (byte)'w' || name[1] != (byte)':') return KnownTag.Other;
        foreach ((byte[] known, KnownTag tag) in KnownTags)
            if (name.SequenceEqual(known)) return tag;
        return KnownTag.Other;
    }

    private void OnTagEnd()
    {
        bool isSelfClosing = _previousWasSlash;

        if (_isClosingTag)
        {
            if (_tag == KnownTag.Text) _isInsideTextRun = false;
            if (_tag == KnownTag.FieldInstruction) _isInsideFieldInstruction = false;
            if (_tag == KnownTag.DeletedContent && _deletedContentDepth > 0) _deletedContentDepth--;
            _depth--;
            if (_tag == KnownTag.Body && _bodyDepth >= 0) { CloseHeading(); BodyEndOffset = _tagStartOffset; _bodyDepth = -1; }
            else if (_tag == KnownTag.ParagraphProperties && _bodyDepth >= 0 && _depth == _bodyDepth + 1) OnParagraphPropertiesClosed();
            return;
        }

        if (_depth == 0) _isRootTagDone = true;

        if (_bodyDepth >= 0 && _depth == _bodyDepth) OnBodyChildStarted();
        else if (_isCapturingAttributes) OnCapturedAttribute(Encoding.ASCII.GetString(_tagAttributes.GetBuffer(), 0, (int)_tagAttributes.Length));

        // A document with several w:body elements (Apache POI's MultipleBodyBug.docx; POI's
        // TestXWPFWordExtractor.testMultipleBodyBug reads all of them): every body's children are
        // blocks, and the body ends at the last </w:body>.
        if (_tag == KnownTag.Body && _bodyDepth < 0) { _bodyDepth = _depth + 1; BodyEndOffset = -1; }
        if (_tag == KnownTag.Text && !isSelfClosing) _isInsideTextRun = true;
        if (_tag == KnownTag.FieldInstruction && !isSelfClosing && _bodyDepth >= 0) _isInsideFieldInstruction = true;
        if (_tag == KnownTag.DeletedContent && !isSelfClosing && _bodyDepth >= 0) _deletedContentDepth++;
        if (!isSelfClosing) _depth++;
        _isCapturingAttributes = false;
    }

    private void OnBodyChildStarted()
    {
        CloseHeading();
        BlockStartOffsets.Add(_tagStartOffset);
        BlockKinds.Add(_tag switch
        {
            KnownTag.Paragraph => WordDocumentBlockKind.Paragraph,
            KnownTag.Table => WordDocumentBlockKind.Table,
            KnownTag.ContentControl => WordDocumentBlockKind.ContentControl,
            _ => WordDocumentBlockKind.Other,
        });
        BlockTextLengths.Add(0);
        BlockStyleIds.Add(-1);
        BlockListLabelIds.Add(-1);
        BlockOpenFieldIds.Add(OpenFieldStackId());
        _currentHeadingLevel = 0;
        _currentStyleId = null;
        _currentNumId = -1;
        _currentLevel = 0;
    }

    private void OnCapturedAttribute(string attributes)
    {
        if (BlockStyleIds.Count == 0) return;
        if (_tag == KnownTag.FieldCharacter)
        {
            if (_deletedContentDepth > 0) return;                // the renderer skips deleted runs: their field characters do not count
            switch (AttributeValue(attributes, "w:fldCharType="))
            {
                case "begin": _openFields.Add((new MemoryStream(), false)); break;
                case "separate" when _openFields.Count > 0: _openFields[_openFields.Count - 1] = (_openFields[_openFields.Count - 1].instruction, true); break;
                case "end" when _openFields.Count > 0: _openFields.RemoveAt(_openFields.Count - 1); break;
            }
            return;
        }
        if (_tag is KnownTag.FootnoteReference or KnownTag.EndnoteReference)
        {
            OnNoteReference(_tag == KnownTag.FootnoteReference ? WordDocumentNoteKind.Footnote : WordDocumentNoteKind.Endnote, attributes);
            return;
        }
        string? value = AttributeValue(attributes, "w:val=");
        if (value is null) return;
        switch (_tag)
        {
            case KnownTag.ParagraphStyle:
                _currentStyleId = value;
                _currentHeadingLevel = _headingLevelOfStyle?.Invoke(value) ?? 0;
                if (!_styleIdLookup.TryGetValue(value, out short tableIndex))
                {
                    tableIndex = (short)StyleIdTable.Count;
                    StyleIdTable.Add(value);
                    _styleIdLookup[value] = tableIndex;
                }
                BlockStyleIds[BlockStyleIds.Count - 1] = tableIndex;
                break;
            case KnownTag.NumberingId: if (int.TryParse(value, out int numId)) _currentNumId = numId; break;
            case KnownTag.NumberingLevel: if (int.TryParse(value, out int level)) _currentLevel = level; break;
        }
    }

    /// <summary>
    /// Word numbers notes in reference order, separately per kind; a reference with
    /// w:customMarkFollows shows the run text after it instead of a number (Number 0).
    /// </summary>
    private void OnNoteReference(WordDocumentNoteKind kind, string attributes)
    {
        if (!int.TryParse(AttributeValue(attributes, "w:id="), out int id)) return;
        string? customMark = AttributeValue(attributes, "w:customMarkFollows=");
        bool isCustomMark = customMark is "1" or "true" or "on";
        int number = isCustomMark ? 0 : ++_noteAutoNumbers[(int)kind];
        NoteReferences.Add(new DocxNoteReference(BlockStartOffsets.Count - 1, kind, id, number));
    }

    /// <summary>
    /// The fields open where a block starts (a HYPERLINK field result can run across paragraphs; a renderer that starts at
    /// this block needs it), interned: a table of contents field open over thousands of blocks is one string.
    /// </summary>
    private int OpenFieldStackId()
    {
        if (_openFields.Count == 0) return -1;
        var stack = new StringBuilder();
        foreach ((MemoryStream instruction, bool isSeparated) in _openFields)
        {
            if (stack.Length > 0) stack.Append('\u0001');
            stack.Append(isSeparated ? 'S' : 'I')
                 .Append(System.Net.WebUtility.HtmlDecode(Encoding.UTF8.GetString(instruction.GetBuffer(), 0, (int)instruction.Length)));
        }
        string key = stack.ToString();
        if (!_openFieldStackLookup.TryGetValue(key, out int id))
        {
            id = OpenFieldStacks.Count;
            OpenFieldStacks.Add(key);
            _openFieldStackLookup[key] = id;
        }
        return id;
    }

    private void CloseHeading()
    {
        if (_currentHeadingLevel > 0 && BlockStartOffsets.Count > 0)
        {
            string text = System.Net.WebUtility.HtmlDecode(Encoding.UTF8.GetString(_headingText.GetBuffer(), 0, (int)_headingText.Length)).Trim();
            if (text.Length > 0) Headings.Add(new WordDocumentHeading(BlockStartOffsets.Count - 1, _currentHeadingLevel, text));
        }
        _currentHeadingLevel = 0;
        _headingText.SetLength(0);
    }

    private void OnParagraphPropertiesClosed()
    {
        if (_listLabelResolver is null || BlockListLabelIds.Count == 0) return;
        DocxListLabel? label = _listLabelResolver(_currentStyleId, _currentNumId, _currentLevel);
        if (label is not { } found) return;
        BlockListLabelIds[BlockListLabelIds.Count - 1] = ListLabels.Count;
        ListLabels.Add(found);
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
