using System.Text;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Streaming byte scan of footnotes.xml / endnotes.xml: records the root's opening tag and, for
/// every child of the root (a w:footnote / w:endnote), its w:id, w:type and byte range — so one
/// note can later be read without the rest of the part. Same chunk-fed state machine idea as
/// <see cref="DocxBodyBlockScanner"/>, reduced to what notes need. Separator notes are dropped.
/// </summary>
internal sealed class DocxNotesPartScanner
{
    private enum State { Text, Tag, Quote, Comment, ProcessingInstruction, DocumentType }

    private State _state = State.Text;
    private long _offset;
    private long _tagStart;
    private int _depth;
    private byte _quote;
    private bool _isClosing;
    private bool _isFirstTagByte;
    private byte _previous;
    private int _markerMatched;
    private readonly DocxDocumentTypeDeclarationSkipper _documentType = new();
    private readonly MemoryStream _tag = new();       // the current tag's bytes (only kept while at depth ≤ 1)
    private long _noteStart = -1;
    private int _noteId;
    private bool _isSeparator;

    public string RootOpenTag { get; private set; } = "";
    public List<int> NoteIds { get; } = new();
    public List<long> NoteStarts { get; } = new();
    public List<long> NoteEnds { get; } = new();

    public void Feed(byte[] chunk, int count)
    {
        for (int i = 0; i < count; i++, _offset++)
        {
            byte b = chunk[i];
            switch (_state)
            {
                case State.Text:
                    if (b != (byte)'<') break;
                    _state = State.Tag;
                    _tagStart = _offset;
                    _isClosing = false;
                    _isFirstTagByte = true;
                    _tag.SetLength(0);
                    _tag.WriteByte(b);
                    break;

                case State.Tag:
                    if (_isFirstTagByte)
                    {
                        _isFirstTagByte = false;
                        if (b == (byte)'?') { _state = State.ProcessingInstruction; _markerMatched = 0; break; }
                        if (b == (byte)'!') { _state = State.Comment; _markerMatched = 0; break; }
                        if (b == (byte)'/') _isClosing = true;
                    }
                    if (_depth <= 1 && _tag.Length < 4096) _tag.WriteByte(b);
                    if (b is (byte)'"' or (byte)'\'') { _quote = b; _state = State.Quote; }
                    else if (b == (byte)'>') { OnTagEnd(isSelfClosing: _previous == (byte)'/'); _state = State.Text; }
                    break;

                case State.Quote:
                    if (_depth <= 1 && _tag.Length < 4096) _tag.WriteByte(b);
                    if (b == _quote) _state = State.Tag;
                    break;

                case State.Comment:
                    if (_markerMatched == 0 && b == (byte)'D' && _previous == (byte)'!') { _state = State.DocumentType; _documentType.Reset(); }
                    else if (b == (byte)'-') _markerMatched = Math.Min(2, _markerMatched + 1);
                    else if (b == (byte)'>' && _markerMatched == 2) _state = State.Text;
                    else if (b == (byte)'>' && _markerMatched == 0) _state = State.Text;    // <![CDATA[ … ]]>
                    else _markerMatched = 0;
                    break;

                case State.DocumentType:
                    if (_documentType.Feed(b)) _state = State.Text;
                    break;

                case State.ProcessingInstruction:
                    if (b == (byte)'?') _markerMatched = 1;
                    else if (b == (byte)'>' && _markerMatched == 1) _state = State.Text;
                    else _markerMatched = 0;
                    break;
            }
            _previous = b;
        }
    }

    private void OnTagEnd(bool isSelfClosing)
    {
        if (_isClosing)
        {
            _depth--;
            if (_depth == 1 && _noteStart >= 0) CloseNote(_offset + 1);
            return;
        }
        if (_depth == 0) RootOpenTag = Encoding.UTF8.GetString(_tag.ToArray());
        else if (_depth == 1)
        {
            string tag = Encoding.UTF8.GetString(_tag.ToArray());
            _noteStart = _tagStart;
            _noteId = int.TryParse(AttributeValue(tag, "w:id"), out int id) ? id : int.MinValue;
            _isSeparator = AttributeValue(tag, "w:type") is "separator" or "continuationSeparator" or "continuationNotice";
            if (isSelfClosing) { CloseNote(_offset + 1); return; }
        }
        if (!isSelfClosing) _depth++;
    }

    private void CloseNote(long end)
    {
        if (!_isSeparator && _noteId != int.MinValue)
        {
            NoteIds.Add(_noteId);
            NoteStarts.Add(_noteStart);
            NoteEnds.Add(end);
        }
        _noteStart = -1;
    }

    private static string? AttributeValue(string tag, string name)
    {
        int at = tag.IndexOf(" " + name + "=", StringComparison.Ordinal);
        if (at < 0) return null;
        int valueStart = at + name.Length + 3;                  // past ' name="'
        if (valueStart > tag.Length) return null;
        char quote = tag[valueStart - 1];
        int valueEnd = tag.IndexOf(quote, valueStart);
        return valueEnd < 0 ? null : tag.Substring(valueStart, valueEnd - valueStart);
    }
}
