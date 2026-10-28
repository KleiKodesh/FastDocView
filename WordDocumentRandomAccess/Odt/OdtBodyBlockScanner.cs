using System.Xml;
using System.Net;
using System.Text;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.Odt;

/// <summary>
/// Streaming byte-level scan of an OpenDocument text body — content.xml of a package, or a whole
/// .fodt. No DOM: it tracks the element stack and records, per block, the byte range of the
/// element to render (a text:p / text:h — lists are flattened to one block per list paragraph —
/// or a whole table:table), its list depth and label (counted here with <see cref="OdtListNumbering"/>),
/// heading level + text, the note references inside it with their bodies' byte ranges, and a rough
/// text length. Before office:body it captures the root open tag (namespace declarations) and the
/// office:font-face-decls / office:styles / office:automatic-styles sections, which are handed to
/// the style cascade as they close and stored in the index. Fed in arbitrary chunks; text and tags
/// are skipped with vectorised span searches; a tag that lies whole inside a chunk is processed
/// in place, one crossing a chunk edge is assembled in a small buffer.
/// </summary>
internal sealed class OdtBodyBlockScanner
{
    private enum Frame : byte
    {
        Root, Skip, Capture, Body, Text, Transparent, Index, List, ListItem,
        Block, Inner, InnerNoText, Note, NoteCitation, NoteBody,
    }

    private enum Name : byte
    {
        Other, Paragraph, Heading, Table, List, ListItem, ListHeader, Section, SectionSource, IndexBody, IndexElement,
        SoftPageBreak, TrackedChanges, NumberedParagraph, Note, NoteCitation, NoteBody, OfficeBody, OfficeText,
        FontFaceDecls, Styles, AutomaticStyles, BinaryData, Annotation,
    }

    private const int MaxTagBytes = 1 << 20;
    private const int MaxHeadingTextBytes = 600;
    private const int MaxCitationBytes = 64;
    private const long MaxCaptureBytes = 64L << 20;

    private readonly OdtStyleCascade _styles;
    private readonly OdtListNumbering _numbering;
    private readonly List<Frame> _frames = new();
    private (byte[] name, Name kind)[][] _namesByLength;      // the names of interest, bucketed by byte length

    // tokenizer state
    private enum State { Text, TagOpen, Tag, Comment, Declaration, ProcessingInstruction, CData }
    private State _state = State.Text;
    private long _offset;                       // absolute offset of the current chunk's first byte
    private long _tagStartOffset;
    private readonly MemoryStream _partialTag = new();
    private byte _quote;
    private int _markerMatched;
    private byte[] _chunk = [];
    private int _chunkCount;

    // capture of style sections (before the body)
    private MemoryStream? _capture;
    private long _capturedUpTo;

    // block state
    private int _currentBlock = -1;
    private int _headingLevel;
    private readonly MemoryStream _headingText = new();
    private WordDocumentNoteKind _noteKind;
    private readonly MemoryStream _citationText = new();
    private string? _citationLabel;
    private long _noteBodyStart, _noteBodyEnd;
    private readonly int[] _noteCounts = new int[2];

    public string RootOpenTag { get; private set; } = "";
    public List<long> BlockStartOffsets { get; } = new();
    public List<long> BlockEndOffsets { get; } = new();
    public List<WordDocumentBlockKind> BlockKinds { get; } = new();
    public List<int> BlockTextLengths { get; } = new();
    public List<byte> BlockListLevels { get; } = new();          // 255 = not a list paragraph
    public List<int> BlockListLabelIds { get; } = new();         // -1 = no label
    public List<string> ListLabels { get; } = new();
    public List<WordDocumentHeading> Headings { get; } = new();
    public List<OdtNoteReference> NoteReferences { get; } = new();
    public bool IsBodyComplete { get; private set; }

    /// <summary>Block-level elements skipped because no rule handles them (element name → count) — for the audit.</summary>
    public Dictionary<string, int> SkippedBlockElements { get; } = new(StringComparer.Ordinal);

    public OdtBodyBlockScanner(OdtStyleCascade styles)
    {
        _styles = styles;
        _numbering = new OdtListNumbering(styles);
        _namesByLength = BuildNames(null);
    }

    // ── tokenizer ─────────────────────────────────────────────────────────────────────

    public void Feed(byte[] chunk, int count)
    {
        _chunk = chunk;
        _chunkCount = count;
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
                    int whole = WholeTagLength(span.Slice(position));
                    if (whole > 0)
                    {
                        OnTag(span.Slice(position, whole), _offset + position + whole);
                        position += whole;
                        break;
                    }
                    // the tag crosses the chunk edge, or is a comment / PI / CDATA
                    ReadOnlySpan<byte> rest = span.Slice(position);
                    if (rest.Length >= 2 && rest[1] == (byte)'!') { _state = State.Comment; _markerMatched = -1; position += 2; break; }
                    if (rest.Length >= 2 && rest[1] == (byte)'?') { _state = State.ProcessingInstruction; _markerMatched = 0; position += 2; break; }
                    if (rest.Length == 1) { _state = State.TagOpen; position++; break; }     // '<' is the chunk's last byte: the next byte decides
                    _partialTag.SetLength(0);
                    _quote = 0;
                    _state = State.Tag;
                    break;
                }

                case State.TagOpen:
                {
                    byte b = span[position];
                    if (b == (byte)'!') { _state = State.Comment; _markerMatched = -1; position++; break; }
                    if (b == (byte)'?') { _state = State.ProcessingInstruction; _markerMatched = 0; position++; break; }
                    _partialTag.SetLength(0);
                    _partialTag.WriteByte((byte)'<');
                    _quote = 0;
                    _state = State.Tag;
                    break;
                }

                case State.Declaration:
                {
                    // "<!DOCTYPE …>" and other declarations end at '>' (no internal subsets in ODF)
                    int close = span.Slice(position).IndexOf((byte)'>');
                    if (close < 0) { position = count; break; }
                    position += close + 1;
                    _state = State.Text;
                    break;
                }

                case State.Tag:
                {
                    // assemble the tag until a '>' outside quotes
                    int start = position;
                    while (position < count)
                    {
                        byte b = span[position++];
                        if (_quote != 0) { if (b == _quote) _quote = 0; }
                        else if (b is (byte)'"' or (byte)'\'') _quote = b;
                        else if (b == (byte)'>')
                        {
                            AppendPartial(chunk, start, position - start);
                            byte[] tag = _partialTag.ToArray();
                            _state = State.Text;
                            if (tag.Length >= 2 && tag[1] == (byte)'!') { }                                                // a short comment / declaration ended
                            else if (tag.Length >= 2 && tag[1] == (byte)'?') { }                                           // short PI ended
                            else OnTag(tag, _offset + position);
                            start = -1;
                            break;
                        }
                    }
                    if (start >= 0) AppendPartial(chunk, start, position - start);
                    break;
                }

                case State.Comment:
                {
                    // entered after "<!": "[" starts CDATA; a comment ends at "-->"
                    byte b = span[position++];
                    if (_markerMatched == -1)
                    {
                        if (b == (byte)'[') { _state = State.CData; _markerMatched = 0; break; }
                        if (b != (byte)'-') { _state = b == (byte)'>' ? State.Text : State.Declaration; break; }
                        _markerMatched = 0;
                    }
                    if (b == (byte)'-') _markerMatched = Math.Min(2, _markerMatched + 1);
                    else if (b == (byte)'>' && _markerMatched == 2) _state = State.Text;
                    else _markerMatched = 0;
                    break;
                }

                case State.CData:
                {
                    // CDATA does not occur in ODF bodies written by office suites; its text is not counted
                    byte b = span[position++];
                    if (b == (byte)']') _markerMatched = Math.Min(2, _markerMatched + 1);
                    else if (b == (byte)'>' && _markerMatched == 2) _state = State.Text;
                    else _markerMatched = 0;
                    break;
                }

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
        if (_capture is not null) CopyCaptured(_offset + count);
        _offset += count;
    }

    private void AppendPartial(byte[] chunk, int start, int length)
    {
        if (_partialTag.Length + length > MaxTagBytes) throw new InvalidDataException("OpenDocument: a tag longer than 1 MB.");
        _partialTag.Write(chunk, start, length);
    }

    /// <summary>Length of a complete element tag starting at '&lt;' in this chunk (balanced quotes), or 0.</summary>
    private static int WholeTagLength(ReadOnlySpan<byte> fromOpen)
    {
        if (fromOpen.Length < 2 || fromOpen[1] is (byte)'!' or (byte)'?') return 0;
        // jump from quote to quote with vectorised searches; a '>' inside a value is not the end
        int position = 1;
        while (true)
        {
            int next = fromOpen.Slice(position).IndexOfAny((byte)'>', (byte)'"', (byte)'\'');
            if (next < 0) return 0;
            position += next;
            byte found = fromOpen[position];
            if (found == (byte)'>') return position + 1;
            int close = fromOpen.Slice(position + 1).IndexOf(found);
            if (close < 0) return 0;
            position += close + 2;
        }
    }

    // ── element events ────────────────────────────────────────────────────────────────

    private void OnTag(ReadOnlySpan<byte> tag, long tagEndOffset)
    {
        bool isClosing = tag[1] == (byte)'/';
        bool isSelfClosing = !isClosing && tag[tag.Length - 2] == (byte)'/';
        int nameStart = isClosing ? 2 : 1;
        int nameEnd = nameStart;
        while (nameEnd < tag.Length && tag[nameEnd] is not ((byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n')) nameEnd++;
        Name name = Classify(tag.Slice(nameStart, nameEnd - nameStart));

        if (isClosing) { OnClose(name, tagEndOffset); return; }
        OnOpen(name, tag, nameEnd, tagEndOffset, isSelfClosing);
        if (isSelfClosing) OnClose(name, tagEndOffset);
    }

    private Name Classify(ReadOnlySpan<byte> qualifiedName)
    {
        if (qualifiedName.Length >= _namesByLength.Length) return Name.Other;
        foreach ((byte[] known, Name name) in _namesByLength[qualifiedName.Length])
            if (qualifiedName.SequenceEqual(known)) return name;
        return Name.Other;
    }

    private Frame Top => _frames.Count == 0 ? Frame.Root : _frames[_frames.Count - 1];

    private void OnOpen(Name name, ReadOnlySpan<byte> tag, int nameEnd, long tagEndOffset, bool isSelfClosing)
    {
        if (_frames.Count == 0)
        {
            RootOpenTag = Encoding.UTF8.GetString(tag.ToArray());
            if (RootOpenTag.EndsWith("/>", StringComparison.Ordinal)) RootOpenTag = RootOpenTag.Substring(0, RootOpenTag.Length - 2) + ">";
            _namesByLength = BuildNames(RootOpenTag);
            _frames.Add(Frame.Root);
            return;
        }

        Frame parent = Top;
        Frame frame;
        switch (parent)
        {
            case Frame.Skip: frame = Frame.Skip; break;
            case Frame.Capture: frame = Frame.Skip; break;
            case Frame.Root:
                if (name is Name.FontFaceDecls or Name.Styles or Name.AutomaticStyles) { frame = Frame.Capture; StartCapture(tag, tagEndOffset); }
                else frame = name == Name.OfficeBody ? Frame.Body : Frame.Skip;
                break;
            case Frame.Body:
                frame = name == Name.OfficeText ? Frame.Text : Frame.Skip;
                break;
            case Frame.Text or Frame.Transparent:
                frame = BlockLevel(name, tag, nameEnd, tagEndOffset, isListItemParagraph: false);
                break;
            case Frame.Index:
                frame = name == Name.IndexBody ? Frame.Transparent : Frame.Skip;      // *-source templates are not content
                break;
            case Frame.List:
                if (name is Name.ListItem or Name.ListHeader)
                {
                    int startValue = -1;
                    if (name == Name.ListItem && Attribute(tag, nameEnd, "text:start-value") is { } start && int.TryParse(start, out int parsed) && parsed >= 0) startValue = parsed;
                    _numbering.OpenItem(name == Name.ListHeader, startValue);
                    frame = Frame.ListItem;
                }
                else frame = Frame.Skip;
                break;
            case Frame.ListItem:
                frame = BlockLevel(name, tag, nameEnd, tagEndOffset, isListItemParagraph: true);
                break;
            default:   // inside a block
                frame = InsideBlock(parent, name, tag, nameEnd);
                break;
        }
        _frames.Add(frame);
    }

    private Frame BlockLevel(Name name, ReadOnlySpan<byte> tag, int nameEnd, long tagEndOffset, bool isListItemParagraph)
    {
        switch (name)
        {
            case Name.Paragraph or Name.Heading:
            {
                string? label = isListItemParagraph ? _numbering.ParagraphStarted() : null;
                StartBlock(WordDocumentBlockKind.Paragraph, isListItemParagraph || _numbering.IsInsideList ? (byte)_numbering.CurrentLevel : (byte)255, label);
                if (name == Name.Heading)
                {
                    _headingLevel = 1;                                           // txtparai.cxx: text:h defaults to outline level 1
                    if (Attribute(tag, nameEnd, "text:outline-level") is { } level && int.TryParse(level, out int parsed) && parsed >= 1) _headingLevel = Math.Min(10, parsed);
                }
                return Frame.Block;
            }
            case Name.Table:
                StartBlock(WordDocumentBlockKind.Table, 255, null);
                return Frame.Block;
            case Name.List:
                _numbering.OpenList(Attribute(tag, nameEnd, "xml:id"), Attribute(tag, nameEnd, "text:style-name"),
                                    Attribute(tag, nameEnd, "text:continue-numbering"), Attribute(tag, nameEnd, "text:continue-list"));
                return Frame.List;
            case Name.Section or Name.NumberedParagraph: return Frame.Transparent;
            case Name.IndexElement: return Frame.Index;
            case Name.SoftPageBreak or Name.SectionSource or Name.TrackedChanges: return Frame.Skip;
            default:
                string skipped = Encoding.UTF8.GetString(tag.Slice(1, nameEnd - 1).ToArray());
                SkippedBlockElements[skipped] = SkippedBlockElements.TryGetValue(skipped, out int count) ? count + 1 : 1;
                return Frame.Skip;
        }
    }

    private Frame InsideBlock(Frame parent, Name name, ReadOnlySpan<byte> tag, int nameEnd)
    {
        if (parent is Frame.InnerNoText) return Frame.InnerNoText;
        if (name is Name.BinaryData or Name.Annotation) return Frame.InnerNoText;
        if (parent == Frame.Note)
        {
            if (name == Name.NoteCitation)
            {
                _citationText.SetLength(0);
                _citationLabel = Attribute(tag, nameEnd, "text:label");
                return Frame.NoteCitation;
            }
            if (name == Name.NoteBody) { _noteBodyStart = _tagStartOffset; return Frame.NoteBody; }
            return Frame.InnerNoText;
        }
        if (name == Name.Note && parent is Frame.Block or Frame.Inner)
        {
            _noteKind = Attribute(tag, nameEnd, "text:note-class") == "endnote" ? WordDocumentNoteKind.Endnote : WordDocumentNoteKind.Footnote;
            _citationText.SetLength(0);
            _citationLabel = null;
            _noteBodyStart = _noteBodyEnd = -1;
            return Frame.Note;
        }
        if (parent is Frame.NoteBody or Frame.NoteCitation) return Frame.InnerNoText;     // note content is not body text
        return Frame.Inner;
    }

    private void OnClose(Name name, long tagEndOffset)
    {
        if (_frames.Count == 0) return;
        Frame frame = _frames[_frames.Count - 1];
        _frames.RemoveAt(_frames.Count - 1);
        switch (frame)
        {
            case Frame.Capture: FinishCapture(tagEndOffset); break;
            case Frame.Text: IsBodyComplete = true; break;
            case Frame.List: _numbering.CloseList(); break;
            case Frame.ListItem: _numbering.CloseItem(); break;
            case Frame.Block: FinishBlock(tagEndOffset); break;
            case Frame.NoteBody: _noteBodyEnd = tagEndOffset; break;
            case Frame.Note: FinishNote(); break;
        }
    }

    private void OnText(ReadOnlySpan<byte> text)
    {
        Frame top = Top;
        if (top is not (Frame.Block or Frame.Inner)) { if (top == Frame.NoteCitation) AppendBounded(_citationText, text, MaxCitationBytes); return; }
        if (_currentBlock < 0) return;
        BlockTextLengths[_currentBlock] += text.Length;
        if (_headingLevel > 0) AppendBounded(_headingText, text, MaxHeadingTextBytes);
    }

    private static void AppendBounded(MemoryStream destination, ReadOnlySpan<byte> text, int limit)
    {
        int room = (int)Math.Min(text.Length, limit - destination.Length);
        for (int i = 0; i < room; i++) destination.WriteByte(text[i]);     // net48 MemoryStream has no span overload; bounded
    }

    // ── blocks and notes ──────────────────────────────────────────────────────────────

    private void StartBlock(WordDocumentBlockKind kind, byte listLevel, string? label)
    {
        BlockStartOffsets.Add(_tagStartOffset);
        BlockEndOffsets.Add(-1);
        BlockKinds.Add(kind);
        BlockTextLengths.Add(0);
        BlockListLevels.Add(listLevel);
        if (label is not null) { BlockListLabelIds.Add(ListLabels.Count); ListLabels.Add(label); }
        else BlockListLabelIds.Add(-1);
        _currentBlock = BlockStartOffsets.Count - 1;
        _headingLevel = 0;
        _headingText.SetLength(0);
    }

    private void FinishBlock(long tagEndOffset)
    {
        if (_currentBlock < 0) return;
        BlockEndOffsets[_currentBlock] = tagEndOffset;
        if (_headingLevel > 0)
        {
            string text = WebUtility.HtmlDecode(Encoding.UTF8.GetString(_headingText.GetBuffer(), 0, (int)_headingText.Length)).Trim();
            if (text.Length > 0) Headings.Add(new WordDocumentHeading(_currentBlock, Math.Min(9, _headingLevel), text));
        }
        _headingLevel = 0;
        _currentBlock = -1;
    }

    /// <summary>
    /// A text:note closed: auto-numbered per class unless its citation has a text:label (custom mark,
    /// Number 0). The label shown is the citation text the producer wrote (LibreOffice writes the
    /// formatted number there), else the custom label, else the configured format.
    /// </summary>
    private void FinishNote()
    {
        if (_currentBlock < 0 || _noteBodyStart < 0 || _noteBodyEnd < 0) return;
        bool isCustom = !string.IsNullOrEmpty(_citationLabel);
        int number = isCustom ? 0 : ++_noteCounts[(int)_noteKind] + _styles.NotesConfiguration(_noteKind).StartOffset;
        string citation = WebUtility.HtmlDecode(Encoding.UTF8.GetString(_citationText.GetBuffer(), 0, (int)_citationText.Length)).Trim();
        string label = citation.Length > 0 ? citation : isCustom ? _citationLabel! : FormatNoteNumber(_noteKind, number);
        NoteReferences.Add(new OdtNoteReference(_currentBlock, _noteKind, number, label, _noteBodyStart, _noteBodyEnd));
    }

    private string FormatNoteNumber(WordDocumentNoteKind kind, int number)
    {
        OdtStyleCascade.OdtNotesConfiguration configuration = _styles.NotesConfiguration(kind);
        return configuration.Prefix + OdtListStyle.FormatNumber(number, configuration.NumberFormat) + configuration.Suffix;
    }

    // ── style sections ────────────────────────────────────────────────────────────────

    private void StartCapture(ReadOnlySpan<byte> openTag, long tagEndOffset)
    {
        _capture = new MemoryStream();
        _capture.Write(openTag.ToArray(), 0, openTag.Length);
        _capturedUpTo = tagEndOffset;
    }

    /// <summary>Copies the current chunk's bytes in [_capturedUpTo, until) into the capture.</summary>
    private void CopyCaptured(long until)
    {
        long from = Math.Max(_capturedUpTo, _offset);
        long to = Math.Min(until, _offset + _chunkCount);
        if (to > from)
        {
            if (_capture!.Length + (to - from) > MaxCaptureBytes) throw new NotSupportedException("OpenDocument: style section larger than 64 MB.");
            _capture.Write(_chunk, (int)(from - _offset), (int)(to - from));
        }
        _capturedUpTo = Math.Max(_capturedUpTo, to);
    }

    private void FinishCapture(long tagEndOffset)
    {
        CopyCaptured(tagEndOffset);
        string section = Encoding.UTF8.GetString(_capture!.GetBuffer(), 0, (int)_capture.Length);
        _capture = null;
        _styles.AddSection(ParseSection(RootOpenTag, section));
    }

    /// <summary>A captured section parsed inside the root tag (for its namespace declarations).</summary>
    public static XElement ParseSection(string rootOpenTag, string section)
    {
        string rootName = OdtXmlFragment.RootElementName(rootOpenTag);
        using var text = new StringReader(rootOpenTag + section + "</" + rootName + ">");
        try
        {
            using XmlReader reader = WordDocumentXml.CreateReader(text, tolerateUndeclaredPrefixes: true, ignoreWhitespace: true);
            return XElement.Load(reader).Elements().First();
        }
        catch (XmlException exception) { throw WordDocumentXml.Malformed("OpenDocument styles", exception); }
    }

    // ── attributes and names ──────────────────────────────────────────────────────────

    /// <summary>The value of a canonical-prefix attribute in an open tag (prefix mapped through the root's declarations).</summary>
    private string? Attribute(ReadOnlySpan<byte> tag, int nameEnd, string canonicalName)
    {
        string qualified = _prefixMap is null ? canonicalName : MapPrefix(canonicalName);
        string text = Encoding.UTF8.GetString(tag.Slice(nameEnd).ToArray());
        int at = 0;
        while (true)
        {
            at = text.IndexOf(qualified, at, StringComparison.Ordinal);
            if (at < 0) return null;
            bool startsAttribute = at > 0 && char.IsWhiteSpace(text[at - 1]);
            int equals = at + qualified.Length;
            while (equals < text.Length && char.IsWhiteSpace(text[equals])) equals++;
            if (startsAttribute && equals < text.Length && text[equals] == '=')
            {
                int quote = equals + 1;
                while (quote < text.Length && char.IsWhiteSpace(text[quote])) quote++;
                if (quote >= text.Length) return null;
                char quoteChar = text[quote];
                int close = text.IndexOf(quoteChar, quote + 1);
                return close < 0 ? null : WebUtility.HtmlDecode(text.Substring(quote + 1, close - quote - 1));
            }
            at += qualified.Length;
        }
    }

    private Dictionary<string, string>? _prefixMap;     // canonical prefix → the document's prefix

    private string MapPrefix(string canonicalName)
    {
        int colon = canonicalName.IndexOf(':');
        string prefix = canonicalName.Substring(0, colon);
        return _prefixMap!.TryGetValue(prefix, out string? actual) ? actual + canonicalName.Substring(colon) : canonicalName;
    }

    private static readonly Dictionary<string, string> CanonicalNamespaces = new(StringComparer.Ordinal)
    {
        ["office"] = OdtStyleCascade.Office.NamespaceName, ["text"] = OdtStyleCascade.Text.NamespaceName,
        ["table"] = OdtStyleCascade.Table.NamespaceName, ["style"] = OdtStyleCascade.Style.NamespaceName,
    };

    /// <summary>The element names of interest, spelled with the prefixes the root declares (LibreOffice uses the canonical ones).</summary>
    private (byte[] name, Name kind)[][] BuildNames(string? rootOpenTag)
    {
        _prefixMap = rootOpenTag is null ? null : new Dictionary<string, string>(StringComparer.Ordinal);
        if (rootOpenTag is not null)
            foreach (KeyValuePair<string, string> canonical in CanonicalNamespaces)
            {
                // the declaration must be "xmlns:prefix=\"uri\"" directly (a default xmlns=\"uri\" keeps the canonical prefix)
                int uri = rootOpenTag.IndexOf("=\"" + canonical.Value + "\"", StringComparison.Ordinal);
                if (uri < 0) continue;
                int nameStart = uri;
                while (nameStart > 0 && !char.IsWhiteSpace(rootOpenTag[nameStart - 1])) nameStart--;
                string attribute = rootOpenTag.Substring(nameStart, uri - nameStart);
                if (attribute.StartsWith("xmlns:", StringComparison.Ordinal)) _prefixMap![canonical.Key] = attribute.Substring(6);
            }
        string P(string canonicalName) => _prefixMap is null ? canonicalName : MapPrefix(canonicalName);
        var names = new Dictionary<string, Name>(StringComparer.Ordinal)
        {
            [P("text:p")] = Name.Paragraph, [P("text:h")] = Name.Heading, [P("table:table")] = Name.Table,
            [P("text:list")] = Name.List, [P("text:list-item")] = Name.ListItem, [P("text:list-header")] = Name.ListHeader,
            [P("text:section")] = Name.Section, [P("text:section-source")] = Name.SectionSource, [P("text:index-body")] = Name.IndexBody,
            [P("text:soft-page-break")] = Name.SoftPageBreak, [P("text:tracked-changes")] = Name.TrackedChanges,
            [P("text:numbered-paragraph")] = Name.NumberedParagraph, [P("text:note")] = Name.Note,
            [P("text:note-citation")] = Name.NoteCitation, [P("text:note-body")] = Name.NoteBody,
            [P("office:body")] = Name.OfficeBody, [P("office:text")] = Name.OfficeText,
            [P("office:font-face-decls")] = Name.FontFaceDecls, [P("office:styles")] = Name.Styles,
            [P("office:automatic-styles")] = Name.AutomaticStyles, [P("office:binary-data")] = Name.BinaryData,
            [P("office:annotation")] = Name.Annotation,
        };
        foreach (string index in new[] { "text:table-of-content", "text:alphabetical-index", "text:illustration-index", "text:table-index",
                                         "text:object-index", "text:user-index", "text:bibliography" })
            names[P(index)] = Name.IndexElement;
        int longest = names.Keys.Max(key => key.Length);
        var buckets = new List<(byte[], Name)>[longest + 1];
        for (int i = 0; i < buckets.Length; i++) buckets[i] = new List<(byte[], Name)>();
        foreach (KeyValuePair<string, Name> entry in names) buckets[entry.Key.Length].Add((Encoding.UTF8.GetBytes(entry.Key), entry.Value));
        return buckets.Select(bucket => bucket.ToArray()).ToArray();
    }
}

/// <summary>A note referenced inside a block: its class, number (0 = custom mark), label and its text:note-body byte range.</summary>
public readonly record struct OdtNoteReference(int BlockIndex, WordDocumentNoteKind Kind, int Number, string Label, long BodyStart, long BodyEnd);
