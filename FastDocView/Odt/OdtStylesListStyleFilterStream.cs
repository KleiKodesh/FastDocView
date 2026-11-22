using System.Text;

namespace FastDocView.Odt;

/// <summary>
/// A read-only view of styles.xml without its text:list-style elements: each one is cut out of the byte stream as it
/// passes, parsed straight from its bytes (<see cref="OdtListStyle.Parse"/>) and handed to a callback, so the XmlReader
/// after this stream only sees the rest of the part. A LibreOffice conversion of a Word file carries thousands of list
/// styles (the sample: 3,058 of 10 levels, 13 MB of a 15 MB styles.xml); reading them through XmlReader (let alone as
/// XML trees) dominated the first open. Bounded memory: one list-style element at a time plus a chunk.
/// The prefixes come from the root element's xmlns declarations; when the text namespace has no prefix of its own (a
/// default xmlns), nothing is cut and the XmlReader path reads the list styles instead.
/// </summary>
internal sealed class OdtStylesListStyleFilterStream : Stream
{
    private const int ChunkSize = 1 << 16;
    private const int MaxRootTagBytes = 1 << 20;

    private readonly Stream _inner;
    private readonly Action<string, OdtListStyle> _onListStyle;
    private readonly Action<string, string, OdtStyleCascade.UnparsedStyle>? _onStyle;
    private readonly byte[] _chunk = new byte[ChunkSize];
    private byte[] _output = new byte[ChunkSize * 2];
    private int _outputStart, _outputEnd;
    private bool _isInnerDone;

    // the root start tag (for the prefixes), then the filter
    private readonly MemoryStream _head = new();
    private bool _isHeadDone;
    private byte[]? _startMarker, _endMarker;        // "<text:list-style", "</text:list-style>"
    private byte[]? _styleStartMarker, _styleEndMarker;   // "<style:style", "</style:style>" (only with a style callback)
    private byte[] _rootOpenTag = [];
    private string _rootElementName = "";
    private bool _isCuttingStyle;                    // the element being cut is a style:style (else a text:list-style)
    private string _textPrefix = "text", _stylePrefix = "style", _foPrefix = "fo";
    private byte[] _carry = [];                      // bytes held back: a possible marker start at a chunk's end
    private MemoryStream? _element;                  // the list-style element being cut (null: none)
    private readonly MemoryStream _elementBuffer = new();    // reused for every element
    private byte[] _joined = new byte[ChunkSize * 2];       // held-back bytes + the next chunk, reused
    private int _elementSearchFrom;
    private readonly OdtListStyle.ParsedLevels _parsedLevels = new();

    /// <param name="onStyle">When given, style:style elements are cut too and handed over unparsed (family, name, bytes):
    /// the cascade parses one when it is first looked up (<see cref="OdtStyleCascade.AddUnparsedStyle"/>).</param>
    public OdtStylesListStyleFilterStream(Stream inner, Action<string, OdtListStyle> onListStyle,
                                          Action<string, string, OdtStyleCascade.UnparsedStyle>? onStyle = null)
    {
        _inner = inner;
        _onListStyle = onListStyle;
        _onStyle = onStyle;
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        while (_outputStart == _outputEnd)
        {
            if (_isInnerDone) return 0;
            Fill();
        }
        int copied = Math.Min(count, _outputEnd - _outputStart);
        Buffer.BlockCopy(_output, _outputStart, buffer, offset, copied);
        _outputStart += copied;
        return copied;
    }

    private void Fill()
    {
        _outputStart = _outputEnd = 0;
        int read = _inner.Read(_chunk, 0, _chunk.Length);
        if (read <= 0)
        {
            _isInnerDone = true;
            if (!_isHeadDone) Emit(_head.GetBuffer(), 0, (int)_head.Length);
            Emit(_carry, 0, _carry.Length);                 // an element left open at the end: dropped, as a broken part
            _carry = [];
            return;
        }
        if (!_isHeadDone) { ReadHead(read); return; }
        Filter(_chunk, read);
    }

    /// <summary>Collects bytes until the root start tag has closed, then learns the prefixes and filters the rest.</summary>
    private void ReadHead(int read)
    {
        _head.Write(_chunk, 0, read);
        byte[] head = _head.GetBuffer();
        int length = (int)_head.Length;
        int rootEnd = RootStartTagEnd(head, length);
        if (rootEnd < 0 && length < MaxRootTagBytes) return;
        _isHeadDone = true;
        if (rootEnd < 0) { Emit(head, 0, length); return; }      // no root tag found in 1 MB: pass everything through
        string rootTag = Encoding.UTF8.GetString(head, 0, rootEnd);
        LearnPrefixes(rootTag);
        int tagStart = rootTag.LastIndexOf('<');
        _rootOpenTag = Encoding.UTF8.GetBytes(rootTag.Substring(tagStart));
        _rootElementName = OdtXmlFragment.RootElementName(rootTag.Substring(tagStart));
        Emit(head, 0, rootEnd);
        byte[] rest = new byte[length - rootEnd];
        Buffer.BlockCopy(head, rootEnd, rest, 0, rest.Length);
        Filter(rest, rest.Length);
    }

    /// <summary>The offset just after the root element's start tag ('&gt;' outside quotes), or -1 when it has not closed yet.</summary>
    internal static int RootStartTagEnd(byte[] bytes, int length)
    {
        int i = 0;
        while (i < length)
        {
            int open = Array.IndexOf(bytes, (byte)'<', i, length - i);
            if (open < 0 || open + 1 >= length) return -1;
            byte next = bytes[open + 1];
            if (next is (byte)'?' or (byte)'!')
            {
                int close = Array.IndexOf(bytes, (byte)'>', open, length - open);
                if (close < 0) return -1;
                i = close + 1;
                continue;
            }
            byte quote = 0;
            for (int j = open + 1; j < length; j++)
            {
                byte b = bytes[j];
                if (quote != 0) { if (b == quote) quote = 0; }
                else if (b is (byte)'"' or (byte)'\'') quote = b;
                else if (b == (byte)'>') return j + 1;
            }
            return -1;
        }
        return -1;
    }

    /// <summary>The prefix the root start tag declares for <paramref name="namespaceUri"/> (xmlns:prefix="uri"); null for none or a default xmlns.</summary>
    internal static string? PrefixOf(string rootTag, string namespaceUri)
    {
        int at = rootTag.IndexOf("=\"" + namespaceUri + "\"", StringComparison.Ordinal);
        if (at < 0) at = rootTag.IndexOf("='" + namespaceUri + "'", StringComparison.Ordinal);
        if (at < 0) return null;
        int start = at;
        while (start > 0 && !char.IsWhiteSpace(rootTag[start - 1])) start--;
        string attribute = rootTag.Substring(start, at - start);
        return attribute.StartsWith("xmlns:", StringComparison.Ordinal) ? attribute.Substring(6) : null;
    }

    private void LearnPrefixes(string rootTag)
    {
        string? text = PrefixOf(rootTag, OdtStyleCascade.Text.NamespaceName), style = PrefixOf(rootTag, OdtStyleCascade.Style.NamespaceName), fo = PrefixOf(rootTag, OdtStyleCascade.Fo.NamespaceName);
        if (text is null || style is null) return;            // no filtering: the XmlReader path reads the list styles
        _textPrefix = text;
        _stylePrefix = style;
        _foPrefix = fo ?? "fo";
        _startMarker = Encoding.UTF8.GetBytes("<" + text + ":list-style");
        _endMarker = Encoding.UTF8.GetBytes("</" + text + ":list-style>");
        if (_onStyle is not null)
        {
            _styleStartMarker = Encoding.UTF8.GetBytes("<" + style + ":style");
            _styleEndMarker = Encoding.UTF8.GetBytes("</" + style + ":style>");
        }
    }

    private void Filter(byte[] input, int count)
    {
        if (_startMarker is null) { Emit(input, 0, count); return; }
        // the held-back bytes go first
        byte[] data = input;
        int length = count;
        if (_carry.Length > 0)
        {
            if (_joined.Length < _carry.Length + count) _joined = new byte[_carry.Length + count];
            data = _joined;
            Buffer.BlockCopy(_carry, 0, data, 0, _carry.Length);
            Buffer.BlockCopy(input, 0, data, _carry.Length, count);
            length = _carry.Length + count;
            _carry = [];
        }
        int position = 0;
        // the next match of each marker in this data (-1: none); looked up again only once the position has passed it
        int nextList = -2, nextStyle = _styleStartMarker is null ? -1 : -2;
        while (position < length)
        {
            if (_element is not null)
            {
                position = CutElement(data, position, length);
                continue;
            }
            if (nextList != -1 && nextList < position) nextList = Find(data, position, length, _startMarker);
            if (nextStyle != -1 && nextStyle < position) nextStyle = Find(data, position, length, _styleStartMarker!);
            bool isStyle = nextStyle >= 0 && (nextList < 0 || nextStyle < nextList);
            int at = isStyle ? nextStyle : nextList;
            byte[] marker = isStyle ? _styleStartMarker! : _startMarker;
            if (at < 0)
            {
                // keep a tail that could be the start of a marker
                int keep = Math.Min(Math.Max(_startMarker.Length, _styleStartMarker?.Length ?? 0), length - position);
                Emit(data, position, length - position - keep);
                _carry = data.AsSpan(length - keep, keep).ToArray();
                return;
            }
            int after = at + marker.Length;
            if (after >= length)
            {
                Emit(data, position, at - position);
                _carry = data.AsSpan(at, length - at).ToArray();
                return;
            }
            byte next = data[after];
            if (next is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'>' or (byte)'/'))
            {
                // another element whose name starts the same (text:list-style-..., style:style-...): not ours
                Emit(data, position, after - position);
                position = after;
                continue;
            }
            Emit(data, position, at - position);
            _elementBuffer.SetLength(0);
            _element = _elementBuffer;
            _isCuttingStyle = isStyle;
            _elementSearchFrom = 0;
            position = at;
        }
    }

    private static int Find(byte[] data, int position, int length, byte[] marker)
    {
        int found = data.AsSpan(position, length - position).IndexOf(marker);
        return found < 0 ? -1 : position + found;
    }

    /// <summary>Appends to the element being cut until it ends (self-closing start tag, or its end tag); returns the new position.
    /// An element that ends inside this chunk is parsed from the chunk itself: copying the rest of the chunk for every element
    /// would copy up to a chunk per list style.</summary>
    private int CutElement(byte[] data, int position, int length)
    {
        MemoryStream element = _element!;
        if (element.Length == 0)
        {
            _elementSearchFrom = 0;
            int direct = ElementEnd(data.AsSpan(position, length - position));
            if (direct >= 0)
            {
                CompleteElement(data.AsSpan(position, direct));
                return position + direct;
            }
            _elementSearchFrom = 0;                         // it continues in the next chunk: collected below
        }
        long before = element.Length;
        element.Write(data, position, length - position);
        byte[] bytes = element.GetBuffer();
        int total = (int)element.Length;
        int end = ElementEnd(bytes.AsSpan(0, total));
        if (end < 0) return length;                         // continues in the next chunk
        CompleteElement(bytes.AsSpan(0, end));
        return position + (int)(end - before);
    }

    private void CompleteElement(ReadOnlySpan<byte> bytes)
    {
        if (_isCuttingStyle)
        {
            _element = null;
            (string? family, string? styleName) = StyleKey(bytes);
            if (family is not null && styleName is not null) _onStyle!(family, styleName, new OdtStyleCascade.UnparsedStyle(bytes.ToArray(), _rootOpenTag, _rootElementName));
            return;
        }
        OdtListStyle listStyle = OdtListStyle.Parse(bytes, _textPrefix, _stylePrefix, _foPrefix, _parsedLevels, out string? name);
        if (name is not null) _onListStyle(name, listStyle);
        _element = null;
    }

    /// <summary>style:family and style:name of a style:style start tag (entity-decoded), as the XML parse would read them.</summary>
    private (string? family, string? name) StyleKey(ReadOnlySpan<byte> element)
    {
        int tagEnd = 0;
        byte quote = 0;
        for (; tagEnd < element.Length; tagEnd++)
        {
            byte b = element[tagEnd];
            if (quote != 0) { if (b == quote) quote = 0; }
            else if (b is (byte)'"' or (byte)'\'') quote = b;
            else if (b == (byte)'>') break;
        }
        ReadOnlySpan<byte> tag = element.Slice(0, tagEnd);
        return (AttributeValue(tag, _stylePrefix + ":family"), AttributeValue(tag, _stylePrefix + ":name"));
    }

    private static string? AttributeValue(ReadOnlySpan<byte> tag, string qualifiedName)
    {
        byte[] name = Encoding.UTF8.GetBytes(qualifiedName);
        int from = 0;
        while (from < tag.Length)
        {
            int at = tag.Slice(from).IndexOf(name);
            if (at < 0) return null;
            at += from;
            int position = at + name.Length;
            // a whole attribute name: preceded by whitespace, followed by optional blanks and '='
            bool isWhole = at > 0 && tag[at - 1] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n';
            while (position < tag.Length && tag[position] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') position++;
            if (!isWhole || position >= tag.Length || tag[position] != (byte)'=') { from = at + name.Length; continue; }
            position++;
            while (position < tag.Length && tag[position] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') position++;
            if (position >= tag.Length || tag[position] is not ((byte)'"' or (byte)'\'')) return null;
            byte quote = tag[position++];
            int end = tag.Slice(position).IndexOf(quote);
            if (end < 0) return null;
            string value = Encoding.UTF8.GetString(tag.Slice(position, end).ToArray());
            return value.IndexOf('&') >= 0 ? System.Net.WebUtility.HtmlDecode(value) : value;
        }
        return null;
    }

    /// <summary>The length of the element in <paramref name="bytes"/> once it is complete, or -1.</summary>
    private int ElementEnd(ReadOnlySpan<byte> bytes)
    {
        int total = bytes.Length;
        if (_elementSearchFrom == 0)
        {
            // the start tag: self-closing ends the element
            byte quote = 0;
            int i = 1;
            for (; i < total; i++)
            {
                byte b = bytes[i];
                if (quote != 0) { if (b == quote) quote = 0; }
                else if (b is (byte)'"' or (byte)'\'') quote = b;
                else if (b == (byte)'>') break;
            }
            if (i >= total) return -1;
            if (bytes[i - 1] == (byte)'/') return i + 1;
            _elementSearchFrom = i + 1;
        }
        byte[] endMarker = _isCuttingStyle ? _styleEndMarker! : _endMarker!;
        int found = bytes.Slice(_elementSearchFrom, total - _elementSearchFrom).IndexOf(endMarker);
        if (found < 0)
        {
            _elementSearchFrom = Math.Max(_elementSearchFrom, total - endMarker.Length + 1);
            return -1;
        }
        return _elementSearchFrom + found + endMarker.Length;
    }

    private void Emit(byte[] source, int offset, int count)
    {
        if (count <= 0) return;
        if (_outputEnd + count > _output.Length) Array.Resize(ref _output, Math.Max(_output.Length * 2, _outputEnd + count));
        Buffer.BlockCopy(source, offset, _output, _outputEnd, count);
        _outputEnd += count;
    }

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override void Flush() { }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing) _inner.Dispose();
        base.Dispose(disposing);
    }
}
