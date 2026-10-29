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
    private readonly byte[] _chunk = new byte[ChunkSize];
    private byte[] _output = new byte[ChunkSize * 2];
    private int _outputStart, _outputEnd;
    private bool _isInnerDone;

    // the root start tag (for the prefixes), then the filter
    private readonly MemoryStream _head = new();
    private bool _isHeadDone;
    private byte[]? _startMarker, _endMarker;        // "<text:list-style", "</text:list-style>"
    private string _textPrefix = "text", _stylePrefix = "style", _foPrefix = "fo";
    private byte[] _carry = [];                      // bytes held back: a possible marker start at a chunk's end
    private MemoryStream? _element;                  // the list-style element being cut (null: none)
    private readonly MemoryStream _elementBuffer = new();    // reused for every element
    private byte[] _joined = new byte[ChunkSize * 2];       // held-back bytes + the next chunk, reused
    private int _elementSearchFrom;
    private readonly OdtListStyle.ParsedLevels _parsedLevels = new();

    public OdtStylesListStyleFilterStream(Stream inner, Action<string, OdtListStyle> onListStyle)
    {
        _inner = inner;
        _onListStyle = onListStyle;
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
        LearnPrefixes(Encoding.UTF8.GetString(head, 0, rootEnd));
        Emit(head, 0, rootEnd);
        byte[] rest = new byte[length - rootEnd];
        Buffer.BlockCopy(head, rootEnd, rest, 0, rest.Length);
        Filter(rest, rest.Length);
    }

    /// <summary>The offset just after the root element's start tag ('&gt;' outside quotes), or -1 when it has not closed yet.</summary>
    private static int RootStartTagEnd(byte[] bytes, int length)
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

    private void LearnPrefixes(string rootTag)
    {
        string? Prefix(string namespaceUri)
        {
            int at = rootTag.IndexOf("=\"" + namespaceUri + "\"", StringComparison.Ordinal);
            if (at < 0) at = rootTag.IndexOf("='" + namespaceUri + "'", StringComparison.Ordinal);
            if (at < 0) return null;
            int start = at;
            while (start > 0 && !char.IsWhiteSpace(rootTag[start - 1])) start--;
            string attribute = rootTag.Substring(start, at - start);
            return attribute.StartsWith("xmlns:", StringComparison.Ordinal) ? attribute.Substring(6) : null;
        }
        string? text = Prefix(OdtStyleCascade.Text.NamespaceName), style = Prefix(OdtStyleCascade.Style.NamespaceName), fo = Prefix(OdtStyleCascade.Fo.NamespaceName);
        if (text is null || style is null) return;            // no filtering: the XmlReader path reads the list styles
        _textPrefix = text;
        _stylePrefix = style;
        _foPrefix = fo ?? "fo";
        _startMarker = Encoding.UTF8.GetBytes("<" + text + ":list-style");
        _endMarker = Encoding.UTF8.GetBytes("</" + text + ":list-style>");
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
        while (position < length)
        {
            if (_element is not null)
            {
                position = CutElement(data, position, length);
                continue;
            }
            int found = data.AsSpan(position, length - position).IndexOf(_startMarker);
            if (found < 0)
            {
                // keep a tail that could be the start of a marker
                int keep = Math.Min(_startMarker.Length, length - position);
                Emit(data, position, length - position - keep);
                _carry = data.AsSpan(length - keep, keep).ToArray();
                return;
            }
            int at = position + found;
            int after = at + _startMarker.Length;
            if (after >= length)
            {
                Emit(data, position, at - position);
                _carry = data.AsSpan(at, length - at).ToArray();
                return;
            }
            byte next = data[after];
            if (next is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'>' or (byte)'/'))
            {
                // another element whose name starts the same (text:list-style-...): not ours
                Emit(data, position, after - position);
                position = after;
                continue;
            }
            Emit(data, position, at - position);
            _elementBuffer.SetLength(0);
            _element = _elementBuffer;
            _elementSearchFrom = 0;
            position = at;
        }
    }

    /// <summary>Appends to the element being cut until it ends (self-closing start tag, or its end tag); returns the new position.</summary>
    private int CutElement(byte[] data, int position, int length)
    {
        MemoryStream element = _element!;
        long before = element.Length;
        element.Write(data, position, length - position);
        byte[] bytes = element.GetBuffer();
        int total = (int)element.Length;
        int end = ElementEnd(bytes, total);
        if (end < 0) return length;                         // continues in the next chunk
        OdtListStyle listStyle = OdtListStyle.Parse(bytes.AsSpan(0, end), _textPrefix, _stylePrefix, _foPrefix, _parsedLevels, out string? name);
        if (name is not null) _onListStyle(name, listStyle);
        _element = null;
        return position + (int)(end - before);
    }

    /// <summary>The length of the element in <paramref name="bytes"/> once it is complete, or -1.</summary>
    private int ElementEnd(byte[] bytes, int total)
    {
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
        int found = bytes.AsSpan(_elementSearchFrom, total - _elementSearchFrom).IndexOf(_endMarker);
        if (found < 0)
        {
            _elementSearchFrom = Math.Max(_elementSearchFrom, total - _endMarker!.Length + 1);
            return -1;
        }
        return _elementSearchFrom + found + _endMarker!.Length;
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
