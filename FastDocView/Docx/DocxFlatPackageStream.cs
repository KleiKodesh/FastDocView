using System.Text;
using System.Xml;

namespace FastDocView.Docx;

/// <summary>
/// A Flat OPC file (Word's "Word XML Document", wdFormatFlatXML: one XML file whose root pkg:package holds every
/// package part as a pkg:part — XML parts inline in pkg:xmlData, binary parts base64 in pkg:binaryData; [MS-OI29500]
/// and Microsoft's "Flat OPC" format description) presented as a zip package of STORED entries, so the .docx reader
/// opens it unchanged. Nothing is copied: one streaming pass over the file at open records where each part's bytes
/// are (the format keeps no directory, so finding the parts needs that pass — the same pass a .fodt gets); then an XML
/// part's entry data is a window onto the file itself (the deflate-free stored path of every .docx reader), and a
/// binary part is decoded from its base64 only when a read reaches it (one decoded part cached). The zip headers,
/// the central directory and a [Content_Types].xml made from the parts' pkg:contentType are synthesized.
/// LibreOffice has no Flat OPC import for Writer (its OOXML detection, oox/source/core/filterdetect.cxx, only looks
/// inside zip packages), so Word is the reference.
/// </summary>
internal sealed class DocxFlatPackageStream : Stream
{
    public const string PackageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";

    private readonly Stream _file;
    private readonly byte[]? _contentTypes;          // synthesized when the package has none
    private readonly Segment[] _segments;
    private readonly long _length;
    private long _position;
    private int _decodedPart = -1;
    private byte[]? _decoded;

    private sealed class Segment
    {
        public long Start;                  // virtual offset
        public long Length;
        public byte[]? Bytes;               // synthesized (headers, directory, content types)
        public long FileOffset;             // an XML part: its bytes in the file
        public long Base64End = -1;         // a binary part: the base64 text runs FileOffset..Base64End in the file
    }

    private sealed record Part(string Name, string ContentType, long ContentStart, long ContentEnd, bool IsBinary, long DecodedLength);

    /// <summary>True when the file's root element is pkg:package (in the package namespace, whatever its prefix). Only the root start tag is read.</summary>
    public static bool IsFlatPackage(Stream file) => RootElement(file) is (PackageNamespace, "package");

    /// <summary>The namespace and local name of the root element; null when the file is not XML up to its root tag.</summary>
    internal static (string ns, string localName)? RootElement(Stream file)
    {
        file.Position = 0;
        try
        {
            using XmlReader reader = XmlReader.Create(file, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, CloseInput = false });
            return reader.MoveToContent() == XmlNodeType.Element ? (reader.NamespaceURI, reader.LocalName) : null;
        }
        catch (XmlException) { return null; }
        catch (DecoderFallbackException) { return null; }
    }

    // The last file's scanned parts, keyed by its path, length and write time: a host asks Decide and then Open for the same
    // file, and the second needs no second pass over it. One entry (bounded); the parts are immutable.
    private static readonly object LastScanGate = new();
    private static (string Key, List<Part> Parts)? _lastScan;
    private static int _scanCount;

    /// <summary>Passes made over Flat OPC files so far (the harness checks that Decide + Open make one).</summary>
    internal static int ScanCount => Volatile.Read(ref _scanCount);

    /// <summary>
    /// Scans the file's parts (one pass) and lays out the virtual zip. The stream owns <paramref name="file"/>. With
    /// <paramref name="filePath"/> (the file <paramref name="file"/> reads), the scan of an unchanged file is reused by the next open.
    /// </summary>
    public DocxFlatPackageStream(Stream file, string? filePath = null)
    {
        _file = file;
        string? key = null;
        if (filePath is not null)
        {
            var info = new FileInfo(filePath);
            key = Path.GetFullPath(filePath) + "|" + info.Length + "|" + info.LastWriteTimeUtc.Ticks;
        }
        List<Part>? scanned = null;
        lock (LastScanGate) if (key is not null && _lastScan is { } last && last.Key == key) scanned = last.Parts;
        if (scanned is null)
        {
            string prefix = PackagePrefix(file);
            scanned = DocxFlatPackageScanner.Scan(file, prefix);
            Interlocked.Increment(ref _scanCount);
            if (key is not null) lock (LastScanGate) _lastScan = (key, scanned);
        }
        var parts = new List<Part>(scanned);
        if (!parts.Any(part => part.Name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)))
        {
            byte[] types = ContentTypes(parts);
            parts.Insert(0, new Part("[Content_Types].xml", "", -1, types.Length, false, types.Length));
            _contentTypes = types;
        }
        var segments = new List<Segment>();
        var directory = new MemoryStream();
        var directoryWriter = new BinaryWriter(directory);
        long offset = 0;
        foreach (Part part in parts)
        {
            byte[] name = Encoding.UTF8.GetBytes(part.Name);
            long size = part.IsBinary ? part.DecodedLength : part.ContentEnd - part.ContentStart;
            if (part.ContentStart < 0) size = part.ContentEnd;
            if (size > uint.MaxValue - 1 || offset > uint.MaxValue - 1) throw new NotSupportedException("Flat XML packages over 4 GB are not supported.");
            var header = new MemoryStream();
            var headerWriter = new BinaryWriter(header);
            headerWriter.Write(0x04034b50u); headerWriter.Write((ushort)20); headerWriter.Write((ushort)0x0800); headerWriter.Write((ushort)0);
            headerWriter.Write(0u);                      // time, date
            headerWriter.Write(0u);                      // CRC: no reader checks it for stored entries
            headerWriter.Write((uint)size); headerWriter.Write((uint)size);
            headerWriter.Write((ushort)name.Length); headerWriter.Write((ushort)0); headerWriter.Write(name);
            Add(segments, ref offset, new Segment { Bytes = header.ToArray() });
            long localHeaderOffset = offset - header.Length;
            if (part.ContentStart < 0) Add(segments, ref offset, new Segment { Bytes = _contentTypes! });
            else if (part.IsBinary) Add(segments, ref offset, new Segment { FileOffset = part.ContentStart, Base64End = part.ContentEnd, Length = size });
            else Add(segments, ref offset, new Segment { FileOffset = part.ContentStart, Length = size });

            directoryWriter.Write(0x02014b50u); directoryWriter.Write((ushort)20); directoryWriter.Write((ushort)20); directoryWriter.Write((ushort)0x0800);
            directoryWriter.Write((ushort)0); directoryWriter.Write(0u); directoryWriter.Write(0u);
            directoryWriter.Write((uint)size); directoryWriter.Write((uint)size);
            directoryWriter.Write((ushort)name.Length); directoryWriter.Write((ushort)0); directoryWriter.Write((ushort)0);
            directoryWriter.Write((ushort)0); directoryWriter.Write((ushort)0); directoryWriter.Write(0u);
            directoryWriter.Write((uint)localHeaderOffset); directoryWriter.Write(name);
        }
        if (parts.Count > ushort.MaxValue) throw new NotSupportedException("Flat XML packages with over 65535 parts are not supported.");
        long directoryOffset = offset;
        directoryWriter.Write(0x06054b50u); directoryWriter.Write((ushort)0); directoryWriter.Write((ushort)0);
        directoryWriter.Write((ushort)parts.Count); directoryWriter.Write((ushort)parts.Count);
        directoryWriter.Write((uint)(directory.Length - 4 - 2 - 2 - 2 - 2)); directoryWriter.Write((uint)directoryOffset); directoryWriter.Write((ushort)0);
        Add(segments, ref offset, new Segment { Bytes = directory.ToArray() });
        _segments = segments.ToArray();
        _length = offset;
    }

    private static void Add(List<Segment> segments, ref long offset, Segment segment)
    {
        segment.Start = offset;
        if (segment.Bytes is not null) segment.Length = segment.Bytes.Length;
        offset += segment.Length;
        segments.Add(segment);
    }

    /// <summary>The prefix the root binds to the package namespace ("" for a default namespace); only UTF-8 files are read.</summary>
    private static string PackagePrefix(Stream file)
    {
        file.Position = 0;
        RequireUtf8(file);
        file.Position = 0;
        using XmlReader reader = XmlReader.Create(file, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, CloseInput = false });
        reader.MoveToContent();
        return reader.Prefix;
    }

    /// <summary>Throws for a file not in UTF-8 (a UTF-16 byte order mark, or an XML declaration naming another encoding): part bytes are handed on as they are.</summary>
    private static void RequireUtf8(Stream file)
    {
        file.Position = 0;
        var head = new byte[256];
        int count = file.Read(head, 0, head.Length);
        if (count >= 2 && ((head[0] == 0xFF && head[1] == 0xFE) || (head[0] == 0xFE && head[1] == 0xFF)))
            throw new NotSupportedException("UTF-16 Flat XML packages are not supported (Word writes them in UTF-8).");
        int skip = count >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF ? 3 : 0;   // UTF-8 byte order mark
        string text = Encoding.ASCII.GetString(head, skip, count - skip);
        int declarationEnd = text.StartsWith("<?xml", StringComparison.Ordinal) ? text.IndexOf("?>", StringComparison.Ordinal) : -1;
        if (declarationEnd < 0) return;
        int at = text.IndexOf("encoding", 0, declarationEnd, StringComparison.Ordinal);
        if (at < 0) return;
        int quote = text.IndexOfAny(['"', '\''], at);
        if (quote < 0 || quote > declarationEnd) return;
        int close = text.IndexOf(text[quote], quote + 1);
        string encoding = close < 0 ? "" : text.Substring(quote + 1, close - quote - 1);
        if (!encoding.Equals("UTF-8", StringComparison.OrdinalIgnoreCase) && !encoding.Equals("UTF8", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException($"Flat XML packages in {encoding} are not supported (Word writes them in UTF-8).");
    }

    /// <summary>A [Content_Types].xml listing each part's pkg:contentType (ECMA-376 Part 2 10.1.2.3 Override elements).</summary>
    private static byte[] ContentTypes(List<Part> parts)
    {
        var text = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        foreach (Part part in parts)
            text.Append("<Override PartName=\"/").Append(System.Security.SecurityElement.Escape(part.Name))
                .Append("\" ContentType=\"").Append(System.Security.SecurityElement.Escape(part.ContentType)).Append("\"/>");
        text.Append("</Types>");
        return Encoding.UTF8.GetBytes(text.ToString());
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        if (count <= 0 || _position >= _length) return 0;
        int index = SegmentAt(_position);
        Segment segment = _segments[index];
        long within = _position - segment.Start;
        int take = (int)Math.Min(count, segment.Length - within);
        if (segment.Bytes is not null) Buffer.BlockCopy(segment.Bytes, (int)within, buffer, offset, take);
        else if (segment.Base64End >= 0)
        {
            byte[] decoded = Decoded(index, segment);
            Buffer.BlockCopy(decoded, (int)within, buffer, offset, take);
        }
        else
        {
            _file.Position = segment.FileOffset + within;
            take = _file.Read(buffer, offset, take);
            if (take <= 0) throw new EndOfStreamException("The Flat XML file ended inside a part.");
        }
        _position += take;
        return take;
    }

    private int SegmentAt(long position)
    {
        int low = 0, high = _segments.Length - 1;
        while (low < high)
        {
            int middle = (low + high + 1) / 2;
            if (_segments[middle].Start <= position) low = middle; else high = middle - 1;
        }
        return low;
    }

    /// <summary>
    /// A binary part's bytes: its base64 text decoded when first reached (the last one kept), streamed through one 64 KB
    /// buffer straight into the part's bytes: memory is the decoded part, not its text as well.
    /// </summary>
    private byte[] Decoded(int index, Segment segment)
    {
        if (_decodedPart == index) return _decoded!;
        _decoded = null;
        _decodedPart = -1;
        var decoded = new byte[segment.Length];
        var chunk = new byte[1 << 16];
        long remaining = segment.Base64End - segment.FileOffset;
        _file.Position = segment.FileOffset;
        int written = 0, quantum = 0, bits = 0, padding = 0;
        while (remaining > 0)
        {
            int read = _file.Read(chunk, 0, (int)Math.Min(chunk.Length, remaining));
            if (read <= 0) throw new EndOfStreamException("The Flat XML file ended inside a binary part.");
            remaining -= read;
            for (int i = 0; i < read; i++)
            {
                byte b = chunk[i];
                if (b <= (byte)' ') continue;                       // line breaks and indentation between the base64 lines
                int value = Base64Value(b);
                if (b == '=') { padding++; value = 0; }
                else if (value < 0 || padding > 0) throw new InvalidDataException("A Flat XML binary part is not valid base64.");
                bits = bits << 6 | value;
                if (++quantum < 4) continue;
                int bytes = 3 - padding;
                if (bytes < 1 || written + bytes > decoded.Length) throw new InvalidDataException("A Flat XML binary part's decoded length is not what its base64 promised.");
                decoded[written++] = (byte)(bits >> 16);
                if (bytes > 1) decoded[written++] = (byte)(bits >> 8);
                if (bytes > 2) decoded[written++] = (byte)bits;
                quantum = 0;
                bits = 0;
            }
        }
        if (quantum != 0) throw new InvalidDataException("A Flat XML binary part is not valid base64.");
        if (written != decoded.Length) throw new InvalidDataException("A Flat XML binary part's decoded length is not what its base64 promised.");
        _decodedPart = index;
        _decoded = decoded;
        return decoded;
    }

    /// <summary>The 6-bit value of a base64 character (RFC 4648 table 1), or -1.</summary>
    private static int Base64Value(byte b) => b switch
    {
        >= (byte)'A' and <= (byte)'Z' => b - 'A',
        >= (byte)'a' and <= (byte)'z' => b - 'a' + 26,
        >= (byte)'0' and <= (byte)'9' => b - '0' + 52,
        (byte)'+' => 62,
        (byte)'/' => 63,
        _ => -1,
    };

    public override long Seek(long offset, SeekOrigin origin)
    {
        _position = origin switch { SeekOrigin.Begin => offset, SeekOrigin.Current => _position + offset, _ => _length + offset };
        return _position;
    }

    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => false;
    public override long Length => _length;
    public override long Position { get => _position; set => _position = value; }
    public override void Flush() { }
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing) _file.Dispose();
        base.Dispose(disposing);
    }

    /// <summary>
    /// The one pass over a Flat OPC file: a byte scanner (no DOM, bounded memory) that finds each pkg:part, its
    /// pkg:name / pkg:contentType, and the byte range of its pkg:xmlData content or the base64 text of its pkg:binaryData
    /// (counting the base64 characters for the decoded length). Inside pkg:xmlData only the closing tag is looked for: a
    /// part's own markup can never hold "&lt;/pkg:xmlData" (a '&lt;' in text is escaped), except inside a comment or
    /// CDATA, which Word never writes there.
    /// </summary>
    private sealed class DocxFlatPackageScanner
    {
        private const int ChunkSize = 1 << 16;
        private readonly Stream _file;
        private readonly byte[] _chunk = new byte[ChunkSize];
        private int _chunkLength, _chunkPosition;
        private long _chunkStart;

        private DocxFlatPackageScanner(Stream file) { _file = file; _file.Position = 0; }

        private long Position => _chunkStart + _chunkPosition;

        private int Next()
        {
            if (_chunkPosition == _chunkLength)
            {
                _chunkStart += _chunkLength;
                _file.Position = _chunkStart;
                _chunkLength = _file.Read(_chunk, 0, ChunkSize);
                _chunkPosition = 0;
                if (_chunkLength <= 0) { _chunkLength = 0; return -1; }
            }
            return _chunk[_chunkPosition++];
        }

        public static List<Part> Scan(Stream file, string prefix)
        {
            var scanner = new DocxFlatPackageScanner(file);
            string qualified = prefix.Length == 0 ? "" : prefix + ":";
            byte[] xmlDataEnd = Encoding.UTF8.GetBytes("</" + qualified + "xmlData");
            byte[] binaryDataEnd = Encoding.UTF8.GetBytes("</" + qualified + "binaryData");
            var parts = new List<Part>();
            string? name = null, contentType = null;
            var tag = new StringBuilder();
            while (true)
            {
                int b = scanner.Next();
                if (b < 0) break;
                if (b != '<') continue;
                tag.Clear();
                if (!scanner.ReadTag(tag)) break;
                string text = tag.ToString();
                if (text.StartsWith("!--", StringComparison.Ordinal) || text.StartsWith("?", StringComparison.Ordinal) || text.StartsWith("!", StringComparison.Ordinal) || text.StartsWith("/", StringComparison.Ordinal)) continue;
                string element = ElementName(text);
                if (element == qualified + "part")
                {
                    name = AttributeValue(text, qualified + "name");
                    contentType = AttributeValue(text, qualified + "contentType") ?? "";
                }
                else if (element == qualified + "xmlData" || element == qualified + "binaryData")
                {
                    bool isBinary = element == qualified + "binaryData";
                    if (name is null) throw new InvalidDataException($"Flat XML: a {element} outside a named part.");
                    long start = scanner.Position;
                    if (text.EndsWith("/", StringComparison.Ordinal)) { parts.Add(new Part(PartName(name), contentType!, start, start, isBinary, 0)); name = null; continue; }
                    (long end, long base64Characters, int padding) = scanner.FindClosingTag(isBinary ? binaryDataEnd : xmlDataEnd, isBinary);
                    if (end < 0) throw new InvalidDataException($"Flat XML: the part {name} has no closing {element} tag.");
                    long decoded = isBinary ? base64Characters / 4 * 3 - padding : end - start;
                    parts.Add(new Part(PartName(name), contentType!, start, end, isBinary, decoded));
                    name = null;
                }
            }
            if (parts.Count == 0) throw new InvalidDataException("Flat XML: the package holds no parts.");
            return parts;
        }

        private static string PartName(string name) => Uri.UnescapeDataString(name.TrimStart('/'));

        /// <summary>Reads a tag after its '&lt;' up to its closing '&gt;' (quotes respected; a comment up to "--&gt;").</summary>
        private bool ReadTag(StringBuilder tag)
        {
            byte quote = 0;
            var bytes = new List<byte>(64);
            while (true)
            {
                int b = Next();
                if (b < 0) return false;
                bool isComment = bytes.Count >= 3 && bytes[0] == '!' && bytes[1] == '-' && bytes[2] == '-';
                if (quote != 0) { if (b == quote) quote = 0; }
                else if (!isComment && (b == '"' || b == '\'')) quote = (byte)b;
                else if (b == '>')
                {
                    if (bytes.Count >= 3 && bytes[0] == '!' && bytes[1] == '-' && bytes[2] == '-'
                        && !(bytes.Count >= 5 && bytes[bytes.Count - 1] == '-' && bytes[bytes.Count - 2] == '-')) { bytes.Add((byte)b); continue; }
                    break;
                }
                bytes.Add((byte)b);
            }
            tag.Append(Encoding.UTF8.GetString(bytes.ToArray()));
            return true;
        }

        /// <summary>The position of the closing tag's '&lt;' (and, for base64, its character count and '=' padding), or -1.</summary>
        private (long end, long base64Characters, int padding) FindClosingTag(byte[] closing, bool countBase64)
        {
            long characters = 0;
            int padding = 0;
            int matched = 0;
            long matchStart = -1;
            while (true)
            {
                int b = Next();
                if (b < 0) return (-1, 0, 0);
                if (b == closing[matched])
                {
                    if (matched == 0) matchStart = Position - 1;
                    if (++matched == closing.Length)
                    {
                        int after = Next();
                        if (after is '>' or ' ' or '\t' or '\r' or '\n')
                        {
                            if (after != '>') while (Next() is int c && c >= 0 && c != '>') { }
                            return (matchStart, characters, padding);
                        }
                        matched = 0;
                    }
                    continue;
                }
                matched = b == closing[0] ? 1 : 0;
                if (matched == 1) { matchStart = Position - 1; continue; }
                if (countBase64 && b > ' ')
                {
                    characters++;
                    if (b == '=') padding++;
                }
            }
        }

        private static string ElementName(string tag)
        {
            int end = 0;
            while (end < tag.Length && !char.IsWhiteSpace(tag[end]) && tag[end] != '/') end++;
            return tag.Substring(0, end);
        }

        private static string? AttributeValue(string tag, string attribute)
        {
            int at = 0;
            while ((at = tag.IndexOf(attribute, at, StringComparison.Ordinal)) >= 0)
            {
                int equals = at + attribute.Length;
                while (equals < tag.Length && char.IsWhiteSpace(tag[equals])) equals++;
                if (at > 0 && char.IsWhiteSpace(tag[at - 1]) && equals < tag.Length && tag[equals] == '=')
                {
                    int quote = equals + 1;
                    while (quote < tag.Length && char.IsWhiteSpace(tag[quote])) quote++;
                    if (quote >= tag.Length) return null;
                    char q = tag[quote];
                    int close = tag.IndexOf(q, quote + 1);
                    if (close < 0) return null;
                    return System.Net.WebUtility.HtmlDecode(tag.Substring(quote + 1, close - quote - 1));
                }
                at += attribute.Length;
            }
            return null;
        }
    }
}
