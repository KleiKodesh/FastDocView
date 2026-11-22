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
internal sealed class DocxFlatPackageStream : DocxVirtualPackageStream
{
    public const string PackageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";

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
    public DocxFlatPackageStream(Stream file, string? filePath = null) : base(file)
    {
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
        var parts = new List<VirtualPart>(scanned.Count + 1);
        if (!scanned.Any(part => part.Name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)))
            parts.Add(VirtualPart.Synthesized("[Content_Types].xml", ContentTypes(scanned)));
        foreach (Part part in scanned)
            parts.Add(part.IsBinary ? VirtualPart.Base64(part.Name, part.ContentStart, part.ContentEnd, part.DecodedLength)
                                    : VirtualPart.FileRange(part.Name, part.ContentStart, part.ContentEnd));
        Layout(parts);
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
    internal static void RequireUtf8(Stream file, string formatName = "Flat XML packages")
    {
        file.Position = 0;
        var head = new byte[256];
        int count = file.Read(head, 0, head.Length);
        if (count >= 2 && ((head[0] == 0xFF && head[1] == 0xFE) || (head[0] == 0xFE && head[1] == 0xFF)))
            throw new NotSupportedException($"UTF-16 {formatName} are not supported (Word writes them in UTF-8).");
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
            throw new NotSupportedException($"{formatName} in {encoding} are not supported (Word writes them in UTF-8).");
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
        /// <remarks>A part's data is skipped a chunk at a time with a vectorised search for the closing tag (a body part is
        /// most of the file: byte-by-byte it was most of a first open); only base64 data is looked at byte by byte, to count it.</remarks>
        private (long end, long base64Characters, int padding) FindClosingTag(byte[] closing, bool countBase64)
        {
            long characters = 0;
            int padding = 0;
            void Count(ReadOnlySpan<byte> data)
            {
                if (!countBase64) return;
                foreach (byte b in data)
                    if (b > ' ')
                    {
                        characters++;
                        if (b == '=') padding++;
                    }
            }
            while (true)
            {
                // the closing tag and the byte after it must fit in what is buffered
                if (_chunkLength - _chunkPosition <= closing.Length && !Refill())
                {
                    Count(_chunk.AsSpan(_chunkPosition, _chunkLength - _chunkPosition));
                    return (-1, 0, 0);
                }
                ReadOnlySpan<byte> data = _chunk.AsSpan(_chunkPosition, _chunkLength - _chunkPosition);
                int found = data.IndexOf(closing);
                if (found < 0 || found + closing.Length >= data.Length)
                {
                    // no complete match here: everything before a possible match at the end is data
                    int safe = found >= 0 ? found : data.Length - closing.Length;
                    Count(data.Slice(0, safe));
                    _chunkPosition += safe;
                    if (!Refill()) { Count(_chunk.AsSpan(_chunkPosition, _chunkLength - _chunkPosition)); return (-1, 0, 0); }
                    continue;
                }
                Count(data.Slice(0, found));
                long matchStart = Position + found;
                byte after = data[found + closing.Length];
                _chunkPosition += found + closing.Length + 1;
                if (after is (byte)'>' or (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n')
                {
                    if (after != '>') while (Next() is int c && c >= 0 && c != '>') { }
                    return (matchStart, characters, padding);
                }
                // a longer name that starts the same: not the closing tag, go on after its '<'
                _chunkPosition -= closing.Length;
            }
        }

        /// <summary>Moves the unread bytes to the buffer's start and reads more after them; false when nothing more was read.</summary>
        private bool Refill()
        {
            int unread = _chunkLength - _chunkPosition;
            Buffer.BlockCopy(_chunk, _chunkPosition, _chunk, 0, unread);
            _chunkStart += _chunkPosition;
            _chunkPosition = 0;
            _file.Position = _chunkStart + unread;
            int read = _file.Read(_chunk, unread, ChunkSize - unread);
            _chunkLength = unread + Math.Max(0, read);
            return read > 0;
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
