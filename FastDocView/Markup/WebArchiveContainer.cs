using System.Text;

namespace FastDocView.Markup;

/// <summary>Where a web archive's HTML part body starts and how it is encoded.</summary>
public sealed record WebArchiveHtmlPart(long BodyStart, MimeTransferEncoding TransferEncoding, string? Charset, string? Boundary);

/// <summary>
/// Reads the MIME structure of a Word web archive (.mht / .mhtml / .doc saved as "Single File Web
/// Page") — headers only, never a whole part. RFC 2045/2046/2557: top-level headers (folded lines
/// unfolded), multipart/related with a boundary parameter (token or quoted-string), parts separated by
/// "--boundary" lines (optional trailing whitespace; "--boundary--" closes), each part headers + blank
/// line + body; a part may itself be multipart (nested, its own boundary). The first text/html part, depth
/// first, is the document (Word writes it first); a single-part archive whose own Content-Type is text/html is
/// the document itself.
/// </summary>
public static class WebArchiveContainer
{
    public static WebArchiveHtmlPart LocateHtmlPart(Stream file)
    {
        file.Position = 0;
        var lines = new MarkupLineReader(file);
        Dictionary<string, string> top = ReadHeaders(lines);
        string topType = top.TryGetValue("content-type", out string? t) ? t : "";
        if (MediaType(topType) == "text/html")
            return new WebArchiveHtmlPart(lines.Position, TransferEncodingOf(top), Parameter(topType, "charset"), null);
        string? boundary = Parameter(topType, "boundary");
        if (!MediaType(topType).StartsWith("multipart/", StringComparison.Ordinal) || string.IsNullOrEmpty(boundary))
            throw new InvalidDataException("Web archive: no multipart boundary and no text/html content.");
        // RFC 2046 §5.1.1: a body part may itself be multipart with its own boundary (Internet Explorer saves
        // multipart/related whose first part is a multipart/alternative holding text/plain + text/html). The open
        // multiparts' boundaries are a stack; a delimiter of an outer one also ends every inner one.
        var boundaries = new List<string> { boundary! };
        while (boundaries.Count > 0 && lines.ReadLine(out ReadOnlySpan<byte> line, out _))
        {
            int level = boundaries.Count - 1;
            bool isClose = false;
            while (level >= 0 && !IsBoundaryLine(line, boundaries[level], out isClose)) level--;
            if (level < 0) continue;
            boundaries.RemoveRange(level + 1, boundaries.Count - level - 1);
            if (isClose) { boundaries.RemoveAt(level); continue; }
            Dictionary<string, string> part = ReadHeaders(lines);
            string type = part.TryGetValue("content-type", out string? pt) ? pt : "text/plain";
            string mediaType = MediaType(type);
            if (mediaType == "text/html")
                return new WebArchiveHtmlPart(lines.Position, TransferEncodingOf(part), Parameter(type, "charset"), boundaries[level]);
            string? inner = Parameter(type, "boundary");
            if (mediaType.StartsWith("multipart/", StringComparison.Ordinal) && !string.IsNullOrEmpty(inner)) boundaries.Add(inner!);
        }
        throw new InvalidDataException("Web archive: no text/html part.");
    }

    /// <summary>
    /// The decoded body of the first leaf part whose headers satisfy <paramref name="matches"/> (headers lower-cased), with its
    /// Content-Type and Content-Location; null when no part matches. Only part headers are read until the match, then that one
    /// body - never another part or the whole archive (a picture is read on request, like every other format's).
    /// </summary>
    public static byte[]? ReadPart(Stream file, Func<IReadOnlyDictionary<string, string>, bool> matches, out string contentType, out string location)
    {
        contentType = location = "";
        file.Position = 0;
        var lines = new MarkupLineReader(file);
        Dictionary<string, string> top = ReadHeaders(lines);
        string topType = top.TryGetValue("content-type", out string? t) ? t : "";
        string? boundary = Parameter(topType, "boundary");
        if (!MediaType(topType).StartsWith("multipart/", StringComparison.Ordinal) || string.IsNullOrEmpty(boundary)) return null;
        var boundaries = new List<string> { boundary! };
        while (boundaries.Count > 0 && lines.ReadLine(out ReadOnlySpan<byte> line, out _))
        {
            int level = boundaries.Count - 1;
            bool isClose = false;
            while (level >= 0 && !IsBoundaryLine(line, boundaries[level], out isClose)) level--;
            if (level < 0) continue;
            boundaries.RemoveRange(level + 1, boundaries.Count - level - 1);
            if (isClose) { boundaries.RemoveAt(level); continue; }
            Dictionary<string, string> part = ReadHeaders(lines);
            string type = part.TryGetValue("content-type", out string? pt) ? pt : "text/plain";
            string? inner = Parameter(type, "boundary");
            if (MediaType(type).StartsWith("multipart/", StringComparison.Ordinal) && !string.IsNullOrEmpty(inner)) { boundaries.Add(inner!); continue; }
            if (!matches(part)) continue;
            contentType = MediaType(type);
            location = part.TryGetValue("content-location", out string? l) ? l : "";
            var decoder = new MimeTransferDecoder(TransferEncodingOf(part));
            var body = new MemoryStream();
            var decoded = new byte[0];
            while (lines.ReadLine(out ReadOnlySpan<byte> bodyLine, out long offset))
            {
                bool ends = false;
                foreach (string open in boundaries) if (IsBoundaryLine(bodyLine, open, out _)) { ends = true; break; }
                if (ends) break;
                byte[] encoded = bodyLine.ToArray();
                if (decoded.Length < encoded.Length + 4) decoded = new byte[encoded.Length + 4];
                int produced = 0;
                decoder.Feed(encoded, 0, encoded.Length, offset, decoded, null, ref produced);
                body.Write(decoded, 0, produced);
            }
            return body.ToArray();
        }
        return null;
    }

    /// <summary>"--boundary" [whitespace] (or the "--boundary--" close-delimiter), RFC 2046 §5.1.1.</summary>
    public static bool IsBoundaryLine(ReadOnlySpan<byte> line, string boundary, out bool isClose)
    {
        isClose = false;
        if (line.Length < boundary.Length + 2 || line[0] != '-' || line[1] != '-') return false;
        for (int i = 0; i < boundary.Length; i++) if (line[2 + i] != boundary[i]) return false;
        int k = 2 + boundary.Length;
        if (k + 1 < line.Length && line[k] == '-' && line[k + 1] == '-') { isClose = true; k += 2; }
        for (; k < line.Length; k++) if (line[k] is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n')) return false;
        return true;
    }

    private static Dictionary<string, string> ReadHeaders(MarkupLineReader lines)
    {
        var headers = new Dictionary<string, string>(StringComparer.Ordinal);
        string? name = null;
        var value = new StringBuilder();
        while (lines.ReadLine(out ReadOnlySpan<byte> raw, out _))
        {
            string line = Encoding.ASCII.GetString(raw.ToArray()).TrimEnd('\r', '\n');
            if (line.Length == 0) break;                                   // end of headers
            if ((line[0] == ' ' || line[0] == '\t') && name is not null) { value.Append(' ').Append(line.Trim()); continue; }   // folded
            if (name is not null && !headers.ContainsKey(name)) headers[name] = value.ToString();
            int colon = line.IndexOf(':');
            if (colon <= 0) { name = null; continue; }
            name = line.Substring(0, colon).Trim().ToLowerInvariant();
            value.Clear().Append(line.Substring(colon + 1).Trim());
        }
        if (name is not null && !headers.ContainsKey(name)) headers[name] = value.ToString();
        return headers;
    }

    private static MimeTransferEncoding TransferEncodingOf(Dictionary<string, string> headers)
    {
        string value = headers.TryGetValue("content-transfer-encoding", out string? v) ? v.Trim().ToLowerInvariant() : "";
        return value switch { "quoted-printable" => MimeTransferEncoding.QuotedPrintable, "base64" => MimeTransferEncoding.Base64, _ => MimeTransferEncoding.Identity };
    }

    private static string MediaType(string contentType)
    {
        int semicolon = contentType.IndexOf(';');
        return (semicolon < 0 ? contentType : contentType.Substring(0, semicolon)).Trim().ToLowerInvariant();
    }

    /// <summary>A Content-Type parameter (value := token / quoted-string, RFC 2045 §5.1; quotes not part of the value).</summary>
    public static string? Parameter(string contentType, string parameterName)
    {
        int i = contentType.IndexOf(';');
        while (i >= 0 && i < contentType.Length)
        {
            i++;
            while (i < contentType.Length && char.IsWhiteSpace(contentType[i])) i++;
            int equals = contentType.IndexOf('=', i);
            if (equals < 0) return null;
            string name = contentType.Substring(i, equals - i).Trim();
            i = equals + 1;
            while (i < contentType.Length && char.IsWhiteSpace(contentType[i])) i++;
            string value;
            if (i < contentType.Length && contentType[i] == '"')
            {
                int close = contentType.IndexOf('"', i + 1);
                if (close < 0) close = contentType.Length;
                value = contentType.Substring(i + 1, close - i - 1);
                i = contentType.IndexOf(';', Math.Min(close, contentType.Length));
            }
            else
            {
                int end = contentType.IndexOf(';', i);
                value = contentType.Substring(i, (end < 0 ? contentType.Length : end) - i).Trim();
                i = end;
            }
            if (name.Equals(parameterName, StringComparison.OrdinalIgnoreCase)) return value;
        }
        return null;
    }
}

/// <summary>Buffered line reader over a stream that tells each line's file offset (lines include their terminator).</summary>
public sealed class MarkupLineReader
{
    private readonly Stream _stream;
    private byte[] _buffer = new byte[1 << 16];
    private int _start, _end;
    private long _bufferFileOffset;

    public MarkupLineReader(Stream stream)
    {
        _stream = stream;
        _bufferFileOffset = stream.Position;
    }

    /// <summary>File offset of the next unread byte.</summary>
    public long Position => _bufferFileOffset + _start;

    /// <summary>Next line up to and including '\n' (or end of file); lines longer than the buffer come in pieces.</summary>
    public bool ReadLine(out ReadOnlySpan<byte> line, out long lineOffset)
    {
        while (true)
        {
            int newline = _buffer.AsSpan(_start, _end - _start).IndexOf((byte)'\n');
            if (newline >= 0 || (_end - _start == _buffer.Length) || !Fill())
            {
                int length = newline >= 0 ? newline + 1 : _end - _start;
                lineOffset = Position;
                line = _buffer.AsSpan(_start, length);
                _start += length;
                return length > 0;
            }
        }
    }

    private bool Fill()
    {
        if (_start > 0)
        {
            Buffer.BlockCopy(_buffer, _start, _buffer, 0, _end - _start);
            _bufferFileOffset += _start;
            _end -= _start;
            _start = 0;
        }
        _stream.Position = _bufferFileOffset + _end;          // the stream may be shared (windows read between two on-demand index steps)
        int read = _stream.Read(_buffer, _end, _buffer.Length - _end);
        if (read <= 0) return false;
        _end += read;
        return true;
    }
}
