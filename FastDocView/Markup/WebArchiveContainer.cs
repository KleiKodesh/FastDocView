using System.Text;

namespace FastDocView.Markup;

/// <summary>Where a web archive's HTML part body starts and how it is encoded.</summary>
public sealed record WebArchiveHtmlPart(long BodyStart, MimeTransferEncoding TransferEncoding, string? Charset, string? Boundary);

/// <summary>
/// Reads the MIME structure of a Word web archive (.mht / .mhtml / .doc saved as "Single File Web
/// Page") — headers only, never a whole part. RFC 2045/2046/2557: top-level headers (folded lines
/// unfolded), multipart/related with a boundary parameter (token or quoted-string), parts separated by
/// "--boundary" lines (optional trailing whitespace; "--boundary--" closes), each part headers + blank
/// line + body. The first text/html part is the document (Word writes it first); a single-part archive
/// whose own Content-Type is text/html is the document itself.
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
        while (lines.ReadLine(out ReadOnlySpan<byte> line, out _))
        {
            if (!IsBoundaryLine(line, boundary!, out bool isClose)) continue;
            if (isClose) break;
            Dictionary<string, string> part = ReadHeaders(lines);
            string type = part.TryGetValue("content-type", out string? pt) ? pt : "text/plain";
            if (MediaType(type) == "text/html")
                return new WebArchiveHtmlPart(lines.Position, TransferEncodingOf(part), Parameter(type, "charset"), boundary);
        }
        throw new InvalidDataException("Web archive: no text/html part.");
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
        int read = _stream.Read(_buffer, _end, _buffer.Length - _end);
        if (read <= 0) return false;
        _end += read;
        return true;
    }
}
