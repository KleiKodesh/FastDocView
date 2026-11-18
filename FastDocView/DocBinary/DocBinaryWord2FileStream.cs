namespace FastDocView.DocBinary;

/// <summary>
/// A Word for Windows 1.x / 2.0 file as its WordDocument stream: those files are no compound file, the FIB, text and tables
/// all lie in the file itself (LibreOffice ww8par.cxx reads them from the one stream). Each view keeps its own position over
/// the shared file, as the compound file's streams do, so the text, table and data readers never move each other; disposing
/// a view leaves the file open.
/// </summary>
public sealed class DocBinaryWord2FileStream : Stream
{
    private readonly Stream _file;
    private long _position;

    public DocBinaryWord2FileStream(Stream file) => _file = file;

    /// <summary>True when the stream starts with a Word 1.x / 2.0 FIB identifier (ww8scan.cxx GetFIBVersion: 0xA59B, 0xA59C, 0xA5DB).</summary>
    public static bool StartsWithWord2Identifier(Stream file)
    {
        file.Position = 0;
        int low = file.ReadByte(), high = file.ReadByte();
        file.Position = 0;
        return low >= 0 && high >= 0 && (low | high << 8) is 0xA59B or 0xA59C or 0xA5DB;
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        _file.Position = _position;
        int read = _file.Read(buffer, offset, (int)Math.Max(0, Math.Min(count, Length - _position)));
        if (read > 0) _position += read;
        return Math.Max(0, read);
    }

    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => false;
    public override long Length => _file.Length;
    public override long Position { get => _position; set => _position = Math.Max(0, Math.Min(Length, value)); }
    public override long Seek(long offset, SeekOrigin origin)
    {
        Position = origin switch { SeekOrigin.Begin => offset, SeekOrigin.Current => _position + offset, _ => Length + offset };
        return _position;
    }
    public override void Flush() { }
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}
