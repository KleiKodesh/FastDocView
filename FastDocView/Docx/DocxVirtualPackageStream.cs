using System.Text;

namespace FastDocView.Docx;

/// <summary>
/// A zip package of STORED entries that exists only as a layout over one source file, so the .docx reader opens a
/// one-file XML format unchanged. Each part is synthesized bytes (headers, the central directory, small parts made at
/// open), a window onto a byte range of the file, base64 text in the file decoded when a read reaches it (one decoded
/// part cached), or a derived part: an entry with <see cref="DerivedPartMethod"/> and no data, whose bytes a subclass
/// produces on demand through <see cref="IDocxDerivedPartSource"/> (a Word 2003 XML body translated block by block).
/// Used by <see cref="DocxFlatPackageStream"/> (Flat OPC) and <see cref="Word2003PackageStream"/> (Word 2003 XML).
/// </summary>
internal abstract class DocxVirtualPackageStream : Stream
{
    /// <summary>The zip compression method given to a derived part (outside the APPNOTE.TXT 4.4.5 range); only this package reads it.</summary>
    public const ushort DerivedPartMethod = 0xF003;

    protected Stream SourceFile { get; }
    private Segment[] _segments = [];
    private long _length;
    private long _position;
    private int _decodedPart = -1;
    private byte[]? _decoded;

    private sealed class Segment
    {
        public long Start;                  // virtual offset
        public long Length;
        public byte[]? Bytes;               // synthesized
        public long FileOffset;             // a file range: its bytes in the file
        public long Base64End = -1;         // a base64 part: the base64 text runs FileOffset..Base64End in the file
    }

    /// <summary>One entry of the virtual zip.</summary>
    protected sealed class VirtualPart
    {
        public string Name { get; private init; } = "";
        public byte[]? Bytes { get; private init; }
        public long FileStart { get; private init; }
        public long FileEnd { get; private init; }
        public bool IsBase64 { get; private init; }
        public bool IsDerived { get; private init; }
        public long Size { get; private init; }

        public static VirtualPart Synthesized(string name, byte[] bytes) => new() { Name = name, Bytes = bytes, Size = bytes.Length };
        public static VirtualPart FileRange(string name, long start, long end) => new() { Name = name, FileStart = start, FileEnd = end, Size = end - start };
        public static VirtualPart Base64(string name, long textStart, long textEnd, long decodedLength)
            => new() { Name = name, FileStart = textStart, FileEnd = textEnd, IsBase64 = true, Size = decodedLength };
        /// <summary>A part produced on demand; <paramref name="estimatedSize"/> stands for its size where the reader weighs parts (index policy, progress).</summary>
        public static VirtualPart Derived(string name, long estimatedSize) => new() { Name = name, IsDerived = true, Size = estimatedSize };
    }

    protected DocxVirtualPackageStream(Stream sourceFile) => SourceFile = sourceFile;

    /// <summary>Lays out the local headers, the parts' data and the central directory.</summary>
    protected void Layout(IReadOnlyList<VirtualPart> parts)
    {
        if (parts.Count > ushort.MaxValue) throw new NotSupportedException("Packages with over 65535 parts are not supported.");
        var segments = new List<Segment>();
        var directory = new MemoryStream();
        var directoryWriter = new BinaryWriter(directory);
        long offset = 0;
        foreach (VirtualPart part in parts)
        {
            byte[] name = Encoding.UTF8.GetBytes(part.Name);
            long size = part.Size;
            long stored = part.IsDerived ? 0 : size;
            ushort method = part.IsDerived ? DerivedPartMethod : (ushort)0;
            if (size > uint.MaxValue - 1 || offset > uint.MaxValue - 1) throw new NotSupportedException("XML packages over 4 GB are not supported.");
            var header = new MemoryStream();
            var headerWriter = new BinaryWriter(header);
            headerWriter.Write(0x04034b50u); headerWriter.Write((ushort)20); headerWriter.Write((ushort)0x0800); headerWriter.Write(method);
            headerWriter.Write(0u);                      // time, date
            headerWriter.Write(0u);                      // CRC: no reader checks it for stored entries
            headerWriter.Write((uint)stored); headerWriter.Write((uint)size);
            headerWriter.Write((ushort)name.Length); headerWriter.Write((ushort)0); headerWriter.Write(name);
            Add(segments, ref offset, new Segment { Bytes = header.ToArray() });
            long localHeaderOffset = offset - header.Length;
            if (part.Bytes is not null) Add(segments, ref offset, new Segment { Bytes = part.Bytes });
            else if (part.IsBase64) Add(segments, ref offset, new Segment { FileOffset = part.FileStart, Base64End = part.FileEnd, Length = size });
            else if (!part.IsDerived) Add(segments, ref offset, new Segment { FileOffset = part.FileStart, Length = size });

            directoryWriter.Write(0x02014b50u); directoryWriter.Write((ushort)20); directoryWriter.Write((ushort)20); directoryWriter.Write((ushort)0x0800);
            directoryWriter.Write(method); directoryWriter.Write(0u); directoryWriter.Write(0u);
            directoryWriter.Write((uint)stored); directoryWriter.Write((uint)size);
            directoryWriter.Write((ushort)name.Length); directoryWriter.Write((ushort)0); directoryWriter.Write((ushort)0);
            directoryWriter.Write((ushort)0); directoryWriter.Write((ushort)0); directoryWriter.Write(0u);
            directoryWriter.Write((uint)localHeaderOffset); directoryWriter.Write(name);
        }
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
        if (segment.Length == 0) return;          // an empty part: nothing to read (keeps segment starts strictly increasing)
        offset += segment.Length;
        segments.Add(segment);
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
            SourceFile.Position = segment.FileOffset + within;
            take = SourceFile.Read(buffer, offset, take);
            if (take <= 0) throw new EndOfStreamException("The XML file ended inside a part.");
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
    /// A base64 part's bytes: its text decoded when first reached (the last one kept), streamed through one 64 KB
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
        SourceFile.Position = segment.FileOffset;
        int written = 0, quantum = 0, bits = 0, padding = 0;
        while (remaining > 0)
        {
            int read = SourceFile.Read(chunk, 0, (int)Math.Min(chunk.Length, remaining));
            if (read <= 0) throw new EndOfStreamException("The XML file ended inside a binary part.");
            remaining -= read;
            for (int i = 0; i < read; i++)
            {
                byte b = chunk[i];
                if (b <= (byte)' ') continue;                       // line breaks and indentation between the base64 lines
                int value = Base64Value(b);
                if (b == '=') { padding++; value = 0; }
                else if (value < 0 || padding > 0) throw new InvalidDataException("A binary part is not valid base64.");
                bits = bits << 6 | value;
                if (++quantum < 4) continue;
                int bytes = 3 - padding;
                if (bytes < 1 || written + bytes > decoded.Length) throw new InvalidDataException("A binary part's decoded length is not what its base64 promised.");
                decoded[written++] = (byte)(bits >> 16);
                if (bytes > 1) decoded[written++] = (byte)(bits >> 8);
                if (bytes > 2) decoded[written++] = (byte)bits;
                quantum = 0;
                bits = 0;
            }
        }
        if (quantum != 0) throw new InvalidDataException("A binary part is not valid base64.");
        if (written != decoded.Length) throw new InvalidDataException("A binary part's decoded length is not what its base64 promised.");
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
        if (disposing) SourceFile.Dispose();
        base.Dispose(disposing);
    }
}

/// <summary>
/// A package whose some parts are derived (<see cref="DocxVirtualPackageStream.DerivedPartMethod"/>): the reader's
/// sequential index pass and its random-access views go through here instead of the zip entry's bytes.
/// </summary>
internal interface IDocxDerivedPartSource
{
    /// <summary>A pass over the part from its start, adding checkpoints to <paramref name="checkpoints"/> as it goes (about every <paramref name="checkpointSpanBytes"/> of output, or the span the part's own form needs).</summary>
    IDocxDerivedPartPass OpenPass(SeekableDeflate.ZipEntryLocation entry, List<SeekableDeflate.DeflateCheckpoint> checkpoints, long checkpointSpanBytes);

    /// <summary>A seekable view of the part over the checkpoints a pass recorded (the list may still be growing).</summary>
    Stream OpenView(SeekableDeflate.ZipEntryLocation entry, IReadOnlyList<SeekableDeflate.DeflateCheckpoint> checkpoints);
}

/// <summary>The pull form of a derived part's pass: fills the buffer with the next bytes, 0 at the end.</summary>
internal interface IDocxDerivedPartPass
{
    int Read(byte[] buffer, int offset, int count);
}
