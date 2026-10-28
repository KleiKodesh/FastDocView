using System.Text;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// Minimal reader for the OLE compound file that wraps a binary .doc ([MS-CFB]): header, FAT
/// (through the DIFAT), directory, mini FAT and mini stream. Streams are exposed as seekable
/// <see cref="Stream"/>s mapped over their sector chains — nothing is copied out of the file, so a
/// 500 MB WordDocument stream costs a few hundred KB of chain table, not 500 MB.
/// </summary>
public sealed class CompoundFileReader : IDisposable
{
    private static readonly byte[] Signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;
    private const int DirectoryEntrySize = 128;

    private readonly Stream _file;
    private readonly int _sectorSize;
    private readonly int _miniSectorSize;
    private readonly uint _miniStreamCutoff;
    private readonly uint[] _fat;
    private readonly uint[] _miniFat;
    private readonly List<DirectoryEntry> _entries = new();
    private readonly Stream _miniStream;

    public sealed record DirectoryEntry(int Id, string Name, byte Type, uint StartSector, long Size, int Left, int Right, int Child);

    public static bool HasSignature(string path)
    {
        using FileStream file = WordDocumentFileStream.OpenShared(path);
        var header = new byte[8];
        return file.Read(header, 0, 8) == 8 && header.AsSpan().SequenceEqual(Signature);
    }

    public CompoundFileReader(string path)
        : this(new WordDocumentFileStream(path)) { }

    /// <summary>Over any seekable stream (owned by this reader); what a WASM host passes in.</summary>
    public CompoundFileReader(Stream file)
    {
        _file = file;
        var header = new byte[512];
        ReadAt(_file, 0, header, 512);
        if (!header.AsSpan(0, 8).SequenceEqual(Signature)) throw new InvalidDataException("Not an OLE compound file.");
        _sectorSize = 1 << BitConverter.ToUInt16(header, 0x1E);
        _miniSectorSize = 1 << BitConverter.ToUInt16(header, 0x20);
        uint fatSectorCount = BitConverter.ToUInt32(header, 0x2C);
        uint firstDirectorySector = BitConverter.ToUInt32(header, 0x30);
        _miniStreamCutoff = BitConverter.ToUInt32(header, 0x38);
        uint firstMiniFatSector = BitConverter.ToUInt32(header, 0x3C);
        uint miniFatSectorCount = BitConverter.ToUInt32(header, 0x40);
        uint firstDifatSector = BitConverter.ToUInt32(header, 0x44);
        uint difatSectorCount = BitConverter.ToUInt32(header, 0x48);

        _fat = LoadFat(header, fatSectorCount, firstDifatSector, difatSectorCount);
        ReadDirectory(firstDirectorySector);
        _miniFat = LoadChainedTable(firstMiniFatSector, miniFatSectorCount);
        DirectoryEntry root = _entries[0];
        _miniStream = new SectorChainStream(_file, _sectorSize, ChainOf(_fat, root.StartSector), root.Size, SectorOffset);
    }

    /// <summary>Opens a stream that is a direct child of the root storage (where Word keeps its streams).</summary>
    public Stream? OpenRootStream(string name)
    {
        DirectoryEntry? entry = FindRootChild(name);
        if (entry is null) return null;
        if (entry.Size < _miniStreamCutoff)
            return new SectorChainStream(_miniStream, _miniSectorSize, ChainOf(_miniFat, entry.StartSector), entry.Size, sector => (long)sector * _miniSectorSize);
        return new SectorChainStream(_file, _sectorSize, ChainOf(_fat, entry.StartSector), entry.Size, SectorOffset);
    }

    public bool HasRootStream(string name) => FindRootChild(name) is not null;

    private DirectoryEntry? FindRootChild(string name)
    {
        // The root's children form a red-black tree; an iterative walk over left/right/child is enough.
        var pending = new Stack<int>();
        if (_entries.Count == 0) return null;
        pending.Push(_entries[0].Child);
        int visited = 0;
        while (pending.Count > 0 && visited++ < _entries.Count * 2)
        {
            int id = pending.Pop();
            if (id < 0 || id >= _entries.Count) continue;
            DirectoryEntry entry = _entries[id];
            if (string.Equals(entry.Name, name, StringComparison.Ordinal)) return entry;
            pending.Push(entry.Left);
            pending.Push(entry.Right);
        }
        return null;
    }

    private long SectorOffset(uint sector) => ((long)sector + 1) * _sectorSize;

    private uint[] LoadFat(byte[] header, uint fatSectorCount, uint firstDifatSector, uint difatSectorCount)
    {
        var fatSectors = new List<uint>();
        for (int i = 0; i < 109 && fatSectors.Count < fatSectorCount; i++)
        {
            uint sector = BitConverter.ToUInt32(header, 0x4C + i * 4);
            if (sector != FreeSector) fatSectors.Add(sector);
        }
        uint difatSector = firstDifatSector;
        var difatBuffer = new byte[_sectorSize];
        for (uint n = 0; n < difatSectorCount && difatSector < 0xFFFFFFFA; n++)
        {
            ReadAt(_file, SectorOffset(difatSector), difatBuffer, _sectorSize);
            int entriesPerSector = _sectorSize / 4 - 1;
            for (int i = 0; i < entriesPerSector && fatSectors.Count < fatSectorCount; i++)
            {
                uint sector = BitConverter.ToUInt32(difatBuffer, i * 4);
                if (sector != FreeSector) fatSectors.Add(sector);
            }
            difatSector = BitConverter.ToUInt32(difatBuffer, entriesPerSector * 4);
        }

        int entriesPerFatSector = _sectorSize / 4;
        var fat = new uint[fatSectors.Count * entriesPerFatSector];
        var sectorBuffer = new byte[_sectorSize];
        for (int s = 0; s < fatSectors.Count; s++)
        {
            ReadAt(_file, SectorOffset(fatSectors[s]), sectorBuffer, _sectorSize);
            for (int i = 0; i < entriesPerFatSector; i++) fat[s * entriesPerFatSector + i] = BitConverter.ToUInt32(sectorBuffer, i * 4);
        }
        return fat;
    }

    private uint[] LoadChainedTable(uint firstSector, uint sectorCount)
    {
        if (sectorCount == 0 || firstSector >= 0xFFFFFFFA) return [];
        List<uint> chain = ChainOf(_fat, firstSector);
        int entriesPerSector = _sectorSize / 4;
        var table = new uint[chain.Count * entriesPerSector];
        var buffer = new byte[_sectorSize];
        for (int s = 0; s < chain.Count; s++)
        {
            ReadAt(_file, SectorOffset(chain[s]), buffer, _sectorSize);
            for (int i = 0; i < entriesPerSector; i++) table[s * entriesPerSector + i] = BitConverter.ToUInt32(buffer, i * 4);
        }
        return table;
    }

    private void ReadDirectory(uint firstDirectorySector)
    {
        var buffer = new byte[_sectorSize];
        int id = 0;
        foreach (uint sector in ChainOf(_fat, firstDirectorySector))
        {
            ReadAt(_file, SectorOffset(sector), buffer, _sectorSize);
            for (int offset = 0; offset + DirectoryEntrySize <= _sectorSize; offset += DirectoryEntrySize, id++)
            {
                int nameLength = BitConverter.ToUInt16(buffer, offset + 64);
                string name = nameLength >= 2 ? Encoding.Unicode.GetString(buffer, offset, nameLength - 2) : "";
                _entries.Add(new DirectoryEntry(
                    id, name, buffer[offset + 66], BitConverter.ToUInt32(buffer, offset + 116),
                    BitConverter.ToUInt32(buffer, offset + 120),
                    BitConverter.ToInt32(buffer, offset + 68), BitConverter.ToInt32(buffer, offset + 72), BitConverter.ToInt32(buffer, offset + 76)));
            }
        }
    }

    private static List<uint> ChainOf(uint[] table, uint firstSector)
    {
        var chain = new List<uint>();
        uint sector = firstSector;
        while (sector < 0xFFFFFFFA && sector < table.Length && chain.Count <= table.Length)
        {
            chain.Add(sector);
            sector = table[sector];
            if (sector == EndOfChain) break;
        }
        return chain;
    }

    private static void ReadAt(Stream stream, long offset, byte[] buffer, int count)
    {
        stream.Position = offset;
        int filled = 0;
        while (filled < count)
        {
            int read = stream.Read(buffer, filled, count - filled);
            if (read <= 0) { Array.Clear(buffer, filled, count - filled); return; }
            filled += read;
        }
    }

    public void Dispose()
    {
        _miniStream.Dispose();
        _file.Dispose();
    }

    /// <summary>A stream's logical bytes mapped onto its sector chain inside an underlying stream.</summary>
    private sealed class SectorChainStream : Stream
    {
        private readonly Stream _underlying;
        private readonly int _sectorSize;
        private readonly List<uint> _chain;
        private readonly long _length;
        private readonly Func<uint, long> _offsetOfSector;
        private long _position;

        public SectorChainStream(Stream underlying, int sectorSize, List<uint> chain, long length, Func<uint, long> offsetOfSector)
        {
            _underlying = underlying;
            _sectorSize = sectorSize;
            _chain = chain;
            _length = Math.Min(length, (long)chain.Count * sectorSize);
            _offsetOfSector = offsetOfSector;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            int total = 0;
            while (count > 0 && _position < _length)
            {
                int sectorIndex = (int)(_position / _sectorSize);
                int inSector = (int)(_position % _sectorSize);
                int toRead = (int)Math.Min(Math.Min(count, _sectorSize - inSector), _length - _position);
                _underlying.Position = _offsetOfSector(_chain[sectorIndex]) + inSector;
                int read = _underlying.Read(buffer, offset, toRead);
                if (read <= 0) break;
                _position += read; offset += read; count -= read; total += read;
            }
            return total;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position { get => _position; set => _position = Math.Max(0, Math.Min(_length, value)); }
        public override long Seek(long offset, SeekOrigin origin)
        {
            Position = origin switch { SeekOrigin.Begin => offset, SeekOrigin.Current => _position + offset, _ => _length + offset };
            return _position;
        }
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
