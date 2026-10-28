namespace WordDocumentRandomAccess.Deflate;

/// <summary>
/// LSB-first bit reader over a bounded region of a stream, with the one feature a plain reader
/// lacks: it can report its exact position (byte + leftover bits) and be re-seated there later —
/// what makes a <see cref="DeflateCheckpoint"/> possible.
/// </summary>
internal sealed class DeflateBitReader
{
    private readonly Stream _input;
    private readonly long _regionStart;
    private readonly long _regionEnd;
    private readonly byte[] _buffer = new byte[1 << 16];
    private int _bufferPosition;
    private int _bufferLength;
    private long _nextLoadPosition;   // absolute stream position of _buffer[_bufferLength]

    private ulong _bitBuffer;
    private int _bitCount;

    public DeflateBitReader(Stream input, long regionStart, long regionLength)
    {
        _input = input;
        _regionStart = regionStart;
        _regionEnd = regionStart + regionLength;
        SeekTo(0, 0, 0);
    }

    /// <summary>Re-seats the reader at a checkpoint position (relative to the region start).</summary>
    public void SeekTo(long compressedBytePosition, int leftoverBitCount, int leftoverBits)
    {
        _nextLoadPosition = _regionStart + compressedBytePosition;
        _bufferPosition = _bufferLength = 0;
        _bitBuffer = (ulong)leftoverBits & ((1UL << leftoverBitCount) - 1);
        _bitCount = leftoverBitCount;
    }

    /// <summary>Position of the first byte not fully consumed, relative to the region start.</summary>
    public long CompressedBytePosition
        => (_nextLoadPosition - _regionStart) - (_bufferLength - _bufferPosition) - (_bitCount / 8);

    public int LeftoverBitCount => _bitCount % 8;

    public int LeftoverBits => (int)(_bitBuffer & ((1UL << LeftoverBitCount) - 1));

    public void EnsureBits(int count)
    {
        while (_bitCount < count)
        {
            if (_bufferPosition == _bufferLength && !Fill())
            {
                // Past the end of the data: feed zero bits so a final peek can still complete.
                // A real code never needs them; a truncated stream errors out on the next symbol.
                if (_bitCount < count) { _bitCount += 8; continue; }
                return;
            }
            _bitBuffer |= (ulong)_buffer[_bufferPosition++] << _bitCount;
            _bitCount += 8;
        }
    }

    public int PeekBits(int count) => (int)(_bitBuffer & ((1UL << count) - 1));

    public void DropBits(int count)
    {
        _bitBuffer >>= count;
        _bitCount -= count;
    }

    public int ReadBits(int count)
    {
        if (count == 0) return 0;
        EnsureBits(count);
        int value = PeekBits(count);
        DropBits(count);
        return value;
    }

    /// <summary>Discards the bits up to the next byte boundary (stored blocks are byte aligned).</summary>
    public void AlignToByte() => DropBits(_bitCount % 8);

    public bool IsAtEnd => _bitCount == 0 && _bufferPosition == _bufferLength && _nextLoadPosition >= _regionEnd;

    private bool Fill()
    {
        long remaining = _regionEnd - _nextLoadPosition;
        if (remaining <= 0) return false;
        int toRead = (int)Math.Min(_buffer.Length, remaining);
        if (_input.Position != _nextLoadPosition) _input.Position = _nextLoadPosition;
        int read = _input.Read(_buffer, 0, toRead);
        if (read <= 0) return false;
        _bufferPosition = 0;
        _bufferLength = read;
        _nextLoadPosition += read;
        return true;
    }
}
