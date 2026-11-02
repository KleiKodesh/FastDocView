namespace FastDocView;

/// <summary>
/// A non-decreasing sequence of block indexes (headings, note references, bookmarks…) in Elias-Fano
/// form: each value's low <c>l</c> bits packed, its high part as a unary code — about
/// 2 + log2(universe / count) bits per value, instead of 32. <see cref="Get"/> (select) walks the unary
/// part from a sampled position (one sample per 64 values), so a random access is a few dozen bit
/// operations; sequential decoding is <see cref="ToArray"/>. Managed only, no intrinsics (net48).
/// </summary>
public sealed class EliasFanoSequence
{
    private const int SampleEvery = 64;

    private readonly int _count;
    private readonly int _lowBits;
    private readonly ulong[] _low;        // packed low bits, _lowBits each
    private readonly ulong[] _high;       // unary: for value i, a 1 at position (value >> _lowBits) + i
    private readonly int[] _samples;      // high-bit position of every SampleEvery-th 1

    public int Count => _count;

    public EliasFanoSequence(IReadOnlyList<int> values)
    {
        _count = values.Count;
        long universe = _count == 0 ? 0 : values[_count - 1];
        for (int i = 1; i < _count; i++) if (values[i] < values[i - 1] || values[i - 1] < 0) throw new ArgumentException("Elias-Fano needs a non-decreasing sequence of non-negative values.");
        _lowBits = 0;
        while (_count > 0 && (universe >> _lowBits) > (long)_count) _lowBits++;        // floor(log2(u / n)) when u > n
        _low = new ulong[((long)_count * _lowBits + 63) / 64 + 1];
        long highLength = _count + ((universe >> _lowBits) + 1);
        _high = new ulong[(highLength + 63) / 64 + 1];
        _samples = new int[(_count + SampleEvery - 1) / SampleEvery];
        ulong lowMask = _lowBits == 64 ? ulong.MaxValue : (1UL << _lowBits) - 1;
        for (int i = 0; i < _count; i++)
        {
            ulong value = (ulong)values[i];
            if (_lowBits > 0)
            {
                long bit = (long)i * _lowBits;
                ulong low = value & lowMask;
                _low[bit >> 6] |= low << (int)(bit & 63);
                if ((bit & 63) + _lowBits > 64) _low[(bit >> 6) + 1] |= low >> (64 - (int)(bit & 63));
            }
            long highPosition = (long)(value >> _lowBits) + i;
            _high[highPosition >> 6] |= 1UL << (int)(highPosition & 63);
            if (i % SampleEvery == 0) _samples[i / SampleEvery] = (int)highPosition;
        }
    }

    private EliasFanoSequence(int count, int lowBits, ulong[] low, ulong[] high)
    {
        _count = count;
        _lowBits = lowBits;
        _low = low;
        _high = high;
        _samples = new int[(count + SampleEvery - 1) / SampleEvery];
        int seen = 0;
        for (long word = 0; word < high.Length && seen < count; word++)
        {
            ulong bits = high[word];
            while (bits != 0)
            {
                int trailing = TrailingZeros(bits);
                if (seen % SampleEvery == 0) _samples[seen / SampleEvery] = (int)(word * 64 + trailing);
                seen++;
                bits &= bits - 1;
            }
        }
        if (seen != count) throw new InvalidDataException("Elias-Fano: high bits do not hold the count.");
    }

    /// <summary>Value number <paramref name="index"/> (select).</summary>
    public int Get(int index)
    {
        if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index));
        int sample = index / SampleEvery;
        long position = _samples[sample];
        int remaining = index - sample * SampleEvery;          // 1s to skip after the sampled one
        long word = position >> 6;
        ulong bits = _high[word] & (ulong.MaxValue << (int)(position & 63));
        while (true)
        {
            int ones = PopCount(bits);
            if (ones > remaining) break;
            remaining -= ones;
            bits = _high[++word];
        }
        for (int i = 0; i < remaining; i++) bits &= bits - 1;   // drop the lower 1s
        long highPosition = word * 64 + TrailingZeros(bits);
        long high = highPosition - index;
        return (int)((high << _lowBits) | (long)Low(index));
    }

    private ulong Low(int index)
    {
        if (_lowBits == 0) return 0;
        long bit = (long)index * _lowBits;
        ulong value = _low[bit >> 6] >> (int)(bit & 63);
        if ((bit & 63) + _lowBits > 64) value |= _low[(bit >> 6) + 1] << (64 - (int)(bit & 63));
        return value & ((1UL << _lowBits) - 1);
    }

    /// <summary>The first index whose value is ≥ <paramref name="value"/> (lower bound), <see cref="Count"/> when none.</summary>
    public int LowerBound(int value)
    {
        int low = 0, high = _count;
        while (low < high)
        {
            int middle = (low + high) / 2;
            if (Get(middle) < value) low = middle + 1; else high = middle;
        }
        return low;
    }

    public int[] ToArray()
    {
        var values = new int[_count];
        for (int i = 0; i < _count; i++) values[i] = Get(i);
        return values;
    }

    public void Write(BinaryWriter writer)
    {
        writer.Write(_count);
        writer.Write((byte)_lowBits);
        writer.Write(_low.Length);
        foreach (ulong word in _low) writer.Write(word);
        writer.Write(_high.Length);
        foreach (ulong word in _high) writer.Write(word);
    }

    public static EliasFanoSequence Read(BinaryReader reader)
    {
        int count = reader.ReadInt32();
        int lowBits = reader.ReadByte();
        if (count < 0 || lowBits > 63) throw new InvalidDataException("Elias-Fano: bad header.");
        var low = new ulong[reader.ReadInt32()];
        for (int i = 0; i < low.Length; i++) low[i] = reader.ReadUInt64();
        var high = new ulong[reader.ReadInt32()];
        for (int i = 0; i < high.Length; i++) high[i] = reader.ReadUInt64();
        if (low.Length != ((long)count * lowBits + 63) / 64 + 1) throw new InvalidDataException("Elias-Fano: low bits do not match the count.");
        return new EliasFanoSequence(count, lowBits, low, high);
    }

    private static int PopCount(ulong value)
    {
        value -= (value >> 1) & 0x5555555555555555UL;
        value = (value & 0x3333333333333333UL) + ((value >> 2) & 0x3333333333333333UL);
        value = (value + (value >> 4)) & 0x0F0F0F0F0F0F0F0FUL;
        return (int)((value * 0x0101010101010101UL) >> 56);
    }

    private static int TrailingZeros(ulong value)
    {
        if (value == 0) return 64;
        int count = 0;
        while ((value & 1) == 0) { value >>= 1; count++; }
        return count;
    }
}
