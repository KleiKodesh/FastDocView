namespace FastDocView;

/// <summary>
/// The bytes most recently read from a part's view, kept as one contiguous range [<see cref="Start"/>, <see cref="End"/>)
/// of bounded size. A read that continues or overlaps the range extends it at either end (dropping bytes from the far
/// end when full), so windows shown again — forward or paged back — come from here instead of a seek back into a
/// deflated part, which resumes from a checkpoint and inflates again. A read elsewhere restarts the range. The buffer
/// may grow up to <c>maxCapacity</c> for one long read; nothing past that is kept.
/// </summary>
internal sealed class WordIndexByteRing
{
    private const int MaxGap = 1 << 16;    // a read this close past the end reads the gap too, so the range stays one piece
    private readonly int _maxCapacity;
    private byte[]? _gapBuffer;
    private byte[] _bytes;
    private long _start = -1;
    private int _length;

    public WordIndexByteRing(int capacity, int maxCapacity)
    {
        _bytes = new byte[capacity];
        _maxCapacity = maxCapacity;
    }

    public long Start => _start;
    public long End => _start < 0 ? -1 : _start + _length;
    public bool IsEmpty => _start < 0;

    public bool Contains(long offset) => _start >= 0 && offset >= _start && offset < _start + _length;

    /// <summary>Copies the held bytes from <paramref name="offset"/> on (up to <paramref name="count"/>); 0 when the offset is not held.</summary>
    public int CopyFrom(long offset, byte[] destination, int destinationOffset, int count)
    {
        if (!Contains(offset)) return 0;
        int available = (int)Math.Min(count, _start + _length - offset);
        Buffer.BlockCopy(_bytes, (int)(offset - _start), destination, destinationOffset, available);
        return available;
    }

    /// <summary>
    /// Fills <paramref name="destination"/> from index <paramref name="filled"/> on, where index 0 is offset
    /// <paramref name="rangeStart"/>: held bytes are copied, the rest is read from the view (only up to where the
    /// held range begins) and retained.
    /// </summary>
    public void Fill(Stream view, long rangeStart, byte[] destination, int filled)
    {
        while (filled < destination.Length)
        {
            long position = rangeStart + filled;
            int copied = CopyFrom(position, destination, filled, destination.Length - filled);
            if (copied > 0) { filled += copied; continue; }
            int toRead = destination.Length - filled;
            if (_start > position) toRead = (int)Math.Min(toRead, _start - position);
            else CloseGapTo(view, position);
            if (view.Position != position) view.Position = position;
            int read = view.Read(destination, filled, toRead);
            if (read <= 0) throw new EndOfStreamException("Part ended inside an element.");
            Retain(position, destination, filled, read);
            filled += read;
        }
    }

    /// <summary>
    /// Before a read a little past the end (the next block of a window, the next span's anchor): reads the bytes between,
    /// so the range stays one piece — the view continues from the end anyway, no seek back.
    /// </summary>
    public void CloseGapTo(Stream view, long offset)
    {
        if (_start < 0 || offset <= End || offset - End > MaxGap) return;
        long gapStart = End;
        int gap = (int)(offset - gapStart);
        _gapBuffer ??= new byte[MaxGap];
        if (view.Position != gapStart) view.Position = gapStart;
        int filled = 0;
        while (filled < gap)
        {
            int read = view.Read(_gapBuffer, filled, gap - filled);
            if (read <= 0) throw new EndOfStreamException("Part ended inside an element.");
            filled += read;
        }
        Retain(gapStart, _gapBuffer, 0, gap);
    }

    /// <summary>Keeps bytes [offset, offset + count) read from the view: appended or prepended when they touch the range, else the range restarts on them.</summary>
    public void Retain(long offset, byte[] source, int from, int count)
    {
        if (count <= 0) return;
        if (count > _bytes.Length && _bytes.Length < _maxCapacity) Grow(count);
        long end = offset + count;
        if (_start >= 0 && offset <= _start && end >= End) Restart(offset, source, from, count);   // covers the whole range
        else if (_start >= 0 && offset <= End && end > End && offset >= _start) Append(source, from + (int)(End - offset), (int)(end - End));
        else if (_start >= 0 && offset < _start && end >= _start) Prepend(source, from, (int)(_start - offset));
        else if (_start >= 0 && offset >= _start && end <= End) { }        // already held
        else Restart(offset, source, from, count);
    }

    private void Grow(int needed)
    {
        var bigger = new byte[Math.Min(_maxCapacity, Math.Max(needed, _bytes.Length * 2))];
        Buffer.BlockCopy(_bytes, 0, bigger, 0, _length);
        _bytes = bigger;
    }

    private void Restart(long offset, byte[] source, int from, int count)
    {
        int kept = Math.Min(count, _bytes.Length);           // a read longer than the ring: its last bytes
        _start = offset + count - kept;
        _length = kept;
        Buffer.BlockCopy(source, from + count - kept, _bytes, 0, kept);
    }

    private void Append(byte[] source, int from, int count)
    {
        if (_length + count > _bytes.Length && _bytes.Length < _maxCapacity) Grow(_length + count);
        if (count >= _bytes.Length) { Restart(End, source, from, count); return; }
        if (_length + count > _bytes.Length)
        {
            int drop = Math.Min(_length, Math.Max(_length + count - _bytes.Length, _bytes.Length / 4));   // drop the oldest, a quarter at a time
            Buffer.BlockCopy(_bytes, drop, _bytes, 0, _length - drop);
            _start += drop;
            _length -= drop;
        }
        Buffer.BlockCopy(source, from, _bytes, _length, count);
        _length += count;
    }

    private void Prepend(byte[] source, int from, int count)
    {
        if (_length + count > _bytes.Length && _bytes.Length < _maxCapacity) Grow(_length + count);
        if (count >= _bytes.Length) { Restart(_start - count, source, from, count); return; }
        int keep = Math.Min(_length, _bytes.Length - count);  // full: the far end goes
        Buffer.BlockCopy(_bytes, 0, _bytes, count, keep);
        Buffer.BlockCopy(source, from, _bytes, 0, count);
        _start -= count;
        _length = keep + count;
    }
}
