namespace FastDocView;

/// <summary>
/// A span anchor's scan state made of two parts: an opaque head stored whole (a parser snapshot, open frames) and a
/// list-numbering state stored as a delta from the previous anchor (<see cref="WordListCounterState"/>). Layout:
/// head length as a varint, the head, then the numbering bytes.
/// </summary>
internal static class WordIndexScanState
{
    public static byte[] Join(byte[] head, byte[] numbering)
    {
        var bytes = new List<byte>(head.Length + numbering.Length + 4);
        IndexBinaryArrays.AppendVarint(bytes, (uint)head.Length);
        bytes.AddRange(head);
        bytes.AddRange(numbering);
        return bytes.ToArray();
    }

    public static (byte[] head, byte[] numbering) Split(byte[] state)
    {
        if (state.Length == 0) return ([], []);
        int position = 0;
        int headLength = (int)IndexBinaryArrays.ReadVarint(state, ref position);
        if (headLength < 0 || headLength > state.Length - position) throw new InvalidDataException("Index: scan state damaged.");
        var head = new byte[headLength];
        Buffer.BlockCopy(state, position, head, 0, headLength);
        var numbering = new byte[state.Length - position - headLength];
        Buffer.BlockCopy(state, position + headLength, numbering, 0, numbering.Length);
        return (head, numbering);
    }

    /// <summary><paramref name="current"/> as a delta from <paramref name="previous"/>: the head whole, the numbering as a <see cref="WordListCounterState.Delta"/>.</summary>
    public static byte[] Delta(byte[] previous, byte[] current)
    {
        (byte[] head, byte[] numbering) = Split(current);
        (_, byte[] previousNumbering) = Split(previous);
        return Join(head, WordListCounterState.Delta(previousNumbering, numbering));
    }

    public static byte[] Apply(byte[] previous, byte[] delta)
    {
        (byte[] head, byte[] numberingDelta) = Split(delta);
        (_, byte[] previousNumbering) = Split(previous);
        return Join(head, WordListCounterState.Apply(previousNumbering, numberingDelta));
    }
}
