namespace FastDocView;

/// <summary>
/// One span of a compact index (every format but .docx, which has its own record carrying a paragraph style): where
/// the rescan of blocks [FirstBlock, next anchor's FirstBlock) starts (<see cref="Position"/> = that block's start - a
/// CP for .doc, a byte offset for the others), the block just before it (its start and kind: a window that starts at
/// this span's first block reads its predecessor without scanning the previous span) and the list-numbering state the
/// scan continues from, as a delta from the previous anchor. Serialised as varint deltas: a dozen bytes per anchor
/// plus the list delta.
/// </summary>
public sealed record WordIndexSpanAnchor(int FirstBlock, long Position, long PreviousBlockPosition, WordDocumentBlockKind PreviousBlockKind, byte[] ListStateDelta)
{
    internal static void WriteAll(BinaryWriter writer, WordIndexSpanAnchor[] anchors)
    {
        var bytes = new List<byte>(anchors.Length * 16);
        IndexBinaryArrays.AppendVarint(bytes, (uint)anchors.Length);
        int previousBlock = 0;
        long previousPosition = 0;
        foreach (WordIndexSpanAnchor anchor in anchors)
        {
            IndexBinaryArrays.AppendVarint(bytes, (uint)(anchor.FirstBlock - previousBlock));
            IndexBinaryArrays.AppendVarint(bytes, (ulong)(anchor.Position - previousPosition));
            IndexBinaryArrays.AppendVarint(bytes, anchor.PreviousBlockPosition < 0 ? 0UL : (ulong)(anchor.Position - anchor.PreviousBlockPosition + 1));   // 0 = no previous block
            bytes.Add((byte)anchor.PreviousBlockKind);
            IndexBinaryArrays.AppendVarint(bytes, (uint)anchor.ListStateDelta.Length);
            bytes.AddRange(anchor.ListStateDelta);
            previousBlock = anchor.FirstBlock;
            previousPosition = anchor.Position;
        }
        writer.Write(bytes.Count);
        writer.Write(bytes.ToArray());
    }

    /// <param name="maxPosition">The text's end (CP count or byte length): every anchor lies in [0, maxPosition].</param>
    internal static WordIndexSpanAnchor[] ReadAll(BinaryReader reader, int blockCount, long maxPosition)
    {
        int byteCount = reader.ReadInt32();
        if (byteCount < 0 || byteCount > reader.BaseStream.Length - reader.BaseStream.Position) throw new InvalidDataException("Index: anchors damaged.");
        byte[] bytes = reader.ReadBytes(byteCount);
        if (bytes.Length != byteCount) throw new EndOfStreamException("Index: anchors truncated.");
        int position = 0;
        int count = (int)IndexBinaryArrays.ReadVarint(bytes, ref position);
        if (count < 0 || (count == 0) != (blockCount == 0)) throw new InvalidDataException("Index: anchors do not match the blocks.");
        var anchors = new WordIndexSpanAnchor[count];
        int previousBlock = 0;
        long previousPosition = 0;
        for (int i = 0; i < count; i++)
        {
            int firstBlock = previousBlock + (int)IndexBinaryArrays.ReadVarint(bytes, ref position);
            long at = previousPosition + (long)IndexBinaryArrays.ReadVarint(bytes, ref position);
            long back = (long)IndexBinaryArrays.ReadVarint(bytes, ref position);
            if (position >= bytes.Length) throw new EndOfStreamException("Index: anchors truncated.");
            var kind = (WordDocumentBlockKind)bytes[position++];
            int stateLength = (int)IndexBinaryArrays.ReadVarint(bytes, ref position);
            if (stateLength < 0 || stateLength > bytes.Length - position) throw new InvalidDataException("Index: anchor list state truncated.");
            var state = new byte[stateLength];
            Buffer.BlockCopy(bytes, position, state, 0, stateLength);
            position += stateLength;
            if (firstBlock >= blockCount || (i > 0 && firstBlock <= previousBlock) || (i == 0 && firstBlock != 0)) throw new InvalidDataException("Index: anchor blocks out of order.");
            if (at < 0 || at > maxPosition || (i > 0 && at <= previousPosition) || (i == 0) != (back == 0) || back - 1 > at) throw new InvalidDataException("Index: anchor positions out of order.");
            anchors[i] = new WordIndexSpanAnchor(firstBlock, at, back == 0 ? -1 : at - (back - 1), kind, state);
            previousBlock = firstBlock;
            previousPosition = at;
        }
        return anchors;
    }

    /// <summary>The anchor (span) a block belongs to: the last anchor whose first block is at or before the block.</summary>
    public static int IndexOfBlock(WordIndexSpanAnchor[] anchors, int blockIndex)
    {
        int low = 0, high = anchors.Length - 1;
        while (low < high)
        {
            int middle = (low + high + 1) / 2;
            if (anchors[middle].FirstBlock <= blockIndex) low = middle; else high = middle - 1;
        }
        return low;
    }

    /// <summary>The anchor (span) a position falls in: the last anchor whose position is at or before it.</summary>
    public static int IndexOfPosition(WordIndexSpanAnchor[] anchors, long position)
    {
        int low = 0, high = anchors.Length - 1;
        while (low < high)
        {
            int middle = (low + high + 1) / 2;
            if (anchors[middle].Position <= position) low = middle; else high = middle - 1;
        }
        return low;
    }
}
