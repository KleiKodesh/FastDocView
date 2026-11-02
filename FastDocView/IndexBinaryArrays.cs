using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace FastDocView;

/// <summary>
/// Array I/O for the sidecar indexes. Two layouts:
/// <list type="bullet">
/// <item><see cref="Write{T}"/>/<see cref="Read{T}"/>: an int count, then the elements' bytes as-is (little-endian, the only
/// byte order the library runs on) — one memory cast per array on read.</item>
/// <item>the compact forms (delta-varint offsets, varint lengths, byte ids with an escape, sparse id pairs): an int count,
/// an int byte length, then the encoded bytes, so a read is still one <c>ReadBytes</c> plus a tight decode loop
/// (123k deltas decode in under 1 ms). Measured on the 87 MB sample they cut the per-block tables ~4x.</item>
/// </list>
/// </summary>
internal static class IndexBinaryArrays
{
    public static void Write<T>(BinaryWriter writer, T[] values) where T : unmanaged
    {
        writer.Write(values.Length);
        writer.Write(MemoryMarshal.AsBytes(values.AsSpan()).ToArray());
    }

    public static T[] Read<T>(BinaryReader reader) where T : unmanaged
    {
        int count = reader.ReadInt32();
        int byteCount = checked(count * Unsafe.SizeOf<T>());
        byte[] bytes = reader.ReadBytes(byteCount);
        if (bytes.Length != byteCount) throw new EndOfStreamException("Index array truncated.");
        return MemoryMarshal.Cast<byte, T>(bytes).ToArray();
    }

    /// <summary>The whole index file in one read, as a reader over memory (indexes are a few MB at most).</summary>
    public static BinaryReader OpenInMemory(string path)
        => new(new MemoryStream(File.ReadAllBytes(path), writable: false), System.Text.Encoding.UTF8);

    // ── strings that compress well (root tags full of namespace declarations, bookmark names): one deflated blob ──

    /// <summary>Strings joined by '\0', UTF-8, deflated: an int count, an int blob length, the blob.</summary>
    public static void WriteCompressedStrings(BinaryWriter writer, IReadOnlyList<string> values)
    {
        var raw = new MemoryStream();
        for (int i = 0; i < values.Count; i++)
        {
            if (values[i].IndexOf('\0') >= 0) throw new InvalidDataException("Index string holds '\0'.");
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(values[i]);
            raw.Write(bytes, 0, bytes.Length);
            if (i + 1 < values.Count) raw.WriteByte(0);
        }
        var blob = new MemoryStream();
        using (var deflate = new System.IO.Compression.DeflateStream(blob, System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true))
            deflate.Write(raw.GetBuffer(), 0, (int)raw.Length);
        writer.Write(values.Count);
        writer.Write((int)blob.Length);
        writer.Write(blob.GetBuffer(), 0, (int)blob.Length);
    }

    public static string[] ReadCompressedStrings(BinaryReader reader)
    {
        int count = reader.ReadInt32();
        int blobLength = reader.ReadInt32();
        if (count < 0 || blobLength < 0) throw new InvalidDataException("Index strings: negative size.");
        byte[] blob = reader.ReadBytes(blobLength);
        if (blob.Length != blobLength) throw new EndOfStreamException("Index strings truncated.");
        if (count == 0) return [];
        var raw = new MemoryStream();
        using (var inflate = new System.IO.Compression.DeflateStream(new MemoryStream(blob, writable: false), System.IO.Compression.CompressionMode.Decompress))
        {
            var chunk = new byte[1 << 14];
            int read;
            while ((read = inflate.Read(chunk, 0, chunk.Length)) > 0) raw.Write(chunk, 0, read);
        }
        string[] values = System.Text.Encoding.UTF8.GetString(raw.GetBuffer(), 0, (int)raw.Length).Split('\0');
        if (values.Length != count) throw new InvalidDataException("Index strings: count mismatch.");
        return values;
    }

    public static void WriteCompressedString(BinaryWriter writer, string value) => WriteCompressedStrings(writer, [value]);

    public static string ReadCompressedString(BinaryReader reader)
    {
        string[] values = ReadCompressedStrings(reader);
        return values.Length == 1 ? values[0] : throw new InvalidDataException("Index string: count mismatch.");
    }

    // ── varint primitives ─────────────────────────────────────────────────────────────

    /// <summary>Appends an unsigned LEB128 varint (7 bits per byte, low group first).</summary>
    public static void AppendVarint(List<byte> buffer, ulong value)
    {
        while (value >= 0x80) { buffer.Add((byte)(value | 0x80)); value >>= 7; }
        buffer.Add((byte)value);
    }

    public static ulong ReadVarint(byte[] bytes, ref int position)
    {
        ulong value = 0;
        int shift = 0;
        while (true)
        {
            if (position >= bytes.Length) throw new EndOfStreamException("Index varint truncated.");
            byte b = bytes[position++];
            value |= (ulong)(b & 0x7F) << shift;
            if (b < 0x80) return value;
            shift += 7;
            if (shift > 63) throw new InvalidDataException("Index varint too long.");
        }
    }

    private static void WriteEncoded(BinaryWriter writer, int count, List<byte> encoded)
    {
        writer.Write(count);
        writer.Write(encoded.Count);
        writer.Write(encoded.ToArray());
    }

    private static byte[] ReadEncoded(BinaryReader reader, out int count)
    {
        count = reader.ReadInt32();
        int byteCount = reader.ReadInt32();
        if (count < 0 || byteCount < 0) throw new InvalidDataException("Index array: negative size.");
        byte[] bytes = reader.ReadBytes(byteCount);
        if (bytes.Length != byteCount) throw new EndOfStreamException("Index array truncated.");
        return bytes;
    }

    // ── non-decreasing offsets: delta varints ─────────────────────────────────────────

    /// <summary>Non-decreasing offsets (block starts) as varint deltas from the previous value (first from 0).</summary>
    public static void WriteDeltaVarints(BinaryWriter writer, long[] values)
    {
        var encoded = new List<byte>(values.Length * 3);
        long previous = 0;
        foreach (long value in values)
        {
            if (value < previous) throw new InvalidDataException("Index offsets must not decrease.");
            AppendVarint(encoded, (ulong)(value - previous));
            previous = value;
        }
        WriteEncoded(writer, values.Length, encoded);
    }

    public static long[] ReadDeltaVarints(BinaryReader reader)
    {
        byte[] bytes = ReadEncoded(reader, out int count);
        var values = new long[count];
        int position = 0;
        long running = 0;
        for (int i = 0; i < count; i++)
        {
            running += (long)ReadVarint(bytes, ref position);
            values[i] = running;
        }
        return values;
    }

    public static void WriteDeltaVarints(BinaryWriter writer, int[] values)
    {
        var longs = new long[values.Length];
        for (int i = 0; i < values.Length; i++) longs[i] = values[i];
        WriteDeltaVarints(writer, longs);
    }

    public static int[] ReadDeltaVarintsInt32(BinaryReader reader)
    {
        long[] longs = ReadDeltaVarints(reader);
        var values = new int[longs.Length];
        for (int i = 0; i < longs.Length; i++) values[i] = checked((int)longs[i]);
        return values;
    }

    // ── block ends: only where a block does not end where the next one starts ───────

    /// <summary>Block ends as (block delta, end) varint pairs for the blocks whose end differs from the next block's start (the last always).</summary>
    public static void WriteBlockEnds(BinaryWriter writer, long[] starts, long[] ends)
    {
        if (ends.Length != starts.Length) throw new InvalidDataException("Index: block ends do not match the starts.");
        var encoded = new List<byte>();
        int previousBlock = 0, set = 0;
        for (int block = 0; block < ends.Length; block++)
        {
            if (block + 1 < starts.Length && ends[block] == starts[block + 1]) continue;
            if (ends[block] < 0) throw new InvalidDataException("Index: negative block end.");
            AppendVarint(encoded, (uint)(block - previousBlock));
            AppendVarint(encoded, (ulong)ends[block]);
            previousBlock = block;
            set++;
        }
        WriteEncoded(writer, set, encoded);
    }

    public static long[] ReadBlockEnds(BinaryReader reader, long[] starts)
    {
        byte[] bytes = ReadEncoded(reader, out int set);
        var ends = new long[starts.Length];
        for (int i = 0; i + 1 < starts.Length; i++) ends[i] = starts[i + 1];
        int position = 0, block = 0;
        for (int i = 0; i < set; i++)
        {
            block += (int)(uint)ReadVarint(bytes, ref position);
            if ((uint)block >= (uint)starts.Length) throw new InvalidDataException("Index block ends out of range.");
            ends[block] = (long)ReadVarint(bytes, ref position);
        }
        return ends;
    }

    // ── small non-negative values: plain varints ─────────────────────────────────────

    public static void WriteVarints(BinaryWriter writer, int[] values)
    {
        var encoded = new List<byte>(values.Length * 2);
        foreach (int value in values) AppendVarint(encoded, (uint)value);
        WriteEncoded(writer, values.Length, encoded);
    }

    public static int[] ReadVarints(BinaryReader reader)
    {
        byte[] bytes = ReadEncoded(reader, out int count);
        var values = new int[count];
        int position = 0;
        for (int i = 0; i < count; i++) values[i] = (int)(uint)ReadVarint(bytes, ref position);
        return values;
    }

    // ── small ids with -1 = none: one byte each, 0xFF escapes to a varint ────────────

    /// <summary>Ids in -1..: -1 → 0xFE, 0..253 → the byte, larger → 0xFF + varint. Style ids are a handful per document.</summary>
    public static void WriteByteIds(BinaryWriter writer, short[] values)
    {
        var encoded = new List<byte>(values.Length);
        foreach (short value in values)
        {
            if (value < -1) throw new InvalidDataException("Index id below -1.");
            if (value == -1) encoded.Add(0xFE);
            else if (value < 0xFE) encoded.Add((byte)value);
            else { encoded.Add(0xFF); AppendVarint(encoded, (uint)value); }
        }
        WriteEncoded(writer, values.Length, encoded);
    }

    public static short[] ReadByteIds(BinaryReader reader)
    {
        byte[] bytes = ReadEncoded(reader, out int count);
        var values = new short[count];
        int position = 0;
        for (int i = 0; i < count; i++)
        {
            if (position >= bytes.Length) throw new EndOfStreamException("Index array truncated.");
            byte b = bytes[position++];
            values[i] = b == 0xFE ? (short)-1 : b == 0xFF ? checked((short)ReadVarint(bytes, ref position)) : b;
        }
        return values;
    }

    // ── ids with -1 = none, 16-bit when the table fits ───────────────────────────────

    /// <summary>Ids in -1..: as ushort (0xFFFF = -1) when every id fits, else as int; a leading width byte says which.</summary>
    public static void WriteIds(BinaryWriter writer, int[] values)
    {
        bool fits16 = true;
        foreach (int value in values) { if (value < -1) throw new InvalidDataException("Index id below -1."); if (value >= 0xFFFF) { fits16 = false; break; } }
        writer.Write((byte)(fits16 ? 2 : 4));
        if (fits16)
        {
            var shorts = new ushort[values.Length];
            for (int i = 0; i < values.Length; i++) shorts[i] = values[i] < 0 ? (ushort)0xFFFF : (ushort)values[i];
            Write(writer, shorts);
        }
        else Write(writer, values);
    }

    public static int[] ReadIds(BinaryReader reader)
    {
        byte width = reader.ReadByte();
        if (width == 4) return Read<int>(reader);
        if (width != 2) throw new InvalidDataException("Index id array: unknown width.");
        ushort[] shorts = Read<ushort>(reader);
        var values = new int[shorts.Length];
        for (int i = 0; i < shorts.Length; i++) values[i] = shorts[i] == 0xFFFF ? -1 : shorts[i];
        return values;
    }

    // ── mostly -1 ids: sorted (block, id) pairs, expanded on read ────────────────────

    /// <summary>Per-block ids where nearly every block has none (-1): only the set blocks are written, as (block delta, id) varint pairs.</summary>
    public static void WriteSparseIds(BinaryWriter writer, int[] values)
    {
        var encoded = new List<byte>();
        int previousBlock = 0, set = 0;
        for (int block = 0; block < values.Length; block++)
        {
            if (values[block] < 0) continue;
            AppendVarint(encoded, (uint)(block - previousBlock));
            AppendVarint(encoded, (uint)values[block]);
            previousBlock = block;
            set++;
        }
        writer.Write(values.Length);
        WriteEncoded(writer, set, encoded);
    }

    public static int[] ReadSparseIds(BinaryReader reader)
    {
        int length = reader.ReadInt32();
        if (length < 0) throw new InvalidDataException("Index array: negative size.");
        byte[] bytes = ReadEncoded(reader, out int set);
        var values = new int[length];
        for (int i = 0; i < length; i++) values[i] = -1;
        int position = 0, block = 0;
        for (int i = 0; i < set; i++)
        {
            block += (int)(uint)ReadVarint(bytes, ref position);
            if ((uint)block >= (uint)length) throw new InvalidDataException("Index sparse ids out of range.");
            values[block] = (int)(uint)ReadVarint(bytes, ref position);
        }
        return values;
    }
}
