using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace FastDocView;

/// <summary>
/// Raw-block array I/O for the sidecar indexes: an int count, then the elements' bytes as-is
/// (little-endian, the only byte order the library runs on). Reading a 123k-block index this way is
/// one memory cast per array instead of hundreds of thousands of BinaryReader calls — the
/// difference between ~80 ms and a few ms on reopen.
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
}
