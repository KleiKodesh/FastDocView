using System.IO.Compression;

namespace FastDocView;

/// <summary>One deflated blob inside an index (a table of small records that compresses well): an int length, the bytes.</summary>
internal static class WordIndexDeflatedBlob
{
    public static void Write(BinaryWriter writer, byte[] raw)
    {
        var blob = new MemoryStream();
        using (var deflate = new DeflateStream(blob, CompressionLevel.Optimal, leaveOpen: true)) deflate.Write(raw, 0, raw.Length);
        writer.Write(raw.Length);
        writer.Write((int)blob.Length);
        writer.Write(blob.GetBuffer(), 0, (int)blob.Length);
    }

    public static byte[] Read(BinaryReader reader)
    {
        int rawLength = reader.ReadInt32();
        int blobLength = reader.ReadInt32();
        if (rawLength < 0 || blobLength < 0) throw new InvalidDataException("Index blob: negative size.");
        byte[] blob = reader.ReadBytes(blobLength);
        if (blob.Length != blobLength) throw new EndOfStreamException("Index blob truncated.");
        var raw = new byte[rawLength];
        using var inflate = new DeflateStream(new MemoryStream(blob, writable: false), CompressionMode.Decompress);
        int filled = 0;
        while (filled < rawLength)
        {
            int read = inflate.Read(raw, filled, rawLength - filled);
            if (read <= 0) throw new InvalidDataException("Index blob: shorter than declared.");
            filled += read;
        }
        return raw;
    }
}
