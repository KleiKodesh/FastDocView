using System.Text;
using WordDocumentRandomAccess.Deflate;
using WordDocumentRandomAccess.Zip;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The sidecar index of one .docx: where word/document.xml sits in the zip, the deflate
/// checkpoints, and every body-level block's uncompressed offset + kind + style + text length.
/// Saved as a small binary file (<see cref="Save"/>/<see cref="Load"/>); about 3 % of the
/// uncompressed XML size at the default 1 MB checkpoint spacing (32 KB window per checkpoint).
/// </summary>
public sealed class DocxRandomAccessIndex
{
    private const string Magic = "KHWORDIDX";
    private const int FormatVersion = 1;

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public ZipEntryLocation DocumentEntry { get; }
    public string RootOpenTag { get; }
    public long BodyEndOffset { get; }
    public IReadOnlyList<DeflateCheckpoint> Checkpoints { get; }
    public long[] BlockStartOffsets { get; }
    public WordDocumentBlockKind[] BlockKinds { get; }
    public int[] BlockTextLengths { get; }
    public short[] BlockStyleIds { get; }
    public string[] StyleIdTable { get; }

    public int BlockCount => BlockStartOffsets.Length;

    public DocxRandomAccessIndex(long sourceFileLength, long sourceLastWriteTicksUtc, ZipEntryLocation documentEntry,
                                 string rootOpenTag, long bodyEndOffset, IReadOnlyList<DeflateCheckpoint> checkpoints,
                                 long[] blockStartOffsets, WordDocumentBlockKind[] blockKinds, int[] blockTextLengths,
                                 short[] blockStyleIds, string[] styleIdTable)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        DocumentEntry = documentEntry;
        RootOpenTag = rootOpenTag;
        BodyEndOffset = bodyEndOffset;
        Checkpoints = checkpoints;
        BlockStartOffsets = blockStartOffsets;
        BlockKinds = blockKinds;
        BlockTextLengths = blockTextLengths;
        BlockStyleIds = blockStyleIds;
        StyleIdTable = styleIdTable;
    }

    public long BlockEndOffset(int blockIndex)
        => blockIndex + 1 < BlockStartOffsets.Length ? BlockStartOffsets[blockIndex + 1] : BodyEndOffset;

    public string? StyleIdOf(int blockIndex)
    {
        short tableIndex = BlockStyleIds[blockIndex];
        return tableIndex >= 0 ? StyleIdTable[tableIndex] : null;
    }

    public bool MatchesSource(string documentPath)
    {
        var info = new FileInfo(documentPath);
        return info.Exists && info.Length == SourceFileLength && info.LastWriteTimeUtc.Ticks == SourceLastWriteTicksUtc;
    }

    public void Save(string indexPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(indexPath)!);
        string temporaryPath = indexPath + ".tmp";
        using (var file = File.Create(temporaryPath))
        using (var writer = new BinaryWriter(file, Encoding.UTF8))
        {
            writer.Write(Magic);
            writer.Write(FormatVersion);
            writer.Write(SourceFileLength);
            writer.Write(SourceLastWriteTicksUtc);
            writer.Write(DocumentEntry.Name);
            writer.Write(DocumentEntry.DataOffset);
            writer.Write(DocumentEntry.CompressedSize);
            writer.Write(DocumentEntry.UncompressedSize);
            writer.Write(DocumentEntry.CompressionMethod);
            writer.Write(DocumentEntry.Crc32);
            writer.Write(RootOpenTag);
            writer.Write(BodyEndOffset);

            writer.Write(Checkpoints.Count);
            foreach (DeflateCheckpoint checkpoint in Checkpoints)
            {
                writer.Write(checkpoint.CompressedBytePosition);
                writer.Write((byte)checkpoint.LeftoverBitCount);
                writer.Write((byte)checkpoint.LeftoverBits);
                writer.Write(checkpoint.UncompressedOffset);
                writer.Write(checkpoint.Window.Length);
                writer.Write(checkpoint.Window);
            }

            writer.Write(StyleIdTable.Length);
            foreach (string styleId in StyleIdTable) writer.Write(styleId);

            writer.Write(BlockStartOffsets.Length);
            foreach (long offset in BlockStartOffsets) writer.Write(offset);
            foreach (WordDocumentBlockKind kind in BlockKinds) writer.Write((byte)kind);
            foreach (int length in BlockTextLengths) writer.Write(length);
            foreach (short styleId in BlockStyleIds) writer.Write(styleId);
        }
        if (File.Exists(indexPath)) File.Delete(indexPath);
        File.Move(temporaryPath, indexPath);
    }

    public static DocxRandomAccessIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            DocxRandomAccessIndex index = Load(indexPath);
            return index.MatchesSource(documentPath) ? index : null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or EndOfStreamException)
        {
            return null;
        }
    }

    public static DocxRandomAccessIndex Load(string indexPath)
    {
        using var file = File.OpenRead(indexPath);
        using var reader = new BinaryReader(file, Encoding.UTF8);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not a word index file.");
        if (reader.ReadInt32() != FormatVersion) throw new InvalidDataException("Word index: unsupported version.");
        long sourceLength = reader.ReadInt64();
        long sourceTicks = reader.ReadInt64();
        var entry = new ZipEntryLocation(reader.ReadString(), reader.ReadInt64(), reader.ReadInt64(), reader.ReadInt64(),
                                         reader.ReadUInt16(), reader.ReadUInt32());
        string rootOpenTag = reader.ReadString();
        long bodyEnd = reader.ReadInt64();

        int checkpointCount = reader.ReadInt32();
        var checkpoints = new List<DeflateCheckpoint>(checkpointCount);
        for (int i = 0; i < checkpointCount; i++)
        {
            long compressedPosition = reader.ReadInt64();
            int leftoverBitCount = reader.ReadByte();
            int leftoverBits = reader.ReadByte();
            long uncompressedOffset = reader.ReadInt64();
            byte[] window = reader.ReadBytes(reader.ReadInt32());
            checkpoints.Add(new DeflateCheckpoint(compressedPosition, leftoverBitCount, leftoverBits, uncompressedOffset, window));
        }

        var styleIdTable = new string[reader.ReadInt32()];
        for (int i = 0; i < styleIdTable.Length; i++) styleIdTable[i] = reader.ReadString();

        int blockCount = reader.ReadInt32();
        var starts = new long[blockCount];
        var kinds = new WordDocumentBlockKind[blockCount];
        var textLengths = new int[blockCount];
        var styleIds = new short[blockCount];
        for (int i = 0; i < blockCount; i++) starts[i] = reader.ReadInt64();
        for (int i = 0; i < blockCount; i++) kinds[i] = (WordDocumentBlockKind)reader.ReadByte();
        for (int i = 0; i < blockCount; i++) textLengths[i] = reader.ReadInt32();
        for (int i = 0; i < blockCount; i++) styleIds[i] = reader.ReadInt16();

        return new DocxRandomAccessIndex(sourceLength, sourceTicks, entry, rootOpenTag, bodyEnd, checkpoints,
                                         starts, kinds, textLengths, styleIds, styleIdTable);
    }
}
