using System.Text;
using WordDocumentRandomAccess.Deflate;
using WordDocumentRandomAccess.Zip;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The index of one .docx body: where word/document.xml sits in the zip, the deflate checkpoints
/// (none in in-memory mode), and every body-level block's uncompressed offset + kind + style +
/// list label + text length. Serialised to/from a <see cref="Stream"/> (the sidecar file on disk,
/// or whatever store a WASM host provides); about 3 % of the uncompressed XML at the default 1 MB
/// checkpoint spacing.
/// </summary>
public sealed class DocxRandomAccessIndex
{
    private const string Magic = "KHWORDIDX";
    private const int FormatVersion = 3;

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
    public int[] BlockListLabelIds { get; }
    public string[] ListLabels { get; }
    public WordDocumentHeading[] Headings { get; }
    public DocxNoteReference[] NoteReferences { get; }      // in body order

    public int BlockCount => BlockStartOffsets.Length;

    public DocxRandomAccessIndex(long sourceFileLength, long sourceLastWriteTicksUtc, ZipEntryLocation documentEntry,
                                 string rootOpenTag, long bodyEndOffset, IReadOnlyList<DeflateCheckpoint> checkpoints,
                                 long[] blockStartOffsets, WordDocumentBlockKind[] blockKinds, int[] blockTextLengths,
                                 short[] blockStyleIds, string[] styleIdTable, int[] blockListLabelIds, string[] listLabels,
                                 WordDocumentHeading[] headings, DocxNoteReference[] noteReferences)
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
        BlockListLabelIds = blockListLabelIds;
        ListLabels = listLabels;
        Headings = headings;
        NoteReferences = noteReferences;
    }

    /// <summary>The note references inside blocks [first, last] — a slice of the body-ordered array.</summary>
    public ReadOnlySpan<DocxNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex)
    {
        int start = LowerBound(firstBlockIndex), end = LowerBound(lastBlockIndex + 1);
        return NoteReferences.AsSpan(start, end - start);
    }

    private int LowerBound(int blockIndex)
    {
        int low = 0, high = NoteReferences.Length;
        while (low < high)
        {
            int middle = (low + high) / 2;
            if (NoteReferences[middle].BlockIndex < blockIndex) low = middle + 1; else high = middle;
        }
        return low;
    }

    public long BlockEndOffset(int blockIndex)
        => blockIndex + 1 < BlockStartOffsets.Length ? BlockStartOffsets[blockIndex + 1] : BodyEndOffset;

    public string? StyleIdOf(int blockIndex)
    {
        short tableIndex = BlockStyleIds[blockIndex];
        return tableIndex >= 0 ? StyleIdTable[tableIndex] : null;
    }

    public string? ListLabelOf(int blockIndex)
    {
        int labelIndex = BlockListLabelIds[blockIndex];
        return labelIndex >= 0 ? ListLabels[labelIndex] : null;
    }

    public bool MatchesSource(long fileLength, long lastWriteTicksUtc)
        => fileLength == SourceFileLength && lastWriteTicksUtc == SourceLastWriteTicksUtc;

    public bool MatchesSource(string documentPath)
    {
        var info = new FileInfo(documentPath);
        return info.Exists && MatchesSource(info.Length, info.LastWriteTimeUtc.Ticks);
    }

    // ── serialisation ─────────────────────────────────────────────────────────────────

    public void Write(Stream destination)
    {
        var writer = new BinaryWriter(destination, Encoding.UTF8, leaveOpen: true);
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
        writer.Write(ListLabels.Length);
        foreach (string label in ListLabels) writer.Write(label);

        writer.Write(BlockStartOffsets.Length);
        foreach (long offset in BlockStartOffsets) writer.Write(offset);
        foreach (WordDocumentBlockKind kind in BlockKinds) writer.Write((byte)kind);
        foreach (int length in BlockTextLengths) writer.Write(length);
        foreach (short styleId in BlockStyleIds) writer.Write(styleId);
        foreach (int labelId in BlockListLabelIds) writer.Write(labelId);

        writer.Write(Headings.Length);
        foreach (WordDocumentHeading heading in Headings) { writer.Write(heading.BlockIndex); writer.Write((byte)heading.Level); writer.Write(heading.Text); }
        writer.Write(NoteReferences.Length);
        foreach (DocxNoteReference reference in NoteReferences)
        {
            writer.Write(reference.BlockIndex); writer.Write((byte)reference.Kind); writer.Write(reference.NoteId); writer.Write(reference.Number);
        }
        writer.Flush();
    }

    public static DocxRandomAccessIndex Read(Stream source)
    {
        var reader = new BinaryReader(source, Encoding.UTF8, leaveOpen: true);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not a word index.");
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
        var listLabels = new string[reader.ReadInt32()];
        for (int i = 0; i < listLabels.Length; i++) listLabels[i] = reader.ReadString();

        int blockCount = reader.ReadInt32();
        var starts = new long[blockCount];
        var kinds = new WordDocumentBlockKind[blockCount];
        var textLengths = new int[blockCount];
        var styleIds = new short[blockCount];
        var labelIds = new int[blockCount];
        for (int i = 0; i < blockCount; i++) starts[i] = reader.ReadInt64();
        for (int i = 0; i < blockCount; i++) kinds[i] = (WordDocumentBlockKind)reader.ReadByte();
        for (int i = 0; i < blockCount; i++) textLengths[i] = reader.ReadInt32();
        for (int i = 0; i < blockCount; i++) styleIds[i] = reader.ReadInt16();
        for (int i = 0; i < blockCount; i++) labelIds[i] = reader.ReadInt32();

        var headings = new WordDocumentHeading[reader.ReadInt32()];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(reader.ReadInt32(), reader.ReadByte(), reader.ReadString());
        var noteReferences = new DocxNoteReference[reader.ReadInt32()];
        for (int i = 0; i < noteReferences.Length; i++)
            noteReferences[i] = new DocxNoteReference(reader.ReadInt32(), (WordDocumentNoteKind)reader.ReadByte(), reader.ReadInt32(), reader.ReadInt32());

        return new DocxRandomAccessIndex(sourceLength, sourceTicks, entry, rootOpenTag, bodyEnd, checkpoints,
                                         starts, kinds, textLengths, styleIds, styleIdTable, labelIds, listLabels, headings, noteReferences);
    }

    // ── sidecar file helpers (file-system hosts) ───────────────────────────────────────

    public void Save(string indexPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(indexPath)!);
        string temporaryPath = indexPath + ".tmp";
        using (var file = File.Create(temporaryPath)) Write(file);
        if (File.Exists(indexPath)) File.Delete(indexPath);
        File.Move(temporaryPath, indexPath);
    }

    public static DocxRandomAccessIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using var file = File.OpenRead(indexPath);
            DocxRandomAccessIndex index = Read(file);
            return index.MatchesSource(documentPath) ? index : null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or EndOfStreamException)
        {
            return null;
        }
    }
}
