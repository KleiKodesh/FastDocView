using System.Text;
using SeekableDeflate;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The index of one .docx body: where word/document.xml sits in the zip, the deflate checkpoints
/// (none in in-memory mode), and every body-level block's uncompressed offset + kind + style +
/// list label + text length — plus the same random-access index for footnotes.xml and endnotes.xml
/// (<see cref="DocxNotesPartIndex"/>), so notes are read on demand like blocks. Serialised to/from a <see cref="Stream"/> (the sidecar file on disk,
/// or whatever store a WASM host provides); about 3 % of the uncompressed XML at the default 1 MB
/// checkpoint spacing.
/// </summary>
public sealed class DocxRandomAccessIndex : IDocxBlockTable
{
    private const string Magic = "KHWORDIDX";
    private const int FormatVersion = 6;     // 6: list labels carry their suffix and level CSS

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
    public DocxListLabel[] ListLabels { get; }
    public WordDocumentHeading[] Headings { get; }
    public DocxNoteReference[] NoteReferences { get; }      // in body order
    public DocxNotesPartIndex? FootnotesPart { get; }
    public DocxNotesPartIndex? EndnotesPart { get; }

    public int BlockCount => BlockStartOffsets.Length;

    long IDocxBlockTable.BlockStartOffset(int blockIndex) => BlockStartOffsets[blockIndex];
    WordDocumentBlockKind IDocxBlockTable.BlockKindOf(int blockIndex) => BlockKinds[blockIndex];
    int IDocxBlockTable.BlockTextLengthOf(int blockIndex) => BlockTextLengths[blockIndex];

    public DocxRandomAccessIndex(long sourceFileLength, long sourceLastWriteTicksUtc, ZipEntryLocation documentEntry,
                                 string rootOpenTag, long bodyEndOffset, IReadOnlyList<DeflateCheckpoint> checkpoints,
                                 long[] blockStartOffsets, WordDocumentBlockKind[] blockKinds, int[] blockTextLengths,
                                 short[] blockStyleIds, string[] styleIdTable, int[] blockListLabelIds, DocxListLabel[] listLabels,
                                 WordDocumentHeading[] headings, DocxNoteReference[] noteReferences,
                                 DocxNotesPartIndex? footnotesPart, DocxNotesPartIndex? endnotesPart)
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
        FootnotesPart = footnotesPart;
        EndnotesPart = endnotesPart;
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

    public DocxListLabel? ListLabelOf(int blockIndex)
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
        WriteEntry(writer, DocumentEntry);
        writer.Write(RootOpenTag);
        writer.Write(BodyEndOffset);
        WriteCheckpoints(writer, Checkpoints);

        writer.Write(StyleIdTable.Length);
        foreach (string styleId in StyleIdTable) writer.Write(styleId);
        // label CSS strings repeat for every item of a level: written once, referenced by number
        var cssTable = new List<string>();
        var cssIds = new Dictionary<string, int>(StringComparer.Ordinal);
        writer.Write(ListLabels.Length);
        foreach (DocxListLabel label in ListLabels)
        {
            if (!cssIds.TryGetValue(label.LevelCss, out int cssId)) { cssId = cssTable.Count; cssTable.Add(label.LevelCss); cssIds[label.LevelCss] = cssId; }
            writer.Write(label.Text); writer.Write((byte)label.Suffix); writer.Write(cssId);
        }
        writer.Write(cssTable.Count);
        foreach (string css in cssTable) writer.Write(css);

        IndexBinaryArrays.Write(writer, BlockStartOffsets);
        IndexBinaryArrays.Write(writer, BlockKinds);
        IndexBinaryArrays.Write(writer, BlockTextLengths);
        IndexBinaryArrays.Write(writer, BlockStyleIds);
        IndexBinaryArrays.Write(writer, BlockListLabelIds);

        writer.Write(Headings.Length);
        foreach (WordDocumentHeading heading in Headings) { writer.Write(heading.BlockIndex); writer.Write((byte)heading.Level); writer.Write(heading.Text); }
        writer.Write(NoteReferences.Length);
        foreach (DocxNoteReference reference in NoteReferences)
        {
            writer.Write(reference.BlockIndex); writer.Write((byte)reference.Kind); writer.Write(reference.NoteId); writer.Write(reference.Number);
        }
        foreach (DocxNotesPartIndex? part in new[] { FootnotesPart, EndnotesPart })
        {
            writer.Write(part is not null);
            part?.Write(writer);
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
        ZipEntryLocation entry = ReadEntry(reader);
        string rootOpenTag = reader.ReadString();
        long bodyEnd = reader.ReadInt64();
        List<DeflateCheckpoint> checkpoints = ReadCheckpoints(reader);

        var styleIdTable = new string[reader.ReadInt32()];
        for (int i = 0; i < styleIdTable.Length; i++) styleIdTable[i] = reader.ReadString();
        var labelParts = new (string text, DocxListLabelSuffix suffix, int cssId)[reader.ReadInt32()];
        for (int i = 0; i < labelParts.Length; i++) labelParts[i] = (reader.ReadString(), (DocxListLabelSuffix)reader.ReadByte(), reader.ReadInt32());
        var cssTable = new string[reader.ReadInt32()];
        for (int i = 0; i < cssTable.Length; i++) cssTable[i] = reader.ReadString();
        var listLabels = new DocxListLabel[labelParts.Length];
        for (int i = 0; i < listLabels.Length; i++)
        {
            (string text, DocxListLabelSuffix suffix, int cssId) = labelParts[i];
            if ((uint)cssId >= (uint)cssTable.Length) throw new InvalidDataException("Word index: list label style out of range.");
            listLabels[i] = new DocxListLabel(text, suffix, cssTable[cssId]);
        }

        long[] starts = IndexBinaryArrays.Read<long>(reader);
        WordDocumentBlockKind[] kinds = IndexBinaryArrays.Read<WordDocumentBlockKind>(reader);
        int[] textLengths = IndexBinaryArrays.Read<int>(reader);
        short[] styleIds = IndexBinaryArrays.Read<short>(reader);
        int[] labelIds = IndexBinaryArrays.Read<int>(reader);

        var headings = new WordDocumentHeading[reader.ReadInt32()];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(reader.ReadInt32(), reader.ReadByte(), reader.ReadString());
        var noteReferences = new DocxNoteReference[reader.ReadInt32()];
        for (int i = 0; i < noteReferences.Length; i++)
            noteReferences[i] = new DocxNoteReference(reader.ReadInt32(), (WordDocumentNoteKind)reader.ReadByte(), reader.ReadInt32(), reader.ReadInt32());
        DocxNotesPartIndex? footnotesPart = reader.ReadBoolean() ? DocxNotesPartIndex.Read(reader) : null;
        DocxNotesPartIndex? endnotesPart = reader.ReadBoolean() ? DocxNotesPartIndex.Read(reader) : null;

        return new DocxRandomAccessIndex(sourceLength, sourceTicks, entry, rootOpenTag, bodyEnd, checkpoints,
                                         starts, kinds, textLengths, styleIds, styleIdTable, labelIds, listLabels, headings, noteReferences,
                                         footnotesPart, endnotesPart);
    }

    // ── shared record layouts (also used by DocxNotesPartIndex) ───────────────────────

    internal static void WriteEntry(BinaryWriter writer, ZipEntryLocation entry)
    {
        writer.Write(entry.Name);
        writer.Write(entry.DataOffset);
        writer.Write(entry.CompressedSize);
        writer.Write(entry.UncompressedSize);
        writer.Write(entry.CompressionMethod);
        writer.Write(entry.Crc32);
    }

    internal static ZipEntryLocation ReadEntry(BinaryReader reader)
        => new(reader.ReadString(), reader.ReadInt64(), reader.ReadInt64(), reader.ReadInt64(), reader.ReadUInt16(), reader.ReadUInt32());

    internal static void WriteCheckpoints(BinaryWriter writer, IReadOnlyList<DeflateCheckpoint> checkpoints)
    {
        writer.Write(checkpoints.Count);
        foreach (DeflateCheckpoint checkpoint in checkpoints)
        {
            writer.Write(checkpoint.CompressedBytePosition);
            writer.Write((byte)checkpoint.LeftoverBitCount);
            writer.Write((byte)checkpoint.LeftoverBits);
            writer.Write(checkpoint.UncompressedOffset);
            writer.Write(checkpoint.Window.Length);
            writer.Write(checkpoint.Window);
        }
    }

    internal static List<DeflateCheckpoint> ReadCheckpoints(BinaryReader reader)
    {
        int count = reader.ReadInt32();
        var checkpoints = new List<DeflateCheckpoint>(count);
        for (int i = 0; i < count; i++)
        {
            long compressedPosition = reader.ReadInt64();
            int leftoverBitCount = reader.ReadByte();
            int leftoverBits = reader.ReadByte();
            long uncompressedOffset = reader.ReadInt64();
            byte[] window = reader.ReadBytes(reader.ReadInt32());
            checkpoints.Add(new DeflateCheckpoint(compressedPosition, leftoverBitCount, leftoverBits, uncompressedOffset, window));
        }
        return checkpoints;
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
            using BinaryReader reader = IndexBinaryArrays.OpenInMemory(indexPath);
            DocxRandomAccessIndex index = Read(reader.BaseStream);
            return index.MatchesSource(documentPath) ? index : null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or EndOfStreamException)
        {
            return null;
        }
    }
}
