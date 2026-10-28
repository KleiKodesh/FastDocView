using System.Text;

namespace WordDocumentRandomAccess.Rtf;

/// <summary>A note referenced in a block: kind, number among the auto-numbered notes of its kind (0 = custom mark), label.</summary>
public readonly record struct RtfNoteReference(int BlockIndex, WordDocumentNoteKind Kind, int Number, string Label);

/// <summary>
/// The index of one RTF file. RTF is plain bytes, so a block is a byte range of the file itself — no checkpoints —
/// plus the group-state stack in effect at its first byte (interned: most blocks share a handful of stacks), so a
/// window parses only its own bytes. Also per block: kind, text length, list label (counted at index time when the
/// paragraph carries no \listtext), heading, note references; and the header tables rendering needs.
/// </summary>
public sealed class RtfRandomAccessIndex
{
    private const string Magic = "KHRTFIDX";
    private const int FormatVersion = 1;

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public RtfDocumentTables Tables { get; }
    public bool IsRightToLeft { get; }
    public long[] BlockStartOffsets { get; }
    public long[] BlockEndOffsets { get; }
    public WordDocumentBlockKind[] BlockKinds { get; }
    public int[] BlockTextLengths { get; }
    public int[] BlockStateIds { get; }
    public byte[][] States { get; }
    public byte[] BlockListLevels { get; }                    // 255 = no label from the index
    public int[] BlockListLabelIds { get; }
    public string[] ListLabels { get; }
    public WordDocumentHeading[] Headings { get; }
    public RtfNoteReference[] NoteReferences { get; }         // body order

    public int BlockCount => BlockStartOffsets.Length;

    public RtfRandomAccessIndex(long sourceFileLength, long sourceLastWriteTicksUtc, RtfDocumentTables tables, bool isRightToLeft,
                                long[] blockStartOffsets, long[] blockEndOffsets, WordDocumentBlockKind[] blockKinds, int[] blockTextLengths,
                                int[] blockStateIds, byte[][] states, byte[] blockListLevels, int[] blockListLabelIds, string[] listLabels,
                                WordDocumentHeading[] headings, RtfNoteReference[] noteReferences)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        Tables = tables;
        IsRightToLeft = isRightToLeft;
        BlockStartOffsets = blockStartOffsets;
        BlockEndOffsets = blockEndOffsets;
        BlockKinds = blockKinds;
        BlockTextLengths = blockTextLengths;
        BlockStateIds = blockStateIds;
        States = states;
        BlockListLevels = blockListLevels;
        BlockListLabelIds = blockListLabelIds;
        ListLabels = listLabels;
        Headings = headings;
        NoteReferences = noteReferences;
    }

    public string? ListLabelOf(int blockIndex)
    {
        int labelIndex = BlockListLabelIds[blockIndex];
        return labelIndex >= 0 ? ListLabels[labelIndex] : null;
    }

    public int ListLevelOf(int blockIndex) => BlockListLevels[blockIndex] == 255 ? -1 : BlockListLevels[blockIndex];

    public ReadOnlySpan<RtfNoteReference> NoteReferencesInBlock(int blockIndex)
    {
        int start = LowerBound(blockIndex), end = LowerBound(blockIndex + 1);
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
        writer.Write(RtfGroupState.Size);                     // the snapshots are raw RtfGroupState bytes
        writer.Write(SourceFileLength);
        writer.Write(SourceLastWriteTicksUtc);
        writer.Write(IsRightToLeft);
        Tables.Write(writer);
        writer.Write(States.Length);
        foreach (byte[] state in States) { writer.Write(state.Length); writer.Write(state); }
        writer.Write(ListLabels.Length);
        foreach (string label in ListLabels) writer.Write(label);
        IndexBinaryArrays.Write(writer, BlockStartOffsets);
        IndexBinaryArrays.Write(writer, BlockEndOffsets);
        IndexBinaryArrays.Write(writer, BlockKinds);
        IndexBinaryArrays.Write(writer, BlockTextLengths);
        IndexBinaryArrays.Write(writer, BlockStateIds);
        IndexBinaryArrays.Write(writer, BlockListLevels);
        IndexBinaryArrays.Write(writer, BlockListLabelIds);
        writer.Write(Headings.Length);
        foreach (WordDocumentHeading heading in Headings) { writer.Write(heading.BlockIndex); writer.Write((byte)heading.Level); writer.Write(heading.Text); }
        writer.Write(NoteReferences.Length);
        foreach (RtfNoteReference note in NoteReferences) { writer.Write(note.BlockIndex); writer.Write((byte)note.Kind); writer.Write(note.Number); writer.Write(note.Label); }
        writer.Flush();
    }

    public static RtfRandomAccessIndex Read(Stream source)
    {
        var reader = new BinaryReader(source, Encoding.UTF8, leaveOpen: true);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not an RTF index.");
        if (reader.ReadInt32() != FormatVersion) throw new InvalidDataException("RTF index: unsupported version.");
        if (reader.ReadInt32() != RtfGroupState.Size) throw new InvalidDataException("RTF index: group-state layout changed.");
        long sourceLength = reader.ReadInt64();
        long sourceTicks = reader.ReadInt64();
        bool isRightToLeft = reader.ReadBoolean();
        RtfDocumentTables tables = RtfDocumentTables.Read(reader);
        var states = new byte[reader.ReadInt32()][];
        for (int i = 0; i < states.Length; i++) states[i] = reader.ReadBytes(reader.ReadInt32());
        var labels = new string[reader.ReadInt32()];
        for (int i = 0; i < labels.Length; i++) labels[i] = reader.ReadString();
        long[] starts = IndexBinaryArrays.Read<long>(reader);
        long[] ends = IndexBinaryArrays.Read<long>(reader);
        WordDocumentBlockKind[] kinds = IndexBinaryArrays.Read<WordDocumentBlockKind>(reader);
        int[] textLengths = IndexBinaryArrays.Read<int>(reader);
        int[] stateIds = IndexBinaryArrays.Read<int>(reader);
        byte[] listLevels = IndexBinaryArrays.Read<byte>(reader);
        int[] labelIds = IndexBinaryArrays.Read<int>(reader);
        var headings = new WordDocumentHeading[reader.ReadInt32()];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(reader.ReadInt32(), reader.ReadByte(), reader.ReadString());
        var notes = new RtfNoteReference[reader.ReadInt32()];
        for (int i = 0; i < notes.Length; i++) notes[i] = new RtfNoteReference(reader.ReadInt32(), (WordDocumentNoteKind)reader.ReadByte(), reader.ReadInt32(), reader.ReadString());
        return new RtfRandomAccessIndex(sourceLength, sourceTicks, tables, isRightToLeft, starts, ends, kinds, textLengths, stateIds, states,
                                        listLevels, labelIds, labels, headings, notes);
    }

    public void Save(string indexPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(indexPath)!);
        string temporaryPath = indexPath + ".tmp";
        using (var file = File.Create(temporaryPath)) Write(file);
        if (File.Exists(indexPath)) File.Delete(indexPath);
        File.Move(temporaryPath, indexPath);
    }

    public static RtfRandomAccessIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader reader = IndexBinaryArrays.OpenInMemory(indexPath);
            RtfRandomAccessIndex index = Read(reader.BaseStream);
            return index.MatchesSource(documentPath) ? index : null;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            return null;                                     // a corrupt sidecar is only a cache miss
        }
    }
}
