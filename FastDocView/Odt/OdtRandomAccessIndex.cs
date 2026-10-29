using System.Text;
using SeekableDeflate;
using FastDocView.Docx;

namespace FastDocView.Odt;

/// <summary>
/// The index of one OpenDocument text: where content.xml sits in the zip and its deflate
/// checkpoints (a .fodt has neither — block offsets are plain file offsets), the root open tag,
/// the captured style sections (font faces, automatic styles; a .fodt's office:styles too), and per
/// block its byte range, kind, list depth + label and text length, plus headings and note
/// references with their bodies' byte ranges. Serialised to/from a <see cref="Stream"/>.
/// </summary>
public sealed class OdtRandomAccessIndex : IOdtBlockTable
{
    private const string Magic = "KHODTIDX";
    private const int FormatVersion = 4;     // 2: section text widths; 3: list label level geometry; 4: bookmarks (name -> block)

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public bool IsFlat { get; }
    public ZipEntryLocation? ContentEntry { get; }           // null for .fodt
    public IReadOnlyList<DeflateCheckpoint> Checkpoints { get; }
    public string RootOpenTag { get; }
    public string[] StyleSections { get; }
    public long[] BlockStartOffsets { get; }
    public long[] BlockEndOffsets { get; }
    public WordDocumentBlockKind[] BlockKinds { get; }
    public int[] BlockTextLengths { get; }
    public byte[] BlockListLevels { get; }                    // 255 = not a list paragraph
    public int[] BlockListLabelIds { get; }
    public string[] ListLabels { get; }
    /// <summary>Per list label: its level geometry (<see cref="OdtListStyle.Level.Layout"/>), parallel to <see cref="ListLabels"/>.</summary>
    public string[] ListLabelLayouts { get; init; } = [];
    public WordDocumentHeading[] Headings { get; }
    public OdtNoteReference[] NoteReferences { get; }        // body order
    /// <summary>Per section: its first block and its page text width in twips (see OdtPageSetup).</summary>
    public int[] SectionFirstBlocks { get; init; } = [];
    public int[] SectionTextWidths { get; init; } = [];
    /// <summary>Bookmarks and reference marks (name -> block), see OdtBodyBlockScanner.BookmarkNames.</summary>
    public WordDocumentBookmarkTable Bookmarks { get; init; } = WordDocumentBookmarkTable.Empty;

    public int BlockCount => BlockStartOffsets.Length;

    public OdtRandomAccessIndex(long sourceFileLength, long sourceLastWriteTicksUtc, bool isFlat, ZipEntryLocation? contentEntry,
                                IReadOnlyList<DeflateCheckpoint> checkpoints, string rootOpenTag, string[] styleSections,
                                long[] blockStartOffsets, long[] blockEndOffsets, WordDocumentBlockKind[] blockKinds, int[] blockTextLengths,
                                byte[] blockListLevels, int[] blockListLabelIds, string[] listLabels,
                                WordDocumentHeading[] headings, OdtNoteReference[] noteReferences)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        IsFlat = isFlat;
        ContentEntry = contentEntry;
        Checkpoints = checkpoints;
        RootOpenTag = rootOpenTag;
        StyleSections = styleSections;
        BlockStartOffsets = blockStartOffsets;
        BlockEndOffsets = blockEndOffsets;
        BlockKinds = blockKinds;
        BlockTextLengths = blockTextLengths;
        BlockListLevels = blockListLevels;
        BlockListLabelIds = blockListLabelIds;
        ListLabels = listLabels;
        Headings = headings;
        NoteReferences = noteReferences;
    }

    long IOdtBlockTable.BlockStartOffset(int blockIndex) => BlockStartOffsets[blockIndex];
    long IOdtBlockTable.BlockEndOffset(int blockIndex) => BlockEndOffsets[blockIndex];
    WordDocumentBlockKind IOdtBlockTable.BlockKindOf(int blockIndex) => BlockKinds[blockIndex];
    int IOdtBlockTable.BlockTextLengthOf(int blockIndex) => BlockTextLengths[blockIndex];

    public string? ListLabelOf(int blockIndex)
    {
        int labelIndex = BlockListLabelIds[blockIndex];
        return labelIndex >= 0 ? ListLabels[labelIndex] : null;
    }

    public string ListLabelLayoutOf(int blockIndex)
    {
        int labelIndex = BlockListLabelIds[blockIndex];
        return labelIndex >= 0 && labelIndex < ListLabelLayouts.Length ? ListLabelLayouts[labelIndex] : "";
    }

    /// <summary>The list depth (0-based) of a list paragraph, or -1.</summary>
    public int ListLevelOf(int blockIndex) => BlockListLevels[blockIndex] == 255 ? -1 : BlockListLevels[blockIndex];

    /// <summary>The note references inside blocks [first, last] — a slice of the body-ordered array.</summary>
    public ReadOnlySpan<OdtNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex)
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
        writer.Write(IsFlat);
        writer.Write(ContentEntry is not null);
        if (ContentEntry is not null) DocxRandomAccessIndex.WriteEntry(writer, ContentEntry);
        DocxRandomAccessIndex.WriteCheckpoints(writer, Checkpoints);
        writer.Write(RootOpenTag);
        WriteStrings(writer, StyleSections);
        WriteStrings(writer, ListLabels);
        WriteStrings(writer, ListLabelLayouts);
        IndexBinaryArrays.Write(writer, BlockStartOffsets);
        IndexBinaryArrays.Write(writer, BlockEndOffsets);
        IndexBinaryArrays.Write(writer, BlockKinds);
        IndexBinaryArrays.Write(writer, BlockTextLengths);
        IndexBinaryArrays.Write(writer, BlockListLevels);
        IndexBinaryArrays.Write(writer, BlockListLabelIds);
        writer.Write(Headings.Length);
        foreach (WordDocumentHeading heading in Headings) { writer.Write(heading.BlockIndex); writer.Write((byte)heading.Level); writer.Write(heading.Text); }
        writer.Write(NoteReferences.Length);
        foreach (OdtNoteReference note in NoteReferences)
        {
            writer.Write(note.BlockIndex); writer.Write((byte)note.Kind); writer.Write(note.Number); writer.Write(note.Label);
            writer.Write(note.BodyStart); writer.Write(note.BodyEnd);
        }
        IndexBinaryArrays.Write(writer, SectionFirstBlocks);
        IndexBinaryArrays.Write(writer, SectionTextWidths);
        Bookmarks.Write(writer);
        writer.Flush();
    }

    public static OdtRandomAccessIndex Read(Stream source)
    {
        var reader = new BinaryReader(source, Encoding.UTF8, leaveOpen: true);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not an OpenDocument index.");
        if (reader.ReadInt32() != FormatVersion) throw new InvalidDataException("OpenDocument index: unsupported version.");
        long sourceLength = reader.ReadInt64();
        long sourceTicks = reader.ReadInt64();
        bool isFlat = reader.ReadBoolean();
        ZipEntryLocation? entry = reader.ReadBoolean() ? DocxRandomAccessIndex.ReadEntry(reader) : null;
        List<DeflateCheckpoint> checkpoints = DocxRandomAccessIndex.ReadCheckpoints(reader);
        string rootOpenTag = reader.ReadString();
        string[] styleSections = ReadStrings(reader);
        string[] listLabels = ReadStrings(reader);
        string[] listLabelLayouts = ReadStrings(reader);
        if (listLabelLayouts.Length != listLabels.Length) throw new InvalidDataException("OpenDocument index: list label geometry does not match the labels.");
        long[] starts = IndexBinaryArrays.Read<long>(reader);
        long[] ends = IndexBinaryArrays.Read<long>(reader);
        WordDocumentBlockKind[] kinds = IndexBinaryArrays.Read<WordDocumentBlockKind>(reader);
        int[] textLengths = IndexBinaryArrays.Read<int>(reader);
        byte[] listLevels = IndexBinaryArrays.Read<byte>(reader);
        int[] labelIds = IndexBinaryArrays.Read<int>(reader);
        var headings = new WordDocumentHeading[reader.ReadInt32()];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(reader.ReadInt32(), reader.ReadByte(), reader.ReadString());
        var notes = new OdtNoteReference[reader.ReadInt32()];
        for (int i = 0; i < notes.Length; i++)
            notes[i] = new OdtNoteReference(reader.ReadInt32(), (WordDocumentNoteKind)reader.ReadByte(), reader.ReadInt32(), reader.ReadString(), reader.ReadInt64(), reader.ReadInt64());
        int[] sectionFirstBlocks = IndexBinaryArrays.Read<int>(reader);
        int[] sectionTextWidths = IndexBinaryArrays.Read<int>(reader);
        if (sectionFirstBlocks.Length != sectionTextWidths.Length || sectionFirstBlocks.Any(block => (uint)block >= (uint)Math.Max(1, starts.Length)))
            throw new InvalidDataException("OpenDocument index: section table does not match the blocks.");
        WordDocumentBookmarkTable bookmarks = WordDocumentBookmarkTable.Read(reader, starts.Length);
        return new OdtRandomAccessIndex(sourceLength, sourceTicks, isFlat, entry, checkpoints, rootOpenTag, styleSections,
                                        starts, ends, kinds, textLengths, listLevels, labelIds, listLabels, headings, notes)
        {
            ListLabelLayouts = listLabelLayouts,
            SectionFirstBlocks = sectionFirstBlocks,
            SectionTextWidths = sectionTextWidths,
            Bookmarks = bookmarks,
        };
    }

    private static void WriteStrings(BinaryWriter writer, string[] values)
    {
        writer.Write(values.Length);
        foreach (string value in values) writer.Write(value);
    }

    private static string[] ReadStrings(BinaryReader reader)
    {
        var values = new string[reader.ReadInt32()];
        for (int i = 0; i < values.Length; i++) values[i] = reader.ReadString();
        return values;
    }

    // ── sidecar file helpers ──────────────────────────────────────────────────────────

    public void Save(string indexPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(indexPath)!);
        string temporaryPath = indexPath + ".tmp";
        using (var file = File.Create(temporaryPath)) Write(file);
        if (File.Exists(indexPath)) File.Delete(indexPath);
        File.Move(temporaryPath, indexPath);
    }

    public static OdtRandomAccessIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader reader = IndexBinaryArrays.OpenInMemory(indexPath);
            OdtRandomAccessIndex index = Read(reader.BaseStream);
            return index.MatchesSource(documentPath) ? index : null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or EndOfStreamException)
        {
            return null;
        }
    }
}
