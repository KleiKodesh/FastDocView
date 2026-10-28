using SeekableDeflate;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The random-access index of footnotes.xml or endnotes.xml: where the part sits in the zip, its
/// deflate checkpoints (none in in-memory mode), the root's opening tag (namespaces for parsing one
/// note alone) and every note's id → uncompressed byte range. Lets a note be read on demand, like a
/// body block, instead of loading the whole part.
/// </summary>
public sealed class DocxNotesPartIndex
{
    public ZipEntryLocation Entry { get; }
    public string RootOpenTag { get; }
    public IReadOnlyList<DeflateCheckpoint> Checkpoints { get; }
    public int[] NoteIds { get; }
    public long[] NoteStarts { get; }
    public long[] NoteEnds { get; }

    private Dictionary<int, int>? _positionById;

    public DocxNotesPartIndex(ZipEntryLocation entry, string rootOpenTag, IReadOnlyList<DeflateCheckpoint> checkpoints,
                              int[] noteIds, long[] noteStarts, long[] noteEnds)
    {
        Entry = entry;
        RootOpenTag = rootOpenTag;
        Checkpoints = checkpoints;
        NoteIds = noteIds;
        NoteStarts = noteStarts;
        NoteEnds = noteEnds;
    }

    /// <summary>The byte range of a note, or null when the part has no note with that id.</summary>
    public (long start, long end)? RangeOf(int noteId)
    {
        _positionById ??= BuildPositions();
        return _positionById.TryGetValue(noteId, out int position) ? (NoteStarts[position], NoteEnds[position]) : null;
    }

    private Dictionary<int, int> BuildPositions()
    {
        var positions = new Dictionary<int, int>(NoteIds.Length);
        for (int i = 0; i < NoteIds.Length; i++) positions[NoteIds[i]] = i;
        return positions;
    }

    public void Write(BinaryWriter writer)
    {
        DocxRandomAccessIndex.WriteEntry(writer, Entry);
        writer.Write(RootOpenTag);
        DocxRandomAccessIndex.WriteCheckpoints(writer, Checkpoints);
        writer.Write(NoteIds.Length);
        for (int i = 0; i < NoteIds.Length; i++) { writer.Write(NoteIds[i]); writer.Write(NoteStarts[i]); writer.Write(NoteEnds[i]); }
    }

    public static DocxNotesPartIndex Read(BinaryReader reader)
    {
        ZipEntryLocation entry = DocxRandomAccessIndex.ReadEntry(reader);
        string rootOpenTag = reader.ReadString();
        List<DeflateCheckpoint> checkpoints = DocxRandomAccessIndex.ReadCheckpoints(reader);
        int count = reader.ReadInt32();
        var ids = new int[count];
        var starts = new long[count];
        var ends = new long[count];
        for (int i = 0; i < count; i++) { ids[i] = reader.ReadInt32(); starts[i] = reader.ReadInt64(); ends[i] = reader.ReadInt64(); }
        return new DocxNotesPartIndex(entry, rootOpenTag, checkpoints, ids, starts, ends);
    }
}
