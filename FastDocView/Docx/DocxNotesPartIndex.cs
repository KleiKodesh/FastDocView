using SeekableDeflate;

namespace FastDocView.Docx;

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
        IndexBinaryArrays.WriteCompressedString(writer, RootOpenTag);
        DeflateCheckpointTable.Write(writer, Checkpoints);
        // ids ascend in Word's output but need not: zigzag deltas; starts ascend; an end is the next start unless stated
        var idBytes = new List<byte>(NoteIds.Length * 2);
        int previousId = 0;
        foreach (int id in NoteIds) { IndexBinaryArrays.AppendVarint(idBytes, DocxRandomAccessIndex.Zigzag(id - previousId)); previousId = id; }
        writer.Write(NoteIds.Length);
        writer.Write(idBytes.Count);
        writer.Write(idBytes.ToArray());
        IndexBinaryArrays.WriteDeltaVarints(writer, NoteStarts);
        IndexBinaryArrays.WriteBlockEnds(writer, NoteStarts, NoteEnds);
    }

    public static DocxNotesPartIndex Read(BinaryReader reader)
    {
        ZipEntryLocation entry = DocxRandomAccessIndex.ReadEntry(reader);
        string rootOpenTag = IndexBinaryArrays.ReadCompressedString(reader);
        List<DeflateCheckpoint> checkpoints = DeflateCheckpointTable.Read(reader);
        int count = reader.ReadInt32();
        byte[] idBytes = reader.ReadBytes(reader.ReadInt32());
        var ids = new int[count];
        int position = 0, id = 0;
        for (int i = 0; i < count; i++) { id += DocxRandomAccessIndex.Unzigzag(IndexBinaryArrays.ReadVarint(idBytes, ref position)); ids[i] = id; }
        long[] starts = IndexBinaryArrays.ReadDeltaVarints(reader);
        long[] ends = IndexBinaryArrays.ReadBlockEnds(reader, starts);
        if (starts.Length != count) throw new InvalidDataException("Word index: notes part tables do not match.");
        return new DocxNotesPartIndex(entry, rootOpenTag, checkpoints, ids, starts, ends);
    }
}
