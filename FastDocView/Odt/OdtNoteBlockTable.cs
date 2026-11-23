namespace FastDocView.Odt;

/// <summary>
/// What the compact table would recover by rescanning a span, stored at index time for each block that holds a note: its
/// byte range, kind, text length, list level / label / layout and field state. A notes page renders those blocks straight
/// from their bytes (<see cref="OdtCompactBlockTable"/> answers from here first) instead of rescanning the ~200 blocks between
/// them. The values are the index pass's own, so the HTML is the one a span rescan gives.
/// </summary>
internal sealed class OdtNoteBlockTable
{
    public static readonly OdtNoteBlockTable Empty = new([], [], [], [], [], [], [], [], [], []);

    public int[] Blocks { get; }               // ascending, distinct
    private readonly long[] _starts, _ends;
    private readonly byte[] _kinds, _listLevels;
    private readonly int[] _textLengths;
    private readonly string?[] _labels;
    private readonly string[] _labelLayouts;
    private readonly int[] _fieldStateIds;
    private readonly OdtFieldState[] _fieldStates;

    private OdtNoteBlockTable(int[] blocks, long[] starts, long[] ends, byte[] kinds, int[] textLengths, byte[] listLevels,
                              string?[] labels, string[] labelLayouts, int[] fieldStateIds, OdtFieldState[] fieldStates)
    {
        Blocks = blocks;
        _starts = starts;
        _ends = ends;
        _kinds = kinds;
        _textLengths = textLengths;
        _listLevels = listLevels;
        _labels = labels;
        _labelLayouts = labelLayouts;
        _fieldStateIds = fieldStateIds;
        _fieldStates = fieldStates;
    }

    /// <summary>The note-holding blocks of a finished index pass (every block below <paramref name="blockCount"/> has ended).</summary>
    public static OdtNoteBlockTable From(OdtBodyBlockScanner scanner, int blockCount)
    {
        int[] blocks = scanner.NoteReferences.Select(reference => reference.BlockIndex).Where(block => block < blockCount).Distinct().OrderBy(block => block).ToArray();
        var fieldStateIdByBytes = new Dictionary<string, int>(StringComparer.Ordinal);      // equal states share one stored copy
        var fieldStates = new List<OdtFieldState>();
        var fieldStateIds = new int[blocks.Length];
        var labels = new string?[blocks.Length];
        var layouts = new string[blocks.Length];
        for (int i = 0; i < blocks.Length; i++)
        {
            int block = blocks[i];
            int labelId = scanner.BlockListLabelIds[block];
            labels[i] = labelId >= 0 ? scanner.ListLabels[labelId] : null;
            layouts[i] = labelId >= 0 ? scanner.ListLabelLayouts[labelId] : "";
            OdtFieldState state = scanner.BlockFieldStates[block];
            var bytes = new List<byte>();
            state.Write(bytes);
            string key = Convert.ToBase64String(bytes.ToArray());
            if (!fieldStateIdByBytes.TryGetValue(key, out int id))
            {
                id = fieldStates.Count;
                fieldStates.Add(state);
                fieldStateIdByBytes[key] = id;
            }
            fieldStateIds[i] = id;
        }
        return new OdtNoteBlockTable(blocks, blocks.Select(block => scanner.BlockStartOffsets[block]).ToArray(), blocks.Select(block => scanner.BlockEndOffsets[block]).ToArray(),
                                     blocks.Select(block => (byte)scanner.BlockKinds[block]).ToArray(), blocks.Select(block => scanner.BlockTextLengths[block]).ToArray(),
                                     blocks.Select(block => scanner.BlockListLevels[block]).ToArray(), labels, layouts, fieldStateIds, fieldStates.ToArray());
    }

    /// <summary>The block's position in this table, or -1 when it holds no note.</summary>
    public int Find(int blockIndex) => Blocks.Length == 0 ? -1 : Math.Max(-1, Array.BinarySearch(Blocks, blockIndex));

    public long StartOf(int at) => _starts[at];
    public long EndOf(int at) => _ends[at];
    public WordDocumentBlockKind KindOf(int at) => (WordDocumentBlockKind)_kinds[at];
    public int TextLengthOf(int at) => _textLengths[at];
    public int ListLevelOf(int at) => _listLevels[at] == 255 ? -1 : _listLevels[at];
    public string? ListLabelOf(int at) => _labels[at];
    public string ListLabelLayoutOf(int at) => _labelLayouts[at];
    public OdtFieldState FieldStateOf(int at) => _fieldStates[_fieldStateIds[at]];

    public void Write(BinaryWriter writer)
    {
        new EliasFanoSequence(Blocks).Write(writer);
        IndexBinaryArrays.WriteDeltaVarints(writer, _starts);
        IndexBinaryArrays.WriteVarints(writer, _ends.Select((end, i) => checked((int)(end - _starts[i]))).ToArray());
        IndexBinaryArrays.WriteVarints(writer, _textLengths);
        IndexBinaryArrays.WriteVarints(writer, _fieldStateIds);
        var raw = new List<byte>();
        foreach (byte kind in _kinds) raw.Add(kind);
        foreach (byte level in _listLevels) raw.Add(level);
        foreach (string? label in _labels) raw.Add(label is null ? (byte)0 : (byte)1);
        IndexBinaryArrays.AppendVarint(raw, (uint)_fieldStates.Length);
        foreach (OdtFieldState state in _fieldStates)
        {
            var bytes = new List<byte>();
            state.Write(bytes);
            IndexBinaryArrays.AppendVarint(raw, (uint)bytes.Count);
            raw.AddRange(bytes);
        }
        WordIndexDeflatedBlob.Write(writer, raw.ToArray());
        IndexBinaryArrays.WriteCompressedStrings(writer, _labels.Select(label => label ?? "").ToArray());
        IndexBinaryArrays.WriteCompressedStrings(writer, _labelLayouts);
    }

    public static OdtNoteBlockTable Read(BinaryReader reader, int blockCount, long lastBlockEnd)
    {
        int[] blocks = EliasFanoSequence.Read(reader).ToArray();
        long[] starts = IndexBinaryArrays.ReadDeltaVarints(reader);
        int[] lengths = IndexBinaryArrays.ReadVarints(reader);
        int[] textLengths = IndexBinaryArrays.ReadVarints(reader);
        int[] fieldStateIds = IndexBinaryArrays.ReadVarints(reader);
        byte[] raw = WordIndexDeflatedBlob.Read(reader);
        string[] labelTexts = IndexBinaryArrays.ReadCompressedStrings(reader);
        string[] layouts = IndexBinaryArrays.ReadCompressedStrings(reader);
        int count = blocks.Length;
        if (starts.Length != count || lengths.Length != count || textLengths.Length != count || fieldStateIds.Length != count || labelTexts.Length != count
            || layouts.Length != count || raw.Length < count * 3 || blocks.Any(block => (uint)block >= (uint)blockCount))
            throw new InvalidDataException("OpenDocument index: note blocks do not match the blocks.");
        var ends = new long[count];
        for (int i = 0; i < count; i++)
        {
            ends[i] = starts[i] + lengths[i];
            if (lengths[i] < 0 || ends[i] > lastBlockEnd || (i > 0 && blocks[i] <= blocks[i - 1])) throw new InvalidDataException("OpenDocument index: note blocks damaged.");
        }
        byte[] kinds = raw.AsSpan(0, count).ToArray();
        byte[] levels = raw.AsSpan(count, count).ToArray();
        var labels = new string?[count];
        for (int i = 0; i < count; i++) labels[i] = raw[2 * count + i] != 0 ? labelTexts[i] : null;
        int position = 3 * count;
        int stateCount = (int)IndexBinaryArrays.ReadVarint(raw, ref position);
        if (stateCount < 0 || stateCount > raw.Length) throw new InvalidDataException("OpenDocument index: note block field states damaged.");
        var states = new OdtFieldState[stateCount];
        for (int i = 0; i < stateCount; i++)
        {
            int length = (int)IndexBinaryArrays.ReadVarint(raw, ref position);
            if (length < 0 || length > raw.Length - position) throw new InvalidDataException("OpenDocument index: note block field states damaged.");
            states[i] = OdtFieldState.Read(raw.AsSpan(position, length).ToArray());
            position += length;
        }
        if (fieldStateIds.Any(id => (uint)id >= (uint)stateCount)) throw new InvalidDataException("OpenDocument index: note block field states damaged.");
        return new OdtNoteBlockTable(blocks, starts, ends, kinds, textLengths, levels, labels, layouts, fieldStateIds, states);
    }
}
