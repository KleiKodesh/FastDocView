namespace FastDocView.Rtf;

/// <summary>
/// What the compact table would recover by rescanning a span, stored at index time for each block that holds a note: its
/// byte range, kind, list label, the fields open at its start and the parser snapshot there (both interned). A notes page
/// renders those blocks straight from their bytes (<see cref="RtfCompactBlockTable"/> answers from here first) instead of
/// rescanning the blocks between them. The values are the index pass's own, so the HTML is the one a span rescan gives.
/// </summary>
internal sealed class RtfNoteBlockTable
{
    public static readonly RtfNoteBlockTable Empty = new([], [], [], [], [], [], [], [], []);

    public int[] Blocks { get; }               // ascending, distinct
    private readonly long[] _starts, _ends;
    private readonly byte[] _kinds;
    private readonly string?[] _labels;
    private readonly int[] _fieldStateIds, _snapshotIds;
    private readonly string[] _fieldStates;
    private readonly byte[][] _snapshots;

    private RtfNoteBlockTable(int[] blocks, long[] starts, long[] ends, byte[] kinds, string?[] labels,
                              int[] fieldStateIds, string[] fieldStates, int[] snapshotIds, byte[][] snapshots)
    {
        Blocks = blocks;
        _starts = starts;
        _ends = ends;
        _kinds = kinds;
        _labels = labels;
        _fieldStateIds = fieldStateIds;
        _fieldStates = fieldStates;
        _snapshotIds = snapshotIds;
        _snapshots = snapshots;
    }

    /// <summary>
    /// The table of <paramref name="blocks"/> (ascending, distinct, all ended). <paramref name="fieldStateIdOf"/> / <paramref name="snapshotIdOf"/>
    /// give the index pass's interned ids, <paramref name="fieldStates"/> / <paramref name="snapshots"/> their values.
    /// </summary>
    public static RtfNoteBlockTable From(int[] blocks, IReadOnlyList<long> starts, IReadOnlyList<long> ends, IReadOnlyList<WordDocumentBlockKind> kinds,
                                         Func<int, string?> labelOf, Func<int, int> fieldStateIdOf, IReadOnlyList<string> fieldStates,
                                         Func<int, int> snapshotIdOf, IReadOnlyList<byte[]> snapshots)
    {
        var fieldIds = new int[blocks.Length];
        var snapshotIds = new int[blocks.Length];
        var keptFields = new List<string>();
        var keptSnapshots = new List<byte[]>();
        var fieldMap = new Dictionary<int, int>();
        var snapshotMap = new Dictionary<int, int>();
        for (int i = 0; i < blocks.Length; i++)
        {
            int field = fieldStateIdOf(blocks[i]);
            if (!fieldMap.TryGetValue(field, out int fieldId)) { fieldId = keptFields.Count; keptFields.Add(fieldStates[field]); fieldMap[field] = fieldId; }
            fieldIds[i] = fieldId;
            int snapshot = snapshotIdOf(blocks[i]);
            if (!snapshotMap.TryGetValue(snapshot, out int snapshotId)) { snapshotId = keptSnapshots.Count; keptSnapshots.Add(snapshots[snapshot]); snapshotMap[snapshot] = snapshotId; }
            snapshotIds[i] = snapshotId;
        }
        return new RtfNoteBlockTable(blocks, blocks.Select(block => starts[block]).ToArray(), blocks.Select(block => ends[block]).ToArray(),
                                     blocks.Select(block => (byte)kinds[block]).ToArray(), blocks.Select(labelOf).ToArray(),
                                     fieldIds, keptFields.ToArray(), snapshotIds, keptSnapshots.ToArray());
    }

    /// <summary>The block's position in this table, or -1 when it holds no note.</summary>
    public int Find(int blockIndex) => Blocks.Length == 0 ? -1 : Math.Max(-1, Array.BinarySearch(Blocks, blockIndex));

    public long StartOf(int at) => _starts[at];
    public long EndOf(int at) => _ends[at];
    public WordDocumentBlockKind KindOf(int at) => (WordDocumentBlockKind)_kinds[at];
    public string? ListLabelOf(int at) => _labels[at];
    public string FieldStateOf(int at) => _fieldStates[_fieldStateIds[at]];
    public byte[] SnapshotOf(int at) => _snapshots[_snapshotIds[at]];

    public void Write(BinaryWriter writer)
    {
        new EliasFanoSequence(Blocks).Write(writer);
        IndexBinaryArrays.WriteDeltaVarints(writer, _starts);
        IndexBinaryArrays.WriteVarints(writer, _ends.Select((end, i) => checked((int)(end - _starts[i]))).ToArray());
        IndexBinaryArrays.WriteVarints(writer, _fieldStateIds);
        IndexBinaryArrays.WriteVarints(writer, _snapshotIds);
        var raw = new List<byte>();
        foreach (byte kind in _kinds) raw.Add(kind);
        foreach (string? label in _labels) raw.Add(label is null ? (byte)0 : (byte)1);
        IndexBinaryArrays.AppendVarint(raw, (uint)_snapshots.Length);
        foreach (byte[] snapshot in _snapshots)
        {
            IndexBinaryArrays.AppendVarint(raw, (uint)snapshot.Length);
            raw.AddRange(snapshot);
        }
        // labels and field states as length-prefixed UTF-8 (they may hold U+0000, which the compressed string lists do not allow)
        foreach (string? label in _labels) AppendText(raw, label ?? "");
        IndexBinaryArrays.AppendVarint(raw, (uint)_fieldStates.Length);
        foreach (string state in _fieldStates) AppendText(raw, state);
        WordIndexDeflatedBlob.Write(writer, raw.ToArray());
    }

    private static void AppendText(List<byte> raw, string text)
    {
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(text);
        IndexBinaryArrays.AppendVarint(raw, (uint)bytes.Length);
        raw.AddRange(bytes);
    }

    private static string ReadText(byte[] raw, ref int position)
    {
        int length = (int)IndexBinaryArrays.ReadVarint(raw, ref position);
        if (length < 0 || length > raw.Length - position) throw new InvalidDataException("RTF index: note block text damaged.");
        string text = System.Text.Encoding.UTF8.GetString(raw, position, length);
        position += length;
        return text;
    }

    public static RtfNoteBlockTable Read(BinaryReader reader, int blockCount, long lastBlockEnd)
    {
        int[] blocks = EliasFanoSequence.Read(reader).ToArray();
        long[] starts = IndexBinaryArrays.ReadDeltaVarints(reader);
        int[] lengths = IndexBinaryArrays.ReadVarints(reader);
        int[] fieldStateIds = IndexBinaryArrays.ReadVarints(reader);
        int[] snapshotIds = IndexBinaryArrays.ReadVarints(reader);
        byte[] raw = WordIndexDeflatedBlob.Read(reader);
        int count = blocks.Length;
        if (starts.Length != count || lengths.Length != count || fieldStateIds.Length != count || snapshotIds.Length != count
            || raw.Length < count * 2 || blocks.Any(block => (uint)block >= (uint)blockCount))
            throw new InvalidDataException("RTF index: note blocks do not match the blocks.");
        var ends = new long[count];
        for (int i = 0; i < count; i++)
        {
            ends[i] = starts[i] + lengths[i];
            if (lengths[i] < 0 || ends[i] > lastBlockEnd || (i > 0 && blocks[i] <= blocks[i - 1])) throw new InvalidDataException("RTF index: note blocks damaged.");
        }
        byte[] kinds = raw.AsSpan(0, count).ToArray();
        int position = 2 * count;
        int snapshotCount = (int)IndexBinaryArrays.ReadVarint(raw, ref position);
        if (snapshotCount < 0 || snapshotCount > raw.Length) throw new InvalidDataException("RTF index: note block snapshots damaged.");
        var snapshots = new byte[snapshotCount][];
        for (int i = 0; i < snapshotCount; i++)
        {
            int length = (int)IndexBinaryArrays.ReadVarint(raw, ref position);
            if (length < 0 || length > raw.Length - position) throw new InvalidDataException("RTF index: note block snapshots damaged.");
            snapshots[i] = raw.AsSpan(position, length).ToArray();
            position += length;
        }
        if (snapshotIds.Any(id => (uint)id >= (uint)snapshotCount)) throw new InvalidDataException("RTF index: note block snapshots damaged.");
        var labels = new string?[count];
        for (int i = 0; i < count; i++)
        {
            string text = ReadText(raw, ref position);
            labels[i] = raw[count + i] != 0 ? text : null;
        }
        int stateCount = (int)IndexBinaryArrays.ReadVarint(raw, ref position);
        if (stateCount < 0 || stateCount > raw.Length) throw new InvalidDataException("RTF index: note block field states damaged.");
        var fieldStates = new string[stateCount];
        for (int i = 0; i < stateCount; i++) fieldStates[i] = ReadText(raw, ref position);
        if (fieldStateIds.Any(id => (uint)id >= (uint)stateCount)) throw new InvalidDataException("RTF index: note block field states damaged.");
        return new RtfNoteBlockTable(blocks, starts, ends, kinds, labels, fieldStateIds, fieldStates, snapshotIds, snapshots);
    }
}
