namespace FastDocView;

/// <summary>
/// The serialised numbering state of a list counter (every list's counters + the start overrides already applied), as
/// the compact indexes store it per span anchor: zigzag varints, a few bytes per list; deltas between anchors hold
/// only the lists that changed. The .doc counter (<see cref="DocBinary.DocBinaryListCounter"/>) and the RTF / ODT
/// counters share it; the .docx counter keeps its own identical copy (DocxHtml.DocxListCounter).
/// </summary>
internal static class WordListCounterState
{
    public static byte[] Serialize(Dictionary<int, int[]> lists, HashSet<(int listId, int level)> overrides)
    {
        if (lists.Count == 0 && overrides.Count == 0) return [];
        var bytes = new List<byte>();
        IndexBinaryArrays.AppendVarint(bytes, (uint)lists.Count);
        foreach (KeyValuePair<int, int[]> list in lists.OrderBy(pair => pair.Key))
        {
            IndexBinaryArrays.AppendVarint(bytes, Zigzag(list.Key));
            IndexBinaryArrays.AppendVarint(bytes, (uint)list.Value.Length);
            foreach (int counter in list.Value) IndexBinaryArrays.AppendVarint(bytes, Zigzag(counter));
        }
        IndexBinaryArrays.AppendVarint(bytes, (uint)overrides.Count);
        foreach ((int listId, int level) in overrides.OrderBy(pair => pair.listId).ThenBy(pair => pair.level))
        {
            IndexBinaryArrays.AppendVarint(bytes, Zigzag(listId));
            IndexBinaryArrays.AppendVarint(bytes, (uint)level);
        }
        return bytes.ToArray();
    }

    public static (Dictionary<int, int[]> lists, HashSet<(int listId, int level)> overrides) Parse(byte[] state)
    {
        var lists = new Dictionary<int, int[]>();
        var overrides = new HashSet<(int, int)>();
        if (state.Length == 0) return (lists, overrides);
        int position = 0;
        int listCount = (int)IndexBinaryArrays.ReadVarint(state, ref position);
        for (int i = 0; i < listCount; i++)
        {
            int listId = Unzigzag(IndexBinaryArrays.ReadVarint(state, ref position));
            int levels = (int)IndexBinaryArrays.ReadVarint(state, ref position);
            if (levels < 0 || levels > 64) throw new InvalidDataException("Index: list state damaged.");
            var counters = new int[levels];
            for (int level = 0; level < levels; level++) counters[level] = Unzigzag(IndexBinaryArrays.ReadVarint(state, ref position));
            lists[listId] = counters;
        }
        int overrideCount = (int)IndexBinaryArrays.ReadVarint(state, ref position);
        for (int i = 0; i < overrideCount; i++)
        {
            int listId = Unzigzag(IndexBinaryArrays.ReadVarint(state, ref position));
            int level = (int)IndexBinaryArrays.ReadVarint(state, ref position);
            overrides.Add((listId, level));
        }
        return (lists, overrides);
    }

    /// <summary><paramref name="current"/> as a delta from <paramref name="previous"/>: the lists that changed or appeared, the overrides newly applied.</summary>
    public static byte[] Delta(byte[] previous, byte[] current)
    {
        (Dictionary<int, int[]> oldLists, HashSet<(int, int)> oldOverrides) = Parse(previous);
        (Dictionary<int, int[]> newLists, HashSet<(int, int)> newOverrides) = Parse(current);
        var changed = new Dictionary<int, int[]>();
        foreach (KeyValuePair<int, int[]> list in newLists)
            if (!oldLists.TryGetValue(list.Key, out int[]? old) || !old.AsSpan().SequenceEqual(list.Value)) changed[list.Key] = list.Value;
        var added = new HashSet<(int, int)>(newOverrides);
        added.ExceptWith(oldOverrides);
        return Serialize(changed, added);
    }

    /// <summary>The full state <paramref name="delta"/> (from <see cref="Delta"/>) makes out of <paramref name="previous"/>.</summary>
    public static byte[] Apply(byte[] previous, byte[] delta)
    {
        (Dictionary<int, int[]> lists, HashSet<(int, int)> overrides) = Parse(previous);
        (Dictionary<int, int[]> changed, HashSet<(int, int)> added) = Parse(delta);
        foreach (KeyValuePair<int, int[]> list in changed) lists[list.Key] = list.Value;
        overrides.UnionWith(added);
        return Serialize(lists, overrides);
    }

    /// <summary>Turns the full states captured at every anchor into deltas from the previous anchor.</summary>
    public static byte[][] Deltas(IReadOnlyList<byte[]> fullStates)
    {
        var deltas = new byte[fullStates.Count][];
        byte[] previous = [];
        for (int i = 0; i < deltas.Length; i++) { deltas[i] = Delta(previous, fullStates[i]); previous = fullStates[i]; }
        return deltas;
    }

    /// <summary>
    /// Serves full states for anchors whose stored states are deltas: every <paramref name="keyEvery"/>-th full state is kept
    /// once computed, the rest apply their deltas from there.
    /// </summary>
    public sealed class DeltaChain
    {
        private readonly byte[][] _deltas;
        private readonly int _keyEvery;
        private readonly byte[]?[] _keyStates;

        public DeltaChain(byte[][] deltas, int keyEvery = 8)
        {
            _deltas = deltas;
            _keyEvery = keyEvery;
            _keyStates = new byte[]?[(deltas.Length + keyEvery - 1) / keyEvery];
        }

        public byte[] StateAt(int anchorIndex)
        {
            int key = anchorIndex / _keyEvery;
            if (_keyStates[key] is null)
            {
                byte[] state = key > 0 ? StateAt(key * _keyEvery - 1) : [];
                _keyStates[key] = Apply(state, _deltas[key * _keyEvery]);
            }
            byte[] current = _keyStates[key]!;
            for (int i = key * _keyEvery + 1; i <= anchorIndex; i++) current = Apply(current, _deltas[i]);
            return current;
        }
    }

    public static ulong Zigzag(int value) => (ulong)(uint)((value << 1) ^ (value >> 31));
    public static int Unzigzag(ulong value) => (int)((uint)value >> 1) ^ -(int)((uint)value & 1);
}
