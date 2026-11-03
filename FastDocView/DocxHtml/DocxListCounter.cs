namespace FastDocView.DocxHtml;

/// <summary>
/// Runs the list counters during the index pass: one counter vector per abstractNum (after w:numStyleLink), start
/// values and per-num start overrides honoured, deeper levels restarted by their w:lvlRestart rule (§17.9.10: absent =
/// restart whenever a higher level is used; 0 = never; N = when one of the levels 1..N is used — the same reading as
/// Open-Xml-PowerTools' PreviousParagraphsForLvlRestart). Returns the finished label ("3.", "ב.", "1.2") and its level
/// definition (suffix, label formatting) for each list paragraph in document order.
/// </summary>
public sealed class DocxListCounter
{
    internal const int MaxLevels = 9;

    private readonly DocxNumberingDefinitions _definitions;
    private readonly Dictionary<int, int[]> _countersByAbstractId = new();
    private readonly HashSet<(int numId, int level)> _appliedStartOverrides = new();

    public DocxListCounter(DocxNumberingDefinitions definitions)
    {
        _definitions = definitions;
    }

    /// <summary>Advances the counter for this paragraph and returns its label, or null when numId is not a list.</summary>
    public string? Next(int numId, int level) => NextWithDefinition(numId, level)?.label;

    /// <summary>
    /// The complete numbering state (every list's counters, the start overrides already applied), as the compact index
    /// stores it per checkpoint: a renderer that restores it and replays the list paragraphs from there produces the same
    /// labels as the one pass that counted them all. Zigzag varints; a few bytes per list.
    /// </summary>
    public byte[] ExportState() => Serialize(_countersByAbstractId, _appliedStartOverrides);

    /// <summary>Replaces the numbering state by one <see cref="ExportState"/> wrote.</summary>
    public void ImportState(byte[] state)
    {
        _countersByAbstractId.Clear();
        _appliedStartOverrides.Clear();
        (Dictionary<int, int[]> lists, HashSet<(int, int)> overrides) = Parse(state);
        foreach (KeyValuePair<int, int[]> list in lists) _countersByAbstractId[list.Key] = list.Value;
        _appliedStartOverrides.UnionWith(overrides);
    }

    /// <summary>
    /// <paramref name="current"/> as a delta from <paramref name="previous"/> (both <see cref="ExportState"/> forms): only the lists
    /// whose counters changed or appeared, and the overrides newly applied — the state only grows along a document, so a
    /// span's delta is a few lists where the full state is every list ever seen.
    /// </summary>
    public static byte[] DeltaState(byte[] previous, byte[] current)
    {
        (Dictionary<int, int[]> oldLists, HashSet<(int, int)> oldOverrides) = Parse(previous);
        (Dictionary<int, int[]> newLists, HashSet<(int, int)> newOverrides) = Parse(current);
        var changedLists = new Dictionary<int, int[]>();
        foreach (KeyValuePair<int, int[]> list in newLists)
            if (!oldLists.TryGetValue(list.Key, out int[]? old) || !old.AsSpan().SequenceEqual(list.Value)) changedLists[list.Key] = list.Value;
        var addedOverrides = new HashSet<(int, int)>(newOverrides);
        addedOverrides.ExceptWith(oldOverrides);
        return Serialize(changedLists, addedOverrides);
    }

    /// <summary>The full state <paramref name="delta"/> (from <see cref="DeltaState"/>) makes out of <paramref name="previous"/>.</summary>
    public static byte[] ApplyDelta(byte[] previous, byte[] delta)
    {
        (Dictionary<int, int[]> lists, HashSet<(int, int)> overrides) = Parse(previous);
        (Dictionary<int, int[]> changed, HashSet<(int, int)> added) = Parse(delta);
        foreach (KeyValuePair<int, int[]> list in changed) lists[list.Key] = list.Value;
        overrides.UnionWith(added);
        return Serialize(lists, overrides);
    }

    private static (Dictionary<int, int[]> lists, HashSet<(int numId, int level)> overrides) Parse(byte[] state)
    {
        var lists = new Dictionary<int, int[]>();
        var overrides = new HashSet<(int, int)>();
        if (state.Length == 0) return (lists, overrides);
        int position = 0;
        int listCount = (int)IndexBinaryArrays.ReadVarint(state, ref position);
        for (int i = 0; i < listCount; i++)
        {
            int abstractId = Unzigzag(IndexBinaryArrays.ReadVarint(state, ref position));
            var counters = new int[MaxLevels];
            for (int level = 0; level < MaxLevels; level++) counters[level] = Unzigzag(IndexBinaryArrays.ReadVarint(state, ref position));
            lists[abstractId] = counters;
        }
        int overrideCount = (int)IndexBinaryArrays.ReadVarint(state, ref position);
        for (int i = 0; i < overrideCount; i++)
        {
            int numId = Unzigzag(IndexBinaryArrays.ReadVarint(state, ref position));
            int level = (int)IndexBinaryArrays.ReadVarint(state, ref position);
            overrides.Add((numId, level));
        }
        return (lists, overrides);
    }

    private static byte[] Serialize(Dictionary<int, int[]> lists, HashSet<(int numId, int level)> overrides)
    {
        if (lists.Count == 0 && overrides.Count == 0) return [];
        var bytes = new List<byte>();
        IndexBinaryArrays.AppendVarint(bytes, (uint)lists.Count);
        foreach (KeyValuePair<int, int[]> list in lists.OrderBy(pair => pair.Key))
        {
            IndexBinaryArrays.AppendVarint(bytes, Zigzag(list.Key));
            foreach (int counter in list.Value) IndexBinaryArrays.AppendVarint(bytes, Zigzag(counter));
        }
        IndexBinaryArrays.AppendVarint(bytes, (uint)overrides.Count);
        foreach ((int numId, int level) in overrides.OrderBy(pair => pair.numId).ThenBy(pair => pair.level))
        {
            IndexBinaryArrays.AppendVarint(bytes, Zigzag(numId));
            IndexBinaryArrays.AppendVarint(bytes, (uint)level);
        }
        return bytes.ToArray();
    }

    private static ulong Zigzag(int value) => (ulong)(uint)((value << 1) ^ (value >> 31));
    private static int Unzigzag(ulong value) => (int)((uint)value >> 1) ^ -(int)((uint)value & 1);

    public (string label, DocxNumberingDefinitions.LevelDefinition? definition)? NextWithDefinition(int numId, int level)
    {
        if (numId <= 0 || _definitions.AbstractIdOf(numId) is not { } abstractId) return null;
        level = Math.Max(0, Math.Min(MaxLevels - 1, level));
        DocxNumberingDefinitions.LevelDefinition?[] levels = _definitions.LevelsOf(numId, abstractId);
        if (!_countersByAbstractId.TryGetValue(abstractId, out int[]? counters))
        {
            counters = new int[MaxLevels];
            for (int i = 0; i < MaxLevels; i++) counters[i] = (levels[i]?.Start ?? 1) - 1;
            _countersByAbstractId[abstractId] = counters;
        }
        if (_definitions.StartOverrideOf(numId, level) is { } start && _appliedStartOverrides.Add((numId, level)))
            counters[level] = start - 1;

        counters[level]++;
        for (int deeper = level + 1; deeper < MaxLevels; deeper++)
        {
            DocxNumberingDefinitions.LevelDefinition? deeperDefinition = levels[deeper];
            int? restart = deeperDefinition?.Restart;
            bool restarts = restart is null || (restart > 0 && level <= restart.Value - 1);
            if (restarts) counters[deeper] = (deeperDefinition?.Start ?? 1) - 1;
        }
        return (_definitions.FormatLabel(numId, abstractId, level, counters), levels[level]);
    }
}
