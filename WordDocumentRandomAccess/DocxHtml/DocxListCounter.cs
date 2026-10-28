namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// Runs the list counters during the index pass: one counter vector per abstractNum, levels deeper
/// than the current one reset, start values and per-num start overrides honoured. Returns the
/// finished label ("3.", "ב.", "1.2") for each list paragraph in document order.
/// </summary>
public sealed class DocxListCounter
{
    private const int MaxLevels = 9;

    private readonly DocxNumberingDefinitions _definitions;
    private readonly Dictionary<int, int[]> _countersByAbstractId = new();
    private readonly HashSet<(int numId, int level)> _appliedStartOverrides = new();

    public DocxListCounter(DocxNumberingDefinitions definitions)
    {
        _definitions = definitions;
    }

    /// <summary>Advances the counter for this paragraph and returns its label, or null when numId is not a list.</summary>
    public string? Next(int numId, int level)
    {
        if (numId <= 0 || _definitions.AbstractIdOf(numId) is not { } abstractId) return null;
        level = Math.Max(0, Math.Min(MaxLevels - 1, level));
        if (!_countersByAbstractId.TryGetValue(abstractId, out int[]? counters))
        {
            counters = new int[MaxLevels];
            for (int i = 0; i < MaxLevels; i++) counters[i] = (_definitions.LevelOf(abstractId, i)?.Start ?? 1) - 1;
            _countersByAbstractId[abstractId] = counters;
        }
        if (_definitions.StartOverrideOf(numId, level) is { } start && _appliedStartOverrides.Add((numId, level)))
            counters[level] = start - 1;

        counters[level]++;
        for (int deeper = level + 1; deeper < MaxLevels; deeper++) counters[deeper] = (_definitions.LevelOf(abstractId, deeper)?.Start ?? 1) - 1;
        return _definitions.FormatLabel(abstractId, level, counters);
    }
}
