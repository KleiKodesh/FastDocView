namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// Runs the list counters during the index pass: one counter vector per abstractNum (after w:numStyleLink), start
/// values and per-num start overrides honoured, deeper levels restarted by their w:lvlRestart rule (§17.9.10: absent =
/// restart whenever a higher level is used; 0 = never; N = when one of the levels 1..N is used — the same reading as
/// Open-Xml-PowerTools' PreviousParagraphsForLvlRestart). Returns the finished label ("3.", "ב.", "1.2") and its level
/// definition (suffix, label formatting) for each list paragraph in document order.
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
    public string? Next(int numId, int level) => NextWithDefinition(numId, level)?.label;

    public (string label, DocxNumberingDefinitions.LevelDefinition? definition)? NextWithDefinition(int numId, int level)
    {
        if (numId <= 0 || _definitions.AbstractIdOf(numId) is not { } abstractId) return null;
        level = Math.Max(0, Math.Min(MaxLevels - 1, level));
        if (!_countersByAbstractId.TryGetValue(abstractId, out int[]? counters))
        {
            counters = new int[MaxLevels];
            for (int i = 0; i < MaxLevels; i++) counters[i] = (_definitions.LevelOf(numId, abstractId, i)?.Start ?? 1) - 1;
            _countersByAbstractId[abstractId] = counters;
        }
        if (_definitions.StartOverrideOf(numId, level) is { } start && _appliedStartOverrides.Add((numId, level)))
            counters[level] = start - 1;

        counters[level]++;
        for (int deeper = level + 1; deeper < MaxLevels; deeper++)
        {
            DocxNumberingDefinitions.LevelDefinition? deeperDefinition = _definitions.LevelOf(numId, abstractId, deeper);
            int? restart = deeperDefinition?.Restart;
            bool restarts = restart is null || (restart > 0 && level <= restart.Value - 1);
            if (restarts) counters[deeper] = (deeperDefinition?.Start ?? 1) - 1;
        }
        return (_definitions.FormatLabel(numId, abstractId, level, counters), _definitions.LevelOf(numId, abstractId, level));
    }
}
