namespace FastDocView.DocBinary;

/// <summary>
/// Runs the .doc list counters during the index pass ([MS-DOC] §2.4.6.4): one counter vector per
/// list (lsid), a level's sequence restarts when a more significant level appears (fNoRestart /
/// ilvlRestartLim honoured), LFOLVL start-at overrides applied the first time their LFO is met.
/// Returns the finished label per list paragraph in document order; the renderer never sees the
/// earlier items. Mirrors <see cref="DocxHtml.DocxListCounter"/>.
/// </summary>
public sealed class DocBinaryListCounter
{
    private const int SkippedLevel = 0x0C;               // sprmPIlvl: the list skips this paragraph
    private const int NotInListMarker = 0xF801;          // sprmPIlfo: explicitly not in a list

    private readonly DocBinaryListDefinitions _definitions;
    private readonly Dictionary<int, int[]> _countersByListId = new();
    private readonly HashSet<(int overrideIndex, int level)> _appliedStartOverrides = new();

    public DocBinaryListCounter(DocBinaryListDefinitions definitions)
    {
        _definitions = definitions;
    }

    /// <summary>Advances the counters for a paragraph and returns its label; null when it is not a list item.</summary>
    /// <param name="rawListFormatOverride">sprmPIlfo as stored (unsigned 16-bit): 1..0x07FE an index, 0xF801 none, 0xF802.. a negated index.</param>
    public string? Next(int rawListFormatOverride, int level)
    {
        if (_definitions.IsEmpty) return null;
        int overrideIndex = rawListFormatOverride;
        if (overrideIndex <= 0 || overrideIndex == NotInListMarker) return null;
        if (overrideIndex >= 0xF802) overrideIndex = 0x10000 - overrideIndex;
        if (level == SkippedLevel) return null;
        if (level < 0 || level >= DocBinaryListDefinitions.MaxLevels) return null;

        DocBinaryListDefinitions.ListOverride? listOverride = _definitions.OverrideAt(overrideIndex);
        if (listOverride is null || _definitions.ListOf(listOverride.ListId) is not { } list) return null;
        DocBinaryListDefinitions.LevelDefinition? definition = _definitions.LevelFor(listOverride, list, level);
        if (definition is null) return null;

        if (!_countersByListId.TryGetValue(list.ListId, out int[]? counters))
        {
            counters = new int[DocBinaryListDefinitions.MaxLevels];
            for (int i = 0; i < counters.Length; i++) counters[i] = (_definitions.LevelFor(listOverride, list, i)?.StartAt ?? 1) - 1;
            _countersByListId[list.ListId] = counters;
        }

        // A deeper level restarts when this level is more significant than its restart limit (its own level unless fNoRestart).
        for (int deeper = level + 1; deeper < counters.Length; deeper++)
        {
            DocBinaryListDefinitions.LevelDefinition? deeperDefinition = _definitions.LevelFor(listOverride, list, deeper);
            int restartLimit = deeperDefinition is { NoRestart: true } ? deeperDefinition.RestartLimitLevel : deeper;
            if (restartLimit > deeper) restartLimit = -1;          // out of range (0xFF seen in files): never restarts
            if (level < restartLimit) counters[deeper] = (deeperDefinition?.StartAt ?? 1) - 1;
        }

        if (!definition.HasNumberSequence) return _definitions.FormatLabel(listOverride, list, level, definition, counters);

        if (DocBinaryListDefinitions.StartOverrideFor(listOverride, level) is { } startOverride && _appliedStartOverrides.Add((overrideIndex, level)))
            counters[level] = startOverride.StartAt - 1;
        counters[level]++;
        return _definitions.FormatLabel(listOverride, list, level, definition, counters);
    }
}
