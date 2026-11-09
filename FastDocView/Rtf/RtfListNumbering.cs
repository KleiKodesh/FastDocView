using System.Text;
using FastDocView.DocxHtml;

namespace FastDocView.Rtf;

/// <summary>
/// List labels counted during the index pass from \listtable / \listoverridetable (RTF 1.5 spec "List table",
/// "List override table"): a paragraph's \lsN names an override (by its \ls number, as LibreOffice maps it), the
/// override names a list by \listid, the list's \listlevel N gives \levelnfc, \levelstartat, \leveltext (its first
/// character is the length; level placeholders sit at the \levelnumbers positions) and \levelnorestart. The first
/// paragraph of an override with an \lfolevel start-at takes that value (spec: "The first element in the document
/// with each list override index takes the start-at value"). Mirrors the .doc counter (DocBinaryListCounter).
/// </summary>
internal sealed class RtfListNumbering
{
    private readonly RtfDocumentTables _tables;
    private readonly Dictionary<int, int[]> _countersByListId = new();
    private readonly HashSet<(int overrideIndex, int level)> _appliedStartOverrides = new();

    public RtfListNumbering(RtfDocumentTables tables) => _tables = tables;

    /// <summary>The complete numbering state (every list's counters, the start overrides applied), as the compact index stores it per span anchor.</summary>
    public byte[] ExportState() => WordListCounterState.Serialize(_countersByListId, _appliedStartOverrides);

    /// <summary>Replaces the numbering state by one <see cref="ExportState"/> wrote.</summary>
    public void ImportState(byte[] state)
    {
        _countersByListId.Clear();
        _appliedStartOverrides.Clear();
        (Dictionary<int, int[]> lists, HashSet<(int, int)> overrides) = WordListCounterState.Parse(state);
        foreach (KeyValuePair<int, int[]> list in lists)
        {
            if (list.Value.Length != 9) throw new InvalidDataException("RTF index: list state has the wrong level count.");
            _countersByListId[list.Key] = list.Value;
        }
        _appliedStartOverrides.UnionWith(overrides);
    }

    /// <summary>\levelnfc / \levelnfcn values are MSONFC numbers (LibreOffice rtfdispatchvalue.cxx getNumberFormat): the .doc table.</summary>
    public static string FormatName(int numberFormat) => numberFormat switch
    {
        255 or 8 or 9 => "none",                                       // rtfdispatchvalue.cxx:39-50 getNumberFormat: 8, 9 undefined in RTF 1.8
        >= 0 and < 255 => DocBinary.DocBinaryListDefinitions.NumberFormatName((byte)numberFormat),
        _ => "decimal",
    };

    /// <summary>The level definition a paragraph's \ls / \ilvl name, or null.</summary>
    public static RtfDocumentTables.ListLevelDefinition? LevelOf(RtfDocumentTables tables, int overrideIndex, int level)
    {
        if (overrideIndex <= 0 || !tables.ListOverrides.TryGetValue(overrideIndex, out RtfDocumentTables.ListOverride? listOverride)) return null;
        if (!tables.Lists.TryGetValue(listOverride.ListId, out RtfDocumentTables.ListDefinition? list) || list.Levels.Length == 0) return null;
        return list.Levels[Math.Max(0, Math.Min(level, list.Levels.Length - 1))];
    }

    /// <summary>Advances the counters for a list paragraph and returns its label; null when the list is unknown.</summary>
    public string? Next(int overrideIndex, int level)
    {
        if (overrideIndex <= 0 || !_tables.ListOverrides.TryGetValue(overrideIndex, out RtfDocumentTables.ListOverride? listOverride)) return null;
        if (!_tables.Lists.TryGetValue(listOverride.ListId, out RtfDocumentTables.ListDefinition? list) || list.Levels.Length == 0) return null;
        level = Math.Min(level, list.Levels.Length - 1);
        if (!_countersByListId.TryGetValue(list.ListId, out int[]? counters))
        {
            counters = new int[9];
            for (int i = 0; i < counters.Length; i++) counters[i] = (i < list.Levels.Length ? list.Levels[i].StartAt : 1) - 1;
            _countersByListId[list.ListId] = counters;
        }
        for (int deeper = level + 1; deeper < list.Levels.Length && deeper < counters.Length; deeper++)
            if (!list.Levels[deeper].NoRestart) counters[deeper] = list.Levels[deeper].StartAt - 1;

        RtfDocumentTables.ListLevelDefinition definition = list.Levels[level];
        if (FormatName(definition.NumberFormat) is not ("bullet" or "none"))
        {
            if (listOverride.StartAtByLevel.TryGetValue(level, out int startAt) && _appliedStartOverrides.Add((overrideIndex, level)))
                counters[level] = startAt - 1;
            counters[level]++;
        }
        return FormatLabel(list, definition, counters, _tables);
    }

    private static string FormatLabel(RtfDocumentTables.ListDefinition list, RtfDocumentTables.ListLevelDefinition definition, int[] counters, RtfDocumentTables tables)
    {
        string text = definition.Text;
        if (FormatName(definition.NumberFormat) == "bullet") return BulletOf(text, tables.FontName(definition.Font));
        var label = new StringBuilder(text.Length + 4);
        for (int i = 0; i < text.Length; i++)
        {
            int position = i + 1;                                   // \levelnumbers count the length character as position 0
            if (Array.IndexOf(definition.NumberPositions, (byte)position) >= 0 && text[i] < 9)
            {
                int placeholderLevel = text[i];
                string format = placeholderLevel < list.Levels.Length ? FormatName(list.Levels[placeholderLevel].NumberFormat) : "decimal";
                if (definition.Legal && format != "decimalZero") format = "decimal";
                label.Append(DocxNumberingDefinitions.FormatNumber(counters[placeholderLevel], format));
            }
            else label.Append(text[i]);
        }
        return label.ToString();
    }

    /// <summary>A symbol-font bullet (U+F0xx) recoded against the level's font (<see cref="Html.WordSymbolFontBullet"/>).</summary>
    public static string BulletOf(string text, string? fontName) => Html.WordSymbolFontBullet.Resolve(text, fontName).text;
}
