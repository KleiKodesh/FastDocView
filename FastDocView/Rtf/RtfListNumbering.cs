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

    /// <summary>The level definition a paragraph's \ls / \ilvl name (an override's own level first), or null.</summary>
    public static RtfDocumentTables.ListLevelDefinition? LevelOf(RtfDocumentTables tables, int overrideIndex, int level)
    {
        if (overrideIndex <= 0 || !tables.ListOverrides.TryGetValue(overrideIndex, out RtfDocumentTables.ListOverride? listOverride)) return null;
        if (!tables.Lists.TryGetValue(listOverride.ListId, out RtfDocumentTables.ListDefinition? list)) return null;
        RtfDocumentTables.ListLevelDefinition[] levels = LevelsOf(tables, list);
        if (levels.Length == 0) return null;
        return LevelAt(listOverride, levels, Math.Max(0, Math.Min(level, levels.Length - 1)));
    }

    /// <summary>
    /// A list's levels: its own, or - when it has none - those of the list style its \liststyleid names (the list whose \listid
    /// that is; Word 16 draws an emptied linked list from the style's levels and still counts it as its own list, checked by printing).
    /// </summary>
    private static RtfDocumentTables.ListLevelDefinition[] LevelsOf(RtfDocumentTables tables, RtfDocumentTables.ListDefinition list)
        => list.Levels.Length == 0 && list.StyleListId != 0 && list.StyleListId != list.ListId
           && tables.Lists.TryGetValue(list.StyleListId, out RtfDocumentTables.ListDefinition? style) ? style.Levels : list.Levels;

    /// <summary>The override's own level for <paramref name="level"/> (an \lfolevel holding a \listlevel), else the list's.</summary>
    private static RtfDocumentTables.ListLevelDefinition LevelAt(RtfDocumentTables.ListOverride listOverride, RtfDocumentTables.ListLevelDefinition[] levels, int level)
        => listOverride.LevelsByLevel.TryGetValue(level, out RtfDocumentTables.ListLevelDefinition? replaced) ? replaced : levels[level];

    /// <summary>Advances the counters for a list paragraph and returns its label; null when the list is unknown.</summary>
    public string? Next(int overrideIndex, int level)
    {
        if (overrideIndex <= 0 || !_tables.ListOverrides.TryGetValue(overrideIndex, out RtfDocumentTables.ListOverride? listOverride)) return null;
        if (!_tables.Lists.TryGetValue(listOverride.ListId, out RtfDocumentTables.ListDefinition? list)) return null;
        RtfDocumentTables.ListLevelDefinition[] levels = LevelsOf(_tables, list);
        if (levels.Length == 0) return null;
        level = Math.Min(level, levels.Length - 1);
        if (!_countersByListId.TryGetValue(list.ListId, out int[]? counters))
        {
            counters = new int[9];
            for (int i = 0; i < counters.Length; i++) counters[i] = (i < levels.Length ? levels[i].StartAt : 1) - 1;   // the list's own starts: counters are shared by its overrides
            _countersByListId[list.ListId] = counters;
        }
        // \levelnorestart: the level never restarts. RTF has no restart-limit level (w:lvlRestart N / ilvlRestartLim): Word 16 saves a
        // limit as \levelnorestart alone and, printing the .rtf, never restarts that level (list-overrides-v1: 2.1.4, where the .docx prints 2.1.1)
        for (int deeper = level + 1; deeper < levels.Length && deeper < counters.Length; deeper++)
        {
            RtfDocumentTables.ListLevelDefinition deeperLevel = LevelAt(listOverride, levels, deeper);
            if (!deeperLevel.NoRestart) counters[deeper] = deeperLevel.StartAt - 1;
        }

        RtfDocumentTables.ListLevelDefinition definition = LevelAt(listOverride, levels, level);
        if (FormatName(definition.NumberFormat) is not ("bullet" or "none"))
        {
            if (listOverride.StartAtByLevel.TryGetValue(level, out int startAt) && _appliedStartOverrides.Add((overrideIndex, level)))
                counters[level] = startAt - 1;
            counters[level]++;
        }
        return FormatLabel(listOverride, levels, definition, counters, _tables);
    }

    private static string FormatLabel(RtfDocumentTables.ListOverride listOverride, RtfDocumentTables.ListLevelDefinition[] levels,
                                      RtfDocumentTables.ListLevelDefinition definition, int[] counters, RtfDocumentTables tables)
    {
        string text = definition.Text;
        // a picture bullet (\levelpicture) a browser can show is drawn as that picture ("lp:" id); else the level's text below
        if (definition.Picture >= 0 && definition.Picture < tables.ListPictures.Count && tables.ListPictures[definition.Picture] is var picture
            && Html.WordPictureHtml.ContentTypeOf(picture.Extension) is not null)
            return Html.WordListLabelHtml.PictureLabel(ListPicturePrefix + definition.Picture.ToString(System.Globalization.CultureInfo.InvariantCulture),
                                                       picture.WidthTwips > 0 ? picture.WidthTwips : 1, picture.HeightTwips > 0 ? picture.HeightTwips : 1);
        if (FormatName(definition.NumberFormat) == "bullet") return BulletOf(text, tables.FontName(definition.Font));
        var label = new StringBuilder(text.Length + 4);
        for (int i = 0; i < text.Length; i++)
        {
            int position = i + 1;                                   // \levelnumbers count the length character as position 0
            if (Array.IndexOf(definition.NumberPositions, (byte)position) >= 0 && text[i] < 9)
            {
                int placeholderLevel = text[i];
                string format = placeholderLevel < levels.Length ? FormatName(LevelAt(listOverride, levels, placeholderLevel).NumberFormat) : "decimal";
                if (definition.Legal && format != "decimalZero") format = "decimal";
                label.Append(DocxNumberingDefinitions.FormatNumber(counters[placeholderLevel], format));
            }
            else label.Append(text[i]);
        }
        return label.ToString();
    }

    /// <summary>The picture id of a list picture bullet: "lp:" and its index in <see cref="RtfDocumentTables.ListPictures"/>.</summary>
    public const string ListPicturePrefix = "lp:";

    /// <summary>A symbol-font bullet (U+F0xx) recoded against the level's font (<see cref="Html.WordSymbolFontBullet"/>).</summary>
    public static string BulletOf(string text, string? fontName) => Html.WordSymbolFontBullet.Resolve(text, fontName).text;
}
