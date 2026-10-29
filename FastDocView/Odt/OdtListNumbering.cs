namespace FastDocView.Odt;

/// <summary>
/// Index-time list numbering in one forward pass, following LibreOffice's import
/// (xmloff/source/text/XMLTextListBlockContext.cxx, XMLTextListItemContext.cxx, txtlists.cxx and
/// XMLTextImportHelper::SetStyleAndAttrs in txtimp.cxx):
///   • a root text:list's xml:id is its list id; a missing id makes a new list;
///   • text:continue-list (root only) continues that list, following the continue chain to its master;
///   • text:continue-numbering="true" without continue-list continues the last processed list when it
///     has the same list style; "false" (or any other value) restarts;
///   • nested lists inherit the style, id and restart flag; the 2nd+ sub-list of one item restarts;
///   • only the first paragraph of a text:list-item is numbered — text:list-header paragraphs and
///     paragraphs after a nested list in the same item are not;
///   • text:start-value on the item sets that item's number.
/// Counters live per master list id; numbering a level clears the deeper levels (not from source:
/// Writer's usual outline behaviour, see NOTES.md).
/// </summary>
internal sealed class OdtListNumbering
{
    public const int Unset = int.MinValue;

    private sealed class ListBlock
    {
        public string Id = "";
        public string? ContinueId;
        public string? StyleName;
        public int Level;
        public bool IsRestartPending;
        public int[] Counters = null!;
    }

    private sealed class ListItem
    {
        public bool IsHeader;
        public int StartValue = -1;
        public bool IsNumberConsumed;
        public int SubListCount;
    }

    private readonly OdtStyleCascade _styles;
    private readonly Stack<object> _contexts = new();            // ListBlock / ListItem, innermost on top
    private readonly Dictionary<string, (string? style, string? continueId)> _processedLists = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int[]> _countersByMasterId = new(StringComparer.Ordinal);
    private string? _lastProcessedListId;
    private string? _listStyleOfLastProcessedList;
    private int _generatedIds;

    public OdtListNumbering(OdtStyleCascade styles) => _styles = styles;

    public bool IsInsideList => _contexts.Count > 0;

    /// <summary>The list depth (0-based) of a paragraph opened now.</summary>
    public int CurrentLevel => _contexts.OfType<ListBlock>().FirstOrDefault()?.Level ?? 0;

    public void OpenList(string? xmlId, string? styleName, string? continueNumbering, string? continueList)
    {
        ListBlock? parent = _contexts.Count > 0 && _contexts.Peek() is ListItem ? _contexts.OfType<ListBlock>().FirstOrDefault() : null;
        ListItem? parentItem = parent is null ? null : (ListItem)_contexts.Peek();
        var block = new ListBlock();
        if (parent is not null)
        {
            parentItem!.SubListCount++;
            block.StyleName = parent.StyleName;
            block.Level = Math.Min(OdtListStyle.MaxLevels - 1, parent.Level + 1);
            block.IsRestartPending = parent.IsRestartPending || parentItem.SubListCount > 1;
            block.Id = parent.Id;
            block.ContinueId = parent.ContinueId;
            block.Counters = parent.Counters;
        }
        bool isContinueNumberingPresent = continueNumbering is not null;
        if (isContinueNumberingPresent) block.IsRestartPending = continueNumbering != "true";
        if (styleName is not null) block.StyleName = styleName;

        if (parent is null)
        {
            block.Id = string.IsNullOrEmpty(xmlId) ? "\u0001generated" + (++_generatedIds) : xmlId!;
            block.ContinueId = string.IsNullOrEmpty(continueList) ? null : continueList;
            if (isContinueNumberingPresent && !block.IsRestartPending && block.ContinueId is null
                && _listStyleOfLastProcessedList == block.StyleName && _lastProcessedListId != block.Id)
                block.ContinueId = _lastProcessedListId;
            if (block.ContinueId is not null)
            {
                if (!_processedLists.ContainsKey(block.ContinueId)) block.ContinueId = null;
                else
                    for (string? next = _processedLists[block.ContinueId].continueId; next is not null; next = _processedLists.TryGetValue(next, out var entry) ? entry.continueId : null)
                        block.ContinueId = next;
            }
            if (!_processedLists.ContainsKey(block.Id))
            {
                _processedLists[block.Id] = (block.StyleName, block.ContinueId);
                _lastProcessedListId = block.Id;
                _listStyleOfLastProcessedList = block.StyleName;
            }
            string masterId = block.ContinueId ?? block.Id;
            if (!_countersByMasterId.TryGetValue(masterId, out int[]? counters))
            {
                counters = NewCounters();
                _countersByMasterId[masterId] = counters;
            }
            block.Counters = counters;
        }
        _contexts.Push(block);
    }

    public void CloseList()
    {
        if (_contexts.Count == 0 || _contexts.Peek() is not ListBlock block) return;
        _contexts.Pop();
        if (_contexts.Count > 0 && _contexts.Peek() is ListItem item)
        {
            item.IsNumberConsumed = true;                  // a paragraph after the nested list is not numbered
            if (_contexts.OfType<ListBlock>().FirstOrDefault() is { } parent) parent.IsRestartPending = block.IsRestartPending;
        }
    }

    public void OpenItem(bool isHeader, int startValue)
    {
        _contexts.Push(new ListItem { IsHeader = isHeader, StartValue = isHeader ? -1 : startValue });
    }

    public void CloseItem()
    {
        if (_contexts.Count > 0 && _contexts.Peek() is ListItem) _contexts.Pop();
    }

    /// <summary>
    /// A text:p / text:h opened directly inside the current list item or header: its label when
    /// numbered (bullet or number), null when it is an unnumbered list paragraph.
    /// </summary>
    public string? ParagraphStarted()
    {
        if (_contexts.Count == 0 || _contexts.Peek() is not ListItem item) return null;
        if (item.IsHeader || item.IsNumberConsumed) return null;
        item.IsNumberConsumed = true;
        ListBlock block = _contexts.OfType<ListBlock>().First();
        OdtListStyle? style = _styles.ListStyle(block.StyleName);
        int level = block.Level;
        int levelStart = style?.LevelAt(level).StartValue ?? 1;
        int[] counters = block.Counters;
        if (item.StartValue >= 0) counters[level] = item.StartValue;
        else if (block.IsRestartPending || counters[level] == Unset) counters[level] = levelStart;
        else counters[level]++;
        block.IsRestartPending = false;
        for (int deeper = level + 1; deeper < counters.Length; deeper++) counters[deeper] = Unset;
        LastLayout = style?.LevelAt(level).Layout ?? "";
        return style?.FormatLabel(level, counters) ?? "•";
    }

    /// <summary>The label-alignment geometry of the level <see cref="ParagraphStarted"/> last numbered (<see cref="OdtListStyle.Level.Layout"/>).</summary>
    public string LastLayout { get; private set; } = "";

    private static int[] NewCounters()
    {
        var counters = new int[OdtListStyle.MaxLevels];
        for (int i = 0; i < counters.Length; i++) counters[i] = Unset;
        return counters;
    }
}
