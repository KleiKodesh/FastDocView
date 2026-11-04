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
        if (style is not null) UsedListStyles[block.StyleName!] = style;
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

    /// <summary>The list styles a label was resolved from (ParagraphStarted): what the compact index keeps so a rescan can count again.</summary>
    public Dictionary<string, OdtListStyle> UsedListStyles { get; } = new(StringComparer.Ordinal);

    /// <summary>Keeps a list style a list inside a text box or shape names (those lists are numbered when they are drawn).</summary>
    public void NoteListStyle(string? name)
    {
        if (name is not null && !UsedListStyles.ContainsKey(name) && _styles.ListStyle(name) is { } style) UsedListStyles[name] = style;
    }

    // -- state for the compact index (per span anchor) --------------------------------

    /// <summary>
    /// The complete numbering state: every list processed (id, style, continuation), every master list's counters, the
    /// last processed list, the generated-id counter and the open list / item contexts (bottom first). A scanner that
    /// imports it and replays the body from that point produces the same labels as the one pass.
    /// </summary>
    public byte[] ExportState()
    {
        using var memory = new MemoryStream();
        using (var writer = new BinaryWriter(memory, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            WriteScalarsAndContexts(writer);
            WriteLists(writer, _processedLists);
            WriteCounters(writer, _countersByMasterId);
        }
        return memory.ToArray();
    }

    public void ImportState(byte[] state)
    {
        _contexts.Clear();
        _processedLists.Clear();
        _countersByMasterId.Clear();
        _lastProcessedListId = null;
        _listStyleOfLastProcessedList = null;
        _generatedIds = 0;
        if (state.Length == 0) return;
        using var reader = new BinaryReader(new MemoryStream(state), System.Text.Encoding.UTF8);
        List<object> contexts = ReadScalarsAndContexts(reader, this);
        foreach (KeyValuePair<string, (string? style, string? continueId)> list in ReadLists(reader)) _processedLists[list.Key] = list.Value;
        foreach (KeyValuePair<string, int[]> counters in ReadCounters(reader)) _countersByMasterId[counters.Key] = counters.Value;
        foreach (object context in contexts)
        {
            if (context is ListBlock block)
            {
                string masterId = block.ContinueId ?? block.Id;
                if (!_countersByMasterId.TryGetValue(masterId, out int[]? counters)) { counters = NewCounters(); _countersByMasterId[masterId] = counters; }
                block.Counters = counters;
            }
            _contexts.Push(context);
        }
    }

    /// <summary>
    /// <paramref name="current"/> as a delta from <paramref name="previous"/> (both <see cref="ExportState"/> forms): the scalars and
    /// contexts whole (small), only the lists newly processed and the master counters that changed or appeared: the
    /// processed-list table only grows along a document, so a span's delta is a few entries where the full state is every
    /// list ever seen.
    /// </summary>
    public static byte[] DeltaState(byte[] previous, byte[] current)
    {
        Parsed old = Parse(previous), now = Parse(current);
        var addedLists = new Dictionary<string, (string? style, string? continueId)>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, (string? style, string? continueId)> list in now.Lists) if (!old.Lists.ContainsKey(list.Key)) addedLists[list.Key] = list.Value;
        var changed = new Dictionary<string, int[]>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, int[]> counters in now.Counters)
            if (!old.Counters.TryGetValue(counters.Key, out int[]? before) || !before.AsSpan().SequenceEqual(counters.Value)) changed[counters.Key] = counters.Value;
        return Assemble(now.Head, addedLists, changed);
    }

    /// <summary>The full state <paramref name="delta"/> (from <see cref="DeltaState"/>) makes out of <paramref name="previous"/>.</summary>
    public static byte[] ApplyDelta(byte[] previous, byte[] delta)
    {
        if (delta.Length == 0) return [];
        Parsed old = Parse(previous), change = Parse(delta);
        foreach (KeyValuePair<string, (string? style, string? continueId)> list in change.Lists) old.Lists[list.Key] = list.Value;
        foreach (KeyValuePair<string, int[]> counters in change.Counters) old.Counters[counters.Key] = counters.Value;
        return Assemble(change.Head, old.Lists, old.Counters);
    }

    private static byte[] Assemble(byte[] head, Dictionary<string, (string? style, string? continueId)> lists, Dictionary<string, int[]> counters)
    {
        using var memory = new MemoryStream();
        using (var writer = new BinaryWriter(memory, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            writer.Write(head);
            WriteLists(writer, lists);
            WriteCounters(writer, counters);
        }
        return memory.ToArray();
    }

    private sealed class Parsed
    {
        public byte[] Head = [];                                   // scalars + contexts, opaque here
        public Dictionary<string, (string? style, string? continueId)> Lists = new(StringComparer.Ordinal);
        public Dictionary<string, int[]> Counters = new(StringComparer.Ordinal);
    }

    private static Parsed Parse(byte[] state)
    {
        var parsed = new Parsed();
        if (state.Length == 0) return parsed;
        using var reader = new BinaryReader(new MemoryStream(state), System.Text.Encoding.UTF8);
        ReadScalarsAndContexts(reader, null);
        int headLength = (int)reader.BaseStream.Position;
        parsed.Head = new byte[headLength];
        Buffer.BlockCopy(state, 0, parsed.Head, 0, headLength);
        foreach (KeyValuePair<string, (string? style, string? continueId)> list in ReadLists(reader)) parsed.Lists[list.Key] = list.Value;
        foreach (KeyValuePair<string, int[]> counters in ReadCounters(reader)) parsed.Counters[counters.Key] = counters.Value;
        return parsed;
    }

    private void WriteScalarsAndContexts(BinaryWriter writer)
    {
        writer.Write(_generatedIds);
        WriteOptional(writer, _lastProcessedListId);
        WriteOptional(writer, _listStyleOfLastProcessedList);
        object[] contexts = _contexts.ToArray();                 // top first
        writer.Write(contexts.Length);
        for (int i = contexts.Length - 1; i >= 0; i--)
        {
            if (contexts[i] is ListBlock block)
            {
                writer.Write((byte)0);
                writer.Write(block.Id);
                WriteOptional(writer, block.ContinueId);
                WriteOptional(writer, block.StyleName);
                writer.Write((byte)block.Level);
                writer.Write(block.IsRestartPending);
            }
            else
            {
                var item = (ListItem)contexts[i];
                writer.Write((byte)1);
                writer.Write(item.IsHeader);
                writer.Write(item.StartValue);
                writer.Write(item.IsNumberConsumed);
                writer.Write(item.SubListCount);
            }
        }
    }

    /// <summary>Reads what <see cref="WriteScalarsAndContexts"/> wrote (into <paramref name="into"/>'s scalars when given); returns the contexts bottom first.</summary>
    private static List<object> ReadScalarsAndContexts(BinaryReader reader, OdtListNumbering? into)
    {
        int generated = reader.ReadInt32();
        string? last = ReadOptional(reader);
        string? lastStyle = ReadOptional(reader);
        if (into is not null) { into._generatedIds = generated; into._lastProcessedListId = last; into._listStyleOfLastProcessedList = lastStyle; }
        int count = reader.ReadInt32();
        if (count < 0 || count > 10_000) throw new InvalidDataException("OpenDocument index: list state damaged.");
        var contexts = new List<object>(count);
        for (int i = 0; i < count; i++)
        {
            byte type = reader.ReadByte();
            if (type == 0)
                contexts.Add(new ListBlock { Id = reader.ReadString(), ContinueId = ReadOptional(reader), StyleName = ReadOptional(reader),
                                             Level = Math.Min(OdtListStyle.MaxLevels - 1, (int)reader.ReadByte()), IsRestartPending = reader.ReadBoolean() });
            else if (type == 1)
                contexts.Add(new ListItem { IsHeader = reader.ReadBoolean(), StartValue = reader.ReadInt32(), IsNumberConsumed = reader.ReadBoolean(), SubListCount = reader.ReadInt32() });
            else throw new InvalidDataException("OpenDocument index: list state damaged.");
        }
        return contexts;
    }

    private static void WriteLists(BinaryWriter writer, Dictionary<string, (string? style, string? continueId)> lists)
    {
        writer.Write(lists.Count);
        foreach (KeyValuePair<string, (string? style, string? continueId)> list in lists.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            writer.Write(list.Key);
            WriteOptional(writer, list.Value.style);
            WriteOptional(writer, list.Value.continueId);
        }
    }

    private static IEnumerable<KeyValuePair<string, (string? style, string? continueId)>> ReadLists(BinaryReader reader)
    {
        int count = reader.ReadInt32();
        if (count < 0) throw new InvalidDataException("OpenDocument index: list state damaged.");
        for (int i = 0; i < count; i++)
        {
            string id = reader.ReadString();
            yield return new KeyValuePair<string, (string?, string?)>(id, (ReadOptional(reader), ReadOptional(reader)));
        }
    }

    private static void WriteCounters(BinaryWriter writer, Dictionary<string, int[]> counters)
    {
        writer.Write(counters.Count);
        foreach (KeyValuePair<string, int[]> list in counters.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            writer.Write(list.Key);
            foreach (int counter in list.Value) writer.Write(counter);
        }
    }

    private static IEnumerable<KeyValuePair<string, int[]>> ReadCounters(BinaryReader reader)
    {
        int count = reader.ReadInt32();
        if (count < 0) throw new InvalidDataException("OpenDocument index: list state damaged.");
        for (int i = 0; i < count; i++)
        {
            string id = reader.ReadString();
            var counters = new int[OdtListStyle.MaxLevels];
            for (int level = 0; level < counters.Length; level++) counters[level] = reader.ReadInt32();
            yield return new KeyValuePair<string, int[]>(id, counters);
        }
    }

    private static void WriteOptional(BinaryWriter writer, string? value)
    {
        writer.Write(value is not null);
        if (value is not null) writer.Write(value);
    }

    private static string? ReadOptional(BinaryReader reader) => reader.ReadBoolean() ? reader.ReadString() : null;
}
