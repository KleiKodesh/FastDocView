using System.Globalization;

namespace FastDocView.Fields;

/// <summary>A computed field value before its switches format it: a number, a date or text.</summary>
public readonly record struct WordFieldValue(string? Text = null, decimal? Number = null, DateTime? Date = null, int Decimals = 0, bool Grouped = false,
                                             string? DefaultDatePicture = null)
{
    public static WordFieldValue OfText(string text) => new(Text: text);
    public static WordFieldValue OfNumber(decimal number) => new(Number: number);
}

/// <summary>A table the formula field is in: cells by zero-based row and column, and the field's own cell.</summary>
public interface IWordFieldTable
{
    int RowCount { get; }
    int ColumnCount(int row);
    /// <summary>The cell's text (null when the cell does not exist).</summary>
    string? CellText(int row, int column);
    int CurrentRow { get; }
    int CurrentColumn { get; }
}

/// <summary>
/// What a field needs from its document to compute a result Word did not save. Every member is optional: a field whose
/// source is missing keeps its cached result (or shows nothing). Set by each format's windowed document.
/// </summary>
public sealed class WordFieldContext
{
    /// <summary>Day / month names and AM / PM for date pictures - Word takes them from the system (see <see cref="WordFieldFormat.FormatDate"/>).</summary>
    public CultureInfo Culture { get; init; } = CultureInfo.CurrentCulture;

    /// <summary>The time DATE / TIME show (null: now).</summary>
    public DateTime? Now { get; init; }

    /// <summary>
    /// Built-in document properties by Word's field name (TITLE, SUBJECT, AUTHOR, KEYWORDS, COMMENTS, LASTSAVEDBY, TEMPLATE,
    /// REVNUM, NUMPAGES, NUMWORDS, NUMCHARS, EDITTIME, COMPANY, MANAGER, CATEGORY ...) and custom ones by their own name
    /// (DOCPROPERTY); null when unknown.
    /// </summary>
    public Func<string, string?>? DocumentProperty { get; init; }

    /// <summary>CREATEDATE / SAVEDATE / PRINTDATE by that name; null when the document does not record it.</summary>
    public Func<string, DateTime?>? DocumentDate { get; init; }

    /// <summary>The file's name (FILENAME; with \p the full path).</summary>
    public string? FileName { get; init; }
    public string? FilePath { get; init; }
    public long? FileSize { get; init; }

    /// <summary>DOCVARIABLE values (.docx w:docVars, .doc SttbfAssoc / docvars); null when absent.</summary>
    public Func<string, string?>? DocumentVariable { get; init; }

    /// <summary>The text a bookmark marks (REF, and a bookmark name in a formula); null when the bookmark is not known here.</summary>
    public Func<string, string?>? BookmarkText { get; init; }

    /// <summary>The block the field being rendered is in (REF / NOTEREF \p, STYLEREF's search start); set by the reader per block.</summary>
    public int? CurrentBlock { get; set; }

    /// <summary>
    /// False while the bookmark list is still being built (an index scanning on): a bookmark not found yet is then unknown,
    /// not missing - the field shows nothing rather than an error, and nothing scans the rest of the document mid-render.
    /// </summary>
    public Func<bool>? IsBookmarkListComplete { get; init; }

    /// <summary>The block a bookmark starts in; null when it is not known (yet).</summary>
    public Func<string, int?>? BookmarkBlock { get; init; }

    /// <summary>A block's list label as shown ("2.a.", "ג."); null when the block is not a list item (REF / STYLEREF \n \r \w).</summary>
    public Func<int, string?>? ListLabelOfBlock { get; init; }

    /// <summary>The mark of the first note referenced in the bookmark's block (NOTEREF); null when none.</summary>
    public Func<string, string?>? NoteMarkOfBookmark { get; init; }

    /// <summary>The nearest paragraph in a style from a block, backward (forward = true: after it): its block and text (STYLEREF); null when none.</summary>
    public Func<string, int, bool, (int block, string text)?>? StyledParagraph { get; init; }

    /// <summary>The XE / TA field codes in the document with their blocks, as known now: an INDEX / TOA Word saved without a result is built from them.</summary>
    public Func<IReadOnlyList<(int block, string code)>>? IndexEntries { get; init; }

    /// <summary>The document's headings as known now: a TOC Word saved without a result is built from them.</summary>
    public Func<IReadOnlyList<WordDocumentHeading>>? Headings { get; init; }

    /// <summary>The table a formula field sits in (null outside a table).</summary>
    public IWordFieldTable? Table { get; set; }

    /// <summary>SEQ counters, by identifier, as they stand where the window starts (from the index); null: counting starts here.</summary>
    public WordFieldSequenceCounters? Sequences { get; set; }

    /// <summary>The current heading level-1..9 numbers for SEQ \s (the heading level that resets the sequence).</summary>
    public Func<int, int>? HeadingCount { get; set; }
}

/// <summary>
/// SEQ field counters (ECMA-376 Part 1 17.16.5.56): each identifier counts up at every SEQ that is not \c; \r n resets it to n;
/// \s n resets it at each heading of level n. Word 16 check (2026-11-04): SEQ Fig → 1, 2; \c → 2; \h → (hidden, counts) ;
/// next → 4; \r 10 → 10; \* roman → xi.
/// </summary>
public sealed class WordFieldSequenceCounters
{
    private readonly Dictionary<string, int> _values = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, int> _resetLevels = new(StringComparer.OrdinalIgnoreCase);

    public WordFieldSequenceCounters Clone()
    {
        var clone = new WordFieldSequenceCounters();
        foreach (var pair in _values) clone._values[pair.Key] = pair.Value;
        foreach (var pair in _resetLevels) clone._resetLevels[pair.Key] = pair.Value;
        return clone;
    }

    /// <summary>The next value of a SEQ field, the counter advanced.</summary>
    public int Advance(WordFieldInstruction instruction)
    {
        string identifier = instruction.Arguments.Count > 0 ? instruction.Arguments[0] : "";
        _values.TryGetValue(identifier, out int value);
        if (instruction.SwitchArgument("\\s") is { Length: > 0 } level && int.TryParse(level, out int resetLevel)) _resetLevels[identifier] = resetLevel;
        if (instruction.SwitchArgument("\\r") is { Length: > 0 } reset && int.TryParse(reset, NumberStyles.Integer, CultureInfo.InvariantCulture, out int resetTo)) value = resetTo;
        else if (!instruction.Has("\\c")) value++;
        _values[identifier] = value;
        return value;
    }

    private const string AutoNumberKey = "\u0002AUTONUM";
    private const string ListNumberKey = "\u0002LISTNUM";

    /// <summary>
    /// AUTONUM / AUTONUMLGL / AUTONUMOUT: one counter shared by the three (Word 16, 2026-11-04: AUTONUM, AUTONUM, AUTONUM \s :,
    /// AUTONUMLGL, AUTONUMOUT in body paragraphs → 1. 2. 3: 4. 5.).
    /// </summary>
    public int AdvanceAutoNumber()
    {
        _values.TryGetValue(AutoNumberKey, out int value);
        _values[AutoNumberKey] = ++value;
        return value;
    }

    /// <summary>
    /// LISTNUM: one set of level counters shared by every list name (Word 16: LISTNUM, LISTNUM, LISTNUM \l 2,
    /// LISTNUM NumberDefault \l 1 \s 5, LISTNUM LegalDefault, LISTNUM OutlineDefault → 1) 2) a) 5) 6. VII.); a level counts
    /// up, \s sets it, deeper levels start over.
    /// </summary>
    public int AdvanceListNumber(int level, int? start)
    {
        string key = ListNumberKey + level.ToString(CultureInfo.InvariantCulture);
        _values.TryGetValue(key, out int value);
        value = start ?? value + 1;
        _values[key] = value;
        for (int deeper = level + 1; deeper <= 9; deeper++) _values.Remove(ListNumberKey + deeper.ToString(CultureInfo.InvariantCulture));
        return value;
    }

    /// <summary>A heading of a level: sequences reset by that level (\s) or a higher one start over.</summary>
    public void Heading(int level)
    {
        foreach (var pair in _resetLevels.ToList())
            if (level <= pair.Value) _values[pair.Key] = 0;
    }

    /// <summary>The state as one string (index storage): identifier=value(:level) joined by U+0003.</summary>
    public string Serialize()
        => string.Join("\u0003", _values.Select(pair => pair.Key + "=" + pair.Value.ToString(CultureInfo.InvariantCulture)
                                                     + (_resetLevels.TryGetValue(pair.Key, out int level) ? ":" + level.ToString(CultureInfo.InvariantCulture) : "")));

    public static WordFieldSequenceCounters Deserialize(string? serialized)
    {
        var counters = new WordFieldSequenceCounters();
        if (string.IsNullOrEmpty(serialized)) return counters;
        foreach (string entry in serialized!.Split('\u0003'))
        {
            int equals = entry.LastIndexOf('=');
            if (equals <= 0) continue;
            string rest = entry.Substring(equals + 1);
            int colon = rest.IndexOf(':');
            if (int.TryParse(colon < 0 ? rest : rest.Substring(0, colon), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
                counters._values[entry.Substring(0, equals)] = value;
            if (colon >= 0 && int.TryParse(rest.Substring(colon + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out int level))
                counters._resetLevels[entry.Substring(0, equals)] = level;
        }
        return counters;
    }

    public bool IsEmpty => _values.Count == 0;
}
