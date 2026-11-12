using System.Globalization;

namespace FastDocView;

/// <summary>When a note kind's numbering starts again (.docx w:numRestart, .doc sprmSRncFtn / sprmSRncEdn, RTF \sftnrst*).</summary>
public enum WordDocumentNoteRestart : byte { Continuous = 0, EachSection = 1, EachPage = 2 }

/// <summary>
/// One section's numbering of one note kind: number format (a .docx w:numFmt value), start value and restart rule.
/// Word numbers by the section's own rule only (checked in Word 16, test-data/word-notes/note-numbering-v1-word.pdf):
/// <list type="bullet">
/// <item>continuous: start + the notes of that kind before it in the whole document (a later section's start is an offset,
///   not a reset: start 10 after four footnotes prints 14);</item>
/// <item>each section: start + the notes of that kind before it in the same section;</item>
/// <item>each page: page-bound - a flow view has no pages, so it counts like each section (a section starts a page). LibreOffice
///   numbers per page in its layout (sw docftn.cxx, FTNNUM_PAGE); we cannot, and say so in docs/UNSUPPORTED-FEATURES.md.</item>
/// </list>
/// A section without note properties uses the defaults (decimal / lower roman, start 1, continuous): Word ignores the
/// .docx settings.xml footnotePr / endnotePr for numbering (same fixture: settings say upper roman from 3, Word prints 1, 2).
/// </summary>
public readonly record struct WordDocumentNoteNumberRule(string Format, int Start, WordDocumentNoteRestart Restart)
{
    public static readonly WordDocumentNoteNumberRule FootnoteDefault = new("decimal", 1, WordDocumentNoteRestart.Continuous);
    public static readonly WordDocumentNoteNumberRule EndnoteDefault = new("lowerRoman", 1, WordDocumentNoteRestart.Continuous);

    public static WordDocumentNoteNumberRule Default(WordDocumentNoteKind kind)
        => kind == WordDocumentNoteKind.Footnote ? FootnoteDefault : EndnoteDefault;

    /// <summary>The number printed for the note that is <paramref name="ordinal"/>-th (1-based, auto-numbered notes only) of its
    /// kind in the document, <paramref name="ordinalsBeforeSection"/> of them before its section.</summary>
    public int Number(int ordinal, int ordinalsBeforeSection)
        => Start + ordinal - 1 - (Restart == WordDocumentNoteRestart.Continuous ? 0 : ordinalsBeforeSection);

    /// <summary>Index form: "format|start|restart".</summary>
    public string Serialize() => Format + "|" + Start.ToString(CultureInfo.InvariantCulture) + "|" + ((int)Restart).ToString(CultureInfo.InvariantCulture);

    public static WordDocumentNoteNumberRule Parse(string text)
    {
        string[] parts = text.Split('|');
        if (parts.Length != 3 || !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int start)
            || !int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int restart) || restart is < 0 or > 2)
            throw new InvalidDataException("Index: a note number rule is malformed.");
        return new(parts[0], start, (WordDocumentNoteRestart)restart);
    }
}

/// <summary>
/// The note numbering of a document's sections, in block order (section i starts at <c>FirstBlocks[i]</c>): turns a note's
/// document ordinal into the label Word prints. Index-time data, no file access.
/// </summary>
public sealed class WordDocumentNoteNumbering
{
    private readonly IReadOnlyList<int> _firstBlocks;
    private readonly IReadOnlyList<WordDocumentNoteNumberRule> _footnotes, _endnotes;

    /// <param name="firstBlocks">Each section's first block, ascending; empty = one section with the defaults.</param>
    public WordDocumentNoteNumbering(IReadOnlyList<int> firstBlocks, IReadOnlyList<WordDocumentNoteNumberRule> footnotes, IReadOnlyList<WordDocumentNoteNumberRule> endnotes)
    {
        if (footnotes.Count != firstBlocks.Count || endnotes.Count != firstBlocks.Count) throw new ArgumentException("One rule per section and kind.");
        _firstBlocks = firstBlocks; _footnotes = footnotes; _endnotes = endnotes;
    }

    public static readonly WordDocumentNoteNumbering Defaults = new([], [], []);

    /// <summary>The section holding <paramref name="blockIndex"/>: its first block and its rule for <paramref name="kind"/>.</summary>
    public (int firstBlock, WordDocumentNoteNumberRule rule) SectionOf(WordDocumentNoteKind kind, int blockIndex)
    {
        int low = 0, high = _firstBlocks.Count - 1, found = -1;
        while (low <= high)
        {
            int middle = (low + high) >> 1;
            if (_firstBlocks[middle] <= blockIndex) { found = middle; low = middle + 1; } else high = middle - 1;
        }
        if (found < 0) return (0, _firstBlocks.Count > 0 ? (kind == WordDocumentNoteKind.Footnote ? _footnotes[0] : _endnotes[0]) : WordDocumentNoteNumberRule.Default(kind));
        return (_firstBlocks[found], kind == WordDocumentNoteKind.Footnote ? _footnotes[found] : _endnotes[found]);
    }

    /// <summary>The printed label of an auto-numbered note (ordinal &gt; 0); "*" for a custom mark (ordinal 0).</summary>
    /// <param name="ordinalsBefore">How many auto-numbered notes of a kind lie in the blocks before a given block (asked only for a restarting section).</param>
    public string Label(WordDocumentNoteKind kind, int ordinal, int blockIndex, Func<WordDocumentNoteKind, int, int> ordinalsBefore)
    {
        if (ordinal <= 0) return "*";
        (int firstBlock, WordDocumentNoteNumberRule rule) = SectionOf(kind, blockIndex);
        int before = rule.Restart == WordDocumentNoteRestart.Continuous ? 0 : ordinalsBefore(kind, firstBlock);
        return DocxHtml.DocxNumberingDefinitions.FormatNumber(rule.Number(ordinal, before), rule.Format);
    }
}
