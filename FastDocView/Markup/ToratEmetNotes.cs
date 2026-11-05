using System.Text;

namespace FastDocView.Markup;

/// <summary>One note line of a Torat Emet book: the bytes [Start, End) of the text (the line and its line break) and the text shown in the header.</summary>
public sealed record ToratEmetNote(long Start, long End, string Text);

/// <summary>
/// The short NOTE lines that sit right after a Torat Emet book's title ("the book is incomplete and unvetted", an editor's credit...). Such a line
/// belongs in the header line above the title, not in the text. The list of known notes is DATA: every text that a corpus analysis found as a plain
/// line between a title and the first section heading in at least 3 books (<c>ToratEmetNotes.Data.cs</c>, generated - regenerate with
/// <c>python -I scripts/analyse-torat-emet-notes.py &lt;books folders&gt; --emit-cs FastDocView/Markup/ToratEmetNotes.Data.cs</c>). A line is matched
/// after <see cref="Normalise"/>: inline tags removed, white space collapsed, trailing punctuation dropped.
/// <para>
/// Where a note is recognised (<see cref="ToratEmetFormat.TryReadHead"/>, on the detection sample only): within the first
/// <see cref="MaximumNonBlankLines"/> non-blank lines after the first title line, with only blank lines, comment lines and other notes between, the
/// line at most <see cref="MaximumLineChars"/> characters long, and a section heading (hash, at, caret or tilde and a space) somewhere later in the
/// sample. A book with a title and no section heading never loses a line this way: its plain text stays text.
/// </para>
/// </summary>
public static partial class ToratEmetNotes
{
    public const int MaximumNonBlankLines = 5;
    public const int MaximumLineChars = 120;

    private static HashSet<string>? _known;   // built on first use: the order of static initialisers across the two partial files is not defined
    private static HashSet<string> Known => _known ??= new HashSet<string>(Listed, StringComparer.Ordinal);

    /// <summary>The comparison form of a line: tags removed, white space collapsed and trimmed, trailing punctuation dropped.</summary>
    public static string Normalise(string line) => TrimTrailingPunctuation(WithoutTags(line));

    /// <summary>True when the line is a listed note; <paramref name="display"/> is the line as the header shows it (tags removed, white space collapsed).</summary>
    public static bool TryMatch(string line, out string display)
    {
        display = WithoutTags(line);
        return display.Length > 0 && display.Length <= MaximumLineChars && Known.Contains(TrimTrailingPunctuation(display));
    }

    private static string WithoutTags(string line)
    {
        var text = new StringBuilder(line.Length);
        bool isSpace = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '<')
            {
                int close = line.IndexOf('>', i + 1);
                int nextOpen = line.IndexOf('<', i + 1);
                if (close > i && (nextOpen < 0 || nextOpen > close)) { i = close; continue; }     // "<...>" holding no other "<": a tag
            }
            if (char.IsWhiteSpace(c)) { isSpace = text.Length > 0; continue; }
            if (isSpace) { text.Append(' '); isSpace = false; }
            text.Append(c);
        }
        return text.ToString();
    }

    private static string TrimTrailingPunctuation(string text)
    {
        int end = text.Length;
        while (end > 0 && text[end - 1] is ' ' or '\t' or '.' or ',' or ';' or ':' or '!' or '?' or '-' or '׃' or '־') end--;
        return end == text.Length ? text : text.Substring(0, end);
    }
}
