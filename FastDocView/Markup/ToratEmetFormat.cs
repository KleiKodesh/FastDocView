using System.Text;

namespace FastDocView.Markup;

/// <summary>
/// The book files of the Torat Emet collection, read as .txt: line 1 is metadata (<c>&amp;Key=Value&amp;Key=Value...</c>), a few
/// metadata lines may follow (blank, <c>// comment</c>, <c>[Section]</c>, <c>key=value</c>), then the text. A line that starts
/// with a mark and a space is a heading (dollar = book title, hash = title, at and caret = sub-title, tilde = chapter) or, for the
/// exclamation mark, plain text whose mark is hidden. Inline formatting is a small whitelist of tags; every other tag is dropped, its
/// text kept. The metadata head is one hidden block of the document (block 0), shown as a small header instead.
/// </summary>
public sealed class ToratEmetHead
{
    /// <summary>Where the head ends (text coordinates): the start of the first line that is not metadata.</summary>
    public long HeadEndOffset { get; }
    public IReadOnlyList<KeyValuePair<string, string>> Metadata { get; }

    public ToratEmetHead(long headEndOffset, IReadOnlyList<KeyValuePair<string, string>> metadata)
    {
        HeadEndOffset = headEndOffset;
        Metadata = metadata;
    }

    public string? Get(string key)
    {
        foreach (KeyValuePair<string, string> pair in Metadata)
            if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase)) return pair.Value;
        return null;
    }

    /// <summary>ForcedBookName: the document's title, null when the file has none (the first title heading is then the title).</summary>
    public string? BookName => NonEmpty(Get("ForcedBookName"));
    public string? Source => NonEmpty(Get("TextSource"));
    public string? Author => NonEmpty(Get("RavMechaber")) ?? NonEmpty(Get("Mehaber"));
    public string? SpecialTitle => NonEmpty(Get("SpecialTitle"));

    private static string? NonEmpty(string? value) => string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
}

public static class ToratEmetFormat
{
    /// <summary>The head is looked for in the first bytes only: the real heads are under 2.5 KB (1,530 files).</summary>
    public const int HeadSearchBytes = 32 * 1024;
    public const int HeadingSearchLines = 100;

    /// <summary>The keys of line 1 (the corpus' whole vocabulary plus the two lowercase ones); an "&amp;word=" that is not one is part of the value.</summary>
    private static readonly string[] KnownKeys =
    {
        "UniqueId", "CosmeticsType", "TextSource", "LastLevelIndex", "ForcedBookName", "GroupId", "SearchInTitles", "PutNewLinesAsIs",
        "PutNewlinesAsIs", "Comments", "HiddenFromIndex", "MixedDisplay", "RavMechaber", "FilesGotFrom", "Haskamot", "SpecialTitle",
        "ChainFolderName", "PasukIndMode", "Mehaber", "CreateIndex", "ParagraphPrefix", "SourceUrl", "Group", "CommentSZ", "rav", "wrap",
    };

    /// <summary>Inline tags that are drawn (bare, never with their attributes); everything else is dropped.</summary>
    internal static readonly HashSet<string> FormattingTags = new(StringComparer.Ordinal) { "b", "i", "u", "big", "small", "sup", "sub", "br" };

    /// <summary>Heading level of a line mark (dollar 2, hash 3, at and caret 4, tilde 5), -1 for the exclamation mark (text, mark hidden), 0 for any other character.</summary>
    public static int LevelOfMark(int mark) => mark switch
    {
        '$' => 2, '#' => 3, '@' => 4, '^' => 4, '~' => 5, '!' => -1, _ => 0,
    };

    /// <summary>True when the line starts with a mark and a space: the mark and the space are not shown.</summary>
    public static bool HasMark(string line, out int level)
    {
        level = line.Length >= 2 && line[1] == ' ' ? LevelOfMark(line[0]) : 0;
        return level != 0;
    }

    /// <summary>
    /// Reads the head from the detection sample of a .txt, or null when the file is not a Torat Emet book: line 1 must start with
    /// an ampersand and hold at least one known "Key=", and a heading line (dollar and a space) must follow within <see cref="HeadingSearchLines"/> lines.
    /// Single-byte and UTF-8 text only (the marks are ASCII bytes).
    /// </summary>
    public static ToratEmetHead? TryReadHead(ReadOnlySpan<byte> sample, int bomLength, Encoding encoding)
    {
        if (encoding.CodePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian) return null;
        if (bomLength >= sample.Length || sample[bomLength] != (byte)'&') return null;
        int length = Math.Min(sample.Length, bomLength + HeadSearchBytes);
        var lines = new List<(int start, int length)>();
        for (int at = bomLength; at < length && lines.Count < HeadingSearchLines + 1;)
        {
            int next = sample.Slice(at, length - at).IndexOfAny((byte)'\r', (byte)'\n');
            if (next < 0) break;                              // the line does not end inside the search window: not a head
            lines.Add((at, next));
            at += next;
            if (at + 1 < length && sample[at] == '\r' && sample[at + 1] == '\n') at++;
            at++;
        }
        if (lines.Count == 0) return null;
        string first = encoding.GetString(sample.Slice(lines[0].start, lines[0].length).ToArray());
        List<KeyValuePair<string, string>> pairs = ParseKeyValueLine(first);
        if (pairs.Count == 0) return null;
        int headLines = 1;
        for (; headLines < lines.Count; headLines++)
        {
            string line = encoding.GetString(sample.Slice(lines[headLines].start, lines[headLines].length).ToArray());
            if (!IsMetadataLine(line, pairs)) break;
        }
        bool hasHeading = false;
        for (int i = Math.Min(headLines, lines.Count); i < lines.Count && !hasHeading; i++)
            hasHeading = lines[i].length >= 2 && sample[lines[i].start] == (byte)'$' && sample[lines[i].start + 1] == (byte)' ';
        if (!hasHeading || headLines >= lines.Count) return null;
        return new ToratEmetHead(lines[headLines].start, pairs);
    }

    /// <summary>"&amp;Key=Value&amp;Key=Value": a key is a known name between an ampersand and "="; everything up to the next known key is the value ('=' and '&amp;' inside it stay).</summary>
    internal static List<KeyValuePair<string, string>> ParseKeyValueLine(string line)
    {
        var pairs = new List<KeyValuePair<string, string>>();
        int valueStart = 0;
        string? key = null;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] != '&') continue;
            int equals = line.IndexOf('=', i + 1);
            if (equals < 0) break;
            string name = line.Substring(i + 1, equals - i - 1);
            if (!IsKnownKey(name)) continue;
            if (key is not null) pairs.Add(new(key, line.Substring(valueStart, i - valueStart).Trim()));
            key = name; valueStart = equals + 1;
        }
        if (key is not null) pairs.Add(new(key, line.Substring(valueStart).Trim()));
        return pairs;
    }

    private static bool IsKnownKey(string name)
    {
        foreach (string known in KnownKeys) if (string.Equals(name, known, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    /// <summary>A line of the head that is clearly metadata: blank, "// comment", "[Section]", "key=value", or a copy of the "&amp;Key=" line (even with tags around it).</summary>
    private static bool IsMetadataLine(string line, List<KeyValuePair<string, string>> pairs)
    {
        string text = line.Trim();
        if (text.Length == 0 || text.StartsWith("//", StringComparison.Ordinal)) return true;
        if (text[0] == '[' && text[text.Length - 1] == ']' && text.IndexOf('[', 1) < 0) return true;
        if (text[0] < 128 && (char.IsLetter(text[0]) || text[0] == '_'))
        {
            int i = 0;
            while (i < text.Length && text[i] < 128 && (char.IsLetterOrDigit(text[i]) || text[i] == '_')) i++;
            int nameEnd = i;
            while (i < text.Length && text[i] == ' ') i++;
            if (i < text.Length && text[i] == '=')
            {
                pairs.Add(new(text.Substring(0, nameEnd), text.Substring(i + 1).Trim().Trim('"')));
                return true;
            }
        }
        for (int amp = text.IndexOf('&'); amp >= 0; amp = text.IndexOf('&', amp + 1))
        {
            int equals = text.IndexOf('=', amp + 1);
            if (equals > 0 && IsKnownKey(text.Substring(amp + 1, equals - amp - 1))) return true;
        }
        return false;
    }

    // -- rendering ------------------------------------------------------------------------

    /// <summary>The metadata header shown in place of the hidden head: the book's name, its author and its source.</summary>
    public static string RenderHeader(ToratEmetHead head, string? firstTitle)
    {
        string? name = head.BookName ?? firstTitle;
        var html = new StringBuilder("<div class=\"text-line te-meta\">");
        if (name is not null) html.Append("<div class=\"te-meta-name\">").Append(Escape(name)).Append("</div>");
        if (head.SpecialTitle is { } special) html.Append("<div class=\"te-meta-line\">").Append(Escape(special)).Append("</div>");
        if (head.Author is { } author) html.Append("<div class=\"te-meta-line\">").Append(Escape(author)).Append("</div>");
        if (head.Source is { } source) html.Append("<div class=\"te-meta-line te-meta-source\">").Append(Escape(source)).Append("</div>");
        return html.Append("</div>").ToString();
    }

    private static string Escape(string text) => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    public const string Stylesheet =
        "\n.word-document .te-meta { white-space: normal; margin: 0 0 1em; padding-bottom: .5em; border-bottom: 1px solid var(--fdv-rule, #ccc); }"
        + "\n.word-document .te-meta-name { font-size: 1.4em; font-weight: bold; }"
        + "\n.word-document .te-meta-line { color: var(--fdv-muted, #666); font-size: .9em; }"
        + "\n.word-document .te-heading > h2, .word-document .te-heading > h3, .word-document .te-heading > h4, .word-document .te-heading > h5 { margin: .3em 0; }";

    /// <summary>
    /// One line of text: the mark (and its space) removed, a heading line wrapped in its h2-h5, the formatting whitelist drawn (the tags still
    /// open from earlier lines first), everything else dropped. Never raw HTML: only tags this method writes.
    /// </summary>
    public static string RenderLine(string openTags, string line)
    {
        int end = line.Length;
        while (end > 0 && (line[end - 1] == '\n' || line[end - 1] == '\r')) end--;
        string content = line.Substring(0, end);
        int level = 0;
        if (HasMark(content, out int markLevel)) { level = markLevel; content = content.Substring(2); }
        string body = RenderInline(openTags, content);
        if (level > 0) return "<div class=\"text-line te-heading\"><h" + level + ">" + (body.Length == 0 ? "<br>" : body) + "</h" + level + "></div>";
        return "<div class=\"text-line\">" + (IsVisiblyEmpty(body) ? "<br>" : body) + "</div>";
    }

    private static bool IsVisiblyEmpty(string html)
    {
        foreach (char c in html) if (c == '<' || !char.IsWhiteSpace(c)) return false;
        return true;
    }

    /// <summary>The line without its mark: what a search matches (the same removal as <see cref="RenderLine"/>).</summary>
    public static string WithoutMark(string line) => HasMark(line, out _) ? line.Substring(2) : line;

    internal static string RenderInline(string openTags, string content)
    {
        string text = openTags + content;
        var html = new StringBuilder(text.Length + 16);
        var open = new List<string>();
        int i = 0;
        while (i < text.Length)
        {
            if (text[i] != '<') { int next = text.IndexOf('<', i); if (next < 0) next = text.Length; html.Append(text, i, next - i); i = next; continue; }
            char d = i + 1 < text.Length ? text[i + 1] : '\0';
            if (d == '!' && string.CompareOrdinal(text, i, "<!--", 0, 4) == 0)
            {
                int close = text.IndexOf("-->", i + 4, Math.Min(text.Length - i - 4, MarkupBlockScanner.MaximumCommentBytes + 3), StringComparison.Ordinal);
                if (close < 0 && text.Length - i - 4 > MarkupBlockScanner.MaximumCommentBytes) { html.Append("&lt;"); i++; continue; }
                i = close < 0 ? text.Length : close + 3;
                continue;
            }
            if (d == '!' || d == '?')
            {
                int close = text.IndexOf('>', i + 2, Math.Min(text.Length - i - 2, MarkupBlockScanner.MaximumCommentBytes + 1));
                if (close < 0 && text.Length - i - 2 > MarkupBlockScanner.MaximumCommentBytes) { html.Append("&lt;"); i++; continue; }
                i = close < 0 ? text.Length : close + 1;
                continue;
            }
            bool isEnd = d == '/';
            char first = isEnd ? (i + 2 < text.Length ? text[i + 2] : '\0') : d;
            int tagEnd = IsTagStart(first) ? FindTagEnd(text, i) : -1;
            if (tagEnd < 0) { html.Append("&lt;"); i++; continue; }
            (string name, _, bool isSelfClosing) = MarkupBlockScanner.ParseTag(text.Substring(i, tagEnd + 1 - i));
            i = tagEnd + 1;
            if (!FormattingTags.Contains(name)) continue;           // a macro or unknown tag: dropped, its text stays
            if (isEnd)
            {
                int at = open.LastIndexOf(name);
                if (at < 0) continue;
                for (int k = open.Count - 1; k >= at; k--) html.Append("</").Append(open[k]).Append('>');
                open.RemoveRange(at, open.Count - at);
            }
            else
            {
                html.Append('<').Append(name).Append('>');
                if (name != "br" && !isSelfClosing) open.Add(name);
            }
        }
        for (int k = open.Count - 1; k >= 0; k--) html.Append("</").Append(open[k]).Append('>');
        return html.ToString();
    }

    /// <summary>A tag name starts with a letter; the collection's size macros (like 90 or 150 between angle brackets) with a digit.</summary>
    private static bool IsTagStart(char c) => c is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9';

    private static int FindTagEnd(string text, int start)
    {
        bool quoted = false; char quote = '\0', lastNonSpace = 'a';
        int limit = Math.Min(text.Length, start + MarkupBlockScanner.MaximumTagUnits);
        for (int i = start + 1; i < limit; i++)
        {
            char c = text[i];
            if (quoted) { if (c == quote) quoted = false; }
            else if (c == '>') return i;
            else if ((c == '"' || c == '\'') && lastNonSpace == '=') { quoted = true; quote = c; }
            if (c > ' ') lastNonSpace = c;
        }
        return -1;
    }
}
