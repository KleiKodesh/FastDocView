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

    /// <summary>False for a book without line-1 metadata: nothing is hidden, block 0 is an ordinary line.</summary>
    public bool HasHead => Metadata.Count > 0;

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
    /// <summary>
    /// The decision is made on the detection sample only (the same bytes on a first open, a one-pass open and a rebuild), so a file's result never
    /// changes. Real heads are under 2.5 KB (1,530 files); in 687 user books with a line-1 option the first heading mark is within 31 lines except
    /// one at line 101 (53 KB in), so heading marks are looked for in <see cref="MarkSearchLines"/> lines of the whole sample.
    /// </summary>
    public const int MarkSearchLines = 200;
    /// <summary>A "$ " title line without line-1 metadata must be within this many lines.</summary>
    public const int TitleSearchLines = 100;

    /// <summary>The keys of line 1 (the corpus' whole vocabulary plus the two lowercase options, wrap and rav); an "&amp;word=" that is not one is part of the value.</summary>
    private static readonly string[] KnownKeys =
    {
        "UniqueId", "CosmeticsType", "TextSource", "LastLevelIndex", "ForcedBookName", "GroupId", "SearchInTitles", "PutNewLinesAsIs",
        "PutNewlinesAsIs", "Comments", "HiddenFromIndex", "MixedDisplay", "RavMechaber", "FilesGotFrom", "Haskamot", "SpecialTitle",
        "ChainFolderName", "PasukIndMode", "Mehaber", "CreateIndex", "ParagraphPrefix", "SourceUrl", "Group", "CommentSZ", "rav", "wrap",
    };

    /// <summary>Inline tags that are drawn (bare, never with their attributes); everything else is dropped.</summary>
    internal static readonly HashSet<string> FormattingTags = new(StringComparer.Ordinal) { "b", "i", "u", "big", "small", "sup", "sub", "br" };

    /// <summary>Heading level of a line mark (dollar 1, hash 2, at and caret 3, tilde 4), -1 for the exclamation mark (text, mark hidden), 0 for any other character.</summary>
    public static int LevelOfMark(int mark) => mark switch
    {
        '$' => 1, '#' => 2, '@' => 3, '^' => 3, '~' => 4, '!' => -1, _ => 0,
    };

    /// <summary>True when the line starts with a mark and a space: the mark and the space are not shown.</summary>
    public static bool HasMark(string line, out int level)
    {
        level = line.Length >= 2 && line[1] == ' ' ? LevelOfMark(line[0]) : 0;
        return level != 0;
    }

    /// <summary>
    /// Reads the head from the detection sample of a .txt, or null when the file is not a Torat Emet book. Two ways in:
    /// (a) line 1 starts with an ampersand and holds a known "Key=" (wrap and rav count): a book when a heading-mark line (dollar, hash, at, caret or
    /// tilde and a space) follows within <see cref="MarkSearchLines"/> lines, or when the line holds only wrap / rav options (a book in the format that
    /// draws plain: line 1 is a hidden head); no title line is needed. (b) no line-1 metadata (and line 1 is not an ampersand line): a dollar title line within
    /// <see cref="TitleSearchLines"/> lines and at least one hash, at or tilde heading line; the head is then empty. Single-byte and UTF-8 text only.
    /// </summary>
    public static ToratEmetHead? TryReadHead(ReadOnlySpan<byte> sample, int bomLength, Encoding encoding)
    {
        if (encoding.CodePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian) return null;
        if (bomLength >= sample.Length) return null;
        var lines = new List<(int start, int length)>();
        for (int at = bomLength; at < sample.Length && lines.Count < MarkSearchLines + 1;)
        {
            int next = sample.Slice(at, sample.Length - at).IndexOfAny((byte)'\r', (byte)'\n');
            if (next < 0) break;                              // the line does not end inside the sample: not counted
            lines.Add((at, next));
            at += next;
            if (at + 1 < sample.Length && sample[at] == '\r' && sample[at + 1] == '\n') at++;
            at++;
        }
        if (lines.Count == 0) return null;
        bool startsWithAmpersand = lines[0].length > 0 && sample[lines[0].start] == (byte)'&';
        List<KeyValuePair<string, string>> pairs = startsWithAmpersand
            ? ParseKeyValueLine(encoding.GetString(sample.Slice(lines[0].start, lines[0].length).ToArray())) : [];
        if (startsWithAmpersand && pairs.Count == 0) return null;   // an unrecognised ampersand line: plain text

        if (pairs.Count == 0)
        {
            bool hasTitle = false, hasSubHeading = false;
            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].length < 2 || sample[lines[i].start + 1] != (byte)' ') continue;
                byte mark = sample[lines[i].start];
                if (mark == (byte)'$' && i < TitleSearchLines) hasTitle = true;
                else if (mark is (byte)'#' or (byte)'@' or (byte)'~') hasSubHeading = true;
            }
            return hasTitle && hasSubHeading ? new ToratEmetHead(lines[0].start, pairs) : null;
        }

        bool optionsOnly = IsOptionsOnly(pairs);
        int headLines = 1;
        for (; headLines < lines.Count; headLines++)
        {
            string line = encoding.GetString(sample.Slice(lines[headLines].start, lines[headLines].length).ToArray());
            if (!IsMetadataLine(line, pairs)) break;
        }
        if (headLines >= lines.Count) return null;
        bool hasMark = false;
        for (int i = headLines; i < lines.Count && !hasMark; i++)
            hasMark = lines[i].length >= 2 && sample[lines[i].start + 1] == (byte)' ' && "$#@^~".IndexOf((char)sample[lines[i].start]) >= 0;
        if (!hasMark && !optionsOnly) return null;
        return new ToratEmetHead(lines[headLines].start, pairs);
    }

    /// <summary>Line 1 holds only the wrap / rav options: such a file is in the format even when it has no heading mark.</summary>
    private static bool IsOptionsOnly(List<KeyValuePair<string, string>> pairs)
    {
        foreach (KeyValuePair<string, string> pair in pairs)
            if (!string.Equals(pair.Key, "wrap", StringComparison.OrdinalIgnoreCase) && !string.Equals(pair.Key, "rav", StringComparison.OrdinalIgnoreCase)) return false;
        return true;
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
        "\n.word-document .te-meta { white-space: normal; margin: 0 0 1em; padding-bottom: .5em; border-bottom: 1px solid #ccc; }"
        + "\n.word-document .te-meta-name { font-size: 1.4em; font-weight: bold; }"
        + "\n.word-document .te-meta-line { color: #666; font-size: .9em; }"
        + "\n.word-document .te-heading > h1, .word-document .te-heading > h2, .word-document .te-heading > h3, .word-document .te-heading > h4 { margin: .3em 0; }";

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
        if (HasMark(content, out int markLevel)) { level = markLevel; content = content.Substring(2); if (markLevel < 0) content = MergeLabelBreak(content); }
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
    public static string WithoutMark(string line) => HasMark(line, out int level) ? (level < 0 ? MergeLabelBreak(line.Substring(2)) : line.Substring(2)) : line;

    /// <summary>A label block holds the label line and the line merged into it: the line break between them becomes one space (a break at the very end stays for the caller to trim).</summary>
    private static string MergeLabelBreak(string content)
    {
        int end = content.Length;
        while (end > 0 && (content[end - 1] == '\n' || content[end - 1] == '\r')) end--;
        int at = content.IndexOfAny(['\r', '\n']);
        if (at < 0 || at >= end) return content;
        int after = at + (content[at] == '\r' && at + 1 < content.Length && content[at + 1] == '\n' ? 2 : 1);
        return content.Substring(0, at) + " " + content.Substring(after);
    }

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
