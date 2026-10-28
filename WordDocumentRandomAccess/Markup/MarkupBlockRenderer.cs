using System.Text;

namespace WordDocumentRandomAccess.Markup;

/// <summary>
/// One block of HTML-ish text → standalone, safe HTML. The block is parsed after the start tags that
/// were open at its start (from the index), with the same tag rules as <see cref="MarkupBlockScanner"/>:
/// a '<' that does not start a tag is text (escaped), comments and &lt;!...&gt; / &lt;?...?&gt; are dropped,
/// script / style / embedded objects are dropped, event-handler attributes and javascript: URLs are
/// removed, images become a placeholder, and every element left open is closed at the end.
/// </summary>
public static class MarkupBlockRenderer
{
    private static readonly HashSet<string> DroppedTags = new(StringComparer.Ordinal)
        { "html", "head", "body", "meta", "link", "base", "iframe", "frame", "frameset", "object", "embed", "applet", "form", "input", "button", "select", "option" };
    private static readonly HashSet<string> DroppedWithContent = new(StringComparer.Ordinal) { "script", "style", "title", "noscript" };
    private static readonly HashSet<string> VoidElements = new(StringComparer.Ordinal)
        { "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr" };

    public const string ImagePlaceholder = "<span class=\"image-placeholder\">[image]</span>";

    public static string RenderLine(string openTags, string line)
    {
        int end = line.Length;
        while (end > 0 && (line[end - 1] == '\n' || line[end - 1] == '\r')) end--;
        string body = Render(openTags, line.Substring(0, end));
        return "<div class=\"text-line\">" + (IsVisiblyEmpty(body) ? "<br>" : body) + "</div>";
    }

    private static bool IsVisiblyEmpty(string html)
    {
        foreach (char c in html) if (c != '<') { if (!char.IsWhiteSpace(c)) return false; } else return false;
        return true;
    }

    public static string Render(string openTags, string content)
    {
        string text = openTags + content;
        var html = new StringBuilder(text.Length + 32);
        var open = new List<string>();
        int i = 0;
        while (i < text.Length)
        {
            char c = text[i];
            if (c != '<') { int next = text.IndexOf('<', i); if (next < 0) next = text.Length; html.Append(text, i, next - i); i = next; continue; }
            char d = i + 1 < text.Length ? text[i + 1] : '\0';
            if (d == '!' && string.CompareOrdinal(text, i, "<!--", 0, 4) == 0)
            {
                int close = text.IndexOf("-->", i + 4, StringComparison.Ordinal);
                i = close < 0 ? text.Length : close + 3;
                continue;
            }
            if (d == '!' || d == '?')
            {
                int close = text.IndexOf('>', i + 2);
                i = close < 0 ? text.Length : close + 1;
                continue;
            }
            bool isEnd = d == '/';
            char first = isEnd ? (i + 2 < text.Length ? text[i + 2] : '\0') : d;
            int tagEnd = IsAsciiLetter(first) ? FindTagEnd(text, i) : -1;
            if (tagEnd < 0) { html.Append("&lt;"); i++; continue; }
            string tag = text.Substring(i, tagEnd + 1 - i);
            i = tagEnd + 1;
            (string name, _, bool isSelfClosing) = MarkupBlockScanner.ParseTag(tag);
            if (isEnd)
            {
                int at = open.LastIndexOf(name);
                if (at < 0) continue;
                for (int k = open.Count - 1; k >= at; k--) html.Append("</").Append(open[k]).Append('>');
                open.RemoveRange(at, open.Count - at);
                continue;
            }
            if (DroppedWithContent.Contains(name) || MarkupBlockScanner.RawTextElements.Contains(name))
            {
                int close = IndexOfEndTag(text, i, name);
                string raw = text.Substring(i, (close < 0 ? text.Length : close) - i);
                if (!DroppedWithContent.Contains(name)) html.Append("<pre>").Append(Escape(raw)).Append("</pre>");    // textarea / xmp: show the text
                i = close < 0 ? text.Length : Math.Max(close, text.IndexOf('>', close) + 1);
                continue;
            }
            if (name == "img" || name == "v:imagedata") { html.Append(ImagePlaceholder); continue; }
            if (DroppedTags.Contains(name)) continue;
            html.Append(Sanitize(tag, name));
            if (!VoidElements.Contains(name) && !isSelfClosing) open.Add(name);
        }
        for (int k = open.Count - 1; k >= 0; k--) html.Append("</").Append(open[k]).Append('>');
        return html.ToString();
    }

    private static bool IsAsciiLetter(char c) => (c | 0x20) >= 'a' && (c | 0x20) <= 'z';

    /// <summary>The '>' closing the tag at <paramref name="start"/> (quotes after '=' respected), or -1 — the scanner's rule.</summary>
    private static int FindTagEnd(string text, int start)
    {
        bool quoted = false; char quote = '\0', lastNonSpace = 'a';
        int limit = Math.Min(text.Length, start + MarkupBlockScanner.MaximumTagBytes);
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

    private static int IndexOfEndTag(string text, int from, string name)
    {
        string pattern = "</" + name;
        return text.IndexOf(pattern, from, StringComparison.OrdinalIgnoreCase);
    }

    private static string Escape(string text) => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    /// <summary>Whitespace, quote or '/' followed by "on" — a possible event-handler attribute (cheap pre-check).</summary>
    private static bool HasEventAttribute(string tag)
    {
        for (int i = tag.IndexOf("on", StringComparison.OrdinalIgnoreCase); i > 0; i = tag.IndexOf("on", i + 2, StringComparison.OrdinalIgnoreCase))
        {
            char before = tag[i - 1];
            if (char.IsWhiteSpace(before) || before == '"' || before == '\'' || before == '/') return true;
        }
        return false;
    }

    /// <summary>The tag as written, minus event handlers (on*) and script URLs.</summary>
    private static string Sanitize(string tag, string name)
    {
        if (!HasEventAttribute(tag) && tag.IndexOf("script", StringComparison.OrdinalIgnoreCase) < 0) return tag;
        var result = new StringBuilder(tag.Length);
        result.Append('<').Append(name);
        int i = 1 + name.Length;
        while (i < tag.Length)
        {
            while (i < tag.Length && (char.IsWhiteSpace(tag[i]) || tag[i] == '/')) i++;
            if (i >= tag.Length || tag[i] == '>') break;
            int nameStart = i;
            while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '=' && tag[i] != '>' && tag[i] != '/') i++;
            string attribute = tag.Substring(nameStart, i - nameStart);
            string value = "";
            int valueStart = i;
            while (i < tag.Length && char.IsWhiteSpace(tag[i])) i++;
            if (i < tag.Length && tag[i] == '=')
            {
                i++;
                while (i < tag.Length && char.IsWhiteSpace(tag[i])) i++;
                if (i < tag.Length && (tag[i] == '"' || tag[i] == '\''))
                {
                    char quote = tag[i];
                    int close = tag.IndexOf(quote, i + 1);
                    if (close < 0) close = tag.Length - 1;
                    value = tag.Substring(i + 1, close - i - 1);
                    i = close + 1;
                }
                else
                {
                    int start = i;
                    while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '>') i++;
                    value = tag.Substring(start, i - start);
                }
            }
            else i = valueStart;
            if (attribute.Length == 0) { i++; continue; }
            if (attribute.StartsWith("on", StringComparison.OrdinalIgnoreCase)) continue;
            string compact = new string(value.Where(ch => !char.IsWhiteSpace(ch) && !char.IsControl(ch)).ToArray());
            if (compact.IndexOf("script:", StringComparison.OrdinalIgnoreCase) >= 0) continue;
            result.Append(' ').Append(attribute);
            if (valueStart != i) result.Append("=\"").Append(value.Replace("\"", "&quot;")).Append('"');
        }
        if (tag.EndsWith("/>", StringComparison.Ordinal)) result.Append(" /");
        return result.Append('>').ToString();
    }
}
