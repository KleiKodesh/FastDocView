using System.Text;

namespace WordDocumentRandomAccess.Markup;

/// <summary>
/// Word's HTML carries its styles as CSS in the head (p.MsoNormal, h1, span.MsoFootnoteReference,
/// @font-face, @page, @list...). The cheap mapping: keep the plain rules, scoped under
/// <c>.word-document</c> so they cannot leak into the host page; drop @-rules (fonts, pages, Word's
/// list definitions) and the &lt;!-- --&gt; / comment wrappers. Unknown mso-* properties are ignored by browsers.
/// </summary>
public static class WebArchiveStylesheet
{
    private const int MaximumCssCharacters = 256 * 1024;

    public static string Scope(string css)
    {
        if (css.Length == 0) return "";
        if (css.Length > MaximumCssCharacters) css = css.Substring(0, MaximumCssCharacters);
        css = css.Replace("<!--", " ").Replace("-->", " ");
        var text = new StringBuilder(css.Length);
        for (int i = 0; i < css.Length; i++)          // strip /* comments */
        {
            if (css[i] == '/' && i + 1 < css.Length && css[i + 1] == '*')
            {
                int close = css.IndexOf("*/", i + 2, StringComparison.Ordinal);
                i = close < 0 ? css.Length : close + 1;
                continue;
            }
            text.Append(css[i]);
        }
        string source = text.ToString();
        var scoped = new StringBuilder();
        int at = 0;
        while (at < source.Length)
        {
            while (at < source.Length && char.IsWhiteSpace(source[at])) at++;
            if (at >= source.Length) break;
            int open = source.IndexOf('{', at);
            if (source[at] == '@')
            {
                int semicolon = source.IndexOf(';', at);
                if (open < 0 || (semicolon >= 0 && semicolon < open)) { at = semicolon < 0 ? source.Length : semicolon + 1; continue; }
                at = MatchingBrace(source, open) + 1;
                continue;
            }
            if (open < 0) break;
            int close = MatchingBrace(source, open);
            string selectors = source.Substring(at, open - at).Trim();
            string body = source.Substring(open + 1, Math.Max(0, close - open - 1)).Trim();
            at = close + 1;
            if (selectors.Length == 0 || selectors.IndexOfAny(['<', '>', '{', '}']) >= 0 || body.IndexOf('<') >= 0) continue;
            bool first = true;
            foreach (string selector in selectors.Split(','))
            {
                string trimmed = selector.Trim();
                if (trimmed.Length == 0) continue;
                if (!first) scoped.Append(", ");
                scoped.Append(".word-document ").Append(trimmed);
                first = false;
            }
            if (!first) scoped.Append(" { ").Append(body).Append(" }\n");
        }
        return scoped.ToString();
    }

    private static int MatchingBrace(string text, int open)
    {
        int depth = 0;
        for (int i = open; i < text.Length; i++)
        {
            if (text[i] == '{') depth++;
            else if (text[i] == '}' && --depth == 0) return i;
        }
        return text.Length - 1;
    }
}
