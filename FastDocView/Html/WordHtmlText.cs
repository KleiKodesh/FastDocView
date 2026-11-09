using System.Globalization;
using System.Net;
using System.Text;

namespace FastDocView.Html;

/// <summary>HTML building helpers shared by both renderers.</summary>
public static class WordHtmlText
{
    public static string Encode(string text) => WebUtility.HtmlEncode(text);

    public static void AppendOpenTag(StringBuilder html, string tag, string? css, string? className, bool? isRightToLeft, string? extraAttributes = null)
    {
        html.Append('<').Append(tag);
        if (className is not null) html.Append(" class=\"").Append(className).Append('"');
        if (isRightToLeft is { } rtl) html.Append(" dir=\"").Append(rtl ? "rtl" : "ltr").Append('"');
        if (!string.IsNullOrEmpty(css)) html.Append(" style=\"").Append(css).Append('"');
        if (extraAttributes is not null) html.Append(' ').Append(extraAttributes);
        html.Append('>');
    }

    /// <summary>
    /// CSS for a right-to-left run inside a left-to-right paragraph: ECMA-376 1st edition, rtl (Right To Left Text; read in the docx4java WordprocessingML reference) — the run's
    /// contents, punctuation included, are laid out right to left (".three ,two ,one :This is a list"). An embedding, not
    /// an isolate (a dir attribute on a span isolates in HTML): consecutive right-to-left runs and the neutrals between
    /// them still order right to left as one stretch, as the Unicode bidi algorithm does for RLE..PDF.
    /// </summary>
    public static string RunDirectionCss(bool isRunRightToLeft, bool isParagraphRightToLeft)
        => isRunRightToLeft && !isParagraphRightToLeft ? "unicode-bidi:embed;direction:rtl;" : "";

    /// <summary>A font face for a quoted CSS font-family inside a style attribute: quotes and markup characters dropped.</summary>
    public static string CssFontName(string face)
    {
        var name = new StringBuilder(face.Length);
        foreach (char c in face) if (c is not ('\'' or '"' or '<' or '>' or '&' or ';' or '\\')) name.Append(c);
        return name.ToString();
    }

    public static string HeadingTagFor(int outlineLevel) => outlineLevel is >= 0 and <= 5 ? "h" + (outlineLevel + 1) : "p";

    /// <summary>Maps Word highlight names (and the .doc ico palette through them) to CSS.</summary>
    public static string? HighlightCss(string? wordName) => wordName?.ToLowerInvariant() switch
    {
        null or "none" => null,
        "darkyellow" => "#808000", "darkgreen" => "#008000", "darkblue" => "#000080", "darkred" => "#800000",
        "darkcyan" => "#008080", "darkmagenta" => "#800080", "darkgray" => "#808080", "lightgray" => "#c0c0c0",
        var other => other,
    };

    /// <summary>
    /// Appends already-encoded run HTML, drawn at a character scale of <paramref name="scalePercent"/> (Word's "Scale",
    /// 1..600, anything else = 100 as LibreOffice reads it: DomainMapper.cxx:2495-2506, ww8par6.cxx:5004-5010). CSS has no
    /// horizontal glyph scale that keeps line wrapping, so each word becomes its own box: an inline-flex box exactly
    /// <c>scale x</c> its natural width (calc-size(max-content, ...)), its text scaled with transform:scaleX and centred in it
    /// (an overflowing flex item centres both ways, so the result is the same in right-to-left and left-to-right text;
    /// flex:none keeps the item at its natural width where the page lets words break anywhere).
    /// Lines still break between words; spaces stay unscaled. Tags inside a word stay inside its box; a space inside an
    /// open tag does not split.
    /// </summary>
    public static void AppendScaled(StringBuilder html, string encoded, int scalePercent)
    {
        if (scalePercent is 0 or 100 or < 1 or > 600) { html.Append(encoded); return; }
        string factor = (scalePercent / 100.0).ToString("0.##", CultureInfo.InvariantCulture);
        string open = "<span class=\"char-scale\" style=\"display:inline-flex;justify-content:center;width:calc-size(max-content,size * " + factor
                      + ");text-decoration:inherit\"><span style=\"display:inline-block;flex:none;transform:scaleX(" + factor + ");text-decoration:inherit\">";
        int depth = 0, wordStart = -1;
        for (int i = 0; i < encoded.Length; i++)
        {
            char c = encoded[i];
            if (c == '<')
            {
                int end = encoded.IndexOf('>', i);
                if (end < 0) end = encoded.Length - 1;
                bool isClose = i + 1 < encoded.Length && encoded[i + 1] == '/';
                if (isClose && depth == 0)
                {
                    // a close tag opened before this text: end the word box first, so the boxes stay balanced
                    if (wordStart >= 0) { html.Append(open).Append(encoded, wordStart, i - wordStart).Append("</span></span>"); wordStart = -1; }
                    html.Append(encoded, i, end - i + 1);
                    i = end;
                    continue;
                }
                if (wordStart < 0) wordStart = i;
                if (isClose) depth--;
                else if (encoded[end - 1] != '/' && !IsVoidTag(encoded, i + 1)) depth++;
                i = end;
                continue;
            }
            bool isSeparator = depth == 0 && c is ' ' or '\t' or '\n';
            if (isSeparator)
            {
                if (wordStart >= 0) { html.Append(open).Append(encoded, wordStart, i - wordStart).Append("</span></span>"); wordStart = -1; }
                html.Append(c);
            }
            else if (wordStart < 0) wordStart = i;
        }
        if (wordStart >= 0) html.Append(open).Append(encoded, wordStart, encoded.Length - wordStart).Append("</span></span>");
    }

    private static bool IsVoidTag(string html, int nameStart)
    {
        int end = nameStart;
        while (end < html.Length && char.IsLetter(html[end])) end++;
        string name = html.Substring(nameStart, end - nameStart);
        return name is "br" or "img" or "wbr";
    }
}
