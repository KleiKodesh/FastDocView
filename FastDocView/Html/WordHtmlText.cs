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
}
