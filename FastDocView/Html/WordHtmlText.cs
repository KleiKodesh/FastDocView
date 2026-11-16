using System.Globalization;
using System.Net;
using System.Text;

namespace FastDocView.Html;

/// <summary>HTML building helpers shared by both renderers.</summary>
public static class WordHtmlText
{
    public static string Encode(string text) => WebUtility.HtmlEncode(text);

    /// <summary>
    /// Encoded Word run text with each stretch of Arabic letters wrapped in a span that gives the browser's ligatures back:
    /// Word draws no ligatures by default (the document's font-variant-ligatures), but Arabic always with them, whatever
    /// w14:ligatures says (Word 16 XPS: Arial's Allah ligature, 4 characters to 1 glyph, with ligatures none and all alike).
    /// One span per word (a space ends it). Text inside tags is left alone; text with no Arabic letter is returned as it is.
    /// </summary>
    public static string WithWordArabicShaping(string encoded)
    {
        int first = -1;
        for (int i = 0; i < encoded.Length; i++) if (IsArabicLetter(encoded[i])) { first = i; break; }
        if (first < 0) return encoded;
        var html = new StringBuilder(encoded.Length + 64);
        bool inTag = false, inArabic = false;
        for (int i = 0; i < encoded.Length; i++)
        {
            char c = encoded[i];
            if (c == '<') inTag = true;
            // marks and joiners stay inside the stretch; a space ends it, so AppendScaled can still box each word on its own
            bool arabic = !inTag && (IsArabicLetter(c) || (inArabic && (c is '‌' or '‍' || char.GetUnicodeCategory(c) == System.Globalization.UnicodeCategory.NonSpacingMark)));
            if (arabic && !inArabic) { html.Append("<span style=\"font-variant-ligatures:normal\">"); inArabic = true; }
            else if (!arabic && inArabic) { html.Append("</span>"); inArabic = false; }
            html.Append(c);
            if (c == '>') inTag = false;
        }
        if (inArabic) html.Append("</span>");
        return html.ToString();
    }

    private static bool IsArabicLetter(char c)
        => c is (>= '؀' and <= 'ۿ') or (>= 'ݐ' and <= 'ݿ') or (>= 'ࢠ' and <= 'ࣿ') or (>= 'ﭐ' and <= '﷿') or (>= 'ﹰ' and <= '﻿');

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
    /// Lines still break between words; spaces are scaled in boxes of their own. Tags inside a word stay inside its box; a space inside an
    /// open tag does not split.
    /// </summary>
    /// <param name="reflectionCss">A -webkit-box-reflect value (.docx w14:reflection) or null. A reflection mirrors a box, and
    /// an inline box that wraps has none to mirror (Chromium draws nothing), so each word gets its own inline-block, trimmed by
    /// text-box to the text above the baseline: the mirror image then starts at the baseline and sits under its own line,
    /// as Word draws it, while lines still break between words and the line height is unchanged.</param>
    public static void AppendScaled(StringBuilder html, string encoded, int scalePercent, string? reflectionCss = null)
    {
        bool isScaled = scalePercent is not (0 or 100 or < 1 or > 600);
        if (!isScaled && reflectionCss is null) { html.Append(encoded); return; }
        string reflection = reflectionCss is null ? "" : ";-webkit-box-reflect:" + reflectionCss;
        string open, close;
        string? spaceOpen = null;
        if (isScaled)
        {
            string factor = (scalePercent / 100.0).ToString("0.##", CultureInfo.InvariantCulture);
            open = "<span class=\"char-scale\" style=\"display:inline-flex;justify-content:center;width:calc-size(max-content,size * " + factor
                   + ");text-decoration:inherit\"><span style=\"display:inline-block;flex:none;transform:scaleX(" + factor + ");text-decoration:inherit\">";
            close = "</span></span>";
            spaceOpen = open.Replace("class=\"char-scale\"", "class=\"char-scale-space\"").Replace("flex:none;", "flex:none;white-space:pre;");
            // the mirror goes on the scaled word itself, trimmed to the baseline as in the unscaled path (spaces have nothing to mirror)
            if (reflectionCss is not null) open = open.Replace("transform:scaleX(" + factor + ");text-decoration:inherit", "transform:scaleX(" + factor + ");text-box:trim-both text alphabetic;text-decoration:inherit" + reflection);
        }
        else
        {
            open = "<span class=\"text-reflection\" style=\"display:inline-block;text-box:trim-both text alphabetic;text-decoration:inherit" + reflection + "\">";
            close = "</span>";
        }
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
                    if (wordStart >= 0) { html.Append(open).Append(encoded, wordStart, i - wordStart).Append(close); wordStart = -1; }
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
                if (wordStart >= 0) { html.Append(open).Append(encoded, wordStart, i - wordStart).Append(close); wordStart = -1; }
                // Word scales the spaces too: a space gets its own scaled box (white-space:pre keeps it from collapsing),
                // and <wbr> after it keeps the line break opportunity the plain space gave
                if (c == ' ' && spaceOpen is not null) html.Append(spaceOpen).Append(' ').Append("</span></span><wbr>");
                else html.Append(c);
            }
            else if (wordStart < 0) wordStart = i;
        }
        if (wordStart >= 0) html.Append(open).Append(encoded, wordStart, encoded.Length - wordStart).Append(close);
    }

    private static bool IsVoidTag(string html, int nameStart)
    {
        int end = nameStart;
        while (end < html.Length && char.IsLetter(html[end])) end++;
        string name = html.Substring(nameStart, end - nameStart);
        return name is "br" or "img" or "wbr";
    }
}
