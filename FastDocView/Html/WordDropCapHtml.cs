using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>
/// A drop cap (an initial dropped over N lines) as HTML, shared by every format. The initial is a span with the class
/// drop-cap and --drop-cap-lines; the stylesheet draws it as a float whose font size is N line heights. CSS initial-letter
/// was tried first: Chromium (Playwright, 2026-11-09) computes it on the span but draws a normal-size letter, so the float is
/// the only form that drops. The gap to the text is the span's end margin.
/// Word keeps the initial in a framed paragraph of its own before the text paragraph (LibreOffice joins the two the same
/// way: DomainMapper_Impl.cxx finishParagraph "(4) old DropCap and no current frame", ww8par.cxx m_bDropCap /
/// m_xPreviousNode); the renderers draw that frame paragraph's runs inside this span at the start of the next paragraph.
/// </summary>
public static class WordDropCapHtml
{
    public static void Open(StringBuilder html, WordParagraphFormatting dropCap)
    {
        int lines = Math.Min(Math.Max(dropCap.DropCapLines, 1), 10);    // Word's Drop Cap dialog offers 1-10 lines
        html.Append("<span class=\"drop-cap").Append(dropCap.IsDropCapInMargin ? " drop-cap-in-margin" : "")
            .Append("\" style=\"--drop-cap-lines:").Append(lines.ToString(CultureInfo.InvariantCulture)).Append(';');
        if (dropCap.DropCapDistanceTwips > 0)
        {
            string gap = WordParagraphFormatting.TwipsToPt(dropCap.DropCapDistanceTwips) + "pt";
            html.Append("margin-inline-end:").Append(gap).Append(";--drop-cap-gap:").Append(gap).Append(';');
        }
        html.Append("\">");
    }

    public static void Close(StringBuilder html) => html.Append("</span>");
}
