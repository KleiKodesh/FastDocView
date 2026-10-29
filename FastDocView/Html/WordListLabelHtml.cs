using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>What separates a list label from the paragraph text: .docx w:suff (ECMA-376 17.9.28), .doc LVLF ixchFollow
/// ([MS-DOC] 2.9.150: 0 tab, 1 space, 2 nothing), RTF \levelfollow, ODF label-followed-by.</summary>
public enum WordListLabelSuffix : byte { Tab, Space, Nothing }

/// <summary>Where the label sits against its position (.docx lvlJc, .doc LVLF jc 0/1/2): Start = begins there, End = ends there
/// (it reaches back towards the start side), Center = centred on it. Logical: in a right-to-left paragraph End reaches right.</summary>
public enum WordListLabelAlignment : byte { Start, Center, End }

/// <summary>
/// Draws a list label where Word puts it. The label starts at the first-line position (start indent + first-line
/// indent; the hanging indent puts it before the text column). After a tab suffix the text starts at the next tab stop
/// after the label's start: a custom stop, the start indent itself when the paragraph hangs (Word's implicit list tab
/// stop - LibreOffice ww8par3.cxx / DomainMapper keep it as the list tab position), else the next default stop. The
/// label box is an inline-block that reaches that stop (min-width), so the text begins exactly there; a label wider
/// than the gap pushes the text on (Word would go to the stop after - not modelled, the label's width is unknown here).
/// </summary>
public static class WordListLabelHtml
{
    public const int DefaultTabStopTwips = 720;   // [MS-DOC] Dop dxaTab / ECMA-376 17.15.1.25 defaultTabStop when absent

    public static void Append(StringBuilder html, string label, string? labelCss, WordListLabelSuffix suffix,
                              WordParagraphFormatting paragraph, int defaultTabStopTwips,
                              WordListLabelAlignment alignment = WordListLabelAlignment.Start, bool isLabelHtml = false)
    {
        html.Append("<span class=\"list-marker\"");
        string css = labelCss ?? "";
        if (suffix == WordListLabelSuffix.Tab && TabGapTwips(paragraph, defaultTabStopTwips) is var gap && gap > 0)
            css += "min-width:" + WordParagraphFormatting.TwipsToPt(gap) + "pt;";
        if (css.Length > 0) html.Append(" style=\"").Append(css).Append('"');
        html.Append('>');
        // An end / centre aligned label: a zero-width flex box whose content overflows towards the start side (flex-end)
        // or both sides (center) - the label ends at / is centred on its position whatever its width.
        if (alignment != WordListLabelAlignment.Start)
            html.Append("<span style=\"display:inline-flex;width:0;white-space:nowrap;justify-content:").Append(alignment == WordListLabelAlignment.End ? "flex-end" : "center").Append("\">");
        html.Append(isLabelHtml ? label : WordHtmlText.Encode(label));
        if (alignment != WordListLabelAlignment.Start) html.Append("</span>");
        html.Append("</span>");
        if (suffix == WordListLabelSuffix.Space) html.Append(' ');
    }

    /// <summary>From the label's start to the tab stop its text goes to, twips.</summary>
    public static int TabGapTwips(WordParagraphFormatting paragraph, int defaultTabStopTwips)
    {
        int start = paragraph.StartIndentTwips;
        int labelStart = start + paragraph.FirstLineIndentTwips;
        int stop = int.MaxValue;
        foreach (WordTabStop tab in paragraph.TabStops)
            if (tab.PositionTwips > labelStart && tab.Alignment is not ("clear" or "bar")) { stop = tab.PositionTwips; break; }
        if (paragraph.FirstLineIndentTwips < 0 && start > labelStart) stop = Math.Min(stop, start);
        if (stop == int.MaxValue)
        {
            int step = defaultTabStopTwips > 0 ? defaultTabStopTwips : DefaultTabStopTwips;
            stop = (labelStart < 0 ? 0 : (labelStart / step + 1) * step);
        }
        return stop - labelStart;
    }

}
