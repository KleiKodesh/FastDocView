using System.Text;

namespace FastDocView.Html;

/// <summary>
/// A table row's height as Word draws it (.doc sprmTDyaRowHeight, RTF \trrh: twips, positive = at least, negative = exact,
/// 0 = auto - LibreOffice ww8par2.cxx:3150-3162, rtfdispatchvalue.cxx:437-451). The &lt;tr&gt; gets the height (a table row
/// height is a minimum in CSS); an exact height also clips each single-row cell's content, as Word does (the .docx renderer does
/// the same for w:trHeight hRule="exact").
/// </summary>
public static class WordTableRowHeight
{
    public static void AppendRowOpen(StringBuilder html, int heightTwips)
    {
        if (heightTwips == 0) { html.Append("<tr>"); return; }
        html.Append("<tr style=\"height:").Append(WordParagraphFormatting.TwipsToPt(Math.Abs(heightTwips))).Append("pt\">");
    }

    /// <summary>A cell's content; under an exact height, clipped to it (a cell spanning several rows is as tall as all of them: not clipped).</summary>
    public static void AppendCellContent(StringBuilder html, string content, int heightTwips, int rowSpan)
    {
        if (heightTwips >= 0 || rowSpan > 1) { html.Append(content); return; }
        html.Append("<div style=\"max-height:").Append(WordParagraphFormatting.TwipsToPt(-heightTwips)).Append("pt;overflow:hidden\">").Append(content).Append("</div>");
    }
}
