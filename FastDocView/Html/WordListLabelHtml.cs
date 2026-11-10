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
/// A Word 6 / 95 style level (.docx w:legacy, ECMA-376 17.9.10; .doc LVLF fWord6 with dxaSpace / dxaIndent). LibreOffice reads
/// w:legacy and drops it (NumberingManager.cxx:119-120) and keeps only dxaIndent as the list tab position for .doc (ww8par3.cxx:976-979),
/// so Word 16 is the reference (make-legacy-lists-fixture.ps1): no tab after the label - the text starts at the furthest of the start
/// indent, label start + IndentTwips, label end + SpaceTwips.
/// </summary>
public readonly record struct WordListLabelLegacy(int SpaceTwips, int IndentTwips);

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
    public const int DefaultTabStopTwips = 720;
    /// <summary>The least room a tab suffix leaves after a label that reaches past its stop (a quarter em, about a space).</summary>
    public const string MinimumTabGapEm = "0.25";   // [MS-DOC] Dop dxaTab / ECMA-376 17.15.1.25 defaultTabStop when absent

    public static void Append(StringBuilder html, string label, string? labelCss, WordListLabelSuffix suffix,
                              WordParagraphFormatting paragraph, int defaultTabStopTwips,
                              WordListLabelAlignment alignment = WordListLabelAlignment.Start, bool isLabelHtml = false, WordListLabelLegacy? legacy = null)
    {
        html.Append("<span class=\"list-marker\"");
        string css = labelCss ?? "";
        if (legacy is { } word6)
        {
            int hanging = Math.Max(0, -paragraph.FirstLineIndentTwips);
            int reach = Math.Max(hanging, word6.IndentTwips);
            if (reach > 0) css += "min-width:" + WordParagraphFormatting.TwipsToPt(reach) + "pt;";
            html.Append(" style=\"").Append(css).Append("\"><span");
            if (word6.SpaceTwips > 0) html.Append(" style=\"padding-inline-end:").Append(WordParagraphFormatting.TwipsToPt(word6.SpaceTwips)).Append("pt\"");
            html.Append('>');
            if (PictureOf(label) is { } legacyPicture) AppendPicture(html, legacyPicture);
            else html.Append(isLabelHtml ? label : WordHtmlText.Encode(label));
            html.Append("</span></span>");
            return;
        }
        if (suffix == WordListLabelSuffix.Tab && TabGapTwips(paragraph, defaultTabStopTwips) is var gap && gap > 0)
            css += "min-width:" + WordParagraphFormatting.TwipsToPt(gap) + "pt;";
        if (css.Length > 0) html.Append(" style=\"").Append(css).Append('"');
        html.Append('>');
        // An end / centre aligned label: a zero-width flex box whose content overflows towards the start side (flex-end)
        // or both sides (center) - the label ends at / is centred on its position whatever its width.
        if (alignment != WordListLabelAlignment.Start)
            html.Append("<span style=\"display:inline-flex;width:0;white-space:nowrap;justify-content:").Append(alignment == WordListLabelAlignment.End ? "flex-end" : "center").Append("\">");
        // A label wider than its gap: Word / LibreOffice search the next tab stop from where the label ENDS (txttab.cxx
        // SwTextFormatter::NewTabPortion: nTabPos = rInf.X(), the stop at the left margin only "inside the hanging indent"),
        // so the text never touches the label. The label's width is unknown here: the tab keeps at least a quarter em,
        // which only shows when the label (nearly) fills its gap.
        bool keepsTabGap = suffix == WordListLabelSuffix.Tab && alignment == WordListLabelAlignment.Start;
        if (keepsTabGap) html.Append("<span style=\"padding-inline-end:").Append(MinimumTabGapEm).Append("em\">");
        if (PictureOf(label) is { } picture) AppendPicture(html, picture);
        else html.Append(isLabelHtml ? label : WordHtmlText.Encode(label));
        if (keepsTabGap) html.Append("</span>");
        if (alignment != WordListLabelAlignment.Start) html.Append("</span>");
        html.Append("</span>");
        if (suffix == WordListLabelSuffix.Space) html.Append(' ');
    }

    /// <summary>Marks a picture bullet's label text: U+FFFC, the picture's width "/" height (any unit, only the ratio counts), U+0001, its picture id.</summary>
    public const char PictureLabelMark = '\uFFFC';

    /// <summary>
    /// The label text of a picture bullet (.docx w:lvlPicBulletId, .doc sprmCPbiIBullet, RTF \levelpicture, ODF
    /// list-level-style-image): drawn as the picture, not as text. Null id: a plain bullet (the picture cannot be shown).
    /// </summary>
    public static string PictureLabel(string? pictureId, double width, double height)
        => pictureId is null || width <= 0 || height <= 0 ? "\u2022"
           : PictureLabelMark + Number(width) + "/" + Number(height) + "\u0001" + pictureId;

    /// <summary>The picture of a <see cref="PictureLabel"/>; null for a text label.</summary>
    public static (string PictureId, double Width, double Height)? PictureOf(string label)
    {
        if (label.Length < 5 || label[0] != PictureLabelMark) return null;
        int slash = label.IndexOf('/'), mark = label.IndexOf('\u0001');
        if (slash < 0 || mark < slash
            || !double.TryParse(label.Substring(1, slash - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double width)
            || !double.TryParse(label.Substring(slash + 1, mark - slash - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double height)) return null;
        return (label.Substring(mark + 1), width, height);
    }

    /// <summary>
    /// Word draws a picture bullet 2 pt less high than the label's font size, whatever the font, keeping only the picture's aspect
    /// ratio, its bottom on the baseline. Measured in Word 16 (PDF export, PyMuPDF; a 2:1 picture bullet of list-bullets-v1.docx at
    /// each size): 8 pt -> 6.0, 10 -> 7.9, 11 -> 9.1, 12 -> 10.05, 14 -> 12.0, 18 -> 16.1, 24 -> 22.05, 36 -> 34.1, 48 -> 46.1 pt high;
    /// Times New Roman and Courier New at 11 / 36 pt the same as Arial. LibreOffice does not follow the font: a fixed 14 pt height
    /// (.docx NumberingManager.cxx:917-926; .doc the default bullet font's height x 12, ww8par.cxx:5355-5362; RTF goes through the
    /// .docx path, rtfdocumentimpl.cxx:3853-3864) - Word is the reference here. The &lt;img&gt; goes through the picture endpoint.
    /// </summary>
    public static void AppendPicture(StringBuilder html, (string PictureId, double Width, double Height) picture)
    {
        html.Append("<img class=\"word-picture word-picture-bullet\" alt=\"\" ").Append(WordPictureHtml.PictureIdAttribute).Append("=\"")
            .Append(WordHtmlText.Encode(picture.PictureId)).Append("\" decoding=\"async\" style=\"height:max(calc(1em - 2pt),1px);width:calc(max(calc(1em - 2pt),1px) * ").Append(Number(picture.Width / picture.Height)).Append(");max-width:none;flex:none;aspect-ratio:")
            .Append(Number(picture.Width)).Append('/').Append(Number(picture.Height)).Append(";vertical-align:baseline\">");
    }

    private static string Number(double value) => Math.Round(value, 3).ToString(CultureInfo.InvariantCulture);

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
