using System.Globalization;
using System.Text;

namespace WordDocumentRandomAccess.Html;

public enum WordParagraphAlignment { Start, Center, End, Justify }

/// <summary>
/// The resolved paragraph formatting, format-neutral (filled from the .docx style cascade + pPr,
/// or from .doc PAPX sprms). Measures are twips (1/20 pt) as both formats store them.
/// </summary>
public sealed class WordParagraphFormatting
{
    public WordParagraphAlignment? Alignment;
    public bool IsRightToLeft;
    public int LeftIndentTwips;
    public int RightIndentTwips;
    public int FirstLineIndentTwips;        // negative = hanging
    public int SpaceBeforeTwips;
    public int SpaceAfterTwips;
    public int LineSpacingTwipsOrAuto;      // >0 with IsLineSpacingAuto → multiple of 240
    public bool IsLineSpacingAuto = true;
    public int OutlineLevel = -1;           // 0-8 → heading 1-9
    public int NumberingId = -1;
    public int NumberingLevel;
    public bool IsInTable;
    public WordRunFormatting RunFormatting = new();

    public WordParagraphFormatting Clone()
    {
        var clone = (WordParagraphFormatting)MemberwiseClone();
        clone.RunFormatting = RunFormatting.Clone();
        return clone;
    }

    public string ToCss()
    {
        var css = new StringBuilder();
        if (Alignment is { } alignment)
            css.Append("text-align:").Append(alignment switch
            {
                WordParagraphAlignment.Center => "center",
                WordParagraphAlignment.Justify => "justify",
                WordParagraphAlignment.End => IsRightToLeft ? "left" : "right",
                _ => IsRightToLeft ? "right" : "left",
            }).Append(';');
        // Word's "left"/"right" indent are start/end in a bidi paragraph; CSS logical properties match.
        if (LeftIndentTwips != 0) css.Append("margin-inline-start:").Append(TwipsToPt(LeftIndentTwips)).Append("pt;");
        if (RightIndentTwips != 0) css.Append("margin-inline-end:").Append(TwipsToPt(RightIndentTwips)).Append("pt;");
        if (FirstLineIndentTwips != 0) css.Append("text-indent:").Append(TwipsToPt(FirstLineIndentTwips)).Append("pt;");
        if (SpaceBeforeTwips != 0) css.Append("margin-top:").Append(TwipsToPt(SpaceBeforeTwips)).Append("pt;");
        if (SpaceAfterTwips != 0) css.Append("margin-bottom:").Append(TwipsToPt(SpaceAfterTwips)).Append("pt;");
        if (LineSpacingTwipsOrAuto > 0)
        {
            if (IsLineSpacingAuto) css.Append("line-height:").Append((LineSpacingTwipsOrAuto / 240.0).ToString("0.##", CultureInfo.InvariantCulture)).Append(';');
            else css.Append("line-height:").Append(TwipsToPt(LineSpacingTwipsOrAuto)).Append("pt;");
        }
        css.Append(RunFormatting.ToCss());
        return css.ToString();
    }

    public static string TwipsToPt(int twips) => (twips / 20.0).ToString("0.#", CultureInfo.InvariantCulture);
}
