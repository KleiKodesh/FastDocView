using System.Globalization;
using System.Text;

namespace FastDocView.Html;

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
    public bool IsSpaceBeforeAuto;          // w:spacing beforeAutospacing: Word's HTML-style automatic spacing (resolved by the renderer)
    public bool IsSpaceAfterAuto;           // w:spacing afterAutospacing
    public int LineSpacingTwipsOrAuto;      // >0 with IsLineSpacingAuto → multiple of 240
    public bool IsLineSpacingAuto = true;
    public int OutlineLevel = -1;           // 0-8 → heading 1-9
    public int NumberingId = -1;
    public int NumberingLevel;
    public bool IsInTable;
    public string? ShadingHex;              // paragraph background "RRGGBB"
    public bool IsContextualSpacing;        // no space between paragraphs of the same style
    public int FirstTabStopTwips;           // 0 = default stops; drives CSS tab-size
    /// <summary>The custom tab stops after every style layer (w:tabs; "clear" removes a stop inherited at the same position), by position.</summary>
    public WordTabStop[] TabStops = [];
    public string? StyleKey;                // what "same style" means for contextual spacing
    public string? BorderTopCss;            // CSS border values ("1pt solid #000"), null = none
    public string? BorderBottomCss;
    public string? BorderBetweenCss;        // w:pBdr/w:between: drawn between paragraphs of one border group
    public bool IsBorderJoinedAbove;        // same border group as the paragraph before: no top border (the between one instead), its spacing inside the box
    public bool IsBorderJoinedBelow;        // same border group as the paragraph after: no bottom border, its spacing inside the box
    public string? BorderStartCss;
    public string? BorderEndCss;
    public WordRunFormatting RunFormatting = new();

    public WordParagraphFormatting Clone()
    {
        var clone = (WordParagraphFormatting)MemberwiseClone();
        clone.RunFormatting = RunFormatting.Clone();
        return clone;
    }

    /// <param name="previousStyleKey">The style of the paragraph before; with contextual spacing on
    /// and the same style, the space between the two collapses as Word does.</param>
    public string ToCss(string? previousStyleKey = null)
    {
        var css = new StringBuilder();
        bool isSpacingSuppressed = IsContextualSpacing && StyleKey is not null && StyleKey == previousStyleKey;
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
        // Contextual spacing: the previous same-style paragraph's space-after is cancelled by an
        // equal negative top margin (sums to zero whether or not the margins collapse).
        // Inside a border group the spacing between the paragraphs is inside the box (padding), so the side borders run on.
        if (isSpacingSuppressed) { if (SpaceAfterTwips != 0) css.Append("margin-top:-").Append(TwipsToPt(SpaceAfterTwips)).Append("pt;"); }
        else if (SpaceBeforeTwips != 0 && !IsBorderJoinedAbove) css.Append("margin-top:").Append(TwipsToPt(SpaceBeforeTwips)).Append("pt;");
        if (SpaceAfterTwips != 0 && !IsBorderJoinedBelow) css.Append("margin-bottom:").Append(TwipsToPt(SpaceAfterTwips)).Append("pt;");
        if (ShadingHex is not null) css.Append("background-color:#").Append(ShadingHex).Append(';');
        if (FirstTabStopTwips > 0) css.Append("tab-size:").Append(TwipsToPt(FirstTabStopTwips)).Append("pt;");
        bool hasBorder = false;
        string? top = IsBorderJoinedAbove ? BorderBetweenCss : BorderTopCss;
        string? bottom = IsBorderJoinedBelow ? null : BorderBottomCss;
        foreach ((string property, string? value) in new[] { ("border-top", top), ("border-bottom", bottom),
                                                              ("border-inline-start", BorderStartCss), ("border-inline-end", BorderEndCss) })
        {
            if (value is null or "none") continue;
            css.Append(property).Append(':').Append(value).Append(';');
            hasBorder = true;
        }
        if (hasBorder || IsBorderJoinedAbove || IsBorderJoinedBelow)
        {
            string paddingTop = IsBorderJoinedAbove ? TwipsToPt(SpaceBeforeTwips) + "pt" : "1pt";
            string paddingBottom = IsBorderJoinedBelow ? TwipsToPt(SpaceAfterTwips) + "pt" : "1pt";
            css.Append("padding:").Append(paddingTop).Append(" 4pt ").Append(paddingBottom).Append(';');
        }
        if (LineSpacingTwipsOrAuto > 0)
        {
            if (IsLineSpacingAuto) css.Append("line-height:").Append((LineSpacingTwipsOrAuto / 240.0).ToString("0.##", CultureInfo.InvariantCulture)).Append(';');
            else css.Append("line-height:").Append(TwipsToPt(LineSpacingTwipsOrAuto)).Append("pt;");
        }
        css.Append(RunFormatting.ToCss(null, WordDecorationOutput.None));   // lines go on the runs: a child cannot remove a parent's line
        return css.ToString();
    }

    /// <summary>
    /// What makes consecutive paragraphs one border group (ECMA-376 §17.3.1.24 pBdr: "the same set of paragraph border
    /// settings"; Open-Xml-PowerTools' CreateBorderDivs also requires the same indents): every border and the indents.
    /// Null when the paragraph has no border at all.
    /// </summary>
    public string? BorderGroupKey
        => BorderTopCss is null or "none" && BorderBottomCss is null or "none" && BorderStartCss is null or "none" && BorderEndCss is null or "none"
           && BorderBetweenCss is null or "none"
            ? null
            : string.Join("|", BorderTopCss, BorderBottomCss, BorderStartCss, BorderEndCss, BorderBetweenCss, LeftIndentTwips, RightIndentTwips);

    public static string TwipsToPt(int twips) => (twips / 20.0).ToString("0.#", CultureInfo.InvariantCulture);
}

/// <summary>A custom tab stop (ECMA-376 §17.3.1.37): position in twips, alignment (ST_TabJc: left, center, right, decimal, bar, num, start, end) and leader (ST_TabTlc).</summary>
public readonly record struct WordTabStop(int PositionTwips, string Alignment, string Leader);
