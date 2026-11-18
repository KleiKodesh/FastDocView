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
    /// <summary>
    /// .doc only: the alignment came from sprmPJc80, whose values are physical ([MS-DOC] 2.6.2: 0 = physically left,
    /// 2 = physically right) — Start/End then mean left/right whatever the direction. sprmPJc is logical.
    /// </summary>
    public bool IsAlignmentPhysical;
    /// <summary>.doc only: the indent came from sprmPDxaLeft80 / sprmPDxaRight80 — physical left / right ([MS-DOC] 2.6.2), not start / end.</summary>
    public bool IsLeftIndentPhysical, IsRightIndentPhysical;
    public int FirstLineIndentTwips;        // negative = hanging
    public int SpaceBeforeTwips;
    public int SpaceAfterTwips;
    public bool IsSpaceBeforeAuto;          // w:spacing beforeAutospacing: Word's HTML-style automatic spacing (resolved by the renderer)
    public bool IsSpaceAfterAuto;           // w:spacing afterAutospacing
    public int LineSpacingTwipsOrAuto;      // >0 with IsLineSpacingAuto → multiple of 240
    public bool IsLineSpacingAuto = true;
    /// <summary>Not auto: the value is the exact line height (.docx lineRule exact, .doc / RTF negative dyaLine / \sl); otherwise it is a minimum (atLeast).</summary>
    public bool IsLineSpacingExact;

    /// <summary>
    /// One "line" of Word's proportional spacing as a multiple of the font size: Word counts auto spacing in 240ths of the
    /// font's single line height (ascent + descent), not of the em. 1.15 is that height for Times New Roman and Arial (the
    /// document stylesheet's base line-height uses the same figure); measured against LibreOffice's PDF: a 12 pt line with
    /// line=240 is 13.8 pt apart, with 276 (1.15) 15.9 pt.
    /// </summary>
    public const double SingleLineHeightEm = 1.15;
    /// <summary>Word 6/95 .doc only: sprmPNLvlAnm (1-9 outline, 10 numbering, 11 bullets, 12 pause; 0 none) and the sprmPAnld ANLD bytes - see DocBinaryWord6Numbering.</summary>
    public int Word6NumberingLevel;
    public byte[]? Word6Numbering;
    public int OutlineLevel = -1;           // 0-8 → heading 1-9
    public int NumberingId = -1;
    public int NumberingLevel;
    public bool IsInTable;
    public string? ShadingHex;              // paragraph background "RRGGBB"
    public bool IsContextualSpacing;        // no space between paragraphs of the same style
    public int FirstTabStopTwips;           // 0 = default stops; drives CSS tab-size
    /// <summary>The custom tab stops after every style layer (w:tabs; "clear" removes a stop inherited at the same position), by position.</summary>
    public WordTabStop[] TabStops = [];
    /// <summary>The document's default tab interval, twips (.docx w:defaultTabStop, .doc Dop dxaTab, RTF \deftab, ODT style:tab-stop-distance); 0 = 720.</summary>
    public int DefaultTabStopTwips;
    /// <summary>
    /// Tab positions count from the start indent, not from the text area's edge: OpenDocument's default (LibreOffice
    /// DocumentSettingId::TABS_RELATIVE_TO_INDENT, txttab.cxx NewTabPortion nTabLeft). Word measures them from the margin.
    /// </summary>
    public bool AreTabStopsRelativeToIndent;
    public string? StyleKey;                // what "same style" means for contextual spacing
    public string? BorderTopCss;            // CSS border values ("1pt solid #000"), null = none
    public string? BorderBottomCss;
    public string? BorderBetweenCss;        // w:pBdr/w:between: drawn between paragraphs of one border group
    public bool IsBorderJoinedAbove;        // same border group as the paragraph before: no top border (the between one instead), its spacing inside the box
    public bool IsBorderJoinedBelow;        // same border group as the paragraph after: no bottom border, its spacing inside the box
    public string? BorderStartCss;
    public string? BorderEndCss;
    /// <summary>
    /// The distance between the text and each border, twips (.docx w:space, RTF \brsp); null = the default padding (1 pt above and
    /// below, 4 pt at the sides). Word measures it from the text to the line (ECMA-376 17.3.4 space: "the spacing offset ... between
    /// the text and this border").
    /// </summary>
    public int? BorderTopSpaceTwips, BorderBottomSpaceTwips, BorderStartSpaceTwips, BorderEndSpaceTwips;
    /// <summary>The border box's shadow (.docx w:shadow, RTF \brdrsh): Word draws it right and below the box, as wide as the line; null = none.</summary>
    public string? BorderShadowColorCss;
    public double BorderShadowWidthPt;
    public WordRunFormatting RunFormatting = new();

    /// <summary>
    /// A drop cap: the number of lines the initial drops (0 = none). Word keeps the initial in a framed paragraph of its own
    /// (.docx w:framePr dropCap / lines / hSpace, .doc sprmPDcs + sprmPDxaFromText, RTF \dropcapli / \dropcapt); the
    /// renderers draw that paragraph at the start of the next one (<see cref="WordDropCapHtml"/>). ODT style:drop-cap is
    /// on the text paragraph itself (its first <see cref="DropCapCharacterCount"/> characters).
    /// </summary>
    public int DropCapLines;
    /// <summary>The initial sits in the margin, not in the text (.docx dropCap="margin", .doc fdct 2, RTF \dropcapt2).</summary>
    public bool IsDropCapInMargin;
    /// <summary>The gap between the initial and the text, twips (.docx hSpace, .doc sprmPDxaFromText, ODT style:distance).</summary>
    public int DropCapDistanceTwips;
    /// <summary>ODT only: how many characters of the paragraph the drop cap takes (style:length; 0 with lines = 1, "word" = -1).</summary>
    public int DropCapCharacterCount;

    /// <summary>Automatic hyphenation is off for this paragraph (.docx suppressAutoHyphens, .doc sprmPFNoAutoHyph, RTF \hyphpar0, ODT fo:hyphenate false).</summary>
    public bool IsAutoHyphenationSuppressed;
    /// <summary>ODT only: fo:hyphenate true on the paragraph's style (ODT has no document switch).</summary>
    public bool IsAutoHyphenationRequested;
    /// <summary>Resolved by the renderer: the document hyphenates automatically and this paragraph does not opt out → CSS hyphens:auto (with the paragraph's lang, which Chromium needs).</summary>
    public bool IsHyphenated;

    /// <summary>The indent on the paragraph's start side, twips (a physical left indent is the end side of a bidi paragraph, see <see cref="ToCss"/>).</summary>
    public int StartIndentTwips
        => (IsLeftIndentPhysical && IsRightToLeft ? 0 : LeftIndentTwips) + (IsRightIndentPhysical && IsRightToLeft ? RightIndentTwips : 0);

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
                WordParagraphAlignment.End => IsRightToLeft && !IsAlignmentPhysical ? "left" : "right",
                _ => IsRightToLeft && !IsAlignmentPhysical ? "right" : "left",
            }).Append(';');
        // Word's "left"/"right" indent are start/end in a bidi paragraph; CSS logical properties match.
        // A physical left indent is the end side of a bidi paragraph (LibreOffice ww8par6.cxx Read_LR swaps them the same way).
        if (LeftIndentTwips != 0) css.Append(IsLeftIndentPhysical && IsRightToLeft ? "margin-inline-end:" : "margin-inline-start:").Append(TwipsToPt(LeftIndentTwips)).Append("pt;");
        if (RightIndentTwips != 0) css.Append(IsRightIndentPhysical && IsRightToLeft ? "margin-inline-start:" : "margin-inline-end:").Append(TwipsToPt(RightIndentTwips)).Append("pt;");
        if (FirstLineIndentTwips != 0) css.Append("text-indent:").Append(TwipsToPt(FirstLineIndentTwips)).Append("pt;");
        // Contextual spacing: the previous same-style paragraph's space-after is cancelled by an
        // equal negative top margin (sums to zero whether or not the margins collapse).
        // Inside a border group the spacing between the paragraphs is inside the box (padding), so the side borders run on.
        if (isSpacingSuppressed) { if (SpaceAfterTwips != 0) css.Append("margin-top:-").Append(TwipsToPt(SpaceAfterTwips)).Append("pt;"); }
        else if (SpaceBeforeTwips != 0 && !IsBorderJoinedAbove) css.Append("margin-top:").Append(TwipsToPt(SpaceBeforeTwips)).Append("pt;");
        if (SpaceAfterTwips != 0 && !IsBorderJoinedBelow) css.Append("margin-bottom:").Append(TwipsToPt(SpaceAfterTwips)).Append("pt;");
        if (ShadingHex is not null) css.Append("background-color:").Append(WordHtmlColor.Css(ShadingHex)).Append(';');
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
            string paddingTop = IsBorderJoinedAbove ? TwipsToPt(SpaceBeforeTwips) + "pt" : Space(BorderTopSpaceTwips, "1pt");
            string paddingBottom = IsBorderJoinedBelow ? TwipsToPt(SpaceAfterTwips) + "pt" : Space(BorderBottomSpaceTwips, "1pt");
            string paddingStart = Space(BorderStartSpaceTwips, "4pt"), paddingEnd = Space(BorderEndSpaceTwips, "4pt");
            if (paddingStart == paddingEnd) css.Append("padding:").Append(paddingTop).Append(' ').Append(paddingStart).Append(' ').Append(paddingBottom).Append(';');
            else css.Append("padding-top:").Append(paddingTop).Append(";padding-bottom:").Append(paddingBottom)
                    .Append(";padding-inline-start:").Append(paddingStart).Append(";padding-inline-end:").Append(paddingEnd).Append(';');
            // the shadow runs along the right and the bottom of the box; inside a border group only the group's last paragraph has the bottom one
            if (BorderShadowColorCss is { } shadow && BorderShadowWidthPt > 0)
            {
                string width = BorderShadowWidthPt.ToString("0.##", CultureInfo.InvariantCulture) + "pt";
                css.Append("box-shadow:").Append(width).Append(' ').Append(IsBorderJoinedBelow ? "0" : width).Append(" 0 0 ").Append(shadow).Append(';');
            }
        }
        if (LineSpacingTwipsOrAuto > 0)
        {
            if (IsLineSpacingAuto) css.Append("line-height:").Append((LineSpacingTwipsOrAuto / 240.0 * SingleLineHeightEm).ToString("0.###", CultureInfo.InvariantCulture)).Append(';');
            else if (IsLineSpacingExact) css.Append("line-height:").Append(TwipsToPt(LineSpacingTwipsOrAuto)).Append("pt;");
            // at least: the single line height unless the stated minimum is larger
            else css.Append("line-height:max(").Append(SingleLineHeightEm.ToString("0.##", CultureInfo.InvariantCulture)).Append("em,").Append(TwipsToPt(LineSpacingTwipsOrAuto)).Append("pt);");
        }
        if (IsHyphenated) css.Append("hyphens:auto;");
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
            : string.Join("|", BorderTopCss, BorderBottomCss, BorderStartCss, BorderEndCss, BorderBetweenCss, LeftIndentTwips, RightIndentTwips,
                          BorderTopSpaceTwips, BorderBottomSpaceTwips, BorderStartSpaceTwips, BorderEndSpaceTwips, BorderShadowColorCss, BorderShadowWidthPt);

    private static string Space(int? twips, string fallback) => twips is { } value ? TwipsToPt(value) + "pt" : fallback;

    /// <summary>The lang attribute for a hyphenated paragraph (Chromium hyphenates only text whose language it knows); null otherwise.</summary>
    public string? HyphenationLanguageAttribute
        => IsHyphenated && RunFormatting.Language is { Length: > 0 } language && language.All(c => char.IsLetterOrDigit(c) || c == '-')
            ? "lang=\"" + language + "\"" : null;

    /// <summary>
    /// What the client-side tab layout needs (the page's tab pass, fastDocViewTabs.ts): <c>data-tabs="D S F R;pos a l;..."</c> with
    /// D the default interval, S the start indent, F the first-line indent (pt), R 1 when stops count from the indent, then the
    /// stops by position (pt): a = l(eft) c(enter) r(ight) d(ecimal) b(ar), l = the leader: "" none, d dot, h hyphen, u underscore,
    /// H heavy, m middle dot. Null when the browser's own tabs already land where Word puts them (no stops, no indent, the
    /// default 36 pt interval: CSS tab-size 36pt counts from the same edge).
    /// </summary>
    public string? TabStopsAttribute
    {
        get
        {
            int interval = DefaultTabStopTwips > 0 ? DefaultTabStopTwips : 720;
            int start = StartIndentTwips;
            if (TabStops.Length == 0 && interval == 720 && start == 0 && FirstLineIndentTwips >= 0) return null;
            var value = new StringBuilder("data-tabs=\"");
            value.Append(TwipsToExactPt(interval)).Append(' ').Append(TwipsToExactPt(start)).Append(' ').Append(TwipsToExactPt(FirstLineIndentTwips))
                 .Append(' ').Append(AreTabStopsRelativeToIndent ? '1' : '0');
            foreach (WordTabStop stop in TabStops)
            {
                char alignment = stop.Alignment switch { "center" => 'c', "right" or "end" => 'r', "decimal" => 'd', "bar" => 'b', "clear" => '\0', _ => 'l' };
                if (alignment == '\0') continue;
                value.Append(';').Append(TwipsToExactPt(stop.PositionTwips)).Append(' ').Append(alignment).Append(' ').Append(stop.Leader switch
                {
                    "dot" => "d", "hyphen" => "h", "underscore" => "u", "heavy" => "H", "middleDot" => "m", _ => "",
                });
            }
            return value.Append('"').ToString();
        }
    }

    /// <summary>The paragraph tag's extra attributes: the hyphenation language and the tab layout data.</summary>
    public string? ParagraphAttributes
        => (HyphenationLanguageAttribute, TabStopsAttribute) switch
        {
            (null, var tabs) => tabs,
            (var language, null) => language,
            (var language, var tabs) => language + " " + tabs,
        };

    private static string TwipsToExactPt(int twips) => (twips / 20.0).ToString("0.##", CultureInfo.InvariantCulture);

    public static string TwipsToPt(int twips) => (twips / 20.0).ToString("0.#", CultureInfo.InvariantCulture);
}

/// <summary>A custom tab stop (ECMA-376 §17.3.1.37): position in twips, alignment (ST_TabJc: left, center, right, decimal, bar, num, start, end) and leader (ST_TabTlc).</summary>
public readonly record struct WordTabStop(int PositionTwips, string Alignment, string Leader);
