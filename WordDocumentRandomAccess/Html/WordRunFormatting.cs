using System.Globalization;
using System.Text;

namespace WordDocumentRandomAccess.Html;

/// <summary>
/// The resolved character formatting of a run, format-neutral: .docx fills it from the style
/// cascade + rPr, .doc from CHPX sprms. <see cref="ToCss"/> emits only what differs from the
/// paragraph's base run formatting, so most runs carry no inline style at all.
/// </summary>
public sealed class WordRunFormatting
{
    public bool IsBold;
    public bool IsItalic;
    public bool IsUnderlined;
    public bool IsUnderlineWordsOnly;       // w:u words: no line under spaces (ST_Underline)
    public string? UnderlineStyleCss;       // CSS text-decoration-style for the underline (double, dotted, dashed, wavy); null = solid
    public bool IsUnderlineThick;           // thick / *Heavy underline patterns
    public string? UnderlineColorHex;       // "RRGGBB" from w:u/@color (+ theme); null = the text color
    public string? EmphasisMarkCss;         // CSS text-emphasis value ("filled dot", "filled sesame", "open circle"); null = none
    public bool IsEmphasisMarkBelow;        // w:em underDot
    public bool IsStrikethrough;
    public bool IsDoubleStrikethrough;      // w:dstrike / sprmCFDStrike
    public bool IsSuperscript;
    public bool IsSubscript;
    public bool IsHidden;
    public bool IsSmallCaps;
    public bool IsAllCaps;
    public int CharacterSpacingTwips;
    public int PositionHalfPoints;          // raised (+) / lowered (−) text, w:position / sprmCHpsPos
    public bool IsOutline;
    public bool IsShadow;
    public bool IsEmboss;
    public bool IsImprint;
    public string? BorderCss;               // a border around the run (w:bdr / sprmCBrc)
    public bool IsRightToLeft;
    public int FontSizeHalfPoints;          // 0 = inherited/default
    public int ComplexScriptFontSizeHalfPoints;
    public string? ColorHex;                // "RRGGBB", null = auto
    public string? HighlightCssColor;
    public string? FontName;
    public string? ComplexScriptFontName;
    public bool IsComplexScript;            // .docx w:cs: complex-script properties without right-to-left direction
    public string? EastAsianFontName;       // .docx rFonts eastAsia: the face Word draws CJK characters with
    public bool IsWebHidden;                // w:webHidden: hidden in a web view (not a toggle property)
    /// <summary>.docx only: the toggle properties per style layer; <c>DocxStyleCascade</c> resolves them into the flags above.</summary>
    public WordToggleLayers Toggles;

    public WordRunFormatting Clone() => (WordRunFormatting)MemberwiseClone();

    /// <summary>Shadow / emboss / imprint as CSS text-shadow (Word draws them as offset copies of the glyphs).</summary>
    private string TextShadowCss()
        => IsShadow ? "1px 1px 1px rgba(0,0,0,.45)"
         : IsEmboss ? "-1px -1px 0 rgba(255,255,255,.8), 1px 1px 0 rgba(0,0,0,.35)"
         : IsImprint ? "1px 1px 0 rgba(255,255,255,.8), -1px -1px 0 rgba(0,0,0,.35)"
         : "";

    /// <summary>
    /// The underline as one CSS declaration set: line, pattern (text-decoration-style), thick / *Heavy (thickness) and
    /// color; empty when not underlined. Always absolute, never relative to a parent: CSS cannot take a parent's
    /// decoration away from a child, so lines are only ever drawn on the run itself (see <see cref="WordDecorationOutput"/>).
    /// </summary>
    public string UnderlineCss()
    {
        if (!IsUnderlined) return "";
        var css = new StringBuilder("text-decoration:underline;");
        if (UnderlineStyleCss is { } style) css.Append("text-decoration-style:").Append(style).Append(';');
        if (IsUnderlineThick) css.Append("text-decoration-thickness:0.15em;");
        if (UnderlineColorHex is not null) css.Append("text-decoration-color:#").Append(UnderlineColorHex).Append(';');
        return css.ToString();
    }

    /// <summary>Single or double strikethrough (w:strike / w:dstrike) as CSS; empty when none.</summary>
    public string StrikeCss()
        => IsDoubleStrikethrough ? "text-decoration:line-through;text-decoration-style:double;"
         : IsStrikethrough ? "text-decoration:line-through;" : "";

    private void AppendDecorations(StringBuilder css, WordDecorationOutput output)
    {
        if (output == WordDecorationOutput.None) return;
        string strike = StrikeCss();
        if (output == WordDecorationOutput.StrikeOnly || !IsUnderlined) { css.Append(strike); return; }
        if (strike.Length == 0) { css.Append(UnderlineCss()); return; }
        // Both lines on one element share one style / thickness / color: the underline's (a renderer that can nest
        // elements asks for StrikeOnly and draws the underline on an inner element instead).
        css.Append(UnderlineCss().Replace("text-decoration:underline;", "text-decoration:underline line-through;"));
    }

    /// <summary>w:em as CSS text-emphasis (+ position for marks drawn below the glyphs); empty when none.</summary>
    private string EmphasisCss()
        => EmphasisMarkCss is null ? "" : "text-emphasis:" + EmphasisMarkCss + ";" + (IsEmphasisMarkBelow ? "text-emphasis-position:under right;" : "");

    /// <summary>Whether the complex-script set (bold / italic / size / font) applies to this run.</summary>
    public bool UsesComplexScriptProperties => IsRightToLeft || IsComplexScript;

    /// <summary>
    /// The font for East Asian characters when it differs from the run's own (null = no switch needed). Word 16 draws
    /// CJK characters of a run without w:rtl / w:cs in the rFonts eastAsia face and every other character — Hebrew and
    /// Arabic letters included — in the ascii face (PDF export check, 2026-10-28); w:hint did not change that.
    /// </summary>
    public string? EastAsianFontSwitch
        => !UsesComplexScriptProperties && !string.IsNullOrEmpty(EastAsianFontName) && EastAsianFontName != EffectiveFontName ? EastAsianFontName : null;

    /// <summary>
    /// CJK ideographs, kana, Hangul, CJK symbols / compatibility / full-width forms — the ranges Word gives the eastAsia font.
    /// The block list is Unicode's (Blocks.txt); which blocks Word treats as East Asian beyond the ideographs is not specified.
    /// </summary>
    public static bool IsEastAsianCharacter(char c)
        => (int)c is (>= 0x1100 and <= 0x11FF) or (>= 0x2E80 and <= 0x9FFF) or (>= 0xA960 and <= 0xA97F)
             or (>= 0xAC00 and <= 0xD7AF) or (>= 0xF900 and <= 0xFAFF) or (>= 0xFE30 and <= 0xFE4F) or (>= 0xFF00 and <= 0xFFEF);

    public int EffectiveFontSizeHalfPoints
        => UsesComplexScriptProperties && ComplexScriptFontSizeHalfPoints > 0 ? ComplexScriptFontSizeHalfPoints
         : FontSizeHalfPoints > 0 ? FontSizeHalfPoints : ComplexScriptFontSizeHalfPoints;

    public string? EffectiveFontName
        => UsesComplexScriptProperties && !string.IsNullOrEmpty(ComplexScriptFontName) ? ComplexScriptFontName : FontName ?? ComplexScriptFontName;

    /// <summary>CSS for the whole formatting (used for the block's base font).</summary>
    public string ToCss() => ToCss(null);

    /// <summary>CSS for what differs from <paramref name="baseline"/>; empty when nothing differs. Text decorations are not
    /// diffed: they are written whenever the run has them (<paramref name="decorations"/> says which).</summary>
    public string ToCss(WordRunFormatting? baseline, WordDecorationOutput decorations = WordDecorationOutput.Combined)
    {
        var css = new StringBuilder();
        if (baseline is null || IsBold != baseline.IsBold) css.Append(IsBold ? "font-weight:bold;" : baseline is null ? "" : "font-weight:normal;");
        if (baseline is null || IsItalic != baseline.IsItalic) css.Append(IsItalic ? "font-style:italic;" : baseline is null ? "" : "font-style:normal;");
        AppendDecorations(css, decorations);
        string emphasis = EmphasisCss();
        if (emphasis.Length > 0 && (baseline is null || emphasis != baseline.EmphasisCss())) css.Append(emphasis);
        else if (emphasis.Length == 0 && baseline is not null && baseline.EmphasisCss().Length > 0) css.Append("text-emphasis:none;");
        if (IsSuperscript && (baseline is null || !baseline.IsSuperscript)) css.Append("vertical-align:super;font-size:smaller;");
        else if (IsSubscript && (baseline is null || !baseline.IsSubscript)) css.Append("vertical-align:sub;font-size:smaller;");
        else if (baseline is not null && (baseline.IsSuperscript || baseline.IsSubscript) && !IsSuperscript && !IsSubscript) css.Append("vertical-align:baseline;");
        if (IsHidden) css.Append("display:none;");
        if (IsSmallCaps && (baseline is null || !baseline.IsSmallCaps)) css.Append("font-variant:small-caps;");
        if (IsAllCaps && (baseline is null || !baseline.IsAllCaps)) css.Append("text-transform:uppercase;");
        if (PositionHalfPoints != 0 && (baseline is null || PositionHalfPoints != baseline.PositionHalfPoints))
            css.Append("vertical-align:").Append((PositionHalfPoints / 2.0).ToString("0.#", CultureInfo.InvariantCulture)).Append("pt;");
        string effects = TextShadowCss();
        if (effects.Length > 0 && (baseline is null || effects != baseline.TextShadowCss())) css.Append("text-shadow:").Append(effects).Append(';');
        if (IsOutline && (baseline is null || !baseline.IsOutline)) css.Append("-webkit-text-stroke:0.6px currentColor;color:transparent;");
        if (BorderCss is not null && BorderCss != "none" && (baseline is null || BorderCss != baseline.BorderCss)) css.Append("border:").Append(BorderCss).Append(';');
        if (CharacterSpacingTwips != 0 && (baseline is null || CharacterSpacingTwips != baseline.CharacterSpacingTwips))
            css.Append("letter-spacing:").Append((CharacterSpacingTwips / 20.0).ToString("0.##", CultureInfo.InvariantCulture)).Append("pt;");

        int size = EffectiveFontSizeHalfPoints;
        if (size > 0 && (baseline is null || size != baseline.EffectiveFontSizeHalfPoints) && !IsSuperscript && !IsSubscript)
            css.Append("font-size:").Append((size / 2.0).ToString("0.#", CultureInfo.InvariantCulture)).Append("pt;");

        if (ColorHex is not null && (baseline is null || ColorHex != baseline.ColorHex)) css.Append("color:#").Append(ColorHex).Append(';');
        else if (ColorHex is null && baseline?.ColorHex is not null) css.Append("color:inherit;");
        if (HighlightCssColor is not null && (baseline is null || HighlightCssColor != baseline.HighlightCssColor)) css.Append("background-color:").Append(HighlightCssColor).Append(';');

        string? font = EffectiveFontName;
        if (font is not null && (baseline is null || font != baseline.EffectiveFontName)) css.Append("font-family:'").Append(WordHtmlText.CssFontName(font)).Append("';");
        return css.ToString();
    }
}

/// <summary>Which text decorations <see cref="WordRunFormatting.ToCss(WordRunFormatting?, WordDecorationOutput)"/> writes.</summary>
public enum WordDecorationOutput
{
    /// <summary>Underline and strikethrough on the same element (when both are on they share the underline's style).</summary>
    Combined,
    /// <summary>Strikethrough only: the renderer draws the underline on inner elements (word-only underline, untouched trailing spaces).</summary>
    StrikeOnly,
    /// <summary>None: a paragraph element, whose lines its runs could never switch off.</summary>
    None,
}
