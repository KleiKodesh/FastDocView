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
    public bool IsStrikethrough;
    public bool IsSuperscript;
    public bool IsSubscript;
    public bool IsHidden;
    public bool IsRightToLeft;
    public int FontSizeHalfPoints;          // 0 = inherited/default
    public int ComplexScriptFontSizeHalfPoints;
    public string? ColorHex;                // "RRGGBB", null = auto
    public string? HighlightCssColor;
    public string? FontName;
    public string? ComplexScriptFontName;

    public WordRunFormatting Clone() => (WordRunFormatting)MemberwiseClone();

    public int EffectiveFontSizeHalfPoints
        => IsRightToLeft && ComplexScriptFontSizeHalfPoints > 0 ? ComplexScriptFontSizeHalfPoints
         : FontSizeHalfPoints > 0 ? FontSizeHalfPoints : ComplexScriptFontSizeHalfPoints;

    public string? EffectiveFontName
        => IsRightToLeft && !string.IsNullOrEmpty(ComplexScriptFontName) ? ComplexScriptFontName : FontName ?? ComplexScriptFontName;

    /// <summary>CSS for the whole formatting (used for the block's base font).</summary>
    public string ToCss() => ToCss(null);

    /// <summary>CSS for what differs from <paramref name="baseline"/>; empty when nothing differs.</summary>
    public string ToCss(WordRunFormatting? baseline)
    {
        var css = new StringBuilder();
        if (baseline is null || IsBold != baseline.IsBold) css.Append(IsBold ? "font-weight:bold;" : baseline is null ? "" : "font-weight:normal;");
        if (baseline is null || IsItalic != baseline.IsItalic) css.Append(IsItalic ? "font-style:italic;" : baseline is null ? "" : "font-style:normal;");
        string decoration = (IsUnderlined ? "underline " : "") + (IsStrikethrough ? "line-through" : "");
        string baselineDecoration = baseline is null ? "" : (baseline.IsUnderlined ? "underline " : "") + (baseline.IsStrikethrough ? "line-through" : "");
        if (decoration != baselineDecoration) css.Append("text-decoration:").Append(decoration.Length == 0 ? "none" : decoration.Trim()).Append(';');
        if (IsSuperscript && (baseline is null || !baseline.IsSuperscript)) css.Append("vertical-align:super;font-size:smaller;");
        else if (IsSubscript && (baseline is null || !baseline.IsSubscript)) css.Append("vertical-align:sub;font-size:smaller;");
        else if (baseline is not null && (baseline.IsSuperscript || baseline.IsSubscript) && !IsSuperscript && !IsSubscript) css.Append("vertical-align:baseline;");
        if (IsHidden) css.Append("display:none;");

        int size = EffectiveFontSizeHalfPoints;
        if (size > 0 && (baseline is null || size != baseline.EffectiveFontSizeHalfPoints) && !IsSuperscript && !IsSubscript)
            css.Append("font-size:").Append((size / 2.0).ToString("0.#", CultureInfo.InvariantCulture)).Append("pt;");

        if (ColorHex is not null && (baseline is null || ColorHex != baseline.ColorHex)) css.Append("color:#").Append(ColorHex).Append(';');
        else if (ColorHex is null && baseline?.ColorHex is not null) css.Append("color:inherit;");
        if (HighlightCssColor is not null && (baseline is null || HighlightCssColor != baseline.HighlightCssColor)) css.Append("background-color:").Append(HighlightCssColor).Append(';');

        string? font = EffectiveFontName;
        if (font is not null && (baseline is null || font != baseline.EffectiveFontName)) css.Append("font-family:'").Append(font.Replace("'", "")).Append("';");
        return css.ToString();
    }
}
