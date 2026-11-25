using System.Text;
using FastDocView.DocxHtml;

namespace FastDocView.Html;

/// <summary>
/// The East Asian font slot of body text, one path for .docx, .doc and .odt: which characters draw in the run's East Asian face
/// (<see cref="WordRunFormatting.EastAsianFontName"/>) and the HTML that switches to it. Word 16's PDF export
/// (test-data/word-east-asian/east-asian-font-v1.pdf, 2026-11-24):
/// <list type="bullet">
/// <item>Han, kana, Hangul, CJK punctuation and full-width forms draw in the East Asian face in every run, right-to-left (w:rtl)
/// and complex-script (w:cs) runs included; the other characters take the run's Latin or complex-script face. LibreOffice does
/// the same (i18nutil scriptchangescanner.cxx, tdf#163660: Asian characters inside right-to-left runs keep the Asian font).</item>
/// <item>w:hint="eastAsia" (.doc sprmCIdctHint 1) moves the characters both scripts share to the East Asian face, only in runs
/// without w:rtl / w:cs: Latin-1 punctuation and symbols (not its letters, not U+00AB), spacing modifiers, Greek, Cyrillic and
/// U+2000-U+2E7F (punctuation, letterlike, arrows, math, box drawing, shapes, dingbats). ASCII, Latin letters (U+00C0-U+024F)
/// and U+00AB stay Latin. LibreOffice moves its WEAK characters the same way (ScriptHintType::Asian, scriptchangescanner.cxx).</item>
/// </list>
/// The character classification is the SmartArt measuring code's (<see cref="DocxDrawingTextFonts.EastAsianScriptOf"/>).
/// Only the HTML changes: the text itself (GetText, search) is the same.
/// </summary>
public static class WordEastAsianText
{
    /// <summary>The face the run's East Asian characters switch to; null when the run needs no switch.</summary>
    public static string? SwitchFont(WordRunFormatting formatting)
        => !string.IsNullOrEmpty(formatting.EastAsianFontName) && formatting.EastAsianFontName != formatting.EffectiveFontName ? formatting.EastAsianFontName : null;

    /// <summary>
    /// The CSS the run's East Asian characters switch to: the East Asian face and, for .odt, the Asian size / bold / italic
    /// (style:font-size-asian / -weight-asian / -style-asian = CharHeightAsian / CharWeightAsian / CharPostureAsian, xmloff
    /// txtprmap.cxx:212-223; LibreOffice draws each script portion in its own set:
    /// sw/source/core/text/porlay.cxx:1363 SwScriptInfo::InitScriptInfo splits by script, SwFont::SetActual picks the
    /// SwFontScript::CJK font, sw/source/core/inc/swfont.hxx). In a right-to-left / complex-script run they keep the Latin
    /// size, bold and italic (w:sz / w:b / w:i, not w:szCs / w:bCs / w:iCs): Word 16 draws them so
    /// (test-data/rtf-east-asian/rtf-east-asian-runs-v1.pdf, one run model for .docx, .doc and .rtf), as LibreOffice's CJK portion
    /// uses the Asian set. The size is relative (em), so a superscript keeps its ratio. Null when nothing differs from the run.
    /// </summary>
    public static string? SwitchCss(WordRunFormatting formatting)
    {
        string? font = SwitchFont(formatting);
        int size = formatting.EastAsianFontSizeHalfPoints > 0 ? formatting.EastAsianFontSizeHalfPoints
                 : formatting.UsesComplexScriptProperties ? formatting.FontSizeHalfPoints : 0;
        int runSize = formatting.EffectiveFontSizeHalfPoints;
        bool isBold = formatting.HasEastAsianBoldItalic ? formatting.IsBoldEastAsian : formatting.IsBold;
        bool isItalic = formatting.HasEastAsianBoldItalic ? formatting.IsItalicEastAsian : formatting.IsItalic;
        bool isSizeSwitch = size > 0 && runSize > 0 && size != runSize;
        bool isBoldSwitch = isBold != formatting.EffectiveIsBold;
        bool isItalicSwitch = isItalic != formatting.EffectiveIsItalic;
        if (font is null && !isSizeSwitch && !isBoldSwitch && !isItalicSwitch) return null;
        var css = new StringBuilder();
        if (font is not null) css.Append("font-family:'").Append(WordHtmlText.CssFontName(font)).Append("';");
        if (isSizeSwitch) css.Append("font-size:").Append(((double)size / runSize).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)).Append("em;");
        if (isBoldSwitch) css.Append(isBold ? "font-weight:bold;" : "font-weight:normal;");
        if (isItalicSwitch) css.Append(isItalic ? "font-style:italic;" : "font-style:normal;");
        return css.ToString().TrimEnd(';');
    }

    /// <summary>Whether the code point draws in the East Asian face in this run.</summary>
    public static bool IsEastAsian(int codePoint, WordRunFormatting formatting)
        => DocxDrawingTextFonts.EastAsianScriptOf(codePoint) is not null
           || (formatting.IsEastAsianHint && !formatting.UsesComplexScriptProperties && IsSharedWithEastAsian(codePoint));

    /// <summary>The characters w:hint="eastAsia" moves (measured in Word 16, see the class summary).</summary>
    public static bool IsSharedWithEastAsian(int c) => c switch
    {
        0xA1 or 0xA4 or 0xA7 or 0xA8 or 0xAA or 0xAD or 0xAF or (>= 0xB0 and <= 0xB4) or (>= 0xB6 and <= 0xBA) or (>= 0xBC and <= 0xBF) or 0xD7 or 0xF7 => true,
        >= 0x02B0 and <= 0x04FF => true,
        >= 0x2000 and <= 0x2E7F => true,
        _ => false,
    };

    /// <summary>The length of the code point at <paramref name="index"/> (2 for a surrogate pair).</summary>
    public static int CodePointAt(string text, int index, out int length)
    {
        if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) { length = 2; return char.ConvertToUtf32(text[index], text[index + 1]); }
        length = 1;
        return text[index];
    }

    /// <summary>The end of the stretch from <paramref name="start"/> whose characters all are (or all are not) East Asian.</summary>
    public static int StretchEnd(string text, int start, WordRunFormatting formatting, out bool isEastAsian)
    {
        isEastAsian = IsEastAsian(CodePointAt(text, start, out int length), formatting);
        int end = start + length;
        while (end < text.Length && IsEastAsian(CodePointAt(text, end, out length), formatting) == isEastAsian) end += length;
        return end;
    }

    /// <summary>The span that opens the East Asian set (<see cref="SwitchCss"/>).</summary>
    public static void AppendOpen(StringBuilder html, string switchCss) => html.Append("<span style=\"").Append(switchCss).Append("\">");

    /// <summary>HTML-encoded <paramref name="text"/> with its East Asian stretches in the run's East Asian set.</summary>
    public static string Encode(string text, WordRunFormatting formatting)
    {
        string? font = SwitchCss(formatting);
        if (font is null || text.Length == 0) return WordHtmlText.Encode(text);
        var html = new StringBuilder(text.Length + 48);
        for (int start = 0; start < text.Length;)
        {
            int end = StretchEnd(text, start, formatting, out bool isEastAsian);
            if (isEastAsian) AppendOpen(html, font);
            html.Append(WordHtmlText.Encode(text.Substring(start, end - start)));
            if (isEastAsian) html.Append("</span>");
            start = end;
        }
        return html.ToString();
    }
}
