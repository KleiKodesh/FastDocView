namespace FastDocView.Html;

/// <summary>
/// Splits text into Latin and complex-script pieces the way LibreOffice picks a portion's font set
/// (i18nutil/source/utility/scriptchangescanner.cxx:233-305, #i16354# / #i89825#): in a right-to-left paragraph every character
/// is complex script except a left-to-right run that holds a strong left-to-right character (English words, with the digits
/// that follow them, UBA W7); in a left-to-right paragraph every character is Latin except a right-to-left run (Hebrew /
/// Arabic, with the digits that follow them, UBA W2). Word 16 draws a .doc / .docx / .rtf run by its own right-to-left flag
/// (which it sets by the same script split when text is typed); formats without that flag (ODT) use this split.
/// Neutral characters between two runs take the paragraph's side (UBA N1 / N2). The split works on one text piece: a
/// left-to-right run broken by a style change is two runs.
/// </summary>
public static class WordTextScript
{
    /// <summary>Hebrew, Arabic, Syriac, Thaana, NKo, Samaritan..., their presentation forms (i18n ScriptType COMPLEX for the RTL scripts).</summary>
    public static bool IsComplexScript(char c)
        => c is (>= '\u0590' and <= '\u08FF') or (>= '\uFB1D' and <= '\uFDFF') or (>= '\uFE70' and <= '\uFEFF');

    /// <summary>u_charDirection == U_LEFT_TO_RIGHT: letters of the LTR scripts and the private-use area (symbol-font codes).</summary>
    public static bool IsStrongLeftToRight(char c)
        => c is >= '\uE000' and <= '\uF8FF' || (char.IsLetter(c) && !IsComplexScript(c));

    /// <summary>European digits (UBA class EN): they join a strong run they follow.</summary>
    private static bool IsEuropeanDigit(char c) => c is >= '0' and <= '9' || c is >= '\uFF10' and <= '\uFF19';

    /// <summary>
    /// The end of the piece that starts at <paramref name="start"/> and whether it draws in the complex-script set.
    /// </summary>
    public static int PieceEnd(string text, int start, bool isParagraphRightToLeft, out bool isComplex)
    {
        // the "other" run: left-to-right in a right-to-left paragraph, right-to-left in a left-to-right one
        bool IsOther(char c) => isParagraphRightToLeft ? IsStrongLeftToRight(c) : IsComplexScript(c);
        bool IsParagraphSide(char c) => isParagraphRightToLeft ? IsComplexScript(c) : IsStrongLeftToRight(c);
        int i = start;
        if (IsOther(text[i]))
        {
            int lastOther = i;
            for (i++; i < text.Length && !IsParagraphSide(text[i]); i++)
                if (IsOther(text[i]) || IsEuropeanDigit(text[i])) lastOther = i;     // no paragraph-side letter since: a digit joins the run
            isComplex = !isParagraphRightToLeft;
            return lastOther + 1;
        }
        while (i < text.Length && !IsOther(text[i])) i++;
        isComplex = isParagraphRightToLeft;
        return i;
    }

    /// <summary>Whether all of <paramref name="text"/> is one piece (no switch needed).</summary>
    public static bool IsSinglePiece(string text, bool isParagraphRightToLeft)
        => text.Length == 0 || PieceEnd(text, 0, isParagraphRightToLeft, out _) == text.Length;
}
